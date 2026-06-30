const express = require('express');
const crypto = require('crypto');
const fs = require('fs');
const http = require('http');
const https = require('https');
const path = require('path');
const { Server } = require('socket.io');

let XLSX = null;
try { XLSX = require('xlsx'); } catch (err) { console.warn(`Excel 词库解析不可用: ${err.message}`); }

const app = express();

const HTTPS_ENABLED = /^(1|true|yes)$/i.test(process.env.HTTPS || process.env.USE_HTTPS || '');
const SSL_CERT_FILE = process.env.SSL_CERT_FILE || path.join(__dirname, '.cert', 'localhost.crt');
const SSL_KEY_FILE = process.env.SSL_KEY_FILE || path.join(__dirname, '.cert', 'localhost.key');
let protocol = 'http', server;

if (HTTPS_ENABLED) {
  try {
    if (!fs.existsSync(SSL_CERT_FILE) || !fs.existsSync(SSL_KEY_FILE)) throw new Error('HTTPS cert files not found');
    server = https.createServer({ cert: fs.readFileSync(SSL_CERT_FILE), key: fs.readFileSync(SSL_KEY_FILE) }, app);
    protocol = 'https';
  } catch (err) { console.warn(`HTTPS unavailable: ${err.message}`); }
}
if (!server) server = http.createServer(app);

const io = new Server(server, { cors: { origin: "*" }, pingTimeout: 60000, pingInterval: 25000 });
app.use(express.static('public'));
app.use(express.json({ limit: '1mb' }));

const CONFIG_FILE = path.join(__dirname, 'config.json');
const MAX_PLAYERS = 20, MAX_SPECTATORS = 30;
const LOG_DIR = path.join(__dirname, 'logs');
const WORD_BANK_DIR = path.join(__dirname, 'wordbanks');
const WORD_BANK_EXTENSIONS = new Set(['.txt', '.csv', '.xls', '.xlsx']);
const wordBankCache = {};

function loadConfig() { try { return JSON.parse(fs.readFileSync(CONFIG_FILE, 'utf8')); } catch (err) { return {}; } }
function isLocalRequest(req) { const r = req.socket.remoteAddress || ''; return r === '127.0.0.1' || r === '::1' || r === '::ffff:127.0.0.1'; }
function getAdminTokenFromRequest(req) { const a = req.headers.authorization || ''; if (a.startsWith('Bearer ')) return a.slice(7).trim(); return req.headers['x-admin-token'] || req.query.token || ''; }
function getLogDateKey(date = new Date()) { return `${date.getFullYear()}-${String(date.getMonth()+1).padStart(2,'0')}-${String(date.getDate()).padStart(2,'0')}`; }
function getPlayerLogPath(date = new Date()) { return path.join(LOG_DIR, `player-activity-${getLogDateKey(date)}.log`); }
function formatDuration(ms) { if (!Number.isFinite(ms) || ms < 0) return null; const s = Math.floor(ms / 1000); return `${String(Math.floor(s/3600)).padStart(2,'0')}:${String(Math.floor((s%3600)/60)).padStart(2,'0')}:${String(s%60).padStart(2,'0')}`; }
function logPlayerActivity(action, details = {}) {
  const now = new Date(); const line = JSON.stringify({ time: now.toISOString(), action, username: details.username || '', roomId: details.roomId || '', roomName: details.roomName || '', role: details.role || '', socketId: details.socketId || '', reason: details.reason || '', duration: details.duration || null }) + '\n';
  fs.mkdir(LOG_DIR, { recursive: true }, (e) => { if (e) return; fs.appendFile(getPlayerLogPath(now), line, () => {}); });
}

const DEFAULT_WORDS = ['苹果','香蕉','猫','狗','太阳','月亮','星星','电脑','手机','书本','汽车','飞机','房子','树','花','鱼','鸟','蛋糕','冰淇淋','篮球','足球','雨伞','眼镜','手表','书包','铅笔','橡皮','桌子','椅子','电视','冰箱','空调','洗衣机','吉他','钢琴','小提琴','跑步','游泳','跳舞','唱歌','画画','吃饭','睡觉','喝水'];

function ensureWordBankDir() { fs.mkdirSync(WORD_BANK_DIR, { recursive: true }); const f = path.join(WORD_BANK_DIR, 'default.txt'); if (!fs.existsSync(f)) fs.writeFileSync(f, DEFAULT_WORDS.join('\n'), 'utf8'); }

function normalizeQuestionEntry(entry) {
  if (!entry || typeof entry !== 'object') return null;
  const type = entry.type === 'choice' ? 'choice' : 'answer';
  const question = String(entry.question || '').trim(); if (!question) return null;
  const answer = String(entry.answer || '').trim(); if (!answer) return null;
  const image = String(entry.image || '').trim() || null;
  const opts = Array.isArray(entry.options) ? entry.options.map(o => String(o || '').trim()).filter(Boolean) : [];
  return { type, question, image, answer, options: type === 'choice' ? opts : [] };
}

function legacyWordToQuestion(raw) {
  const w = String(raw || '').trim(); if (!w) return null;
  const parts = w.split('/');
  return { type: 'answer', question: `猜词（${parts[0].length}个字）`, image: null, answer: w, options: [] };
}

function parseTextWords(content) {
  const raw = String(content || '').split(/[\n\r,，]+/).map(w => w.trim()).filter(Boolean);
  const seen = new Set();
  return raw.reduce((arr, w) => { const q = legacyWordToQuestion(w); if (q && !seen.has(q.answer)) { seen.add(q.answer); arr.push(q); } return arr; }, []);
}

function parseExcelWords(buffer) {
  if (!XLSX) throw new Error('excel_parser_unavailable');
  const wb = XLSX.read(buffer, { type: 'buffer' });
  const sn = wb.SheetNames[0]; if (!sn) return [];
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: '' });
  if (rows.length < 2) return [];
  const questions = [], seen = new Set();
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i]; if (!Array.isArray(row) || row.length < 4) continue;
    const type = String(row[0] || '').trim(); if (type !== '选择' && type !== '简答') continue;
    const qText = String(row[1] || '').trim(); if (!qText) continue;
    const cellImg = String(row[2] || '').trim() || null;
    const answer = String(row[3] || '').trim(); if (!answer) continue;
    const opts = []; for (let j = 4; j < row.length; j++) { const o = String(row[j] || '').trim(); if (o) opts.push(o); }
    const qType = type === '选择' ? 'choice' : 'answer';
    if (qType === 'choice' && opts.length < 2) continue;
    const entry = normalizeQuestionEntry({ type: qType, question: qText, image: cellImg, answer, options: opts });
    if (!entry) continue;
    entry._row = i;
    const key = `${entry.type}|${entry.question}`; if (seen.has(key)) continue;
    seen.add(key); questions.push(entry);
  }
  return questions;
}

async function extractExcelImages(buffer) {
  const imageMap = {};
  try {
    const JSZip = require('jszip');
    const uploadDir = path.join(__dirname, 'public', 'uploads');
    fs.mkdirSync(uploadDir, { recursive: true });
    const zip = await JSZip.loadAsync(buffer);

    let drawingPath = null;
    const srFile = zip.file('xl/worksheets/_rels/sheet1.xml.rels');
    if (srFile) { const xml = await srFile.async('text'); const m = xml.match(/Target="\.\.\/drawings\/(drawing\d+\.xml)"/); if (m) drawingPath = 'xl/drawings/' + m[1]; }
    if (!drawingPath) { for (const n of Object.keys(zip.files)) { if (/^xl\/drawings\/drawing\d+\.xml$/i.test(n)) { drawingPath = n; break; } } }
    if (!drawingPath) return imageMap;

    const relsPath = drawingPath.replace('xl/drawings/', 'xl/drawings/_rels/').replace('.xml', '.xml.rels');
    const relsFile = zip.file(relsPath);
    const rIdMap = {};
    if (relsFile) {
      const xml = await relsFile.async('text'); const re = /Id="(rId\d+)"[^>]*Target="([^"]+)"/g; let m;
      while ((m = re.exec(xml))) { const t = m[2].replace(/^\.\.\//, ''); rIdMap[m[1]] = t.startsWith('media/') ? 'xl/' + t : t; }
    }

    const dFile = zip.file(drawingPath); if (!dFile) return imageMap;
    const dXml = await dFile.async('text');
    const aRe = /<(?:xdr:)?(?:twoCellAnchor|oneCellAnchor)[^>]*>([\s\S]*?)<\/(?:xdr:)?(?:twoCellAnchor|oneCellAnchor)>/g; let am;
    while ((am = aRe.exec(dXml))) {
      const chunk = am[1];
      const rowM = chunk.match(/<(?:xdr:)?row>(\d+)<\/(?:xdr:)?row>/);
      const embedM = chunk.match(/r:embed="(rId\d+)"/);
      if (!rowM || !embedM) continue;
      const rowIdx = parseInt(rowM[1], 10);
      const mediaPath = rIdMap[embedM[1]]; if (!mediaPath) continue;
      const mFile = zip.file(mediaPath); if (!mFile) continue;
      const imgData = await mFile.async('nodebuffer'); if (!imgData || imgData.length === 0) continue;
      const em = mediaPath.match(/\.(\w+)$/i); const ext = em ? em[1].toLowerCase() : 'png';
      const fname = `img_${Date.now()}_${Math.random().toString(36).substr(2,6)}.${ext}`;
      fs.writeFileSync(path.join(uploadDir, fname), imgData);
      imageMap[rowIdx] = '/uploads/' + fname;
    }
  } catch (e) { console.warn('图片提取失败:', e.message); }
  return imageMap;
}

function safeWordBankId(fileName) { if (typeof fileName !== 'string') return ''; const b = path.basename(fileName); return WORD_BANK_EXTENSIONS.has(path.extname(b).toLowerCase()) ? b : ''; }

async function loadWordBank(fileName) {
  ensureWordBankDir();
  const id = safeWordBankId(fileName); if (!id) throw new Error('invalid_word_bank');
  if (wordBankCache[id]) return wordBankCache[id];
  const fp = path.join(WORD_BANK_DIR, id); if (!fs.existsSync(fp)) throw new Error('word_bank_not_found');
  const ext = path.extname(id).toLowerCase();
  const buffer = fs.readFileSync(fp);
  let questions;
  if (ext === '.xls' || ext === '.xlsx') {
    questions = parseExcelWords(buffer);
    if (questions.length > 0) { const im = await extractExcelImages(buffer); for (const q of questions) { if (!q.image && im[q._row] != null) q.image = im[q._row]; } }
  } else { questions = parseTextWords(buffer.toString('utf8')); }
  if (questions.length === 0) throw new Error('word_bank_empty');
  const result = { id, name: path.basename(id, ext), questions, count: questions.length };
  wordBankCache[id] = result;
  return result;
}

async function listWordBanks() {
  ensureWordBankDir();
  const files = fs.readdirSync(WORD_BANK_DIR).filter(f => WORD_BANK_EXTENSIONS.has(path.extname(f).toLowerCase())).sort((a,b) => a.localeCompare(b, 'zh-CN'));
  const results = [];
  for (const f of files) {
    try { const b = await loadWordBank(f); results.push({ id: b.id, name: b.name, count: b.count }); }
    catch (err) { results.push({ id: f, name: path.basename(f, path.extname(f)), count: 0, error: err.message }); }
  }
  return results;
}

function getRoomQuestions(room) {
  if (room.customQuestions.length > 0) return room.customQuestions;
  const id = room.wordBankId || 'default.txt';
  if (wordBankCache[id]) return wordBankCache[id].questions;
  return DEFAULT_WORDS.map(w => legacyWordToQuestion(w)).filter(Boolean);
}

function getRandomQuestion(room) {
  const list = getRoomQuestions(room); if (list.length === 0) return null;
  return list[Math.floor(Math.random() * list.length)];
}
function normalizeAnswer(text) { return text.replace(/\s+/g, '').toLowerCase(); }
function isAnswerMatched(submitted, question) { const n = normalizeAnswer(submitted); return question.answer.split('/').some(a => normalizeAnswer(a) === n); }
function hasReferee(room) { return room.players.some(p => p.isReferee && !room.disconnectedPlayers[p.id]); }

ensureWordBankDir();

const rooms = {};
function generateRoomId() { let id; do { id = Math.random().toString(36).substring(2,8).toUpperCase(); } while (rooms[id]); return id; }
function createReconnectToken() { return crypto.randomBytes(16).toString('hex'); }

function createRoom(roomName) {
  const rid = generateRoomId();
  rooms[rid] = { id: rid, name: roomName || '房间'+rid, hostId: null, players: [], playerTokens: {}, disconnectedPlayers: {}, gameState: 'waiting', currentQuestion: null, currentBuzzerId: null, pendingJudgment: false, judgmentAnswer: '', currentRound: 1, timer: 30, timerInterval: null, scores: {}, helpUsed: {}, gameSettings: { answerTime: 30, buzzAnswerTime: 10, totalRounds: 5 }, customQuestions: [], wordBankId: 'default.txt', wordSourceName: 'default', wordSourceType: 'builtin' };
  return rid;
}

function getRoomBySocket(socket) { for (const rid in rooms) { if (rooms[rid].players.some(p => p.id === socket.id)) return rooms[rid]; } return null; }

setInterval(() => {
  const now = Date.now();
  for (const rid in rooms) { const room = rooms[rid]; let changed = false;
    for (const id in room.disconnectedPlayers) { if (now - room.disconnectedPlayers[id].disconnectTime > 5*60*1000) { const d = room.disconnectedPlayers[id]; logPlayerActivity('logout', { username: d.player?.name, roomId: rid, roomName: room.name, role: d.player?.isSpectator?'spectator':(d.player?.isReferee?'referee':'player'), socketId: id, reason: 'disconnect_timeout', duration: formatDuration(now-(d.player?.joinedAt||d.disconnectTime)) }); delete room.disconnectedPlayers[id]; delete room.scores[id]; delete room.playerTokens[id]; room.players = room.players.filter(p => p.id !== id); changed = true; } }
    if (!room.players.some(p => !room.disconnectedPlayers[p.id])) { delete rooms[rid]; continue; }
    if (changed) { io.to(rid).emit('playerLeft', { hostId: getHostId(room), players: room.players.filter(p => !room.disconnectedPlayers[p.id]) }); io.emit('roomListUpdate', getRoomListForBroadcast()); }
  }
}, 60000);

function getActivePlayers(room) { return room.players.filter(p => !p.isSpectator && !p.isReferee && !room.disconnectedPlayers[p.id]); }
function getHostId(room) { const a = getActivePlayers(room); const o = room.players.filter(p => !room.disconnectedPlayers[p.id]); const ch = room.hostId ? room.players.find(p => p.id === room.hostId) : null; if (ch && !room.disconnectedPlayers[ch.id]) return ch.id; room.hostId = a[0]?.id || o[0]?.id || null; return room.hostId; }
function getOnlinePlayers(room) { return room.players.filter(p => !room.disconnectedPlayers[p.id]); }
function getPlayerPublicState(room) { return { players: getOnlinePlayers(room), scores: room.scores, hostId: getHostId(room) }; }
function getWordSourcePayload(room) { return { hasCustomQuestions: room.customQuestions.length>0, questionCount: getRoomQuestions(room).length, wordSourceName: room.wordSourceName||(room.customQuestions.length?'自定义词库':'default'), wordSourceType: room.wordSourceType||(room.customQuestions.length?'custom':'builtin'), wordBankId: room.wordBankId||'default.txt' }; }
function canModerateRoom(socket, room) { const p = room.players.find(pp => pp.id === socket.id && !room.disconnectedPlayers[pp.id]); return !!p && p.isReferee; }
function getModerationText(v, fb) { const t = typeof v === 'string' ? v.trim() : ''; return t || fb; }
function emitPlayerState(room) { io.to(room.id).emit('playerStateUpdate', getPlayerPublicState(room)); io.emit('roomListUpdate', getRoomListForBroadcast()); }

function setRoomQuestionList(room, questions, sourceName, sourceType, sourceId = '') {
  const nq = questions.map(q => normalizeQuestionEntry(q)).filter(Boolean); if (nq.length === 0) return { ok: false, error: 'word_bank_empty' };
  room.customQuestions = nq; room.wordSourceName = sourceName || '自定义词库'; room.wordSourceType = sourceType || 'custom'; room.wordBankId = sourceType === 'builtin' && sourceId ? sourceId : room.wordBankId;
  const p = { count: nq.length, isUsingCustom: true, sourceName: room.wordSourceName, sourceType: room.wordSourceType, sourceId: room.wordBankId };
  io.to(room.id).emit('wordListUpdated', p); return { ok: true, ...p };
}

async function setRoomWordBank(room, fileName) {
  let bank; try { bank = await loadWordBank(fileName); } catch (err) { return { ok: false, error: err.message }; }
  room.customQuestions = []; room.wordBankId = bank.id; room.wordSourceName = bank.name; room.wordSourceType = 'builtin';
  const p = { count: bank.count, isUsingCustom: false, sourceName: bank.name, sourceType: 'builtin', sourceId: bank.id };
  io.to(room.id).emit('wordListUpdated', p); return { ok: true, ...p };
}

function broadcastRoomMessage(room, message, actorName = '观察面板') { const t = String(message||'').trim(); if (!t) return { ok: false, error: 'empty_message' }; io.to(room.id).emit('chat', { name: actorName, message: t, isSystem: true, isBroadcast: true }); io.to(room.id).emit('roomBroadcast', { message: t, actorName, roomId: room.id }); return { ok: true, message: t }; }
function closeRoom(room, reason, actorName = '观察面板') { const m = getModerationText(reason, '房间已被关闭。'); clearInterval(room.timerInterval); io.to(room.id).emit('roomClosed', { roomId: room.id, message: m, actorName }); for (const p of getOnlinePlayers(room)) { logPlayerActivity('logout', { username: p.name, roomId: room.id, roomName: room.name, role: p.isSpectator?'spectator':(p.isReferee?'referee':'player'), socketId: p.id, reason: `room_closed_by_${actorName}`, duration: formatDuration(Date.now()-(p.joinedAt||Date.now())) }); const ps = io.sockets.sockets.get(p.id); if (ps) { ps.leave(room.id); ps.leave(room.id+':spectators'); } } delete rooms[room.id]; io.emit('roomListUpdate', getRoomListForBroadcast()); return { ok: true, message: m }; }
function warnPlayerInRoom(room, targetId, reason, actorName = '裁判') { const t = room.players.find(p => p.id===targetId && !room.disconnectedPlayers[p.id]); if (!t) return { ok: false, error: 'target_not_found' }; io.to(t.id).emit('moderationWarning', { message: getModerationText(reason,'请遵守房间规则。'), actorName, roomId: room.id }); return { ok: true, targetName: t.name, message: getModerationText(reason,'') }; }

function kickPlayerFromRoom(room, targetId, reason, actorName = '裁判') {
  const idx = room.players.findIndex(p => p.id===targetId && !room.disconnectedPlayers[p.id]); if (idx===-1) return { ok: false, error: 'target_not_found' };
  const player = room.players[idx], message = getModerationText(reason,'你已被移出房间。'), wasPlaying = room.gameState==='playing', wasBuzzer = room.currentBuzzerId===targetId;
  const ts = io.sockets.sockets.get(targetId); if (ts) { ts.emit('moderationKicked',{message,actorName,roomId:room.id}); ts.leave(room.id); ts.leave(room.id+':spectators'); }
  logPlayerActivity('kick',{username:player.name,roomId:room.id,roomName:room.name,role:player.isSpectator?'spectator':(player.isReferee?'referee':'player'),socketId:targetId,reason:`moderated_by_${actorName}`,duration:formatDuration(Date.now()-(player.joinedAt||Date.now()))});
  room.players.splice(idx,1); delete room.scores[targetId]; delete room.playerTokens[targetId]; delete room.disconnectedPlayers[targetId];
  if (room.hostId===targetId) { room.hostId=null; getHostId(room); } if (room.currentBuzzerId===targetId) room.currentBuzzerId=null; if (room.pendingJudgment) room.pendingJudgment=false;
  io.to(room.id).emit('chat',{name:'系统',message:`${actorName} 已将 ${player.name} 踢出房间。`,isSystem:true});
  io.to(room.id).emit('playerLeft',{playerId:targetId,players:getOnlinePlayers(room),hostId:getHostId(room)}); io.emit('roomListUpdate',getRoomListForBroadcast());
  if (getActivePlayers(room).length<1 && wasPlaying) { io.to(room.id).emit('chat',{name:'系统',message:'玩家不足，游戏已结束。',isSystem:true}); endGame(room); }
  else if (wasPlaying && wasBuzzer) { clearInterval(room.timerInterval); io.to(room.id).emit('chat',{name:'系统',message:'抢答玩家被移出，本题跳过。',isSystem:true}); room.currentRound++; room.currentBuzzerId=null; room.pendingJudgment=false; setTimeout(()=>startQuestionRound(room),2000); }
  return { ok: true, targetName: player.name, message };
}

// ========== Quiz Game Flow ==========

function startGame(room, socket) {
  if (room.gameState==='finished') { clearInterval(room.timerInterval); room.currentRound=1; room.currentBuzzerId=null; room.currentQuestion=null; room.pendingJudgment=false; room.judgmentAnswer=''; room.scores={}; room.helpUsed={}; room.gameState='waiting'; for (const id in room.disconnectedPlayers) { delete room.scores[id]; delete room.playerTokens[id]; room.players=room.players.filter(p=>p.id!==id); } room.disconnectedPlayers={}; getActivePlayers(room).forEach(p=>room.scores[p.id]=0); for (const p of room.players) { if (p.isReferee) delete room.scores[p.id]; } }
  const active = getActivePlayers(room); if (active.length<1) { if (socket) socket.emit('errorMessage','人数不足！'); return; }
  room.gameState='playing'; room.currentRound=1; room.currentBuzzerId=null; room.currentQuestion=null; room.pendingJudgment=false; room.judgmentAnswer=''; room.scores={}; room.helpUsed={}; active.forEach(p=>room.scores[p.id]=0); for (const p of room.players) { if (p.isReferee) delete room.scores[p.id]; }
  io.to(room.id).emit('gameStarted',{settings:room.gameSettings,players:room.players,scores:room.scores,hostId:getHostId(room)}); io.emit('roomListUpdate',getRoomListForBroadcast());
  setTimeout(()=>startQuestionRound(room),2000);
}

function startQuestionRound(room) {
  if (room.gameState==='finished') return; if (room.currentRound > room.gameSettings.totalRounds) { endGame(room); return; }
  const q = getRandomQuestion(room); if (!q) { endGame(room); return; }
  room.currentQuestion = q; room.currentBuzzerId = null; room.pendingJudgment = false; room.judgmentAnswer = ''; room.timer = room.gameSettings.answerTime;
  for (const p of room.players) { if (room.disconnectedPlayers[p.id]) continue; io.to(p.id).emit('questionRound', { round: room.currentRound, totalRounds: room.gameSettings.totalRounds, currentTimer: room.timer, question: sanitizeQuestionForPlayer(q, p, room) }); }
  room.timerInterval = setInterval(() => { if (room.gameState==='finished') return; room.timer--; io.to(room.id).emit('timerUpdate', room.timer); if (room.timer<=0) endQuestionRoundTimeout(room); }, 1000);
}

function sanitizeQuestionForPlayer(q, player, room) {
  const base = { type: q.type, question: q.question, image: q.image, wordLength: q.type==='answer' ? q.answer.split('/')[0].length : 0 };
  if (q.type==='choice') base.options = q.options;
  if (player.isReferee) { base.answer = q.answer; if (q.type==='choice') base.options = q.options; }
  base.isReferee = player.isReferee; base.isSpectator = player.isSpectator; base.isPlayer = !player.isSpectator && !player.isReferee;
  base.canBuzz = !player.isSpectator && !player.isReferee && !room.pendingJudgment; base.buzzerId = room.currentBuzzerId || null;
  return base;
}

function processBuzzIn(room, socketId) {
  if (room.gameState!=='playing') return { ok: false, error: '游戏未进行' }; if (room.currentBuzzerId) return { ok: false, error: '已有人抢答' }; if (room.pendingJudgment) return { ok: false, error: '等待裁判判定中' };
  const player = room.players.find(p => p.id===socketId && !room.disconnectedPlayers[p.id]); if (!player || player.isSpectator || player.isReferee) return { ok: false, error: '无权抢答' };
  room.currentBuzzerId = socketId; clearInterval(room.timerInterval); room.timer = room.gameSettings.buzzAnswerTime;
  io.to(room.id).emit('buzzResult', { buzzerId: socketId, buzzerName: player.name }); io.to(socketId).emit('buzzAccepted', { question: room.currentQuestion, answerTime: room.timer });
  room.timerInterval = setInterval(() => { if (room.gameState==='finished') return; room.timer--; io.to(room.id).emit('timerUpdate', room.timer); if (room.timer<=0) { if (room.pendingJudgment) handleJudgmentTimeout(room); else handleAnswer(room, socketId, room.currentQuestion.type==='choice' ? -1 : ''); } }, 1000);
  return { ok: true };
}

function handleAnswer(room, socketId, answer) {
  if (room.gameState==='finished') return; if (room.currentBuzzerId !== socketId) return;
  const q = room.currentQuestion; if (!q) return; clearInterval(room.timerInterval); const player = room.players.find(p => p.id===socketId);
  if (q.type==='choice') {
    const optIndex = typeof answer==='number' ? answer : parseInt(answer,10); let correct = false;
    const letterMatch = /^[A-Za-z]$/.test(q.answer);
    if (letterMatch) { correct = (optIndex === q.answer.toUpperCase().charCodeAt(0)-65); }
    else { const co = q.options[optIndex]||''; correct = normalizeAnswer(co)===normalizeAnswer(q.answer); }
    applyAnswerResult(room, socketId, player, correct, q.options[optIndex]||'', q.answer);
  } else {
    const text = String(answer||'').trim();
    if (hasReferee(room)) { room.pendingJudgment=true; room.judgmentAnswer=text; clearInterval(room.timerInterval); room.timer=15; const refs=room.players.filter(p=>p.isReferee&&!room.disconnectedPlayers[p.id]); io.to(room.id).emit('chat',{name:'系统',message:`${player?player.name:'玩家'} 已提交答案，等待裁判判定...`,isSystem:true}); for (const ref of refs) io.to(ref.id).emit('judgeRequest',{buzzerName:player?player.name:'玩家',submittedAnswer:text,referenceAnswer:q.answer,timer:room.timer}); room.timerInterval=setInterval(()=>{if(room.gameState==='finished')return;room.timer--;io.to(room.id).emit('timerUpdate',room.timer);if(room.timer<=0)handleJudgmentTimeout(room);},1000); return; }
    applyAnswerResult(room, socketId, player, isAnswerMatched(text,q), text, q.answer);
  }
}

function handleJudgmentTimeout(room) { if (!room.pendingJudgment) return; const bid=room.currentBuzzerId, player=room.players.find(p=>p.id===bid); room.pendingJudgment=false; applyAnswerResult(room,bid,player,false,room.judgmentAnswer,room.currentQuestion?room.currentQuestion.answer:''); }

function applyAnswerResult(room, socketId, player, correct, submittedAnswer, correctAnswer) {
  if (correct) { room.scores[socketId]=(room.scores[socketId]||0)+10; io.to(room.id).emit('answerResult',{buzzerId:socketId,buzzerName:player?player.name:'',correct:true,answer:correctAnswer,submittedAnswer,scores:room.scores}); io.to(room.id).emit('chat',{name:'系统',message:player?`${player.name} 回答正确！答案【${correctAnswer}】+10分`:`回答正确！+10分`,isSystem:true}); }
  else { room.scores[socketId]=(room.scores[socketId]||0)-5; io.to(room.id).emit('answerResult',{buzzerId:socketId,buzzerName:player?player.name:'',correct:false,answer:correctAnswer,submittedAnswer,scores:room.scores}); io.to(room.id).emit('chat',{name:'系统',message:player?`${player.name} 回答错误（输入: ${submittedAnswer||'超时'}）！正确答案【${correctAnswer}】-5分`:`回答错误！-5分`,isSystem:true}); }
  room.currentBuzzerId=null; room.pendingJudgment=false; room.judgmentAnswer=''; clearInterval(room.timerInterval); room.currentRound++;
  setTimeout(()=>startQuestionRound(room),3000);
}

function endQuestionRoundTimeout(room) { if (room.gameState==='finished') return; clearInterval(room.timerInterval); io.to(room.id).emit('questionTimeout',{answer:room.currentQuestion?room.currentQuestion.answer:''}); io.to(room.id).emit('chat',{name:'系统',message:`时间到！无人抢答，正确答案是【${room.currentQuestion?room.currentQuestion.answer:'?'}】`,isSystem:true}); room.currentRound++; room.currentBuzzerId=null; room.pendingJudgment=false; room.judgmentAnswer=''; setTimeout(()=>startQuestionRound(room),3000); }
function endGame(room) { clearInterval(room.timerInterval); room.gameState='finished'; const op=room.players.filter(p=>!room.disconnectedPlayers[p.id]); const pr=op.filter(p=>!p.isSpectator&&!p.isReferee).map(p=>({id:p.id,name:p.name,score:room.scores[p.id]||0})).sort((a,b)=>b.score-a.score); io.to(room.id).emit('gameEnd',{rankings:pr}); for (const id in room.disconnectedPlayers) { delete room.scores[id]; delete room.playerTokens[id]; room.players=room.players.filter(p=>p.id!==id); } room.disconnectedPlayers={}; io.emit('roomListUpdate',getRoomListForBroadcast()); }
function restartGame(room) { room.gameState='waiting'; clearInterval(room.timerInterval); room.currentRound=1; room.currentBuzzerId=null; room.currentQuestion=null; room.pendingJudgment=false; room.judgmentAnswer=''; room.scores={}; room.helpUsed={}; getActivePlayers(room).forEach(p=>room.scores[p.id]=0); for (const p of room.players) { if (p.isReferee) delete room.scores[p.id]; } io.to(room.id).emit('gameRestarted',{players:room.players,settings:room.gameSettings,hostId:getHostId(room),...getWordSourcePayload(room)}); io.emit('roomListUpdate',getRoomListForBroadcast()); }

function getRoomListForBroadcast() { return Object.values(rooms).map(room=>({ id:room.id,name:room.name,playerCount:room.players.filter(p=>!p.isSpectator&&!p.isReferee&&!room.disconnectedPlayers[p.id]).length,spectatorCount:room.players.filter(p=>p.isSpectator&&!room.disconnectedPlayers[p.id]).length,refereeCount:room.players.filter(p=>p.isReferee&&!room.disconnectedPlayers[p.id]).length,gameState:room.gameState,wordSourceName:room.wordSourceName||(room.customQuestions.length?'自定义词库':'default'),wordCount:getRoomQuestions(room).length })); }

function getAdminStatus() { return { ok:true,updatedAt:new Date().toISOString(),socketCount:io.engine.clientsCount,roomCount:Object.keys(rooms).length,onlineUserCount:Object.values(rooms).reduce((s,r)=>s+r.players.filter(p=>!r.disconnectedPlayers[p.id]).length,0),wordBanks:(()=>{try{return fs.readdirSync(WORD_BANK_DIR).filter(f=>WORD_BANK_EXTENSIONS.has(path.extname(f).toLowerCase())).map(f=>({id:f,name:path.basename(f,path.extname(f)),count:0}));}catch(e){return[];}})(),rooms:Object.values(rooms).map(r=>{const op=r.players.filter(p=>!r.disconnectedPlayers[p.id]);return{id:r.id,name:r.name,gameState:r.gameState,currentRound:r.currentRound,totalRounds:r.gameSettings.totalRounds,currentQuestionType:r.currentQuestion?r.currentQuestion.type:'',currentQuestion:r.currentQuestion?r.currentQuestion.question:'',currentBuzzerId:(r.players.find(p=>p.id===r.currentBuzzerId)||{}).id||'',timer:r.timer,counts:{online:op.length,players:op.filter(p=>!p.isSpectator&&!p.isReferee).length,spectators:op.filter(p=>p.isSpectator).length,referees:op.filter(p=>p.isReferee).length,disconnected:Object.keys(r.disconnectedPlayers).length},users:op.map(p=>({id:p.id,name:p.name,role:p.isReferee?'referee':(p.isSpectator?'spectator':'player'),isHost:getHostId(r)===p.id,isBuzzer:r.currentBuzzerId===p.id,score:r.scores[p.id]??null,joinedAt:p.joinedAt||null}))}})}; }

app.get('/admin/status',(req,res)=>{if(!isLocalRequest(req)){res.status(403).json({ok:false,error:'localhost_only'});return;}const c=loadConfig();if(!c.adminToken||getAdminTokenFromRequest(req)!==c.adminToken){res.status(401).json({ok:false,error:'unauthorized'});return;}res.json(getAdminStatus());});
app.post('/admin/moderate',(req,res)=>{if(!isLocalRequest(req)){res.status(403).json({ok:false,error:'localhost_only'});return;}const c=loadConfig();if(!c.adminToken||getAdminTokenFromRequest(req)!==c.adminToken){res.status(401).json({ok:false,error:'unauthorized'});return;}const{action,roomId,targetId,reason}=req.body||{};const room=rooms[roomId];if(!room){res.status(404).json({ok:false,error:'room_not_found'});return;}let r;if(action==='warn')r=warnPlayerInRoom(room,targetId,reason,'观察面板');else if(action==='kick')r=kickPlayerFromRoom(room,targetId,reason,'观察面板');else if(action==='broadcast')r=broadcastRoomMessage(room,reason,'观察面板');else if(action==='closeRoom')r=closeRoom(room,reason,'观察面板');else if(action==='endTurn'){if(room.gameState!=='playing')r={ok:false,error:'round_not_playing'};else{io.to(room.id).emit('chat',{name:'系统',message:'观察面板提前结束了当前题目。',isSystem:true});endQuestionRoundTimeout(room);r={ok:true};}}else r={ok:false,error:'unknown_action'};if(!r.ok){res.status(400).json(r);return;}res.json(r);});

function removePlayerFromRoom(socket, room, isDisconnect = false) {
  const idx = room.players.findIndex(p => p.id===socket.id); if (idx===-1) return;
  const player = room.players[idx], wasPlaying = room.gameState==='playing', wasBuzzer = room.currentBuzzerId===socket.id, eventAt = Date.now();
  logPlayerActivity(isDisconnect?'disconnect':'logout',{username:player.name,roomId:room.id,roomName:room.name,role:player.isSpectator?'spectator':(player.isReferee?'referee':'player'),socketId:socket.id,reason:isDisconnect?'socket_disconnect':'leave_room',duration:formatDuration(eventAt-(player.joinedAt||eventAt))});
  if (isDisconnect) { room.disconnectedPlayers[socket.id]={player,oldScores:room.scores[socket.id],reconnectToken:room.playerTokens[socket.id],disconnectTime:Date.now()}; if (wasPlaying&&wasBuzzer) io.to(room.id).emit('chat',{name:'系统',message:`抢答玩家 ${player.name} 断线了`,isSystem:true}); }
  else { socket.leave(room.id); socket.leave(room.id+':spectators'); room.players.splice(idx,1); delete room.scores[socket.id]; delete room.playerTokens[socket.id]; delete room.disconnectedPlayers[socket.id]; if (room.hostId===socket.id) { room.hostId=null; getHostId(room); } if (room.currentBuzzerId===socket.id) room.currentBuzzerId=null; if (room.pendingJudgment) room.pendingJudgment=false; socket.emit('roomLeft',{roomId:room.id}); if (getOnlinePlayers(room).length===0) { delete rooms[room.id]; io.emit('roomListUpdate',getRoomListForBroadcast()); return; } }
  io.to(room.id).emit('playerLeft',{playerId:socket.id,players:room.players.filter(p=>!room.disconnectedPlayers[p.id]),hostId:getHostId(room)}); io.emit('roomListUpdate',getRoomListForBroadcast());
  if (!isDisconnect) { if (getActivePlayers(room).length<1 && wasPlaying) { io.to(room.id).emit('chat',{name:'系统',message:'玩家不足，游戏已结束。',isSystem:true}); endGame(room); } else if (wasPlaying && wasBuzzer) { clearInterval(room.timerInterval); io.to(room.id).emit('chat',{name:'系统',message:'抢答玩家离开，本题跳过。',isSystem:true}); room.currentRound++; room.currentBuzzerId=null; room.pendingJudgment=false; setTimeout(()=>startQuestionRound(room),2000); } }
}

// ===================== Socket Events =====================
io.on('connection', (socket) => {
  console.log('用户连接:', socket.id);

  socket.on('reconnectAttempt', (payload, legacyOldId) => {
    const isObj = payload && typeof payload === 'object';
    const roomId = isObj ? payload.roomId : payload, oldId = isObj ? payload.playerId : legacyOldId, reconnectToken = isObj ? payload.reconnectToken : null;
    const room = rooms[roomId]; if (!room) { socket.emit('reconnectFailed',{reason:'roomNotFound',message:'房间不存在或已结束'}); return; }
    const d = room.disconnectedPlayers[oldId]; if (!d) { socket.emit('reconnectFailed',{reason:'playerNotFound',message:'未找到可恢复的断线玩家'}); return; }
    if (!reconnectToken || d.reconnectToken !== reconnectToken) { socket.emit('reconnectFailed',{reason:'tokenMismatch',message:'重连凭证无效'}); return; }
    const { player, oldScores } = d; delete room.disconnectedPlayers[oldId];
    const pi = room.players.findIndex(p => p.id===oldId); if (pi===-1) { socket.emit('reconnectFailed',{reason:'playerNotFound',message:'玩家已被移出房间'}); return; }
    room.players[pi].id = socket.id; if (room.hostId===oldId) room.hostId=socket.id; if (room.currentBuzzerId===oldId) room.currentBuzzerId=socket.id;
    if (!player.isSpectator && !player.isReferee) room.scores[socket.id] = oldScores||0; delete room.scores[oldId];
    const newToken = createReconnectToken(); room.playerTokens[socket.id]=newToken; delete room.playerTokens[oldId];
    socket.join(roomId); if (player.isSpectator) socket.join(roomId+':spectators');
    io.to(roomId).emit('playerReconnected',{oldId,newPlayer:room.players[pi],players:room.players.filter(p=>!room.disconnectedPlayers[p.id]),scores:room.scores,hostId:getHostId(room),isCurrentBuzzer:room.currentBuzzerId===socket.id});
    const pr = room.players[pi];
    socket.emit('reconnectSuccess',{player:pr,gameState:room.gameState,players:room.players.filter(p=>!room.disconnectedPlayers[p.id]),settings:room.gameSettings,scores:room.scores,isHost:getHostId(room)===socket.id,currentRound:room.currentRound,totalRounds:room.gameSettings.totalRounds,currentQuestion:room.gameState==='playing'?sanitizeQuestionForPlayer(room.currentQuestion||{},pr,room):null,currentTimer:room.timer,isSpectator:pr.isSpectator,isReferee:pr.isReferee,isBuzzer:room.currentBuzzerId===socket.id,...getWordSourcePayload(room),roomId:room.id,reconnectToken:newToken,pendingJudgment:room.pendingJudgment&&pr.isReferee});
  });

  socket.on('getRoomList', () => socket.emit('roomList', getRoomListForBroadcast()));

  socket.on('createRoom', (roomName, playerName) => {
    const oldRoom = getRoomBySocket(socket); if (oldRoom) removePlayerFromRoom(socket, oldRoom, false);
    const roomId = createRoom(roomName), token = createReconnectToken();
    const player = { id: socket.id, name: playerName, isSpectator: false, isReferee: false, joinedAt: Date.now() };
    rooms[roomId].players.push(player); rooms[roomId].hostId = socket.id; rooms[roomId].scores[socket.id] = 0; rooms[roomId].playerTokens[socket.id] = token;
    socket.join(roomId);
    socket.emit('roomJoined', { roomId, roomName: rooms[roomId].name, isHost: true, player, players: rooms[roomId].players, hostId: socket.id, ...getWordSourcePayload(rooms[roomId]), reconnectToken: token, settings: rooms[roomId].gameSettings });
    socket.emit('reconnectInfo', { roomId, playerId: socket.id, reconnectToken: token });
    logPlayerActivity('login', { username: player.name, roomId, roomName: rooms[roomId].name, role: 'player', socketId: socket.id, reason: 'create_room' });
    io.emit('roomListUpdate', getRoomListForBroadcast());
  });

  socket.on('joinRoom', (roomId, playerName, asSpectator = false) => {
    const room = rooms[roomId]; if (!room) { socket.emit('errorMessage', { type: 'notFound', message: '房间不存在' }); return; }
    if (room.players.find(p => p.name === playerName)) { socket.emit('errorMessage', { type: 'nameConflict', message: `房间内已有玩家叫"${playerName}"，请修改昵称` }); return; }
    const oldRoom = getRoomBySocket(socket); if (oldRoom && oldRoom.id !== roomId) removePlayerFromRoom(socket, oldRoom, false);
    if (room.gameState !== 'waiting' || asSpectator) asSpectator = true;
    else if (getActivePlayers(room).length >= MAX_PLAYERS) { asSpectator = true; socket.emit('chat', { name: '系统', message: '玩家已满，自动转为观战', isSystem: true }); }
    if (asSpectator && room.players.filter(p => p.isSpectator && !room.disconnectedPlayers[p.id]).length >= MAX_SPECTATORS) { socket.emit('errorMessage', { type: 'full', message: '观众人数已满' }); return; }
    const player = { id: socket.id, name: playerName, isSpectator: asSpectator, isReferee: false, joinedAt: Date.now() }, token = createReconnectToken();
    room.players.push(player); if (!asSpectator) room.scores[socket.id] = 0; room.playerTokens[socket.id] = token;
    socket.join(roomId); if (asSpectator) socket.join(roomId + ':spectators');
    const hostId = getHostId(room);
    io.to(roomId).emit(asSpectator ? 'spectatorJoined' : 'playerJoined', { player, players: room.players.filter(p => !room.disconnectedPlayers[p.id]), hostId, ...getWordSourcePayload(room) });
    socket.emit('stateUpdate', { gameState: room.gameState, players: room.players.filter(p => !room.disconnectedPlayers[p.id]), settings: room.gameSettings, isHost: hostId === socket.id, ...getWordSourcePayload(room), roomId: room.id, reconnectToken: token, playerId: socket.id });
    socket.emit('reconnectInfo', { roomId: room.id, playerId: socket.id, reconnectToken: token });
    logPlayerActivity('login', { username: player.name, roomId: room.id, roomName: room.name, role: asSpectator ? 'spectator' : 'player', socketId: socket.id, reason: asSpectator ? 'join_as_spectator' : 'join_room' });
    if (room.gameState === 'playing') {
      socket.emit('gameInProgress', { gameState: room.gameState, currentRound: room.currentRound, totalRounds: room.gameSettings.totalRounds, currentTimer: room.timer, isSpectator: asSpectator, isPlayer: !asSpectator, isReferee: false, isBuzzer: room.currentBuzzerId === socket.id, playerId: socket.id, question: room.currentQuestion ? sanitizeQuestionForPlayer(room.currentQuestion, player, room) : null });
    }
    io.emit('roomListUpdate', getRoomListForBroadcast());
  });

  socket.on('leaveRoom', () => { const room = getRoomBySocket(socket); if (room) removePlayerFromRoom(socket, room, false); });

  socket.on('switchToSpectator', () => { const room = getRoomBySocket(socket); if (!room) return; const idx = room.players.findIndex(p => p.id===socket.id); if (idx===-1) return; const player = room.players[idx]; if (player.isSpectator) return; player.isSpectator=true; player.isReferee=false; delete room.scores[socket.id]; socket.join(room.id+':spectators'); if (room.currentBuzzerId===socket.id) room.currentBuzzerId=null; io.to(room.id).emit('playerSwitched',{player,players:room.players.filter(p=>!room.disconnectedPlayers[p.id]),isNowPlayer:false,isNowReferee:false,hostId:getHostId(room)}); });
  socket.on('switchToPlayer', () => { const room = getRoomBySocket(socket); if (!room) return; const idx = room.players.findIndex(p => p.id===socket.id); if (idx===-1) return; const player = room.players[idx]; if (!player.isSpectator && !player.isReferee) return; const wasReferee=player.isReferee; player.isSpectator=false; player.isReferee=false; room.scores[socket.id]=0; socket.leave(room.id+':spectators'); if (!room.hostId||(wasReferee&&room.hostId===socket.id)) room.hostId=socket.id; io.to(room.id).emit('playerSwitched',{player,players:room.players.filter(p=>!room.disconnectedPlayers[p.id]),isNowPlayer:true,isNowReferee:false,hostId:getHostId(room)}); });
  socket.on('switchToReferee', () => { const room = getRoomBySocket(socket); if (!room || getHostId(room)!==socket.id) return; const idx = room.players.findIndex(p => p.id===socket.id); if (idx===-1) return; const player = room.players[idx]; if (player.isReferee) return; const wasSpectator=player.isSpectator; player.isReferee=true; player.isSpectator=false; delete room.scores[socket.id]; if (wasSpectator) socket.leave(room.id+':spectators'); if (room.currentBuzzerId===socket.id) room.currentBuzzerId=null; io.to(room.id).emit('playerSwitched',{player,players:room.players.filter(p=>!room.disconnectedPlayers[p.id]),isNowPlayer:false,isNowReferee:true,hostId:getHostId(room)}); });

  socket.on('updateSettings', (ns) => { const room = getRoomBySocket(socket); if (!room || getHostId(room)!==socket.id) return; room.gameSettings = { ...room.gameSettings, answerTime: Math.max(10,Math.min(120,Number(ns.answerTime)||room.gameSettings.answerTime)), buzzAnswerTime: Math.max(5,Math.min(30,Number(ns.buzzAnswerTime)||room.gameSettings.buzzAnswerTime)), totalRounds: Math.max(1,Math.min(50,Number(ns.totalRounds)||room.gameSettings.totalRounds)) }; io.to(room.id).emit('settingsUpdated', room.gameSettings); });

  socket.on('getWordBanks', async (ack) => { const r = { ok: true, wordBanks: await listWordBanks() }; if (typeof ack==='function') ack(r); else socket.emit('wordBanksList', r); });

  socket.on('selectWordBank', async (wordBankId, ack) => { const room = getRoomBySocket(socket); if (!room || getHostId(room)!==socket.id) { const r = { ok: false, error: 'not_host' }; if (typeof ack==='function') ack(r); else socket.emit('wordListRejected', r); return; } const r = await setRoomWordBank(room, wordBankId); if (typeof ack==='function') ack(r); if (!r.ok) socket.emit('wordListRejected', r); io.emit('roomListUpdate', getRoomListForBroadcast()); });

  socket.on('updateWordList', async (data, ack) => {
    const room = getRoomBySocket(socket); if (!room || getHostId(room)!==socket.id) { const r = { ok: false, error: 'not_host' }; if (typeof ack==='function') ack(r); else socket.emit('wordListRejected', r); return; }
    let questions; if (Array.isArray(data) && data.length>0 && typeof data[0]==='object') questions = data; else { const words = Array.isArray(data) ? data : String(data||'').split(/[\n\r,，]+/); questions = words.map(w => legacyWordToQuestion(String(w||''))).filter(Boolean); }
    const r = setRoomQuestionList(room, questions, '上传词库', 'upload');
    if (typeof ack==='function') ack(r); if (!r.ok) socket.emit('wordListRejected', r); io.emit('roomListUpdate', getRoomListForBroadcast());
  });

  socket.on('startGame', () => { const room = getRoomBySocket(socket); if (room && getHostId(room)===socket.id) startGame(room, socket); });
  socket.on('restartGame', () => { const room = getRoomBySocket(socket); if (room && getHostId(room)===socket.id) restartGame(room); });
  socket.on('forceEndGame', () => { const room = getRoomBySocket(socket); if (!room || getHostId(room)!==socket.id) return; if (room.gameState==='playing') { io.to(room.id).emit('chat',{name:'系统',message:'房主强制结束了游戏',isSystem:true}); endGame(room); } });
  socket.on('buzzIn', () => { const room = getRoomBySocket(socket); if (!room || room.gameState!=='playing') return; const r = processBuzzIn(room, socket.id); if (!r.ok) socket.emit('buzzRejected', { error: r.error }); });
  socket.on('submitAnswer', (answer) => { const room = getRoomBySocket(socket); if (!room || room.gameState!=='playing') return; if (room.currentBuzzerId!==socket.id) return; if (room.currentQuestion && room.currentQuestion.type==='choice') return; handleAnswer(room, socket.id, String(answer||'').trim()); });
  socket.on('submitChoice', (optionIndex) => { const room = getRoomBySocket(socket); if (!room || room.gameState!=='playing') return; if (room.currentBuzzerId!==socket.id) return; if (!room.currentQuestion || room.currentQuestion.type!=='choice') return; handleAnswer(room, socket.id, parseInt(optionIndex,10)); });
  socket.on('judgeAnswer', (correct) => { const room = getRoomBySocket(socket); if (!room || !room.pendingJudgment) return; const player = room.players.find(p => p.id===socket.id && !room.disconnectedPlayers[p.id]); if (!player || !player.isReferee) return; room.pendingJudgment=false; clearInterval(room.timerInterval); const buzzer = room.players.find(p => p.id===room.currentBuzzerId), sa=room.judgmentAnswer, ca=room.currentQuestion?room.currentQuestion.answer:''; applyAnswerResult(room, room.currentBuzzerId, buzzer, !!correct, sa, ca); io.to(room.id).emit('chat',{name:'系统',message:`裁判 ${player.name} 判定${correct?'正确':'错误'}`,isSystem:true}); });

  socket.on('requestHelp', () => { const room = getRoomBySocket(socket); if (!room||room.gameState!=='playing') return; if (room.currentBuzzerId!==socket.id) return; if (room.helpUsed[socket.id]) { socket.emit('helpRejected',{error:'你已经使用过场外求助了'}); return; } const spectators=room.players.filter(p=>p.isSpectator&&!room.disconnectedPlayers[p.id]); if (spectators.length===0) { socket.emit('helpRejected',{error:'没有可用的观众'}); return; } const helper=spectators[Math.floor(Math.random()*spectators.length)]; room.helpUsed[socket.id]=true; const buzzer=room.players.find(p=>p.id===socket.id); io.to(socket.id).emit('helpSent',{helperName:helper.name}); io.to(helper.id).emit('helpRequested',{buzzerName:buzzer?buzzer.name:'玩家',buzzerId:socket.id}); io.to(room.id).emit('chat',{name:'系统',message:`${buzzer?buzzer.name:'玩家'} 使用了场外求助，等待观众 ${helper.name} 的提示...`,isSystem:true}); });
  socket.on('submitHelp', (answer) => { const room = getRoomBySocket(socket); if (!room||room.gameState!=='playing') return; const bid=room.currentBuzzerId; if (!bid) return; const helper=room.players.find(p=>p.id===socket.id&&!room.disconnectedPlayers[p.id]); if (!helper||!helper.isSpectator) return; const text=String(answer||'').trim(); if (!text) return; io.to(bid).emit('helpAnswer',{helperName:helper.name,answer:text}); io.to(socket.id).emit('helpDelivered',{buzzerName:(room.players.find(p=>p.id===bid)||{}).name||'玩家'}); io.to(room.id).emit('chat',{name:'系统',message:`观众 ${helper.name} 已向抢答者发送提示`,isSystem:true}); });
  socket.on('forceEndTurn', () => { const room = getRoomBySocket(socket); if (!room||room.gameState!=='playing') return; const player=room.players.find(p=>p.id===socket.id); if (!player||!canModerateRoom(socket,room)) return; clearInterval(room.timerInterval); io.to(room.id).emit('chat',{name:'系统',message:`${player.name} 提前结束了当前题目。`,isSystem:true}); endQuestionRoundTimeout(room); });
  socket.on('moderatePlayer', (payload={}) => { const room = getRoomBySocket(socket); if (!room||!canModerateRoom(socket,room)) return; const actor=room.players.find(p=>p.id===socket.id), targetId=payload.targetId; if (!targetId||targetId===socket.id) return; let r; if (payload.action==='warn') { r=warnPlayerInRoom(room,targetId,payload.reason,actor?.name||'裁判'); if (r.ok) socket.emit('moderationActionResult',{ok:true,message:`已警告 ${r.targetName}`}); } else if (payload.action==='kick') { r=kickPlayerFromRoom(room,targetId,payload.reason,actor?.name||'裁判'); if (r.ok) socket.emit('moderationActionResult',{ok:true,message:`已踢出 ${r.targetName}`}); } else return; if (!r.ok) socket.emit('moderationActionResult',{ok:false,message:'操作失败'}); });
  socket.on('chat', (message) => { const room = getRoomBySocket(socket); if (!room) return; const player=room.players.find(p=>p.id===socket.id); if (!player) return; if (player.isReferee) { io.to(room.id).emit('chat',{name:player.name,message,isReferee:true}); return; } if (player.isSpectator) io.to(room.id+':spectators').emit('chat',{name:player.name,message,spectatorChat:true}); else io.to(room.id).emit('chat',{name:player.name,message}); });
  socket.on('disconnect', () => { const room = getRoomBySocket(socket); if (room) removePlayerFromRoom(socket, room, true); });
});

const PORT = process.env.PORT || 3000, HOST = process.env.HOST || '127.0.0.1';
server.listen(PORT, HOST, async () => {
  console.log(`服务器运行在 ${protocol}://${HOST==='127.0.0.1'?'localhost':HOST}:${PORT}`);
  // preload default word bank
  try { await loadWordBank('default.txt'); console.log('默认词库已加载'); } catch (e) { console.warn('默认词库加载失败:', e.message); }
});
