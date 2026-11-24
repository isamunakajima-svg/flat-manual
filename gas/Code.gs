/**
 * FLAT 動画マニュアル（たたき台） - Google Apps Script
 *
 * スクリプトプロパティに設定するもの（プロジェクトの設定 → スクリプト プロパティ）
 *   ADMIN_KEY      : 管理画面のパスワード（好きな文字列）
 *   GEMINI_API_KEY : Google AI Studio で作ったAPIキー
 *   GEMINI_MODEL   : 使うモデル名（省略時は下の DEFAULT_MODEL）
 *
 * 初回は setup() を一度だけ実行すると、必要なシートが作られます。
 */

const DEFAULT_MODEL = 'gemini-3.6-flash';
const LANGS = ['ja', 'en', 'vi'];          // 字幕・手順の言語。先頭が元の言語
const LANG_NAMES = { ja: '日本語', en: 'English', vi: 'Tiếng Việt' };
const GEMINI_BASE = 'https://generativelanguage.googleapis.com';

const SHEETS = {
  manuals: ['id', 'title', 'youtubeId', 'status', 'createdAt', 'updatedAt', 'memo'],
  steps: ['manualId', 'idx', 'start'].concat(flatLang_(['title', 'desc'])),
  subs: ['manualId', 'idx', 'start', 'end'].concat(LANGS.map(l => 'text_' + l)),
  quiz: ['manualId', 'idx', 'question', 'choices', 'answer'],
  staff: ['name', 'active'],
  logs: ['sessionId', 'ts', 'staff', 'manualId', 'watchedSec', 'duration', 'maxPos', 'completed', 'stepViews', 'quizScore', 'quizTotal', 'confirmed'],
  reports: ['ts', 'staff', 'manualId', 'subIdx', 'note', 'done']
};

function flatLang_(fields) {
  const out = [];
  LANGS.forEach(l => fields.forEach(f => out.push(f + '_' + l)));
  return out;
}

/* ========== 初期設定 ========== */

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SHEETS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const header = SHEETS[name];
    sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });
  const staff = ss.getSheetByName('staff');
  if (staff.getLastRow() < 2) {
    staff.getRange(2, 1, 2, 2).setValues([['テスト太郎', true], ['テスト花子', true]]);
  }
}

/* ========== 入口 ========== */

function doGet() {
  return json_({ ok: true, app: 'flat-manual' });
}

function doPost(e) {
  let req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'bad request' });
  }
  try {
    const action = req.action;
    const pub = PUBLIC_ACTIONS[action];
    const adm = ADMIN_ACTIONS[action];
    if (pub) return json_({ ok: true, data: pub(req) });
    if (adm) {
      if (!isAdmin_(req.key)) return json_({ ok: false, error: '管理パスワードが違います' });
      return json_({ ok: true, data: adm(req) });
    }
    return json_({ ok: false, error: 'unknown action: ' + action });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function prop_(k) {
  return PropertiesService.getScriptProperties().getProperty(k);
}

function isAdmin_(key) {
  const k = prop_('ADMIN_KEY');
  return !!k && key === k;
}

/* ========== シート読み書き ========== */

function sheet_(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function readAll_(name) {
  const sh = sheet_(name);
  const values = sh.getDataRange().getValues();
  const header = values.shift();
  return values.map(row => {
    const o = {};
    header.forEach((h, i) => o[h] = row[i]);
    return o;
  });
}

function rowFrom_(name, obj) {
  return SHEETS[name].map(h => (obj[h] === undefined || obj[h] === null) ? '' : obj[h]);
}

/** manualId が一致する行を全部入れ替える */
function replaceRows_(name, manualId, objs) {
  const sh = sheet_(name);
  const header = SHEETS[name];
  const values = sh.getDataRange().getValues();
  values.shift();
  const keep = values.filter(r => String(r[0]) !== String(manualId));
  const add = objs.map(o => rowFrom_(name, Object.assign({ manualId: manualId }, o)));
  const all = keep.concat(add);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, header.length).clearContent();
  if (all.length) sh.getRange(2, 1, all.length, header.length).setValues(all);
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ========== マニュアル ========== */

function listManuals_(includeDrafts) {
  return readAll_('manuals')
    .filter(m => m.id && (includeDrafts || m.status === 'published'))
    .map(m => ({ id: String(m.id), title: m.title, youtubeId: m.youtubeId, status: m.status, updatedAt: m.updatedAt }))
    .sort((a, b) => String(b.updatedAt).localeCompare(String(a.updatedAt)));
}

function getManual_(id, includeDrafts) {
  const m = readAll_('manuals').find(r => String(r.id) === String(id));
  if (!m) throw new Error('マニュアルが見つかりません');
  if (!includeDrafts && m.status !== 'published') throw new Error('このマニュアルはまだ公開されていません');
  const byIdx = (a, b) => Number(a.idx) - Number(b.idx);
  const steps = readAll_('steps').filter(r => String(r.manualId) === String(id)).sort(byIdx).map(r => {
    const o = { start: Number(r.start) || 0, title: {}, desc: {} };
    LANGS.forEach(l => { o.title[l] = r['title_' + l] || ''; o.desc[l] = r['desc_' + l] || ''; });
    return o;
  });
  const subs = readAll_('subs').filter(r => String(r.manualId) === String(id)).sort(byIdx).map(r => {
    const o = { start: Number(r.start) || 0, end: Number(r.end) || 0, text: {} };
    LANGS.forEach(l => o.text[l] = r['text_' + l] || '');
    return o;
  });
  const quiz = readAll_('quiz').filter(r => String(r.manualId) === String(id)).sort(byIdx).map(r => ({
    question: r.question,
    choices: String(r.choices || '').split('|').map(s => s.trim()).filter(s => s),
    answer: Number(r.answer) || 0
  }));
  return {
    id: String(m.id), title: m.title, youtubeId: m.youtubeId, status: m.status, memo: m.memo,
    steps: steps, subs: subs, quiz: quiz, langs: LANGS, langNames: LANG_NAMES
  };
}

function saveManual_(m) {
  return withLock_(() => {
    const sh = sheet_('manuals');
    const now = new Date().toISOString();
    const rows = readAll_('manuals');
    let id = m.id;
    let rowIndex = -1;
    if (id) rowIndex = rows.findIndex(r => String(r.id) === String(id));
    if (!id) id = Utilities.getUuid().slice(0, 8);
    const base = rowIndex >= 0 ? rows[rowIndex] : { createdAt: now, status: 'draft' };
    const rec = Object.assign({}, base, {
      id: id,
      title: m.title !== undefined ? m.title : base.title,
      youtubeId: m.youtubeId !== undefined ? parseYoutubeId_(m.youtubeId) : base.youtubeId,
      status: m.status || base.status || 'draft',
      memo: m.memo !== undefined ? m.memo : base.memo,
      updatedAt: now
    });
    const row = rowFrom_('manuals', rec);
    if (rowIndex >= 0) sh.getRange(rowIndex + 2, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);

    if (m.steps) replaceRows_('steps', id, m.steps.map((s, i) => {
      const o = { idx: i, start: round1_(s.start) };
      LANGS.forEach(l => { o['title_' + l] = (s.title || {})[l] || ''; o['desc_' + l] = (s.desc || {})[l] || ''; });
      return o;
    }));
    if (m.subs) replaceRows_('subs', id, m.subs.map((s, i) => {
      const o = { idx: i, start: round1_(s.start), end: round1_(s.end) };
      LANGS.forEach(l => o['text_' + l] = (s.text || {})[l] || '');
      return o;
    }));
    if (m.quiz) replaceRows_('quiz', id, m.quiz.map((q, i) => ({
      idx: i, question: q.question, choices: (q.choices || []).join(' | '), answer: q.answer
    })));
    return { id: id };
  });
}

function deleteManual_(id) {
  return withLock_(() => {
    ['steps', 'subs', 'quiz'].forEach(n => replaceRows_(n, id, []));
    const sh = sheet_('manuals');
    const rows = readAll_('manuals');
    const i = rows.findIndex(r => String(r.id) === String(id));
    if (i >= 0) sh.deleteRow(i + 2);
    return { deleted: id };
  });
}

function parseYoutubeId_(s) {
  s = String(s || '').trim();
  if (!s) return '';
  const m = s.match(/(?:v=|youtu\.be\/|shorts\/|embed\/|live\/)([A-Za-z0-9_-]{11})/);
  if (m) return m[1];
  if (/^[A-Za-z0-9_-]{11}$/.test(s)) return s;
  throw new Error('YouTubeのURLが読み取れません');
}

function round1_(n) {
  return Math.round((Number(n) || 0) * 10) / 10;
}

/* ========== 視聴記録 ========== */

function logView_(req) {
  if (!req.sessionId || !req.manualId) return { skipped: true };
  return withLock_(() => {
    const sh = sheet_('logs');
    const rec = {
      sessionId: req.sessionId, ts: new Date().toISOString(), staff: req.staff || '(未選択)',
      manualId: req.manualId, watchedSec: round1_(req.watchedSec), duration: round1_(req.duration),
      maxPos: round1_(req.maxPos), completed: !!req.completed,
      stepViews: JSON.stringify(req.stepViews || {}),
      quizScore: req.quizScore === undefined ? '' : req.quizScore,
      quizTotal: req.quizTotal === undefined ? '' : req.quizTotal,
      confirmed: !!req.confirmed
    };
    const found = sh.getRange('A:A').createTextFinder(String(req.sessionId)).matchEntireCell(true).findNext();
    const row = rowFrom_('logs', rec);
    if (found) {
      const r = found.getRow();
      const old = sh.getRange(r, 1, 1, row.length).getValues()[0];
      // 一度ついた「完了」「確認」「テスト結果」は消さない
      const h = SHEETS.logs;
      ['completed', 'confirmed'].forEach(k => { const i = h.indexOf(k); if (old[i] === true) row[i] = true; });
      ['quizScore', 'quizTotal'].forEach(k => { const i = h.indexOf(k); if (row[i] === '' && old[i] !== '') row[i] = old[i]; });
      sh.getRange(r, 1, 1, row.length).setValues([row]);
    } else {
      sh.appendRow(row);
    }
    return { saved: true };
  });
}

/** スタッフ × マニュアルの一覧（一番よく見た回の数字） */
function progressReport_() {
  const manuals = listManuals_(true);
  const staff = readAll_('staff').filter(s => s.name && s.active !== false).map(s => String(s.name));
  const cells = {};
  readAll_('logs').forEach(l => {
    const key = l.staff + '\u0000' + l.manualId;
    const dur = Number(l.duration) || 0;
    const pct = dur ? Math.min(100, Math.round(Number(l.watchedSec) / dur * 100)) : 0;
    const c = cells[key] || { pct: 0, completed: false, confirmed: false, quiz: '', sessions: 0, last: '' };
    c.pct = Math.max(c.pct, pct);
    c.completed = c.completed || l.completed === true;
    c.confirmed = c.confirmed || l.confirmed === true;
    if (l.quizTotal !== '' && l.quizTotal !== undefined) c.quiz = l.quizScore + '/' + l.quizTotal;
    c.sessions++;
    if (String(l.ts) > c.last) c.last = String(l.ts);
    cells[key] = c;
  });
  return {
    manuals: manuals.map(m => ({ id: m.id, title: m.title })),
    staff: staff,
    cells: cells
  };
}

/** どの手順がよく見られているか（マニュアル単位） */
function stepStats_(manualId) {
  const counts = {};
  readAll_('logs').filter(l => String(l.manualId) === String(manualId)).forEach(l => {
    let sv = {};
    try { sv = JSON.parse(l.stepViews || '{}'); } catch (e) {}
    Object.keys(sv).forEach(k => counts[k] = (counts[k] || 0) + Number(sv[k] || 0));
  });
  return counts;
}

function addReport_(req) {
  sheet_('reports').appendRow(rowFrom_('reports', {
    ts: new Date().toISOString(), staff: req.staff || '', manualId: req.manualId,
    subIdx: req.subIdx, note: String(req.note || '').slice(0, 500), done: false
  }));
  return { saved: true };
}

function listReports_(manualId) {
  return readAll_('reports').filter(r => !manualId || String(r.manualId) === String(manualId));
}

/* ========== Gemini：動画のアップロード（分割） ========== */

function geminiKey_() {
  const k = prop_('GEMINI_API_KEY');
  if (!k) throw new Error('GEMINI_API_KEY が設定されていません');
  return k;
}

function headerOf_(res, name) {
  const h = res.getAllHeaders();
  const key = Object.keys(h).find(k => k.toLowerCase() === name.toLowerCase());
  return key ? h[key] : '';
}

/** アップロードの開始。uploadUrl と、分割サイズの単位を返す */
function uploadStart_(req) {
  const res = UrlFetchApp.fetch(GEMINI_BASE + '/upload/v1beta/files', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-goog-api-key': geminiKey_(),
      'X-Goog-Upload-Protocol': 'resumable',
      'X-Goog-Upload-Command': 'start',
      'X-Goog-Upload-Header-Content-Length': String(req.size),
      'X-Goog-Upload-Header-Content-Type': req.mimeType
    },
    payload: JSON.stringify({ file: { display_name: String(req.name || 'manual-video').slice(0, 100) } }),
    muteHttpExceptions: true
  });
  if (res.getResponseCode() >= 300) throw new Error('アップロード開始に失敗: ' + res.getContentText().slice(0, 300));
  const url = headerOf_(res, 'X-Goog-Upload-URL');
  if (!url) throw new Error('アップロード先URLが取得できませんでした');
  return { uploadUrl: url, granularity: Number(headerOf_(res, 'X-Goog-Upload-Chunk-Granularity')) || 8388608 };
}

/** 分割した1片を送る。最後の1片で file 情報が返る */
function uploadChunk_(req) {
  if (String(req.uploadUrl).indexOf(GEMINI_BASE + '/upload/') !== 0) throw new Error('不正なアップロード先');
  const bytes = Utilities.base64Decode(req.data);
  const res = UrlFetchApp.fetch(req.uploadUrl, {
    method: 'post',
    contentType: 'application/octet-stream',
    headers: {
      'X-Goog-Upload-Command': req.last ? 'upload, finalize' : 'upload',
      'X-Goog-Upload-Offset': String(req.offset)
    },
    payload: bytes,
    muteHttpExceptions: true
  });
  if (res.getResponseCode() >= 300) throw new Error('アップロードに失敗: ' + res.getContentText().slice(0, 300));
  if (!req.last) return { received: bytes.length };
  const f = JSON.parse(res.getContentText()).file;
  return { name: f.name, uri: f.uri, mimeType: f.mimeType, state: f.state };
}

/** Gemini側の処理が終わったか（ACTIVE になったら使える） */
function fileStatus_(req) {
  if (!/^files\/[A-Za-z0-9_-]+$/.test(String(req.name))) throw new Error('不正なファイル名');
  const res = UrlFetchApp.fetch(GEMINI_BASE + '/v1beta/' + req.name, {
    headers: { 'x-goog-api-key': geminiKey_() }, muteHttpExceptions: true
  });
  if (res.getResponseCode() >= 300) throw new Error('状態の確認に失敗: ' + res.getContentText().slice(0, 300));
  const f = JSON.parse(res.getContentText());
  return { state: f.state, uri: f.uri, mimeType: f.mimeType };
}

/* ========== Gemini：下書きの生成 ========== */

function callGemini_(parts, schemaHint) {
  const model = prop_('GEMINI_MODEL') || DEFAULT_MODEL;
  const res = UrlFetchApp.fetch(GEMINI_BASE + '/v1beta/models/' + model + ':generateContent', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': geminiKey_() },
    payload: JSON.stringify({
      contents: [{ role: 'user', parts: parts }],
      generationConfig: {
        responseMimeType: 'application/json',
        temperature: 0.3,
        mediaResolution: 'MEDIA_RESOLUTION_LOW'
      }
    }),
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  const body = res.getContentText();
  if (code === 429) throw new Error('Geminiの無料枠の回数上限に達しました。少し時間をおいて試してください');
  if (code >= 300) throw new Error('Geminiのエラー(' + code + '): ' + body.slice(0, 400));
  const data = JSON.parse(body);
  const text = (((data.candidates || [])[0] || {}).content || {}).parts
    ? data.candidates[0].content.parts.map(p => p.text || '').join('')
    : '';
  if (!text) throw new Error('Geminiから返事がありませんでした' + (schemaHint ? '（' + schemaHint + '）' : ''));
  return parseJsonLoose_(text);
}

function parseJsonLoose_(text) {
  try { return JSON.parse(text); } catch (e) {}
  const m = text.match(/[\[{][\s\S]*[\]}]/);
  if (m) return JSON.parse(m[0]);
  throw new Error('AIの返事を読み取れませんでした');
}

const DRAFT_PROMPT = [
  'あなたはパン屋の現場教育マニュアルを作る担当です。',
  '添付の作業動画から、新人スタッフ向けの動画マニュアルの下書きを日本語で作ってください。',
  '',
  '{{SOURCE}}',
  '',
  '次の形のJSONだけを返してください（説明文は不要）。時間はすべて動画の先頭からの秒数（小数1桁まで）。',
  '{',
  '  "title": "マニュアルのタイトル（20字以内）",',
  '  "steps": [ { "start": 0, "title": "手順の見出し（15字以内）", "desc": "やることとコツ（60字以内）" } ],',
  '  "subtitles": [ { "start": 0.0, "end": 3.5, "text": "字幕（1行20字程度まで）" } ],',
  '  "quiz": [ { "question": "確認の問題", "choices": ["選択肢1", "選択肢2", "選択肢3"], "answer": 0 } ]',
  '}',
  '',
  'ルール:',
  '- steps は作業の区切りごとに 3〜10 個。start は昇順で、最初は 0。',
  '- subtitles は動画全体をカバーし、重ならないように。字幕は短く、読みやすく。',
  '- quiz は大事なポイントについて 2〜3 問。answer は正解の選択肢の番号（0始まり）。',
  '- 動画やメモから分からない数値（温度・時間・分量）は作らない。',
].join('\n');

const SOURCE_TEXT = {
  speech: '動画の中で作業者が話している説明を、字幕と手順の主な材料にしてください。言い間違いや口ぐせは整えてください。',
  memo: '動画は無言です。映像と、下の作業メモを照らし合わせて、場面ごとの字幕と手順を作ってください。メモにあるコツや数値は必ず生かしてください。\n\n作業メモ:\n{{MEMO}}',
  voiceover: '動画は無言で、別に録音した説明の音声が添付されています（動画の先頭と同時に録音を始めたもの）。音声の時間をそのまま動画の時間として使い、音声の説明を字幕と手順の主な材料にしてください。'
};

/** 動画（と音声・メモ）から下書きを作り、下書きとして保存する */
function generateDraft_(req) {
  if (!req.video || !req.video.uri) throw new Error('動画がありません');
  let source = SOURCE_TEXT[req.mode] || SOURCE_TEXT.speech;
  source = source.replace('{{MEMO}}', String(req.memo || '(メモなし)').slice(0, 3000));
  const parts = [{ fileData: { mimeType: req.video.mimeType, fileUri: req.video.uri } }];
  if (req.mode === 'voiceover' && req.audio && req.audio.uri) {
    parts.push({ fileData: { mimeType: req.audio.mimeType, fileUri: req.audio.uri } });
  }
  parts.push({ text: DRAFT_PROMPT.replace('{{SOURCE}}', source) });

  const out = callGemini_(parts, '下書き');
  const ja = LANGS[0];
  const steps = (out.steps || []).map(s => {
    const o = { start: Number(s.start) || 0, title: {}, desc: {} };
    o.title[ja] = s.title || ''; o.desc[ja] = s.desc || '';
    return o;
  });
  const subs = (out.subtitles || []).map(s => {
    const o = { start: Number(s.start) || 0, end: Number(s.end) || 0, text: {} };
    o.text[ja] = s.text || '';
    return o;
  });
  const quiz = (out.quiz || []).map(q => ({
    question: q.question || '', choices: (q.choices || []).map(String), answer: Number(q.answer) || 0
  }));
  return saveManual_({
    id: req.id || '',
    title: req.title || out.title || '新しいマニュアル',
    youtubeId: req.youtubeId || '',
    status: 'draft',
    memo: req.memo || '',
    steps: steps, subs: subs, quiz: quiz
  });
}

/** 日本語から他の言語へ翻訳して保存する（手順と字幕） */
function translateManual_(req) {
  const m = getManual_(req.id, true);
  const src = LANGS[0];
  const targets = (req.langs && req.langs.length ? req.langs : LANGS.slice(1)).filter(l => LANGS.indexOf(l) > 0);
  const texts = [];
  m.steps.forEach(s => { texts.push(s.title[src]); texts.push(s.desc[src]); });
  m.subs.forEach(s => texts.push(s.text[src]));
  if (!texts.length) throw new Error('翻訳する文章がありません');

  targets.forEach(lang => {
    const prompt = [
      'パン屋の作業マニュアルの文章です。次のJSON配列の各要素を ' + LANG_NAMES[lang] + ' に翻訳してください。',
      '新人スタッフ向けに、短く分かりやすく。空文字はそのまま空文字に。',
      '同じ長さ・同じ順番のJSON配列（文字列のみ）だけを返してください。',
      JSON.stringify(texts)
    ].join('\n');
    const arr = callGemini_([{ text: prompt }], '翻訳');
    if (!Array.isArray(arr) || arr.length !== texts.length) throw new Error(LANG_NAMES[lang] + 'の翻訳の数が合いませんでした。もう一度試してください');
    let k = 0;
    m.steps.forEach(s => { s.title[lang] = arr[k++]; s.desc[lang] = arr[k++]; });
    m.subs.forEach(s => s.text[lang] = arr[k++]);
  });
  saveManual_({ id: m.id, steps: m.steps, subs: m.subs });
  return { translated: targets };
}

/* ========== API一覧 ========== */

const PUBLIC_ACTIONS = {
  list: () => listManuals_(false),
  get: req => getManual_(req.id, false),
  staff: () => readAll_('staff').filter(s => s.name && s.active !== false).map(s => String(s.name)),
  log: req => logView_(req),
  report: req => addReport_(req)
};

const ADMIN_ACTIONS = {
  ping: () => ({ admin: true }),
  listAll: () => listManuals_(true),
  getDraft: req => getManual_(req.id, true),
  save: req => saveManual_(req.manual),
  remove: req => deleteManual_(req.id),
  uploadStart: req => uploadStart_(req),
  uploadChunk: req => uploadChunk_(req),
  fileStatus: req => fileStatus_(req),
  generate: req => generateDraft_(req),
  translate: req => translateManual_(req),
  progress: () => progressReport_(),
  stepStats: req => stepStats_(req.id),
  reports: req => listReports_(req.id)
};
