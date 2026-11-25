/**
 * FLAT 動画マニュアル（たたき台） - Google Apps Script
 *
 * スクリプトプロパティに設定するもの（プロジェクトの設定 → スクリプト プロパティ）
 *   ADMIN_KEY      : 管理画面のパスワード（好きな文字列）
 *   GEMINI_API_KEY : Google AI Studio で作ったAPIキー
 *   GEMINI_MODEL   : 使うモデル名（省略時は下の DEFAULT_MODEL）
 *   LINE_CHANNEL_ID: LIFFを作ったLINEログインチャネルのチャネルID
 *                    （入れるとLINEでの本人確認が有効になる。空なら名前を選ぶ方式のまま）
 *   LINE_MESSAGING_TOKEN: 承認待ちの通知を送るLINE公式アカウントのチャネルアクセストークン（長期）
 *   NOTIFY_USER_IDS : 通知を受け取る人のLINEユーザーID（staffシートの lineUserId。複数ならカンマ区切り）
 *
 * 初回は setup() を一度だけ実行すると、必要なシートが作られます。
 */

const DEFAULT_MODEL = 'gemini-3.6-flash';
const LANGS = ['ja', 'en', 'vi'];          // 字幕・手順の言語。先頭が元の言語
const LANG_NAMES = { ja: '日本語', en: 'English', vi: 'Tiếng Việt' };
const GEMINI_BASE = 'https://generativelanguage.googleapis.com';
const ADMIN_PAGE_URL = 'https://isamunakajima-svg.github.io/flat-manual/admin.html';

const SHEETS = {
  manuals: ['id', 'title', 'youtubeId', 'status', 'createdAt', 'updatedAt', 'memo', 'note'],
  steps: ['manualId', 'idx', 'start'].concat(flatLang_(['title', 'desc'])),
  subs: ['manualId', 'idx', 'start', 'end'].concat(LANGS.map(l => 'text_' + l)),
  quiz: ['manualId', 'idx', 'question', 'choices', 'answer'],
  staff: ['name', 'active', 'lineUserId', 'lineName', 'approved', 'registeredAt'],
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
  // active と approved をチェックボックスにする
  staff.getRange('B2:B300').insertCheckboxes();
  staff.getRange('E2:E300').insertCheckboxes();
}

/** 列が増えたときに見出しだけ追加する（setup をやり直さなくても動くように） */
function ensureHeaders_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ['staff', 'logs', 'manuals'].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    const header = SHEETS[name];
    if (sh.getLastColumn() < header.length) {
      sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
    }
  });
}

/* ========== 入口 ========== */

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.op === 'approve') return approveFromLink_(p.t);
  return json_({ ok: true, app: 'flat-manual' });
}

/* ========== LINEの通知の「承認する」ボタンから開くページ ========== */

function htmlPage_(title, body) {
  const html = '<div style="font-family:-apple-system,\'Hiragino Sans\',sans-serif;padding:32px 20px;max-width:480px;margin:auto;text-align:center">' +
    '<h2 style="font-size:20px">' + title + '</h2><p style="color:#555;line-height:1.7">' + body + '</p></div>';
  return HtmlService.createHtmlOutput(html).setTitle('FLAT マニュアル')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function escHtml_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/** 承認リンクを作る（7日間有効、1回だけ使える） */
function makeApproveLink_(name) {
  const token = Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('approve_' + token,
    JSON.stringify({ name: String(name), exp: Date.now() + 7 * 24 * 3600 * 1000 }));
  return ScriptApp.getService().getUrl() + '?op=approve&t=' + token;
}

function approveFromLink_(token) {
  if (!/^[0-9a-f]{32}$/.test(String(token || ''))) return htmlPage_('リンクが正しくありません', '');
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty('approve_' + token);
  if (!raw) return htmlPage_('このリンクは使えません', 'すでに使われたか、期限が切れています。管理画面から承認してください。<br><a href="' + ADMIN_PAGE_URL + '">管理画面を開く</a>');
  const info = JSON.parse(raw);
  if (Date.now() > info.exp) {
    props.deleteProperty('approve_' + token);
    return htmlPage_('期限が切れています', '管理画面から承認してください。<br><a href="' + ADMIN_PAGE_URL + '">管理画面を開く</a>');
  }
  const result = withLock_(() => {
    const row = staffRows_().find(r => String(r.name) === info.name);
    if (!row || !row.lineUserId) return 'missing';
    if (row.active !== true) return 'inactive';
    if (row.approved === true) return 'already';
    sheet_('staff').getRange(row._row, SHEETS.staff.indexOf('approved') + 1).setValue(true);
    return 'ok';
  });
  props.deleteProperty('approve_' + token);
  const name = escHtml_(info.name);
  if (result === 'ok') return htmlPage_('✅ 承認しました', name + ' さんがマニュアルを見られるようになりました。<br>この画面は閉じて大丈夫です。');
  if (result === 'already') return htmlPage_('承認済みです', name + ' さんはすでに承認されています。');
  if (result === 'inactive') return htmlPage_('承認できません', name + ' さんは無効になっています。管理画面で確認してください。');
  return htmlPage_('承認できません', name + ' さんの登録が見つかりません（LINE連携が外された可能性があります）。');
}

function doPost(e) {
  let req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'bad request' });
  }
  try {
    ensureHeaders_();
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
    .map(m => ({ id: String(m.id), title: m.title, youtubeId: m.youtubeId, status: m.status, updatedAt: m.updatedAt, note: m.note || '' }))
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
    id: String(m.id), title: m.title, youtubeId: m.youtubeId, status: m.status, memo: m.memo, note: m.note || '',
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
      note: m.note !== undefined ? m.note : (base.note || ''),
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
  PropertiesService.getScriptProperties().deleteProperty('job_' + id);
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

/* ========== だれが見ているか（LINEでの本人確認） ========== */

function lineEnabled_() {
  return !!prop_('LINE_CHANNEL_ID');
}

/** LIFFのIDトークンをLINEに確認してもらう（結果は30分キャッシュ） */
function verifyLine_(idToken) {
  if (!idToken) throw new Error('AUTH:LINEでログインしてください');
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, idToken)
    .map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
  const cache = CacheService.getScriptCache();
  const hit = cache.get('tok_' + digest);
  if (hit) return JSON.parse(hit);
  const res = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify', {
    method: 'post',
    payload: { id_token: idToken, client_id: prop_('LINE_CHANNEL_ID') },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) throw new Error('AUTH:ログインの有効期限が切れました');
  const d = JSON.parse(res.getContentText());
  const who = { userId: d.sub, lineName: d.name || '' };
  cache.put('tok_' + digest, JSON.stringify(who), 1800);
  return who;
}

function staffRows_() {
  return readAll_('staff').map((r, i) => Object.assign(r, { _row: i + 2 })).filter(r => r.name);
}

/**
 * リクエストの送り主を調べる。
 * status: ok（見てよい）/ unregistered（名前をまだ選んでいない）/ pending（承認待ち）/ blocked（無効）
 */
function identify_(req) {
  if (isAdmin_(req.key)) return { status: 'ok', name: '管理者', admin: true };
  if (!lineEnabled_()) {
    // LINE未設定のあいだは、これまでどおり画面で選んだ名前を使う
    return { status: 'ok', name: req.staff || '(未選択)', legacy: true };
  }
  const who = verifyLine_(req.idToken);
  const row = staffRows_().find(r => String(r.lineUserId) === who.userId);
  if (!row) return { status: 'unregistered', lineName: who.lineName, userId: who.userId };
  if (row.active !== true) return { status: 'blocked', name: row.name };
  if (row.approved !== true) return { status: 'pending', name: row.name };
  return { status: 'ok', name: String(row.name), userId: who.userId };
}

function requireViewer_(req) {
  const me = identify_(req);
  if (me.status === 'blocked') throw new Error('BLOCKED:このアカウントでは閲覧できません');
  if (me.status === 'pending') throw new Error('PENDING:承認待ちです');
  if (me.status === 'unregistered') throw new Error('UNREGISTERED:名前を選んでください');
  return me;
}

/** 画面を開いたときの確認。名前を選ぶ必要があれば候補も返す */
function me_(req) {
  const me = identify_(req);
  const out = { status: me.status, name: me.name || '', lineName: me.lineName || '', mode: me.admin ? 'admin' : (me.legacy ? 'legacy' : 'line') };
  const notifyIds = String(prop_('NOTIFY_USER_IDS') || '').split(',').map(x => x.trim());
  out.canEdit = !!me.admin || (me.status === 'ok' && !!me.userId && notifyIds.indexOf(me.userId) >= 0);
  if (me.status === 'unregistered') {
    out.candidates = staffRows_().filter(r => r.active === true && !r.lineUserId).map(r => String(r.name));
  }
  if (me.legacy) {
    out.candidates = staffRows_().filter(r => r.active === true).map(r => String(r.name));
  }
  return out;
}

/** 初回：名前を選んでLINEと結び付ける（承認待ちになる） */
function register_(req) {
  if (!lineEnabled_()) throw new Error('LINEの設定がまだです');
  const who = verifyLine_(req.idToken);
  return withLock_(() => {
    const rows = staffRows_();
    if (rows.some(r => String(r.lineUserId) === who.userId)) throw new Error('すでに登録されています');
    const row = rows.find(r => String(r.name) === String(req.name));
    if (!row || row.active !== true) throw new Error('その名前は選べません');
    if (row.lineUserId) throw new Error('その名前はほかの人が登録済みです。管理者に伝えてください');
    const sh = sheet_('staff');
    const h = SHEETS.staff;
    sh.getRange(row._row, h.indexOf('lineUserId') + 1).setValue(who.userId);
    sh.getRange(row._row, h.indexOf('lineName') + 1).setValue(who.lineName);
    sh.getRange(row._row, h.indexOf('approved') + 1).setValue(false);
    sh.getRange(row._row, h.indexOf('registeredAt') + 1).setValue(new Date().toISOString());
    notifyApproval_(row.name, who.lineName);
    return { status: 'pending', name: row.name };
  });
}

/** 承認待ちの通知（「承認する」ボタン付き） */
function notifyApproval_(name, lineName) {
  let link = '';
  try { link = makeApproveLink_(name); } catch (e) { console.warn('承認リンクを作れませんでした: ' + e); }
  const text = name + ' さん（LINEの名前：' + (lineName || '不明') + '）が承認を求めています。';
  const actions = [];
  if (link) actions.push({ type: 'uri', label: '承認する', uri: link });
  actions.push({ type: 'uri', label: '管理画面を開く', uri: ADMIN_PAGE_URL });
  return notifyAdmins_('【FLAT マニュアル】\n' + text + '\n\n管理画面\n' + ADMIN_PAGE_URL, {
    type: 'template',
    altText: '【FLAT マニュアル】' + text,
    template: { type: 'buttons', text: ('【FLAT マニュアル】' + text).slice(0, 160), actions: actions }
  });
}

/** 管理者のLINEに通知する。設定がなければ何もしない。失敗しても登録は止めない */
function notifyAdmins_(text, richMessage) {
  const token = prop_('LINE_MESSAGING_TOKEN');
  const ids = String(prop_('NOTIFY_USER_IDS') || '').split(',').map(x => x.trim()).filter(x => x);
  if (!token || !ids.length) return false;
  try {
    const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/multicast', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify({ to: ids, messages: [richMessage || { type: 'text', text: text.slice(0, 4900) }] }),
      muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) console.warn('LINE通知に失敗: ' + res.getContentText());
    return res.getResponseCode() === 200;
  } catch (e) {
    console.warn('LINE通知に失敗: ' + e);
    return false;
  }
}

/** 通知の設定を確かめる用。エディタでこの関数を選んで「実行」すると、テストの通知が届く */
function testNotify() {
  const ok = notifyAdmins_('【FLAT マニュアル】通知のテストです。これが届いていれば設定はOKです。');
  // ボタン付きの見た目の確認（押しても、テスト用の名前なので承認はされません）
  notifyAdmins_('', {
    type: 'template', altText: '【FLAT マニュアル】ボタン付き通知のテスト',
    template: { type: 'buttons', text: '【FLAT マニュアル】ボタン付き通知のテストです',
      actions: [{ type: 'uri', label: '管理画面を開く', uri: ADMIN_PAGE_URL }] }
  });
  console.log(ok ? '送りました' : '送れませんでした（LINE_MESSAGING_TOKEN と NOTIFY_USER_IDS、公式アカウントを友だち追加しているかを確認してください）');
}

/* --- 管理画面からのスタッフ管理 --- */

function staffAdminList_() {
  return staffRows_().map(r => ({
    name: String(r.name), active: r.active === true, approved: r.approved === true,
    linked: !!r.lineUserId, lineName: r.lineName || '', registeredAt: r.registeredAt ? String(r.registeredAt) : ''
  }));
}

function staffUpdate_(req) {
  return withLock_(() => {
    const row = staffRows_().find(r => String(r.name) === String(req.name));
    if (!row) throw new Error('スタッフが見つかりません');
    const sh = sheet_('staff');
    const h = SHEETS.staff;
    const set = (k, v) => sh.getRange(row._row, h.indexOf(k) + 1).setValue(v);
    if (req.op === 'approve') set('approved', true);
    else if (req.op === 'disable') set('active', false);
    else if (req.op === 'enable') set('active', true);
    else if (req.op === 'unlink') { set('lineUserId', ''); set('lineName', ''); set('approved', false); set('registeredAt', ''); }
    else throw new Error('不明な操作');
    return staffAdminList_();
  });
}

function staffAdd_(req) {
  const name = String(req.name || '').trim();
  if (!name) throw new Error('名前を入れてください');
  return withLock_(() => {
    if (staffRows_().some(r => String(r.name) === name)) throw new Error('同じ名前がすでにあります');
    sheet_('staff').appendRow(rowFrom_('staff', { name: name, active: true, approved: false }));
    return staffAdminList_();
  });
}

/* ========== 視聴記録 ========== */

function logView_(req) {
  if (!req.sessionId || !req.manualId) return { skipped: true };
  return withLock_(() => {
    const sh = sheet_('logs');
    const rec = {
      sessionId: req.sessionId, ts: new Date().toISOString(), staff: req._viewer || req.staff || '(未選択)',
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
    ts: new Date().toISOString(), staff: req._viewer || req.staff || '', manualId: req.manualId,
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
    note: '',
    steps: steps, subs: subs, quiz: quiz
  });
}

/* ========== 下書きを裏で作る（画面を閉じても進む） ========== */

/** 動画を送り終わったら呼ぶ。マニュアルを「作成中」で作り、タイマーで続きをやる */
function queueDraft_(req) {
  if (!req.video || !req.video.name) throw new Error('動画がありません');
  const saved = saveManual_({
    title: req.title || '（作成中）', youtubeId: req.youtubeId || '', status: 'processing',
    memo: req.memo || '', note: 'AIが下書きを作っています'
  });
  const job = {
    id: saved.id, title: req.title || '', youtubeId: req.youtubeId || '', mode: req.mode,
    memo: String(req.memo || '').slice(0, 3000), video: req.video, audio: req.audio || null,
    queuedAt: Date.now(), tries: 0
  };
  PropertiesService.getScriptProperties().setProperty('job_' + saved.id, JSON.stringify(job));
  scheduleJobs_(5);
  return { id: saved.id, queued: true };
}

function scheduleJobs_(seconds) {
  const exists = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'processJobs');
  if (!exists) ScriptApp.newTrigger('processJobs').timeBased().after(seconds * 1000).create();
}

/** タイマーから呼ばれる。待っている下書きを順に作る */
function processJobs() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'processJobs')
    .forEach(t => ScriptApp.deleteTrigger(t));
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  const started = Date.now();
  let waiting = false;
  Object.keys(all).filter(k => k.indexOf('job_') === 0).forEach(key => {
    if (Date.now() - started > 4 * 60 * 1000) { waiting = true; return; } // 時間切れ前に次回へ回す
    const job = JSON.parse(all[key]);
    try {
      // Geminiの読み込みが終わっているか
      const files = [job.video].concat(job.audio ? [job.audio] : []);
      for (const f of files) {
        const st = fileStatus_({ name: f.name });
        if (st.state === 'FAILED') throw new Error('動画の読み込みに失敗しました（形式が合わない可能性があります）');
        if (st.state !== 'ACTIVE') {
          if (Date.now() - job.queuedAt > 30 * 60 * 1000) throw new Error('動画の読み込みが30分たっても終わりませんでした');
          waiting = true;
          return;
        }
        f.uri = st.uri || f.uri;
        f.mimeType = st.mimeType || f.mimeType;
      }
      generateDraft_({
        id: job.id, title: job.title, youtubeId: job.youtubeId, mode: job.mode, memo: job.memo,
        video: job.video, audio: job.audio
      });
      props.deleteProperty(key);
      const m = readAll_('manuals').find(r => String(r.id) === String(job.id)) || {};
      notifyDone_(job.id, m.title || job.title, '');
    } catch (err) {
      const msg = String(err && err.message || err);
      job.tries = (job.tries || 0) + 1;
      // 回数上限などの一時的なエラーは、3回まで時間をおいてやり直す
      if (job.tries < 3 && /回数上限|429|500|503|timed out|タイムアウト/i.test(msg)) {
        props.setProperty(key, JSON.stringify(job));
        waiting = true;
        return;
      }
      props.deleteProperty(key);
      saveManual_({ id: job.id, status: 'error', note: msg.slice(0, 300) });
      notifyDone_(job.id, job.title || '（タイトルなし）', msg);
    }
  });
  if (waiting) scheduleJobs_(60);
}

/** 下書きができた（または失敗した）ことをLINEで知らせる */
function notifyDone_(id, title, errorMsg) {
  const editUrl = ADMIN_PAGE_URL.replace('admin.html', 'edit.html') + '?id=' + encodeURIComponent(id);
  const text = errorMsg
    ? '【FLAT マニュアル】「' + title + '」の下書きを作れませんでした。\n' + errorMsg
    : '【FLAT マニュアル】「' + title + '」の下書きができました。確認して公開してください。';
  notifyAdmins_(text, {
    type: 'template', altText: text.slice(0, 400),
    template: { type: 'buttons', text: text.slice(0, 160),
      actions: [{ type: 'uri', label: errorMsg ? '管理画面を開く' : '編集する', uri: errorMsg ? ADMIN_PAGE_URL : editUrl }] }
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
  me: req => me_(req),
  register: req => register_(req),
  list: req => { requireViewer_(req); return listManuals_(false); },
  get: req => { requireViewer_(req); return getManual_(req.id, false); },
  staff: () => {
    if (lineEnabled_()) return [];
    return readAll_('staff').filter(s => s.name && s.active !== false).map(s => String(s.name));
  },
  log: req => { req._viewer = requireViewer_(req).name; return logView_(req); },
  report: req => { req._viewer = requireViewer_(req).name; return addReport_(req); }
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
  queue: req => queueDraft_(req),
  translate: req => translateManual_(req),
  progress: () => progressReport_(),
  stepStats: req => stepStats_(req.id),
  reports: req => listReports_(req.id),
  staffList: () => staffAdminList_(),
  staffUpdate: req => staffUpdate_(req),
  staffAdd: req => staffAdd_(req)
};
