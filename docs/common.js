/* 共通の小さな道具 */

// LINEの通知などから https://liff.line.me/…/edit.html?id=… を開くと、最初は一覧のページに
// liff.state=/edit.html?id=… が付いて届く。管理パスワードのある端末でも目的のページに移るように、ここで先に移動する。
(function () {
  try {
    const st = new URLSearchParams(location.search).get('liff.state');
    const m = st && st.match(/([a-z]+\.html(?:\?[^#]*)?)/);
    const here = location.pathname.split('/').pop() || 'index.html';
    if (m && m[1].split('?')[0] !== here) {
      location.replace(location.pathname.replace(/[^/]*$/, '') + m[1] + location.hash);
    }
  } catch (e) {}
})();

(function () {
  const GAS_URL = (window.FLAT_MANUAL_CONFIG || {}).GAS_URL;

  const LIFF_ID = (window.FLAT_MANUAL_CONFIG || {}).LIFF_ID || '';
  const auth = { idToken: '', mode: '' };

  // GASへの問い合わせ。Content-Type を付けない（text/plain）ので事前確認が要らない
  // まとめて先に取っておいた返事（同じ呼び出しが来たら、通信せずにこれを使う）
  const preloaded = {};
  const preKey = (action, params) => {
    const p = Object.assign({}, params || {}); delete p.key; delete p.idToken;
    return action + '|' + JSON.stringify(p);
  };
  // 何回分もの呼び出しを、1回の通信にまとめる。calls: [['list'], ['get', { id }], …]
  async function batch(calls) {
    const res = await rawApi('batch', { calls: calls.map(c => ({ action: c[0], params: c[1] || {} })) });
    calls.forEach((c, i) => { preloaded[preKey(c[0], c[1])] = res[i]; });
    return res;
  }

  async function api(action, params) {
    const k = preKey(action, params);
    if (preloaded[k]) {
      const r = preloaded[k]; delete preloaded[k];
      if (r.ok) return r.data;
      await handleAuth(r.error || 'エラー');
    }
    return rawApi(action, params);
  }

  async function handleAuth(msg) {
    // LINEのログインが切れていたら、ログインし直す
    if (msg.indexOf('AUTH:') === 0 && window.liff && LIFF_ID) {
      try { liff.logout(); } catch (e) {}
      liff.login({ redirectUri: location.href });
      await new Promise(() => {});
    }
    throw new Error(msg.replace(/^(AUTH|BLOCKED|PENDING|UNREGISTERED):/, ''));
  }

  async function rawApi(action, params) {
    const extra = {};
    if (auth.idToken) extra.idToken = auth.idToken;
    const adminKey = store.get('adminKey', '');
    if (adminKey && !('key' in (params || {}))) extra.key = adminKey;
    const body = JSON.stringify(Object.assign({ action: action }, extra, params || {}));
    const res = await fetch(GAS_URL, { method: 'POST', body: body, redirect: 'follow' });
    if (!res.ok) throw new Error('通信エラー (' + res.status + ')');
    const json = await res.json();
    if (!json.ok) await handleAuth(json.error || 'エラー');
    return json.data;
  }

  // ページを閉じるときの記録用（返事は待たない）
  function beacon(action, params) {
    const extra = {};
    if (auth.idToken) extra.idToken = auth.idToken;
    const adminKey = store.get('adminKey', '');
    if (adminKey) extra.key = adminKey;
    const body = JSON.stringify(Object.assign({ action: action }, extra, params || {}));
    try {
      if (navigator.sendBeacon && navigator.sendBeacon(GAS_URL, new Blob([body], { type: 'text/plain' }))) return;
    } catch (e) {}
    try { fetch(GAS_URL, { method: 'POST', body: body, keepalive: true }); } catch (e) {}
  }

  // ブラウザに小さな設定を残す（使えない環境でも動くように）
  const store = {
    get(k, d) { try { const v = localStorage.getItem('flatManual.' + k); return v === null ? d : v; } catch (e) { return d; } },
    set(k, v) { try { localStorage.setItem('flatManual.' + k, v); } catch (e) {} },
    remove(k) { try { localStorage.removeItem('flatManual.' + k); } catch (e) {} }
  };

  function fmtTime(sec) {
    sec = Math.max(0, Number(sec) || 0);
    const m = Math.floor(sec / 60);
    const s = Math.floor(sec % 60);
    return m + ':' + String(s).padStart(2, '0');
  }

  // "1:23" や "83" を秒に
  function parseTime(str) {
    str = String(str || '').trim();
    if (!str) return 0;
    if (str.indexOf(':') >= 0) {
      const [m, s] = str.split(':');
      return (Number(m) || 0) * 60 + (Number(s) || 0);
    }
    return Number(str) || 0;
  }

  function esc(s) {
    return String(s === undefined || s === null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function qs(name) {
    return new URLSearchParams(location.search).get(name);
  }

  function uid() {
    return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
  }

  function toast(msg, isError) {
    let el = document.getElementById('toast');
    if (!el) {
      el = document.createElement('div');
      el.id = 'toast';
      document.body.appendChild(el);
    }
    el.textContent = msg;
    el.className = 'show' + (isError ? ' error' : '');
    clearTimeout(el._t);
    el._t = setTimeout(() => { el.className = ''; }, isError ? 6000 : 2500);
  }

  // YouTube IFrame API を1回だけ読み込む
  let ytPromise = null;
  function loadYouTube() {
    if (ytPromise) return ytPromise;
    ytPromise = new Promise(resolve => {
      if (window.YT && window.YT.Player) return resolve(window.YT);
      const prev = window.onYouTubeIframeAPIReady;
      window.onYouTubeIframeAPIReady = function () { if (prev) prev(); resolve(window.YT); };
      const s = document.createElement('script');
      s.src = 'https://www.youtube.com/iframe_api';
      document.head.appendChild(s);
    });
    return ytPromise;
  }

  // 今の時間にあたる要素の番号（list は start 昇順）
  function indexAt(list, t, useEnd) {
    let found = -1;
    for (let i = 0; i < list.length; i++) {
      if (list[i].start <= t + 0.05) found = i;
      else break;
    }
    if (useEnd && found >= 0 && t >= (list[found].end || Infinity)) return -1;
    return found;
  }

  function loadScript(src) {
    return new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = src; s.onload = resolve; s.onerror = () => reject(new Error('読み込みに失敗: ' + src));
      document.head.appendChild(s);
    });
  }

  /**
   * スタッフ用の画面を開くときの確認。
   * 見てよいときは { name, mode } を返す。名前選び・承認待ち・閲覧不可のときは gate に案内を出して止まる。
   * mode: 'line'（LINEで本人確認）/ 'legacy'（画面で名前を選ぶ）/ 'admin'（管理パスワードのある端末）
   */
  // LINEでログインしてIDトークンを持つ（LIFF_ID が無ければ何もしない）
  async function lineLogin() {
    if (!LIFF_ID) return false;
    await loadScript('https://static.line-scdn.net/liff/edge/2/sdk.js');
    await liff.init({ liffId: LIFF_ID });
    if (new URLSearchParams(location.search).has('liff.state')) await new Promise(() => {}); // LIFFが目的のページへ移動中
    if (!liff.isLoggedIn()) { liff.login({ redirectUri: location.href }); await new Promise(() => {}); }
    auth.idToken = liff.getIDToken() || '';
    return true;
  }

  // extra：ログインの確認と一緒に取っておく呼び出し（例 [['list'], ['meta']]）
  async function ready(gateEl, extra) {
    if (LIFF_ID && !store.get('adminKey', '')) await lineLogin();
    if (extra && extra.length) { try { await batch([['me']].concat(extra)); } catch (e) {} }
    const me = await api('me');
    auth.mode = me.mode;
    if (me.status === 'ok') {
      if (me.mode === 'legacy') me.name = store.get('staff', '');
      return me;
    }
    const show = html => {
      gateEl.innerHTML = html; gateEl.classList.remove('hidden');
      const l = document.getElementById('loading'); if (l) l.classList.add('hidden');
    };
    if (me.status === 'blocked') {
      show('<div class="card"><h2>閲覧できません</h2><p class="muted">このアカウントではマニュアルを見られません。</p></div>');
    } else if (me.status === 'pending') {
      show(`<div class="card"><h2>承認待ちです</h2>
        <p>${esc(me.name)} さんとして登録しました。管理者が承認すると見られるようになります。</p>
        <button id="gateReload" style="width:100%">承認されたか確認する</button></div>`);
      document.getElementById('gateReload').onclick = () => location.reload();
    } else {
      const opts = (me.candidates || []).map(n => `<option>${esc(n)}</option>`).join('');
      show(`<div class="card"><h2>はじめに、あなたの名前を選んでください</h2>
        <p class="muted">${me.lineName ? 'LINEの名前：' + esc(me.lineName) + '<br>' : ''}一度選ぶと、次からは自動で記録されます。</p>
        <select id="gateName"><option value="">選んでください</option>${opts}</select>
        <button class="primary" id="gateSend" style="width:100%;margin-top:10px">この名前で登録する</button>
        <p class="muted" style="margin-top:8px">自分の名前がないときは、管理者に伝えてください。</p></div>`);
      document.getElementById('gateSend').onclick = async () => {
        const name = document.getElementById('gateName').value;
        if (!name) { toast('名前を選んでください', true); return; }
        if (!window.confirm(name + ' さんで登録します。よろしいですか？')) return;
        try { await api('register', { name }); location.reload(); }
        catch (e) { toast(e.message, true); }
      };
    }
    await new Promise(() => {});
  }

    // タグを「#クロワッサン #成形」の形で表示
  function tagsHtml(tags) {
    return (tags || []).map(t => `<span class="tag" data-tag="${esc(t)}">#${esc(t)}</span>`).join('');
  }
  function parseTags(str) {
    const out = [];
    String(str || '').split(/[\s,、，#＃]+/).map(t => t.trim()).filter(t => t).forEach(t => { if (out.indexOf(t) < 0) out.push(t); });
    return out;
  }

  // レシピの写真：GAS経由で読み込む（同じ写真は1回だけ）
  // 写真：一度見たものはスマホの中にしまっておき、まだのものは数枚ずつまとめて取りに行く
  const imgCache = {};
  let dbP = null;
  function db() {
    if (!dbP) dbP = new Promise(resolve => {
      try {
        const r = indexedDB.open('flatManualPhotos', 1);
        r.onupgradeneeded = () => r.result.createObjectStore('p');
        r.onsuccess = () => resolve(r.result);
        r.onerror = () => resolve(null);
      } catch (e) { resolve(null); }
    });
    return dbP;
  }
  async function dbGet(id) {
    const d = await db(); if (!d) return '';
    return new Promise(resolve => {
      try { const q = d.transaction('p').objectStore('p').get(id); q.onsuccess = () => resolve(q.result || ''); q.onerror = () => resolve(''); }
      catch (e) { resolve(''); }
    });
  }
  async function dbPut(id, v) {
    const d = await db(); if (!d || !v) return;
    try { d.transaction('p', 'readwrite').objectStore('p').put(v, id); } catch (e) {}
  }
  let queue = [], timer = null;
  function flush() {
    timer = null;
    const items = queue; queue = [];
    for (let i = 0; i < items.length; i += 6) {
      const part = items.slice(i, i + 6);
      rawApi('images', { ids: part.map(x => x.id) }).then(map => {
        part.forEach(x => { const v = (map || {})[x.id]; if (v) { dbPut(x.id, v); x.ok(v); } else x.ng(new Error('写真がありません')); });
      }).catch(e => part.forEach(x => x.ng(e)));
    }
  }
  function image(id) {
    if (!id) return Promise.resolve('');
    if (!imgCache[id]) imgCache[id] = dbGet(id).then(hit => hit || new Promise((ok, ng) => {
      queue.push({ id, ok, ng });
      if (!timer) timer = setTimeout(flush, 30);
    })).catch(e => { delete imgCache[id]; throw e; });
    return imgCache[id];
  }
  // <img data-photo="ID"> を順に読み込む
  function loadPhotos(root) {
    (root || document).querySelectorAll('img[data-photo]').forEach(el => {
      const id = el.getAttribute('data-photo');
      if (!id || el.dataset.loaded === id) return;
      el.dataset.loaded = id;
      image(id).then(src => { el.src = src; }).catch(() => { el.alt = '写真を読み込めませんでした'; });
    });
  }
  // スマホの写真を縮めて data URL にする（向きはブラウザが直してくれる）
  function resizeImage(fileOrSrc, max, quality) {
    return new Promise((resolve, reject) => {
      const img = new Image();
      img.onload = () => {
        const r = Math.min(1, max / Math.max(img.naturalWidth, img.naturalHeight));
        const c = document.createElement('canvas');
        c.width = Math.round(img.naturalWidth * r); c.height = Math.round(img.naturalHeight * r);
        c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
        resolve(c.toDataURL('image/jpeg', quality || 0.82));
        if (typeof fileOrSrc !== 'string') URL.revokeObjectURL(img.src);
      };
      img.onerror = () => reject(new Error('画像を読み込めませんでした'));
      img.src = typeof fileOrSrc === 'string' ? fileOrSrc : URL.createObjectURL(fileOrSrc);
    });
  }
  // 分量を何倍かにする（数字でないものはそのまま）
  function scaleAmount(amount, ratio) {
    const n = Number(String(amount).replace(/,/g, ''));
    if (amount === '' || isNaN(n) || ratio === 1) return amount;
    const v = n * ratio;
    return String(v >= 100 ? Math.round(v) : Math.round(v * 10) / 10);
  }

  // 一覧のサムネイル
  function thumbHtml(m) {
    if (m.type === 'recipe') return m.coverThumb ? `<img class="thumb" src="${esc(m.coverThumb)}" alt="">` : '<div class="thumb-icon">📷</div>';
    return m.youtubeId ? `<img class="thumb" src="https://i.ytimg.com/vi/${esc(m.youtubeId)}/mqdefault.jpg" alt="">` : '<div class="thumb-icon">🎬</div>';
  }

  window.FM = { batch, thumbHtml, image, loadPhotos, resizeImage, scaleAmount, api, beacon, store, fmtTime, parseTime, esc, qs, uid, toast, loadYouTube, indexAt, ready, lineLogin, auth, tagsHtml, parseTags, LIFF_ID };
})();
