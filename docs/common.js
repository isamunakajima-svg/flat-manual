/* 共通の小さな道具 */
(function () {
  const GAS_URL = (window.FLAT_MANUAL_CONFIG || {}).GAS_URL;

  // GASへの問い合わせ。Content-Type を付けない（text/plain）ので事前確認が要らない
  async function api(action, params) {
    const body = JSON.stringify(Object.assign({ action: action }, params || {}));
    const res = await fetch(GAS_URL, { method: 'POST', body: body, redirect: 'follow' });
    if (!res.ok) throw new Error('通信エラー (' + res.status + ')');
    const json = await res.json();
    if (!json.ok) throw new Error(json.error || 'エラー');
    return json.data;
  }

  // ページを閉じるときの記録用（返事は待たない）
  function beacon(action, params) {
    const body = JSON.stringify(Object.assign({ action: action }, params || {}));
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

  window.FM = { api, beacon, store, fmtTime, parseTime, esc, qs, uid, toast, loadYouTube, indexAt };
})();
