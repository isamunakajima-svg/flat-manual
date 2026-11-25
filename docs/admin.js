/* 管理画面の共通処理：パスワードと、動画のアップロード */
(function () {
  function getKey() { return FM.store.get('adminKey', ''); }

  async function adminApi(action, params) {
    return FM.api(action, Object.assign({ key: getKey() }, params || {}));
  }

  // パスワードが無い・違うときはログイン欄を出す
  async function requireLogin(onReady) {
    const box = document.getElementById('loginBox');
    const main = document.getElementById('adminMain');
    const tryIt = async () => {
      try {
        await adminApi('ping');
        box.classList.add('hidden');
        main.classList.remove('hidden');
        onReady();
      } catch (e) {
        box.classList.remove('hidden');
        main.classList.add('hidden');
        if (getKey()) FM.toast(e.message, true);
      }
    };
    document.getElementById('loginBtn').onclick = () => {
      FM.store.set('adminKey', document.getElementById('loginKey').value);
      tryIt();
    };
    if (getKey()) tryIt(); else box.classList.remove('hidden');
  }

  function toBase64(blob) {
    return new Promise((resolve, reject) => {
      const r = new FileReader();
      r.onload = () => resolve(String(r.result).split(',')[1] || '');
      r.onerror = () => reject(r.error);
      r.readAsDataURL(blob);
    });
  }

  // Geminiが受け付ける形式名にそろえる
  function geminiMime(file) {
    const t = (file.type || '').toLowerCase();
    if (t === 'video/quicktime' || /\.mov$/i.test(file.name || '')) return 'video/mov';
    if (t) return t;
    return 'video/mp4';
  }

  /**
   * ファイルを分割してGASに送り、GASからGeminiに渡す。
   * onProgress(0〜1, メッセージ)
   */
  async function uploadToGemini(file, label, onProgress, opts) {
    opts = opts || {};
    const mimeType = file.mimeType || geminiMime(file);
    const start = await adminApi('uploadStart', { size: file.size, mimeType, name: label });
    const gran = start.granularity || 8388608;
    const chunk = Math.max(gran, Math.floor(8388608 / gran) * gran);
    let offset = 0, info = null;
    while (offset < file.size) {
      const end = Math.min(file.size, offset + chunk);
      const last = end >= file.size;
      const data = await toBase64(file.slice(offset, end));
      onProgress && onProgress(offset / file.size, `${label}を送っています… ${Math.round(offset / file.size * 100)}%`);
      const r = await adminApi('uploadChunk', { uploadUrl: start.uploadUrl, offset, last, data });
      if (last) info = r;
      offset = end;
    }
    if (opts.noWait) return { name: info.name, uri: info.uri, mimeType: info.mimeType || mimeType }; // 読み込み待ちはGAS側でする
    onProgress && onProgress(1, `${label}をAIが読み込んでいます…`);
    for (let i = 0; i < 90; i++) {
      if (info.state === 'ACTIVE') return { uri: info.uri, mimeType: info.mimeType || mimeType };
      if (info.state === 'FAILED') throw new Error(label + 'の読み込みに失敗しました（形式が合わない可能性があります）');
      await new Promise(r => setTimeout(r, 3000));
      const s = await adminApi('fileStatus', { name: info.name });
      info = Object.assign(info, s);
    }
    throw new Error(label + 'の読み込みが終わりませんでした');
  }

  /* --- 録音した音声を、どのブラウザでも読めるWAV（16kHz・モノラル）にする --- */
  async function toWav(blob) {
    const buf = await blob.arrayBuffer();
    const Ctx = window.AudioContext || window.webkitAudioContext;
    const ctx = new Ctx();
    const decoded = await new Promise((res, rej) => ctx.decodeAudioData(buf, res, rej));
    const rate = 16000;
    const off = new OfflineAudioContext(1, Math.ceil(decoded.duration * rate), rate);
    const src = off.createBufferSource();
    src.buffer = decoded;
    src.connect(off.destination);
    src.start();
    const rendered = await off.startRendering();
    try { ctx.close(); } catch (e) {}
    const pcm = rendered.getChannelData(0);
    const out = new DataView(new ArrayBuffer(44 + pcm.length * 2));
    const w = (o, s) => { for (let i = 0; i < s.length; i++) out.setUint8(o + i, s.charCodeAt(i)); };
    w(0, 'RIFF'); out.setUint32(4, 36 + pcm.length * 2, true); w(8, 'WAVE');
    w(12, 'fmt '); out.setUint32(16, 16, true); out.setUint16(20, 1, true); out.setUint16(22, 1, true);
    out.setUint32(24, rate, true); out.setUint32(28, rate * 2, true); out.setUint16(32, 2, true); out.setUint16(34, 16, true);
    w(36, 'data'); out.setUint32(40, pcm.length * 2, true);
    for (let i = 0; i < pcm.length; i++) {
      const v = Math.max(-1, Math.min(1, pcm[i]));
      out.setInt16(44 + i * 2, v < 0 ? v * 0x8000 : v * 0x7fff, true);
    }
    const wav = new Blob([out], { type: 'audio/wav' });
    wav.mimeType = 'audio/wav';
    return wav;
  }

  window.FMA = { getKey, adminApi, requireLogin, uploadToGemini, toWav };
})();
