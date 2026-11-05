// 인트라넷 화면 ↔ 확장 프로그램(background) 중계.
// 화면과는 window.postMessage({__wlx:true, from:'page'|'us', ...}) 로, background와는 포트로 주고받는다.
(function () {
  'use strict';
  var VERSION = chrome.runtime.getManifest().version;
  var ORIGIN = location.origin;
  function post(msg) { msg.__wlx = true; msg.from = 'us'; msg.via = 'ext'; window.postMessage(msg, ORIGIN); }
  function b64ToBuf(b64) { var bin = atob(b64), u = new Uint8Array(bin.length); for (var i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i); return u.buffer; }
  window.addEventListener('message', function (ev) {
    if (ev.origin !== ORIGIN || ev.source !== window) return;
    var m = ev.data;
    if (!m || typeof m !== 'object' || !m.__wlx || m.from !== 'page') return;
    if (m.type === 'wlx-ping') { post({ type: 'wlx-ready', version: VERSION }); return; }
    if (m.type !== 'wlx-run') return;
    var port = chrome.runtime.connect({ name: 'wlx' });
    port.onMessage.addListener(function (r) {
      if (r.type === 'wl-file') { r.buf = b64ToBuf(r.b64); delete r.b64; }
      post(r);
      if (r.type === 'wl-done') port.disconnect();
    });
    port.onDisconnect.addListener(function () { /* 끝 */ });
    port.postMessage({ type: 'run', date: m.date, jobs: m.jobs });
  });
  post({ type: 'wlx-ready', version: VERSION });
})();
