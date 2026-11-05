// ==UserScript==
// @name         우신 인트라넷 - workline 바로 가져오기
// @namespace    https://wooriwoosin.github.io/woosin/
// @version      1.0.0
// @description  인트라넷 실적집계에서 [⚡ 바로 가져오기] 버튼 하나로 workline 로우데이터(접수일·개통일·상품권·인증)를 가져옵니다.
// @match        https://wooriwoosin.github.io/woosin/*
// @grant        GM_xmlhttpRequest
// @connect      workline.kr
// @run-at       document-idle
// @updateURL    https://wooriwoosin.github.io/woosin/workline-import.user.js
// @downloadURL  https://wooriwoosin.github.io/woosin/workline-import.user.js
// ==/UserScript==

// 인트라넷 화면(github.io)은 브라우저 보안상 workline 데이터를 직접 못 읽는다.
// 이 스크립트는 탬퍼몽키 권한(GM_xmlhttpRequest)으로 workline.kr에 요청을 보내고(이 브라우저의
// workline 로그인 쿠키 사용), 받은 엑셀 데이터를 인트라넷 화면에 넘겨준다.
// 무엇을 받을지(주소 목록)는 인트라넷이 알려주며, workline.kr 의 /api/ 경로만 허용한다.
(function () {
  'use strict';
  var VERSION = '1.0.0';
  var WL = 'https://workline.kr';
  var ORIGIN = location.origin;

  function post(msg) { msg.__wlx = true; msg.from = 'us'; window.postMessage(msg, ORIGIN); }
  function safe(u) {
    if (typeof u !== 'string' || !/^\/api\/[A-Za-z0-9_\-\/]+(\?[A-Za-z0-9_\-=&.%]*)?$/.test(u)) throw new Error('허용되지 않은 주소');
    return WL + u;
  }
  function sleep(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }
  function req(path, type) {
    return new Promise(function (resolve, reject) {
      GM_xmlhttpRequest({
        method: 'GET', url: safe(path), responseType: type, timeout: 120000, anonymous: false,
        headers: { 'Accept': type === 'json' ? 'application/json' : '*/*' },
        onload: function (r) { resolve(r); },
        onerror: function () { reject(new Error('workline 연결 실패')); },
        ontimeout: function () { reject(new Error('workline 응답 시간 초과')); }
      });
    });
  }
  async function getJson(path) {
    var r = await req(path, 'text');
    if (r.status === 401 || r.status === 403) throw new Error('workline 로그인이 필요해요');
    var j = null; try { j = JSON.parse(r.responseText); } catch (e) {}
    if (!j) {
      // 로그인 페이지 등 HTML이 오면 로그인 풀린 것
      if (/<html/i.test(String(r.responseText || ''))) throw new Error('workline 로그인이 필요해요');
      throw new Error('응답 형식 오류 (' + r.status + ')');
    }
    if ((r.status < 200 || r.status >= 300) && r.status !== 202) throw new Error((j.error && (j.error.message || j.error)) || ('오류 ' + r.status));
    return j;
  }
  async function runJob(j) {
    var s = await getJson(j.start);
    var job = (s && (s.job || s.data)) || {};
    var id = (job.jobId != null ? job.jobId : job.id);
    if (id == null) throw new Error('엑셀 작업번호를 못 받았어요');
    for (var i = 0; ; i++) {
      var st = String(job.status || '');
      if (job.error || /fail|error|cancel/i.test(st)) throw new Error('엑셀 생성 실패: ' + (job.error || st));
      if (job.downloadable === true || /^(done|complete|completed|success|succeeded|finished)$/i.test(st)) break;
      if (i >= 200) throw new Error('엑셀 생성 시간 초과');
      post({ type: 'wl-progress', kind: j.kind, msg: j.label + ' 엑셀 만드는 중' + (job.percent != null ? ' ' + job.percent + '%' : '') + (job.queueAhead ? ' (대기 ' + job.queueAhead + '건)' : '') });
      await sleep(1500);
      var p = await getJson(j.poll.replace('{id}', encodeURIComponent(id)));
      job = (p && (p.job || p.data)) || job;
    }
    var r = await req(j.download.replace('{id}', encodeURIComponent(id)), 'arraybuffer');
    if (r.status === 401 || r.status === 403) throw new Error('workline 로그인이 필요해요');
    if (r.status < 200 || r.status >= 300) throw new Error('파일 받기 실패 (' + r.status + ')');
    return r.response;
  }
  var busy = false;
  async function run(m) {
    if (busy) { post({ type: 'wl-error', kind: '', msg: '이미 가져오는 중이에요' }); return; }
    busy = true;
    var jobs = m.jobs || [], ok = 0;
    for (var k = 0; k < jobs.length; k++) {
      var j = jobs[k];
      post({ type: 'wl-progress', kind: j.kind, msg: j.label + ' 요청 중...' });
      try {
        var buf = await runJob(j);
        post({ type: 'wl-file', date: m.date, kind: j.kind, buf: buf });
        ok++;
      } catch (e) {
        var msg = (e && e.message) || String(e), login = /로그인/.test(msg);
        post({ type: 'wl-error', kind: j.kind, msg: j.label + ': ' + msg, login: login });
        if (login) { post({ type: 'wl-done', date: m.date, ok: ok, total: jobs.length, login: true }); busy = false; return; } // 로그인 풀림 → 나머지도 안 되므로 중단
      }
    }
    post({ type: 'wl-done', date: m.date, ok: ok, total: jobs.length });
    busy = false;
  }
  window.addEventListener('message', function (ev) {
    if (ev.origin !== ORIGIN) return;
    var m = ev.data;
    if (!m || typeof m !== 'object' || !m.__wlx || m.from !== 'page') return;
    if (m.type === 'wlx-ping') post({ type: 'wlx-ready', version: VERSION });
    else if (m.type === 'wlx-run') run(m);
  });
  post({ type: 'wlx-ready', version: VERSION });
})();
