// workline.kr 요청 담당 (확장 프로그램 권한으로 이 브라우저의 workline 로그인 쿠키 사용).
// 인트라넷이 넘겨준 작업 목록(엑셀 내보내기 → 진행 확인 → 다운로드)을 차례로 실행하고 파일을 돌려준다.
// workline.kr 의 /api/ 경로만 허용한다.
const WL = 'https://workline.kr';
const INTRANET = 'https://wooriwoosin.github.io/woosin/';

function safe(u) {
  if (typeof u !== 'string' || !/^\/api\/[A-Za-z0-9_\-\/]+(\?[A-Za-z0-9_\-=&.%]*)?$/.test(u)) throw new Error('허용되지 않은 주소');
  return WL + u;
}
const sleep = ms => new Promise(r => setTimeout(r, ms));
function bufToB64(buf) {
  const bytes = new Uint8Array(buf); let s = '';
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  return btoa(s);
}
async function getJson(path) {
  const r = await fetch(safe(path), { credentials: 'include', headers: { 'Accept': 'application/json' } });
  if (r.status === 401 || r.status === 403) throw new Error('workline 로그인이 필요해요');
  const t = await r.text(); let j = null; try { j = JSON.parse(t); } catch (e) {}
  if (!j) { if (/<html/i.test(t)) throw new Error('workline 로그인이 필요해요'); throw new Error('응답 형식 오류 (' + r.status + ')'); }
  if (!r.ok && r.status !== 202) throw new Error((j.error && (j.error.message || j.error)) || ('오류 ' + r.status));
  return j;
}
async function runJob(j, send) {
  const s = await getJson(j.start);
  let job = (s && (s.job || s.data)) || {};
  const id = (job.jobId != null ? job.jobId : job.id);
  if (id == null) throw new Error('엑셀 작업번호를 못 받았어요');
  for (let i = 0; ; i++) {
    const st = String(job.status || '');
    if (job.error || /fail|error|cancel/i.test(st)) throw new Error('엑셀 생성 실패: ' + (job.error || st));
    if (job.downloadable === true || /^(done|complete|completed|success|succeeded|finished)$/i.test(st)) break;
    if (i >= 200) throw new Error('엑셀 생성 시간 초과');
    send({ type: 'wl-progress', kind: j.kind, msg: j.label + ' 엑셀 만드는 중' + (job.percent != null ? ' ' + job.percent + '%' : '') + (job.queueAhead ? ' (대기 ' + job.queueAhead + '건)' : '') });
    await sleep(1500);
    const p = await getJson(j.poll.replace('{id}', encodeURIComponent(id)));
    job = (p && (p.job || p.data)) || job;
  }
  const r = await fetch(safe(j.download.replace('{id}', encodeURIComponent(id))), { credentials: 'include' });
  if (r.status === 401 || r.status === 403) throw new Error('workline 로그인이 필요해요');
  if (!r.ok) throw new Error('파일 받기 실패 (' + r.status + ')');
  return r.arrayBuffer();
}
chrome.runtime.onConnect.addListener(port => {
  if (port.name !== 'wlx') return;
  const from = (port.sender && (port.sender.url || '')) || '';
  if (from.indexOf(INTRANET) !== 0) { port.disconnect(); return; } // 인트라넷 화면에서 온 요청만
  let alive = true; port.onDisconnect.addListener(() => { alive = false; });
  const send = m => { if (alive) try { port.postMessage(m); } catch (e) {} };
  port.onMessage.addListener(async m => {
    if (!m || m.type !== 'run') return;
    const jobs = m.jobs || []; let ok = 0;
    for (const j of jobs) {
      send({ type: 'wl-progress', kind: j.kind, msg: j.label + ' 요청 중...' });
      try {
        const buf = await runJob(j, send);
        send({ type: 'wl-file', date: m.date, kind: j.kind, b64: bufToB64(buf) });
        ok++;
      } catch (e) {
        const msg = (e && e.message) || String(e), login = /로그인/.test(msg);
        send({ type: 'wl-error', kind: j.kind, msg: j.label + ': ' + msg, login });
        if (login) { send({ type: 'wl-done', date: m.date, ok, total: jobs.length, login: true }); return; }
      }
    }
    send({ type: 'wl-done', date: m.date, ok, total: jobs.length });
  });
});
