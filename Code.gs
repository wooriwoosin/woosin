// ============================================================
//  업무관리 시스템 — Google Apps Script Web App
//  배포: 웹 앱으로 배포 > 액세스 권한: 모든 사용자
// ============================================================

var SS_ID = '12vPbzL9mZn1gSEXDCdP0cDE0zYSQQLDcx5hTO6OgFod72Lq5SRYp2ddM'; // ← 배포 후 스프레드시트 ID 입력 (URL의 /d/XXXX/edit 부분)
var SS = null;

function getSheet(name) {
  if (!SS) SS = SpreadsheetApp.openById(SS_ID);
  var sh = SS.getSheetByName(name);
  if (!sh) { sh = SS.insertSheet(name); initSheet(sh, name); }
  return sh;
}

// ── 시트 초기화 (헤더 행 설정) ───────────────────────────
function initSheet(sh, name) {
  var headers = {
    '직원':      ['id','pw','name','role','color','hire','phone','team','rank','duty',
                  'cid1','cpw1','idms1','cid2','cpw2','idms2','bph','doc','s1','pcpw','pcpwDate'],
    '실적':      ['id','empId','empName','date','attendStatus','rows'],
    '연차':      ['id','empId','empName','date','type'],
    '연차공지':  ['notice'],
    'IP그룹':    ['id','type','name','prefix','range','subnet','gateway','dns','dns2','entries'],
    '변경로그':  ['dt','by','desc'],
    '자리배치':  ['id','label','empId'],
    '설정':      ['key','value']
  };
  if (headers[name]) {
    sh.getRange(1, 1, 1, headers[name].length).setValues([headers[name]]);
    sh.getRange(1, 1, 1, headers[name].length).setFontWeight('bold').setBackground('#E8EAED');
  }
}

// ── HTTP 진입점 ──────────────────────────────────────────
function doPost(e) {
  return handleRequest(e);
}
function doGet(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  var result;
  try {
    var action, data;
    // GET 파라미터 방식 (프론트엔드에서 URLSearchParams로 전달)
    if (e.parameter && e.parameter.action) {
      action = e.parameter.action;
      data = e.parameter.data ? JSON.parse(e.parameter.data) : {};
    }
    // POST body 방식 (fallback)
    else if (e.postData && e.postData.contents) {
      var body = JSON.parse(e.postData.contents);
      action = body.action;
      data = body;
    }
    else {
      result = { ok: false, error: '파라미터 없음' };
      return output(result);
    }
    result = route(action, data);
  } catch(err) {
    result = { ok: false, error: err.message + ' / ' + err.stack };
  }
  return output(result);
}

function output(result) {
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── 라우터 ────────────────────────────────────────────────
function route(action, data) {
  switch(action) {
    // 인증
    case 'ping':           return { ok: true };
    case 'login':          return login(data);
    // 직원
    case 'getEmployees':   return getEmployees(false);
    case 'getAllUsers':     return getEmployees(true);
    case 'getUser':        return getUser(data.id);
    case 'addEmployee':    return addEmployee(data.emp);
    case 'deleteEmployee': return deleteEmployee(data.id);
    case 'updateEmpAdmin': return updateEmpAdmin(data.id, data.fields);
    case 'updateEmpSelf':  return updateEmpSelf(data.id, data.fields);
    // 실적
    case 'savePerf':       return savePerf(data.r);
    case 'getPerfs':       return getPerfs(data);
    case 'deletePerf':     return deletePerf(data.empId, data.date);
    // 연차
    case 'saveAnnual':     return saveAnnual(data.a);
    case 'getAnnuals':     return getAnnuals(data.ym);
    // IP
    case 'getIPGroups':    return getIPGroups();
    case 'saveIPGroup':    return saveIPGroup(data.group);
    case 'deleteIPGroup':  return deleteIPGroup(data.id);
    // 변경로그
    case 'addChangeLog':   return addChangeLog(data);
    case 'getChangeLogs':  return getChangeLogs(data.days);
    // 연차공지
    case 'getCalNotice':   return getCalNotice();
    case 'saveCalNotice':  return saveCalNotice(data.notice);
    // 자리배치
    case 'getSeats':       return getSeats();
    case 'saveAllSeats':   return saveAllSeats(data.seats);
    // 설정
    case 'getSettings':    return getSettings();
    case 'saveSettings':   return saveSettings(data);
    default: return { ok: false, error: '알 수 없는 액션: ' + action };
  }
}

// ════════════════════════════════════════════════════════════
//  인증
// ════════════════════════════════════════════════════════════
function login(data) {
  var sh = getSheet('직원');
  var rows = sheetToObjects(sh);
  var u = rows.find(function(r){ return r.id === data.id; });
  if (!u) return { ok: false, error: '존재하지 않는 아이디입니다.' };
  if (u.pw && u.pw !== data.pw) return { ok: false, error: '비밀번호가 틀렸습니다.' };
  var safe = Object.assign({}, u); delete safe.pw;
  return { ok: true, user: safe };
}

// ════════════════════════════════════════════════════════════
//  직원
// ════════════════════════════════════════════════════════════
function getEmployees(includeAdmin) {
  var sh = getSheet('직원');
  var rows = sheetToObjects(sh);
  var list = includeAdmin ? rows : rows.filter(function(r){ return r.role !== 'admin'; });
  list = list.map(function(u){ var s=Object.assign({},u); delete s.pw; return s; });
  return { ok: true, employees: list };
}

function getUser(id) {
  var sh = getSheet('직원');
  var rows = sheetToObjects(sh);
  var u = rows.find(function(r){ return r.id === id; });
  if (!u) return { ok: false, error: '직원을 찾을 수 없습니다.' };
  return { ok: true, user: u };
}

function addEmployee(emp) {
  var sh = getSheet('직원');
  var rows = sheetToObjects(sh);
  if (rows.find(function(r){ return r.id === emp.id; }))
    return { ok: false, error: '이미 존재하는 아이디입니다.' };
  var headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var row = headers.map(function(h){ return emp[h] !== undefined ? emp[h] : ''; });
  sh.appendRow(row);
  return { ok: true };
}

function deleteEmployee(id) {
  var sh = getSheet('직원');
  deleteRowById(sh, 'id', id);
  return { ok: true };
}

function updateEmpAdmin(id, fields) {
  var sh = getSheet('직원');
  var headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var data = sh.getDataRange().getValues();
  var idCol = headers.indexOf('id');
  for (var i = 1; i < data.length; i++) {
    if (data[i][idCol] == id) {
      headers.forEach(function(h, ci) {
        if (h === 'id' && fields.newId) { data[i][ci] = fields.newId; return; }
        if (fields[h] !== undefined) data[i][ci] = fields[h];
      });
      sh.getDataRange().setValues(data);
      return { ok: true };
    }
  }
  return { ok: false, error: '직원을 찾을 수 없습니다.' };
}

function updateEmpSelf(id, fields) {
  // 본인이 수정 가능한 필드만 허용
  var allowed = ['cid1','cpw1','idms1','cid2','cpw2','idms2','bph','doc','s1','pcpw','pcpwDate'];
  var safe = {};
  allowed.forEach(function(k){ if (fields[k] !== undefined) safe[k] = fields[k]; });
  return updateEmpAdmin(id, safe);
}

// ════════════════════════════════════════════════════════════
//  실적
// ════════════════════════════════════════════════════════════
function savePerf(r) {
  var sh = getSheet('실적');
  // 같은 empId + date 기존 행 삭제
  deleteRowWhere(sh, function(row, headers) {
    var ei = headers.indexOf('empId'), di = headers.indexOf('date');
    return row[ei] == r.empId && row[di] == r.date;
  });
  var id = 'p_' + r.empId + '_' + r.date;
  sh.appendRow([
    id, r.empId, r.empName, r.date, r.attendStatus || '',
    JSON.stringify(r.rows || [])
  ]);
  return { ok: true };
}

function getPerfs(data) {
  var sh = getSheet('실적');
  var rows = sheetToObjects(sh);
  if (data && data.empId) rows = rows.filter(function(r){ return r.empId === data.empId; });
  if (data && data.ym)    rows = rows.filter(function(r){ return r.date && r.date.indexOf(data.ym) === 0; });
  rows = rows.map(function(r){
    try { r.rows = typeof r.rows === 'string' ? JSON.parse(r.rows) : (r.rows || []); } catch(e){ r.rows=[]; }
    return r;
  });
  return { ok: true, perfs: rows };
}

function deletePerf(empId, date) {
  var sh = getSheet('실적');
  deleteRowWhere(sh, function(row, headers) {
    var ei = headers.indexOf('empId'), di = headers.indexOf('date');
    return row[ei] == empId && row[di] == date;
  });
  return { ok: true };
}

// ════════════════════════════════════════════════════════════
//  연차
// ════════════════════════════════════════════════════════════
function saveAnnual(a) {
  var sh = getSheet('연차');
  // 기존 같은 empId+date 삭제
  deleteRowWhere(sh, function(row, headers) {
    var ei = headers.indexOf('empId'), di = headers.indexOf('date');
    return row[ei] == a.empId && row[di] == a.date;
  });
  if (a.type) { // type 있을 때만 저장 (빈 type = 삭제)
    var id = 'a_' + a.empId + '_' + a.date;
    sh.appendRow([id, a.empId, a.empName || '', a.date, a.type]);
  }
  return { ok: true };
}

function getAnnuals(ym) {
  var sh = getSheet('연차');
  var rows = sheetToObjects(sh);
  if (ym) rows = rows.filter(function(r){ return r.date && r.date.indexOf(ym) === 0; });
  return { ok: true, annuals: rows };
}

// ════════════════════════════════════════════════════════════
//  IP 그룹
// ════════════════════════════════════════════════════════════
function getIPGroups() {
  var sh = getSheet('IP그룹');
  var rows = sheetToObjects(sh);
  rows = rows.map(function(r){
    try { r.entries = typeof r.entries === 'string' ? JSON.parse(r.entries) : (r.entries || []); } catch(e){ r.entries=[]; }
    return r;
  });
  return { ok: true, groups: rows };
}

function saveIPGroup(group) {
  var sh = getSheet('IP그룹');
  if (group.id) deleteRowById(sh, 'id', group.id);
  var id = group.id || 'g_' + Date.now();
  sh.appendRow([
    id, group.type || '', group.name || '', group.prefix || '',
    group.range || '', group.subnet || '', group.gateway || '',
    group.dns || '', group.dns2 || '',
    JSON.stringify(group.entries || [])
  ]);
  return { ok: true };
}

function deleteIPGroup(id) {
  var sh = getSheet('IP그룹');
  deleteRowById(sh, 'id', id);
  return { ok: true };
}

// ════════════════════════════════════════════════════════════
//  변경 로그
// ════════════════════════════════════════════════════════════
function addChangeLog(data) {
  var sh = getSheet('변경로그');
  sh.appendRow([data.dt, data.by, data.desc]);
  return { ok: true };
}

function getChangeLogs(days) {
  var sh = getSheet('변경로그');
  var rows = sheetToObjects(sh);
  rows = rows.reverse(); // 최신순
  if (days) {
    var cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - days);
    var cutStr = Utilities.formatDate(cutoff, 'Asia/Seoul', 'yyyy-MM-dd');
    rows = rows.filter(function(r){ return r.dt && r.dt.slice(0,10) >= cutStr; });
  }
  return { ok: true, logs: rows };
}

// ════════════════════════════════════════════════════════════
//  연차 공지
// ════════════════════════════════════════════════════════════
function getCalNotice() {
  var sh = getSheet('연차공지');
  var data = sh.getDataRange().getValues();
  // 헤더 제외하고 2행부터 읽기
  var notice = '';
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) notice = data[i][0];
  }
  return { ok: true, notice: notice };
}

function saveCalNotice(notice) {
  var sh = getSheet('연차공지');
  // 2행에 저장 (1행은 헤더)
  if (sh.getLastRow() < 2) sh.appendRow([notice]);
  else sh.getRange(2, 1).setValue(notice);
  return { ok: true };
}

// ════════════════════════════════════════════════════════════
//  자리 배치
// ════════════════════════════════════════════════════════════
var DEFAULT_SEATS = [
  {id:'s2-L1',label:'2층 좌①',empId:''},
  {id:'s2-L2',label:'2층 좌②',empId:''},
  {id:'s2-R1',label:'2층 우①',empId:''},
  {id:'s2-R2',label:'2층 우②',empId:''},
  {id:'s2-R3',label:'2층 우③',empId:''},
  {id:'s1-L1',label:'1층 좌①',empId:''},
  {id:'s1-L2',label:'1층 좌②',empId:''},
  {id:'s1-L3',label:'1층 좌③',empId:''},
  {id:'s1-L4',label:'1층 좌④',empId:''},
  {id:'s1-L5',label:'1층 좌⑤',empId:''},
  {id:'s1-R1',label:'1층 우①',empId:''},
  {id:'s1-R2',label:'1층 우②',empId:''},
  {id:'s1-R3',label:'1층 우③',empId:''},
  {id:'s1-R4',label:'1층 우④',empId:''},
  {id:'s1-R5',label:'1층 우⑤',empId:''},
  {id:'s1-SOLO',label:'1층 단독',empId:''}
];

function getSeats() {
  var sh = getSheet('자리배치');
  var rows = sheetToObjects(sh);
  if (!rows.length) {
    // 최초 실행 시 기본값 입력
    DEFAULT_SEATS.forEach(function(s){
      sh.appendRow([s.id, s.label, s.empId]);
    });
    rows = DEFAULT_SEATS;
  }
  // object 형태로 변환 {id: {label, empId}}
  var result = {};
  rows.forEach(function(r){ result[r.id] = { label: r.label, empId: r.empId || '' }; });
  return { ok: true, seats: result };
}

function saveAllSeats(seats) {
  var sh = getSheet('자리배치');
  var headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var data = sh.getDataRange().getValues();
  var idCol = headers.indexOf('id');
  var empCol = headers.indexOf('empId');
  for (var i = 1; i < data.length; i++) {
    var sid = data[i][idCol];
    if (seats[sid] !== undefined) data[i][empCol] = seats[sid];
  }
  sh.getDataRange().setValues(data);
  return { ok: true };
}

// ════════════════════════════════════════════════════════════
//  설정
// ════════════════════════════════════════════════════════════
function getSettings() {
  var sh = getSheet('설정');
  var rows = sheetToObjects(sh);
  var result = {};
  rows.forEach(function(r){ result[r.key] = r.value; });
  return { ok: true, settings: result };
}

function saveSettings(data) {
  var sh = getSheet('설정');
  var headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var existing = sh.getDataRange().getValues();
  var keyCol = headers.indexOf('key');
  Object.keys(data).forEach(function(k) {
    var found = false;
    for (var i = 1; i < existing.length; i++) {
      if (existing[i][keyCol] == k) {
        sh.getRange(i+1, 2).setValue(data[k]);
        found = true; break;
      }
    }
    if (!found) sh.appendRow([k, data[k]]);
  });
  return { ok: true };
}

// ════════════════════════════════════════════════════════════
//  유틸 함수
// ════════════════════════════════════════════════════════════

/** 시트 전체를 객체 배열로 변환 */
function sheetToObjects(sh) {
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0];
  return data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i){ obj[h] = row[i] === '' ? '' : row[i]; });
    return obj;
  });
}

/** id 컬럼으로 행 삭제 */
function deleteRowById(sh, idCol, idVal) {
  var data = sh.getDataRange().getValues();
  var headers = data[0];
  var ci = headers.indexOf(idCol);
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][ci] == idVal) { sh.deleteRow(i + 1); }
  }
}

/** 조건 함수로 행 삭제 */
function deleteRowWhere(sh, condFn) {
  var data = sh.getDataRange().getValues();
  var headers = data[0];
  for (var i = data.length - 1; i >= 1; i--) {
    if (condFn(data[i], headers)) sh.deleteRow(i + 1);
  }
}

// ── 전체 시트 초기화 (최초 1회 실행) ─────────────────────
function setupAllSheets() {
  var sheetNames = ['직원','실적','연차','연차공지','IP그룹','변경로그','자리배치','설정'];
  sheetNames.forEach(function(name) {
    var sh = getSheet(name);
    Logger.log('시트 준비 완료: ' + name);
  });
  // 설정 기본값
  saveSettings({
    ACHIEVE_TARGET_DAY: 120,
    WIRELESS_TARGET_DAY: 20
  });
  Logger.log('초기화 완료!');
}
