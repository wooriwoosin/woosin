// ============================================================
//  업무관리 시스템 — Google Apps Script Web App
//  배포: 웹 앱으로 배포 > 액세스 권한: 모든 사용자
// ============================================================

var SS_ID = '12b0lR6rHHAVprUBbhOju6XZadHNI8Oiw9qb8aAmVsnE'; // ← 스프레드시트 ID
var SS = null;

// ── 스킬 항목 기본값 (설정 시트에 skillCatalog 키가 없을 때 최초 1회 자동 생성) ──
var DEFAULT_SKILL_CATALOG = [
  { cat: '인증 및 간편신청서', items: ['신규', '약정갱신', '상품권/청구매체(SO처리포함)', '간편신청서 작성'] },
  { cat: '유선 기본접수', items: ['인터넷/TV', 'OSS', '(신규) 일반전화, 인터넷전화', '외국인', '사업자 (개인사업자, 법인사업자, 고유번호사업자)'] },
  { cat: '유선 심화접수', items: ['(번호이동) 일반전화, 인터넷전화', '센트릭스', '기가아이즈', '약정갱신'] },
  { cat: '무선관련', items: ['유심개통', 'SO처리 (개통,선약,결합 등)'] },
  { cat: '기본SO처리', items: ['취소처리', '청구매체/자동이체 변경', '일정변경', '결합변경'] },
  { cat: '그외', items: ['인입전화응대 (고객,거래처,기사님,설치점,고객센터,그외 등등)', 'OSS확인', '웹회신,카톡,웍스,채널 응대', 'SRTT(VOC)처리', '이미징처리', '신입교육', '정책관련 업무', '정산관련 업무', '개통확인', '검수(일일검수,모바일,패밀리 등 검수)'] }
];

/** 스킬 카탈로그를 위 DEFAULT_SKILL_CATALOG 내용으로 강제 초기화 (설정 탭에서 수정한 내역이 있어도 덮어씀)
 *  Apps Script 편집기에서 이 함수를 한 번 실행하면 즉시 반영됩니다. */
function resetSkillCatalogToDefault() {
  saveSettings({ skillCatalog: JSON.stringify(DEFAULT_SKILL_CATALOG) });
  Logger.log('스킬 카탈로그가 기본값으로 초기화되었습니다.');
}

function getSheet(name) {
  if (!SS) SS = SpreadsheetApp.openById(SS_ID);
  var sh = SS.getSheetByName(name);
  if (!sh) { sh = SS.insertSheet(name); initSheet(sh, name); }
  if (name === '직원') { ensureColumns(sh, ['skills', 'ext', 'subUserName','wUserName','wcid','wcpw','widms','katalkId','katalkPw','retiredAt','acctDisposed','teamCardAdmin','superAdmin']); seedTeamCardAdmins(sh); seedSuperAdmin(sh); } // 기존 시트에 컬럼 자동 추가 (마이그레이션)
  if (name === '실적집계') { renameColumn(sh, '신규접수', '유선신규'); ensureColumns(sh, ['기가아이즈', '고정업무시간', '약정갱신_인업셀', '약정갱신_인약갱', '약정갱신_티약갱', '약정갱신_티업셀', '상품권']); }
  return sh;
}

/** 팀카드관리 권한 초기 부여 - 이미 명시적으로 값이 설정된 사람은 건드리지 않고, 한 번도 설정 안 된(빈칸) 경우에만 지정 이름 3명에게 Y를 부여 */
function seedTeamCardAdmins(sh) {
  var TEAMCARD_SEED_NAMES = ['김해성', '김민정', '김석환'];
  var headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var nameCol = headers.indexOf('name'), tcaCol = headers.indexOf('teamCardAdmin');
  if (nameCol === -1 || tcaCol === -1) return;
  var data = sh.getDataRange().getValues();
  var changed = false;
  for (var i = 1; i < data.length; i++) {
    var nm = (data[i][nameCol] || '').toString();
    if (TEAMCARD_SEED_NAMES.indexOf(nm) !== -1 && !data[i][tcaCol]) {
      sh.getRange(i + 1, tcaCol + 1).setValue('Y');
      changed = true;
    }
  }
}

/** 대표관리자(슈퍼관리자) 초기 부여 - 김민정. 이미 명시적으로 값이 설정된 사람은 건드리지 않음 */
function seedSuperAdmin(sh) {
  var SUPER_ADMIN_SEED_NAMES = ['김민정'];
  var headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var nameCol = headers.indexOf('name'), saCol = headers.indexOf('superAdmin');
  if (nameCol === -1 || saCol === -1) return;
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var nm = (data[i][nameCol] || '').toString();
    if (SUPER_ADMIN_SEED_NAMES.indexOf(nm) !== -1 && !data[i][saCol]) {
      sh.getRange(i + 1, saCol + 1).setValue('Y');
    }
  }
}

/** 시트에 특정 헤더 컬럼들이 없으면 맨 뒤에 추가 (기존 데이터는 보존) */
function ensureColumns(sh, cols) {
  var lastCol = sh.getLastColumn();
  var headers = lastCol > 0 ? sh.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  var toAdd = cols.filter(function(c){ return headers.indexOf(c) === -1; });
  if (toAdd.length) {
    sh.getRange(1, lastCol + 1, 1, toAdd.length).setValues([toAdd]);
    sh.getRange(1, lastCol + 1, 1, toAdd.length).setFontWeight('bold').setBackground('#E8EAED');
  }
}

/** 기존 시트의 헤더 이름이 바뀐 경우(예: 신규접수→유선신규) 셀 값을 그대로 갈아치움 (데이터는 그대로 유지) */
function renameColumn(sh, oldName, newName) {
  var lastCol = sh.getLastColumn();
  if (lastCol < 1) return;
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idx = headers.indexOf(oldName);
  if (idx !== -1 && headers.indexOf(newName) === -1) {
    sh.getRange(1, idx + 1).setValue(newName);
  }
}

/** skills 필드를 항상 객체로 정규화 (문자열로 저장돼 있으면 JSON 파싱) */
function parseSkills(raw) {
  if (!raw) return {};
  if (typeof raw === 'object') return raw;
  try { return JSON.parse(raw); } catch(e) { return {}; }
}

// ── 시트 초기화 (헤더 행 설정) ───────────────────────────
function initSheet(sh, name) {
  var headers = {
    '직원':      ['id','pw','name','role','color','hire','phone','team','rank','duty',
                  'cid1','cpw1','idms1','cid2','cpw2','idms2','bph','doc','s1','pcpw','pcpwDate',
                  'skills','ext','subUserName','wUserName','wcid','wcpw','widms','katalkId','katalkPw','retiredAt','acctDisposed','teamCardAdmin'],
    '실적':      ['id','empId','empName','date','attendStatus','rows'],
    '실적집계':  ['date','empId','empName','team','유선신규','약정갱신','약정갱신_인업셀','약정갱신_인약갱','약정갱신_티약갱','약정갱신_티업셀','유선기타','기가아이즈','무선','상품권',
                  '발신콜','발신시간','수신콜','수신시간','땡김콜','땡김시간','고정업무','고정업무시간','연차유형','VOC'],
    '연차':      ['id','empId','empName','date','type'],
    '연차공지':  ['notice'],
    'IP그룹':    ['id','type','name','prefix','range','subnet','gateway','dns','dns2','entries'],
    '변경로그':  ['dt','by','desc'],
    '접속이력':  ['dt','empId','empName','ip'],
    '자리배치':  ['id','label','empId'],
    '설정':      ['key','value'],
    '팀카드':    ['id','name','issuer','last4','note','active'],
    '팀카드사용내역': ['id','yearMonth','date','cardId','empId','empName','desc','amount','note'],
    '공지사항':  ['id','dt','authorId','author','title','body','pinned'],
    '공지이미지': ['id','noticeId','seq','part','dataUrl']
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
    case 'bulkUpdateAccounts': return bulkUpdateAccounts(data.rows, data.skipBlanks);
    case 'getLatestPerfDate': return getLatestPerfDate();
    // 실적
    case 'savePerf':       return savePerf(data.r);
    case 'getPerfs':       return getPerfs(data);
    case 'publishDailyPerf': return publishDailyPerf(data);
    case 'getDailyPerf':   return getDailyPerf(data.date);
    case 'getPerfSummary': return getPerfSummary(data.startDate, data.endDate);
    case 'deletePerf':     return deletePerf(data.empId, data.date);
    // 아카이브 (오래된 데이터 - "이전 기록 더보기")
    case 'getPerfsArchive':      return getPerfsArchive(data);
    case 'getChangeLogsArchive': return getChangeLogsArchive(data.startDate, data.endDate);
    case 'getAccessLogArchive':  return getAccessLogArchive(data.startDate, data.endDate);
    // 연차
    case 'saveAnnual':     return saveAnnual(data.a);
    case 'getAnnuals':     return getAnnuals(data.ym);
    case 'getAnnualsArchive': return getAnnualsArchive(data.ym);
    // IP
    case 'getTeamCards':   return getTeamCards();
    case 'addTeamCard':    return addTeamCard(data.card);
    case 'updateTeamCard': return updateTeamCard(data.id, data.fields);
    case 'deleteTeamCard': return deleteTeamCard(data.id);
    case 'getTeamCardUsage': return getTeamCardUsage(data.yearMonth);
    case 'addTeamCardUsage': return addTeamCardUsage(data.row);
    case 'updateTeamCardUsage': return updateTeamCardUsage(data.id, data.fields);
    case 'deleteTeamCardUsage': return deleteTeamCardUsage(data.id);
    case 'getTeamCardBudget': return getTeamCardBudget(data.yearMonth);
    case 'saveTeamCardBudget': return saveTeamCardBudget(data.yearMonth, data.amount);
    case 'getIPGroups':    return getIPGroups();
    case 'saveIPGroup':    return saveIPGroup(data.group);
    case 'deleteIPGroup':  return deleteIPGroup(data.id);
    // 변경로그
    case 'addChangeLog':   return addChangeLog(data);
    case 'getChangeLogs':  return getChangeLogs(data.days);
    case 'getChangeLogRange': return getChangeLogRange(data.startDate, data.endDate);
    case 'logAccess':      return logAccess(data.empId, data.empName, data.ip);
    case 'getAccessLog':   return getAccessLog(data.startDate, data.endDate);
    // 연차공지
    case 'getCalNotice':   return getCalNotice();
    case 'saveCalNotice':  return saveCalNotice(data.notice);
    // 공지사항 (전 직원 열람 · 관리자 작성)
    case 'getNotices':     return getNotices();
    case 'saveNotice':     return saveNotice(data.notice);
    case 'deleteNotice':   return deleteNotice(data.id);
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
  safe.skills = parseSkills(safe.skills);
  return { ok: true, user: safe };
}

// ════════════════════════════════════════════════════════════
//  직원
// ════════════════════════════════════════════════════════════
function getEmployees(includeAdmin) {
  var sh = getSheet('직원');
  var rows = sheetToObjects(sh);
  var list = includeAdmin ? rows : rows.filter(function(r){ return r.role !== 'admin'; });
  list = list.map(function(u){ var s=Object.assign({},u); delete s.pw; s.skills=parseSkills(s.skills); return s; });
  return { ok: true, employees: list };
}

function getUser(id) {
  var sh = getSheet('직원');
  var rows = sheetToObjects(sh);
  var u = rows.find(function(r){ return r.id === id; });
  if (!u) return { ok: false, error: '직원을 찾을 수 없습니다.' };
  u.skills = parseSkills(u.skills);
  return { ok: true, user: u };
}

function addEmployee(emp) {
  var sh = getSheet('직원');
  var rows = sheetToObjects(sh);
  if (rows.find(function(r){ return r.id === emp.id; }))
    return { ok: false, error: '이미 존재하는 아이디입니다.' };
  if (emp.skills && typeof emp.skills === 'object') emp.skills = JSON.stringify(emp.skills);
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
  if (fields.skills !== undefined && typeof fields.skills === 'object') {
    fields.skills = JSON.stringify(fields.skills);
  }
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
  // 본인이 수정 가능한 필드만 허용 (스킬은 관리자 전용이라 여기 포함하지 않음)
  var allowed = ['cid1','cpw1','idms1','cid2','cpw2','idms2','subUserName','wUserName','wcid','wcpw','widms','katalkId','katalkPw','pcpw','pcpwDate','ext'];
  var safe = {};
  allowed.forEach(function(k){ if (fields[k] !== undefined) safe[k] = fields[k]; });
  return updateEmpAdmin(id, safe);
}

/** 계정정보 엑셀 일괄업로드 - 이름(name)으로 기존 직원과 매칭해서 업데이트, 매칭 안 되면 unmatched로 반환 */
function bulkUpdateAccounts(rows, skipBlanks) {
  var sh = getSheet('직원');
  var data = sheetToObjects(sh);
  var headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var updated = [], unmatched = [];
  (rows||[]).forEach(function(r){
    var nm = (r.name||'').trim();
    if (!nm) return;
    var idx = -1;
    for (var i = 0; i < data.length; i++) { if ((data[i].name||'').trim() === nm) { idx = i; break; } }
    if (idx === -1) { unmatched.push(nm); return; }
    var empId = data[idx].id;
    var fields = {};
    ['subUserName','wUserName','wcid','wcpw','widms','katalkId','katalkPw','cid1','cpw1','idms1','cid2','cpw2','idms2'].forEach(function(k){
      if (r[k] === undefined) return;
      if (skipBlanks && r[k] === '') return; // 엑셀 업로드: 빈 칸은 기존 값 유지
      fields[k] = r[k]; // 수동 전체저장: 빈 칸이면 실제로 지움
    });
    updateEmpAdmin(empId, fields);
    updated.push(nm);
  });
  return { ok: true, updated: updated, unmatched: unmatched };
}

// ════════════════════════════════════════════════════════════
//  실적
// ════════════════════════════════════════════════════════════
function savePerf(r) {
  var sh = getSheet('실적');
  // 같은 empId + date 기존 행 삭제
  deleteRowWhere(sh, function(row, headers) {
    var ei = headers.indexOf('empId'), di = headers.indexOf('date');
    return row[ei] == r.empId && dstr(row[di]) == r.date;
  });
  var id = 'p_' + r.empId + '_' + r.date;
  sh.appendRow([
    id, r.empId, r.empName, r.date, r.attendStatus || '',
    JSON.stringify(r.rows || [])
  ]);
  return { ok: true };
}

// 실적 조회 - empId/ym/기간(s~e) 필터를 서버단에서 적용해서 필요한 만큼만 내려준다
// (예전엔 s/e를 프론트에서 보내도 서버가 무시하고 전체를 다 내려줬음 - 데이터가 쌓일수록 느려지는 원인 중 하나였음)
function filterPerfRows(rows, data) {
  if (data && data.empId) rows = rows.filter(function(r){ return r.empId === data.empId; });
  if (data && data.ym)    rows = rows.filter(function(r){ return r.date && r.date.indexOf(data.ym) === 0; });
  if (data && data.s)     rows = rows.filter(function(r){ return r.date && r.date >= data.s; });
  if (data && data.e)     rows = rows.filter(function(r){ return r.date && r.date <= data.e; });
  return rows.map(function(r){
    try { r.rows = typeof r.rows === 'string' ? JSON.parse(r.rows) : (r.rows || []); } catch(e){ r.rows=[]; }
    return r;
  });
}
function getPerfs(data) {
  var sh = getSheet('실적');
  var rows = filterPerfRows(sheetToObjects(sh), data);
  return { ok: true, perfs: rows };
}

function deletePerf(empId, date) {
  var sh = getSheet('실적');
  deleteRowWhere(sh, function(row, headers) {
    var ei = headers.indexOf('empId'), di = headers.indexOf('date');
    return row[ei] == empId && dstr(row[di]) == date;
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
    return row[ei] == a.empId && dstr(row[di]) == a.date;
  });
  if (a.type) { // type 있을 때만 저장 (빈 type = 삭제)
    var id = 'a_' + a.empId + '_' + a.date;
    sh.appendRow([id, a.empId, a.empName || '', a.date, a.type]);
  }
  return { ok: true };
}

/** 관리자가 실적집계 도구에서 "게시" 했을 때: 해당 날짜의 기존 집계행을 지우고 새로 씀 (덮어쓰기/upsert) */
function publishDailyPerf(data) {
  var sh = getSheet('실적집계');
  var headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var dateIdx = headers.indexOf('date');
  var lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    var dateVals = sh.getRange(2, dateIdx+1, lastRow-1, 1).getValues();
    for (var i = dateVals.length - 1; i >= 0; i--) {
      if (dstr(dateVals[i][0]) === data.date) sh.deleteRow(i+2);
    }
  }
  (data.rows||[]).forEach(function(r){
    var row = headers.map(function(h){
      if (h === 'date') return data.date;
      return r[h] !== undefined ? r[h] : '';
    });
    sh.appendRow(row);
  });
  return { ok: true, count: (data.rows||[]).length };
}

/** 실적집계 시트에서 게시된 날짜 중 가장 최근 날짜를 반환 (오늘 데이터가 아직 없을 때 전체실적 초기화면용) */
function getLatestPerfDate() {
  var sh = getSheet('실적집계');
  var rows = sheetToObjects(sh);
  if (!rows.length) return { ok: true, date: null };
  var dates = rows.map(function(r){ return r.date; }).filter(Boolean).sort();
  return { ok: true, date: dates[dates.length - 1] };
}

/** 특정 날짜에 게시된 실적집계 결과 조회 */
function getDailyPerf(date) {
  var sh = getSheet('실적집계');
  var rows = sheetToObjects(sh);
  rows = rows.filter(function(r){ return r.date === date; });
  return { ok: true, rows: rows };
}

/** 기간(주간/월간 등) 실적 조회 - startDate~endDate(둘 다 포함) 사이의 게시된 실적집계 행을 그대로 반환 (합산은 클라이언트에서) */
function getPerfSummary(startDate, endDate) {
  var sh = getSheet('실적집계');
  var rows = sheetToObjects(sh);
  rows = rows.filter(function(r){ return r.date && r.date >= startDate && r.date <= endDate; });
  return { ok: true, rows: rows };
}

function getAnnuals(ym) {
  var sh = getSheet('연차');
  var rows = sheetToObjects(sh);
  if (ym) rows = rows.filter(function(r){ return r.date && r.date.indexOf(ym) === 0; });
  return { ok: true, annuals: rows };
}

/** 연차 - 보관함(아카이브)에서 월 조회. 연차달력에서 오래된 달로 넘어갈 때 자동으로 같이 호출됨 */
function getAnnualsArchive(ym) {
  var sh = getArchiveSheetIfExists('연차');
  if (!sh) return { ok: true, annuals: [] };
  var rows = sheetToObjects(sh);
  if (ym) rows = rows.filter(function(r){ return r.date && r.date.indexOf(ym) === 0; });
  return { ok: true, annuals: rows };
}

// ════════════════════════════════════════════════════════════
//  팀카드관리 (teamCardAdmin 권한자 전용)
// ════════════════════════════════════════════════════════════
function getTeamCards() {
  var sh = getSheet('팀카드');
  return { ok: true, cards: sheetToObjects(sh) };
}
function addTeamCard(card) {
  var sh = getSheet('팀카드');
  var id = 'card_' + Date.now();
  sh.appendRow([id, card.name || '', card.issuer || '', card.last4 || '', card.note || '', 'Y']);
  return { ok: true, id: id };
}
function updateTeamCard(id, fields) {
  var sh = getSheet('팀카드');
  var headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var data = sh.getDataRange().getValues();
  var idCol = headers.indexOf('id');
  for (var i = 1; i < data.length; i++) {
    if (data[i][idCol] == id) {
      headers.forEach(function(h, ci){ if (fields[h] !== undefined) data[i][ci] = fields[h]; });
      sh.getDataRange().setValues(data);
      return { ok: true };
    }
  }
  return { ok: false, error: '카드를 찾을 수 없습니다.' };
}
function deleteTeamCard(id) {
  var sh = getSheet('팀카드');
  deleteRowById(sh, 'id', id);
  return { ok: true };
}
function getTeamCardUsage(yearMonth) {
  var sh = getSheet('팀카드사용내역');
  var rows = sheetToObjects(sh);
  if (yearMonth) rows = rows.filter(function(r){ return r.date && r.date.indexOf(yearMonth) === 0; });
  return { ok: true, rows: rows };
}
function addTeamCardUsage(row) {
  var sh = getSheet('팀카드사용내역');
  var id = 'tcu_' + Date.now() + '_' + Math.floor(Math.random()*1000);
  sh.appendRow([id, row.yearMonth||'', row.date||'', row.cardId||'', row.empId||'', row.empName||'', row.desc||'', row.amount||0, row.note||'']);
  return { ok: true, id: id };
}
function updateTeamCardUsage(id, fields) {
  var sh = getSheet('팀카드사용내역');
  var headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var data = sh.getDataRange().getValues();
  var idCol = headers.indexOf('id');
  for (var i = 1; i < data.length; i++) {
    if (data[i][idCol] == id) {
      headers.forEach(function(h, ci){ if (fields[h] !== undefined) data[i][ci] = fields[h]; });
      sh.getDataRange().setValues(data);
      return { ok: true };
    }
  }
  return { ok: false, error: '내역을 찾을 수 없습니다.' };
}
function deleteTeamCardUsage(id) {
  var sh = getSheet('팀카드사용내역');
  deleteRowById(sh, 'id', id);
  return { ok: true };
}
/** 월별 통합 예산은 설정 시트에 JSON({yearMonth: amount})으로 저장 (JSON_SETTINGS_KEYS에 등록됨) */
function getTeamCardBudget(yearMonth) {
  var s = getSettings().settings;
  var map = s.teamCardBudget || {};
  return { ok: true, amount: map[yearMonth] || 0 };
}
function saveTeamCardBudget(yearMonth, amount) {
  var s = getSettings().settings;
  var map = s.teamCardBudget || {};
  map[yearMonth] = amount;
  saveSettings({ teamCardBudget: map });
  return { ok: true };
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

/** 기간(startDate~endDate) 변경로그 조회 - 대표관리자 전용 로그 메뉴용 */
function getChangeLogRange(startDate, endDate) {
  var sh = getSheet('변경로그');
  var rows = sheetToObjects(sh);
  rows = rows.filter(function(r){ return r.dt && r.dt.slice(0,10) >= startDate && r.dt.slice(0,10) <= endDate; });
  rows.reverse(); // 최신순
  return { ok: true, logs: rows };
}

/** 접속(로그인) 기록 - 클라이언트가 ipify 등으로 알아낸 자신의 공인IP를 같이 보내줌 (Apps Script는 호출자 IP를 직접 알 수 없음) */
function logAccess(empId, empName, ip) {
  var sh = getSheet('접속이력');
  var dt = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
  sh.appendRow([dt, empId || '', empName || '', ip || '']);
  return { ok: true };
}

/** 기간별 접속이력 조회 - 대표관리자 전용 */
function getAccessLog(startDate, endDate) {
  var sh = getSheet('접속이력');
  var rows = sheetToObjects(sh);
  rows = rows.filter(function(r){ return r.dt && r.dt.slice(0,10) >= startDate && r.dt.slice(0,10) <= endDate; });
  rows.reverse(); // 최신순
  return { ok: true, logs: rows };
}

/** 변경로그 - 보관함(아카이브)에서 기간 조회. "이전 기록 더보기" 클릭 시에만 호출됨 */
function getChangeLogsArchive(startDate, endDate) {
  var sh = getArchiveSheetIfExists('변경로그');
  if (!sh) return { ok: true, logs: [] };
  var rows = sheetToObjects(sh);
  rows = rows.filter(function(r){ return r.dt && r.dt.slice(0,10) >= startDate && r.dt.slice(0,10) <= endDate; });
  rows.reverse();
  return { ok: true, logs: rows };
}
/** 접속이력 - 보관함(아카이브)에서 기간 조회. "이전 기록 더보기" 클릭 시에만 호출됨 */
function getAccessLogArchive(startDate, endDate) {
  var sh = getArchiveSheetIfExists('접속이력');
  if (!sh) return { ok: true, logs: [] };
  var rows = sheetToObjects(sh);
  rows = rows.filter(function(r){ return r.dt && r.dt.slice(0,10) >= startDate && r.dt.slice(0,10) <= endDate; });
  rows.reverse();
  return { ok: true, logs: rows };
}
/** 실적 - 보관함(아카이브)에서 조회(empId/ym). "이전 기록 더보기" 클릭 시에만 호출됨 */
function getPerfsArchive(data) {
  var sh = getArchiveSheetIfExists('실적');
  if (!sh) return { ok: true, perfs: [] };
  var rows = filterPerfRows(sheetToObjects(sh), data);
  return { ok: true, perfs: rows };
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
//  공지사항 (전 직원 열람 · 관리자 작성)
//  이미지는 base64 dataUrl로 저장하되, 구글시트 셀 한도(5만자)를 넘지 않도록
//  '공지이미지' 시트에 이미지 1장을 여러 조각(part)으로 나눠 저장하고 읽을 때 다시 합친다.
// ════════════════════════════════════════════════════════════
var NOTICE_IMG_CHUNK = 45000; // 셀당 dataUrl 조각 최대 길이 (5만자 한도 여유)

function getNotices() {
  var sh = getSheet('공지사항');
  var notices = sheetToObjects(sh);
  var imgs = sheetToObjects(getSheet('공지이미지'));
  // noticeId → seq → [part0, part1, ...]
  var byNotice = {};
  imgs.forEach(function(im){
    var nid = im.noticeId; if (!nid) return;
    var seq = Number(im.seq) || 0, part = Number(im.part) || 0;
    (byNotice[nid] = byNotice[nid] || {});
    (byNotice[nid][seq] = byNotice[nid][seq] || []);
    byNotice[nid][seq][part] = im.dataUrl || '';
  });
  notices.forEach(function(n){
    var m = byNotice[n.id] || {};
    var seqs = Object.keys(m).map(Number).sort(function(a,b){ return a-b; });
    n.images = seqs.map(function(s){ return (m[s] || []).join(''); }).filter(function(u){ return u; });
    n.pinned = (n.pinned === 'Y' || n.pinned === true);
  });
  // 정렬: 고정(pinned) 먼저, 그 다음 작성일시(dt) 최신순
  notices.sort(function(a,b){
    if (a.pinned !== b.pinned) return a.pinned ? -1 : 1;
    return String(b.dt).localeCompare(String(a.dt));
  });
  return { ok: true, notices: notices };
}

function saveNotice(n) {
  n = n || {};
  var sh = getSheet('공지사항');
  var id = n.id || ('notice_' + Date.now());
  var dt = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
  if (n.id) {
    // 수정: 기존 작성일시(dt)는 유지
    var existing = sheetToObjects(sh).find(function(r){ return r.id === n.id; });
    if (existing && existing.dt) dt = existing.dt;
    deleteRowById(sh, 'id', n.id);
  }
  sh.appendRow([id, dt, n.authorId || '', n.author || '', n.title || '', n.body || '', n.pinned ? 'Y' : '']);
  // 이미지 교체: 기존 조각 삭제 후 새로 저장
  var imgSh = getSheet('공지이미지');
  deleteRowWhere(imgSh, function(row, headers){ return row[headers.indexOf('noticeId')] === id; });
  (n.images || []).forEach(function(dataUrl, seq){
    dataUrl = String(dataUrl || '');
    var part = 0;
    for (var pos = 0; pos < dataUrl.length; pos += NOTICE_IMG_CHUNK) {
      imgSh.appendRow([id + '_' + seq + '_' + part, id, seq, part, dataUrl.substr(pos, NOTICE_IMG_CHUNK)]);
      part++;
    }
  });
  return { ok: true, id: id };
}

function deleteNotice(id) {
  var sh = getSheet('공지사항');
  deleteRowById(sh, 'id', id);
  var imgSh = getSheet('공지이미지');
  deleteRowWhere(imgSh, function(row, headers){ return row[headers.indexOf('noticeId')] === id; });
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
// 배열/객체 형태로 저장되는 설정 키 목록 — 구글시트가 배열을 문자열로 바꿔버리는 문제를 막기 위해
// 저장할 때 JSON.stringify, 읽을 때 JSON.parse를 항상 적용한다.
var JSON_SETTINGS_KEYS = {
  skillCatalog: DEFAULT_SKILL_CATALOG,
  perfaggFixdur: [],
  TEAMS_FLAT: ['우신기술사업부','우신1팀'],
  ETC_ITEMS: [],
  ETC_TYPES: {},
  SCORE_WEIGHTS: {},
  teamCardBudget: {}
};

function getSettings() {
  var sh = getSheet('설정');
  var rows = sheetToObjects(sh);
  var result = {};
  rows.forEach(function(r){ result[r.key] = r.value; });
  Object.keys(JSON_SETTINGS_KEYS).forEach(function(k){
    var def = JSON_SETTINGS_KEYS[k];
    if (result[k]) {
      if (typeof result[k] === 'string') {
        try { result[k] = JSON.parse(result[k]); } catch(e) { result[k] = def; }
      }
      // 이미 배열/객체면 그대로 둠 (드문 경우지만 방어적으로 처리)
    } else {
      result[k] = def;
      if (k === 'skillCatalog') saveSettings({ skillCatalog: def });
    }
  });
  return { ok: true, settings: result };
}

function saveSettings(data) {
  Object.keys(JSON_SETTINGS_KEYS).forEach(function(k){
    if (data[k] !== undefined && typeof data[k] === 'object') {
      data[k] = JSON.stringify(data[k]);
    }
  });
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
//  아카이브 (오래된 데이터 보관) — 원본 시트가 무한정 커지는 걸 막아서
//  실적/변경로그/접속이력처럼 계속 쌓이기만 하는 시트를 빠르게 유지하기 위함.
//  archiveOldData()를 Apps Script 트리거로 "매월 1일"에 자동 실행되게 등록해서 쓴다.
// ════════════════════════════════════════════════════════════

// 원본 시트별 보관 기준(며칠 지나면 아카이브로 옮길지) - dateCol: 기준이 되는 날짜 컬럼명
var ARCHIVE_RULES = [
  { sheet: '실적',    dateCol: 'date', keepDays: 90 },  // 최근 3개월
  { sheet: '변경로그', dateCol: 'dt',   keepDays: 30 },  // 최근 1개월
  { sheet: '접속이력', dateCol: 'dt',   keepDays: 30 },  // 최근 1개월
  { sheet: '연차',    dateCol: 'date', keepDays: 180 }  // 최근 6개월 (연차달력은 옛날 달도 넘겨보므로 넉넉하게)
];

/** 아카이브 전용 스프레드시트를 가져오거나(없으면) 새로 만들어서 설정 시트에 ID를 저장해둔다.
 *  한 번 만들어지면 이후엔 계속 같은 파일을 재사용한다. */
function getArchiveSS() {
  var s = getSettings().settings;
  var id = s.archiveSsId;
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) { /* 파일이 삭제된 경우 등 - 아래에서 새로 생성 */ }
  }
  var ss = SpreadsheetApp.create('업무관리 아카이브 (자동보관)');
  saveSettings({ archiveSsId: ss.getId() });
  return ss;
}

/** 아카이브에서 해당 이름의 시트를 가져온다. 없으면 새로 만들고(첫 생성 시 기본 "Sheet1"을 재사용) 헤더를 세팅한다. */
function getArchiveSheet(name, headers) {
  var ss = getArchiveSS();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    var sheets = ss.getSheets();
    if (sheets.length === 1 && sheets[0].getName() === 'Sheet1' && sheets[0].getLastRow() === 0) {
      sh = sheets[0]; sh.setName(name);
    } else {
      sh = ss.insertSheet(name);
    }
  }
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#E8EAED');
  }
  return sh;
}

/** "이전 기록 더보기" 조회용 - 아카이브가 아직 한 번도 안 만들어졌으면(archiveSsId 없음) 굳이 새로 만들지 않고 null 반환 */
function getArchiveSheetIfExists(name) {
  var s = getSettings().settings;
  if (!s.archiveSsId) return null;
  try {
    var ss = SpreadsheetApp.openById(s.archiveSsId);
    return ss.getSheetByName(name) || null;
  } catch (e) {
    return null;
  }
}

/** dateCol 값이 cutoff(yyyy-MM-dd)보다 이전인 행을 아카이브로 옮기고 원본에서 제거.
 *  한 줄씩 삭제/추가하지 않고 "남길 행만 골라 한 번에 다시 쓰는" 방식(batch)으로 처리한다. */
function archiveSheetByDateCol(sheetName, dateColName, cutoff) {
  var sh = getSheet(sheetName);
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return 0;
  var headers = data[0];
  var colIdx = headers.indexOf(dateColName);
  if (colIdx === -1) return 0;
  var keep = [], toArchive = [];
  for (var i = 1; i < data.length; i++) {
    var v = dstr(data[i][colIdx]);
    var d = String(v || '').slice(0, 10);
    if (d && d < cutoff) toArchive.push(data[i]); else keep.push(data[i]);
  }
  if (!toArchive.length) return 0;

  // 아카이브로 이동 (한 번에 batch append)
  var archSh = getArchiveSheet(sheetName, headers);
  var archLast = archSh.getLastRow();
  archSh.getRange(archLast + 1, 1, toArchive.length, headers.length).setValues(toArchive);

  // 원본은 남길 행만 다시 쓰기 (한 줄씩 deleteRow 반복하지 않음)
  sh.getRange(1, 1, sh.getLastRow(), headers.length).clearContent();
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#E8EAED');
  if (keep.length) sh.getRange(2, 1, keep.length, headers.length).setValues(keep);

  return toArchive.length;
}

/** 매월 1회(트리거) 실행 - 실적/변경로그/접속이력에서 보관 기준을 넘긴 데이터를 아카이브로 이동.
 *  Apps Script 편집기 좌측 "트리거" 메뉴에서 이 함수를 "월 단위 타이머 → 매월 1일"로 등록해두면 자동 실행된다. */
function archiveOldData() {
  var tz = Session.getScriptTimeZone() || 'Asia/Seoul';
  var summary = [];
  ARCHIVE_RULES.forEach(function(rule) {
    var cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - rule.keepDays);
    var cutoff = Utilities.formatDate(cutoffDate, tz, 'yyyy-MM-dd');
    var cnt = archiveSheetByDateCol(rule.sheet, rule.dateCol, cutoff);
    summary.push(rule.sheet + ' ' + cnt + '건 (기준일 이전: ' + cutoff + ')');
  });
  Logger.log('아카이브 완료 - ' + summary.join(' / '));
  return { ok: true, summary: summary };
}

// ════════════════════════════════════════════════════════════
//  유틸 함수
// ════════════════════════════════════════════════════════════

/** 시트 전체를 객체 배열로 변환 */
/** 셀 값이 Date 객체로 들어와도 항상 문자열로 통일 (시트 셀 서식이 자동으로 날짜/시간으로 바뀌는 문제 방지)
 *  시간 정보가 없으면 'yyyy-MM-dd', 있으면 'yyyy-MM-dd HH:mm:ss'로 변환 */
function dstr(v) {
  if (v instanceof Date) {
    var tz = Session.getScriptTimeZone() || 'Asia/Seoul';
    // 스프레드시트가 "시간만 있는 값"(예: 00:17:23)을 자동으로 날짜형으로 바꿔버리면
    // 엑셀/시트 기준일인 1899-12-30이 붙어서 저장된다 - 이 경우는 날짜 없이 시간만 포맷
    if (v.getFullYear() === 1899 && v.getMonth() === 11 && v.getDate() === 30) {
      return Utilities.formatDate(v, tz, 'HH:mm:ss');
    }
    var hasTime = v.getHours() || v.getMinutes() || v.getSeconds();
    return Utilities.formatDate(v, tz, hasTime ? 'yyyy-MM-dd HH:mm:ss' : 'yyyy-MM-dd');
  }
  return v;
}

function sheetToObjects(sh) {
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0];
  return data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i){ var v = dstr(row[i]); obj[h] = v === '' ? '' : v; });
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
  var sheetNames = ['직원','실적','실적집계','연차','연차공지','IP그룹','변경로그','접속이력','자리배치','설정','팀카드','팀카드사용내역','공지사항','공지이미지'];
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
