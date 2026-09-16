/**
 * 〈어떻게 해야 했을까?〉 좌석 배정 — 구글 시트 백엔드
 *
 * 시트 구성 (없으면 자동으로 만들어집니다)
 *   명단 : ID | 성함 | 소속 | 구분                                  ← 명단의 원본. 여기서 직접 고쳐도 됩니다.
 *   배정 : 좌석 | 좌석코드 | 성함 | 소속 | 구분 | ID | 배정시각       ← 앱이 통째로 덮어씁니다. 손대지 마세요.
 *
 * 배포 전 준비
 *   1) 스프레드시트 > 확장 프로그램 > Apps Script 에 이 코드를 붙여넣습니다.
 *   2) 프로젝트 설정 > 스크립트 속성에 TOKEN 을 추가하고 값을 정합니다.
 *   3) 배포 > 새 배포 > 웹 앱 / 실행: 나 / 액세스: 모든 사용자
 *   4) 나온 /exec 주소와 TOKEN 을 index.html 위쪽 두 줄에 넣습니다.
 *
 * 코드를 고친 뒤에는 [배포 관리]에서 '새 버전'으로 다시 배포해야 반영됩니다.
 */

var SH_ROSTER = '명단';
var SH_SEATS = '배정';
var ROSTER_HEAD = ['ID', '성함', '소속', '구분'];
var SEATS_HEAD = ['좌석', '좌석코드', '성함', '소속', '구분', 'ID', '배정시각'];

function doPost(e) {
  try {
    var req = JSON.parse(e.postData.contents);
    if (req.token !== getToken()) return out({ ok: false, error: 'bad_token' });

    switch (req.action) {
      case 'bootstrap':    return out(bootstrap());
      case 'rosterUpsert': return out(rosterUpsert(req.person));
      case 'rosterDelete': return out(rosterDelete(req.id));
      case 'saveSeats':    return out(saveSeats(req.rows));
      default:             return out({ ok: false, error: 'unknown_action' });
    }
  } catch (err) {
    return out({ ok: false, error: String(err) });
  }
}

function doGet() {
  return out({ ok: true, message: '좌석 배정 백엔드가 연결되어 있습니다.' });
}

function getToken() {
  return PropertiesService.getScriptProperties().getProperty('TOKEN');
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function sheet(name, head) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** 명단과 배정 현황을 한 번에 돌려줍니다. 앱을 열 때 호출합니다. */
function bootstrap() {
  return { ok: true, roster: readRoster(), seats: readSeats(), at: new Date().toISOString() };
}

function readRoster() {
  var sh = sheet(SH_ROSTER, ROSTER_HEAD);
  if (sh.getLastRow() < 2) return [];
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
  var list = [];
  vals.forEach(function (r) {
    var name = String(r[1] || '').trim();
    if (!name) return;                         // 빈 줄은 건너뜁니다
    list.push({
      id: String(r[0] || '').trim() || ('r' + Utilities.getUuid().slice(0, 8)),
      name: name,
      org: String(r[2] || '').trim(),
      group: String(r[3] || '').trim() || '회원'
    });
  });
  return list;
}

function readSeats() {
  var sh = sheet(SH_SEATS, SEATS_HEAD);
  if (sh.getLastRow() < 2) return [];
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues();
  var list = [];
  vals.forEach(function (r) {
    var code = String(r[1] || '').trim();
    var id = String(r[5] || '').trim();
    if (!code || !id) return;
    list.push({
      seat: code,
      pid: id,
      name: String(r[2] || '').trim(),
      org: String(r[3] || '').trim(),
      group: String(r[4] || '').trim(),
      at: r[6] ? String(r[6]) : ''
    });
  });
  return list;
}

/** ID가 있으면 그 줄을 고치고, 없으면 새 줄로 추가합니다. */
function rosterUpsert(p) {
  if (!p || !p.name) return { ok: false, error: 'no_name' };
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sheet(SH_ROSTER, ROSTER_HEAD);
    var id = String(p.id || '').trim() || ('x' + Utilities.getUuid().slice(0, 8));
    var row = findRow(sh, id);
    var values = [[id, p.name, p.org || '', p.group || '회원']];
    if (row > 0) {
      sh.getRange(row, 1, 1, 4).setValues(values);
    } else {
      sh.getRange(sh.getLastRow() + 1, 1, 1, 4).setValues(values);
    }
    return { ok: true, id: id };
  } finally {
    lock.releaseLock();
  }
}

function rosterDelete(id) {
  if (!id) return { ok: false, error: 'no_id' };
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sheet(SH_ROSTER, ROSTER_HEAD);
    var row = findRow(sh, String(id));
    if (row > 0) sh.deleteRow(row);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function findRow(sh, id) {
  if (sh.getLastRow() < 2) return -1;
  var ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === id) return i + 2;
  }
  return -1;
}

/** 배정 현황 전체를 받아 '배정' 시트를 다시 씁니다. 같은 내용을 여러 번 보내도 안전합니다. */
function saveSeats(rows) {
  rows = rows || [];
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sheet(SH_SEATS, SEATS_HEAD);
    sh.clear();
    sh.getRange(1, 1, 1, SEATS_HEAD.length).setValues([SEATS_HEAD]).setFontWeight('bold');
    if (rows.length) sh.getRange(2, 1, rows.length, SEATS_HEAD.length).setValues(rows);
    sh.setFrozenRows(1);
    return { ok: true, assigned: rows.length, savedAt: new Date().toISOString() };
  } finally {
    lock.releaseLock();
  }
}
