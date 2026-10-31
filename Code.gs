/**
 * 삼성중학교 공개 수업 참관 신청 → 구글 시트 자동 취합
 * 배포: 확장 프로그램 > Apps Script > 배포 > 새 배포 > 웹 앱
 *       실행 사용자: 나 / 액세스 권한: 모든 사용자
 */
var SHEET_ID = '1fMP-P4emzD7yUO7QEdXKzsxfF7Yx8zYuyF5gTtqBudk';
var SHEET_NAME = '참관신청'; // 없으면 자동 생성
var HEADERS = ['접수시각', '소속', '이름', '참관 희망 수업', '수업나눔 참석'];
var REVIEW_SHEET_NAME = '참관소감'; // 없으면 자동 생성
var REVIEW_HEADERS = ['접수시각', '소속', '이름', '참관한 수업', '참관 소감'];

function getSheet_(name, headers) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sh;
}

function now_() {
  return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
}

function clean_(v, max) {
  v = String(v == null ? '' : v).trim().slice(0, max || 100);
  // 수식 주입 방지
  if (/^[=+\-@]/.test(v)) v = "'" + v;
  return v;
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var d = JSON.parse(e.postData.contents);
    var org = clean_(d.org), name = clean_(d.name);
    if (!org || !name) {
      return out_({ ok: false, error: 'missing' });
    }
    if (d.type === 'review') {
      var text = clean_(d.text, 2000);
      if (!text) return out_({ ok: false, error: 'missing' });
      getSheet_(REVIEW_SHEET_NAME, REVIEW_HEADERS)
        .appendRow([now_(), org, name, clean_(d.cls), text]);
      return out_({ ok: true });
    }
    getSheet_(SHEET_NAME, HEADERS).appendRow([
      now_(), org, name, clean_(d.cls), clean_(d.share, 10)
    ]);
    return out_({ ok: true });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

/**
 * 교사용 조회 (admin.html)
 * 비밀번호: Apps Script 왼쪽 '프로젝트 설정' > 스크립트 속성 > ADMIN_PW 에 입력
 * 요청: ?action=list&pw=비밀번호
 */
function readTab_(name, headers) {
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
  var rows = [];
  if (sh && sh.getLastRow() > 1) {
    rows = sh.getRange(2, 1, sh.getLastRow() - 1, headers.length).getDisplayValues();
  }
  return { headers: headers, rows: rows };
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.action !== 'list') return out_({ ok: true, msg: 'alive' });
  var pw = PropertiesService.getScriptProperties().getProperty('ADMIN_PW');
  if (!pw || String(p.pw || '') !== pw) {
    Utilities.sleep(800); // 무작위 대입 억제
    return out_({ ok: false, error: 'auth' });
  }
  return out_({
    ok: true,
    apply: readTab_(SHEET_NAME, HEADERS),
    review: readTab_(REVIEW_SHEET_NAME, REVIEW_HEADERS)
  });
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}
