/**
 * 삼성중학교 공개 수업 참관 신청 → 구글 시트 자동 취합
 * 배포: 확장 프로그램 > Apps Script > 배포 > 새 배포 > 웹 앱
 *       실행 사용자: 나 / 액세스 권한: 모든 사용자
 */
var SHEET_ID = '1fMP-P4emzD7yUO7QEdXKzsxfF7Yx8zYuyF5gTtqBudk';
var SHEET_NAME = '참관신청'; // 없으면 자동 생성
var HEADERS = ['접수시각', '소속', '이름', '참관 희망 수업', '수업나눔 참석'];

function getSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
  return sh;
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
    getSheet_().appendRow([
      Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'),
      org, name, clean_(d.cls), clean_(d.share, 10)
    ]);
    return out_({ ok: true });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function doGet() { return out_({ ok: true, msg: 'alive' }); }

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}
