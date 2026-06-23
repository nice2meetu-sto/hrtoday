/**
 * 인력현황 대시보드 — Apps Script (GET 전용 JSON API)
 *
 * 사용법
 * 1) 데이터가 든 구글시트(시트 탭: '인원', '제도')에서 확장 프로그램 > Apps Script 열기
 * 2) 이 코드를 붙여넣고 저장
 * 3) 배포 > 새 배포 > 유형: 웹 앱
 *    - 실행 계정: 나
 *    - 액세스 권한: 모든 사용자
 * 4) 배포 후 나오는 '웹 앱 URL'을 복사해서 index.html 의 SCRIPT_URL 에 붙여넣기
 *
 * 반환 형식:
 *   { ok:true, data:{ 인원:[{사번,사원명,...}], 제도:[{사번,제도,...}] } }
 *
 * 날짜 컬럼(입사일/퇴사일/시작일/종료일)은 시트에서 '텍스트' 서식의
 * YYYY-MM-DD 문자열로 관리하세요. getDisplayValues()가 보이는 그대로
 * 문자열로 내보내므로 시간대/직렬화 문제가 없습니다.
 */

const SHEET_NAMES = { 인원: '인원', 제도: '제도' };

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const data = {};
    Object.keys(SHEET_NAMES).forEach(function (key) {
      data[key] = readSheet(ss.getSheetByName(SHEET_NAMES[key]));
    });
    return json({ ok: true, data: data });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function readSheet(sheet) {
  if (!sheet) return [];
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return [];
  const headers = values[0].map(function (h) { return String(h).trim(); });
  return values.slice(1)
    .filter(function (row) { return row.some(function (c) { return String(c).trim() !== ''; }); })
    .map(function (row) {
      const obj = {};
      headers.forEach(function (h, i) { obj[h] = String(row[i]).trim(); });
      return obj;
    });
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
