/**
 * 제로초이스 무인 편의점 - Google Apps Script
 *
 * 시트 열 구조:
 *   A=매장정보  B=카테고리  C=상품이름  D=바코드
 *   E=단위      F=재고수량  G=입고가    H=판매가  I=과세여부
 *
 * 웹 앱 배포 후 URL을 앱의 "구글 시트 연동" 란에 붙여넣으세요.
 * 배포 설정: 실행 계정 = 나, 액세스 권한 = 모든 사용자
 */

const SHEET_NAME = 'Sheet1'; // 실제 시트 탭 이름으로 변경
const COL_NAME   = 3;  // C열: 상품이름
const COL_QTY    = 6;  // F열: 재고수량
const HEADER_ROW = 1;  // 헤더 행 수 (데이터 시작 행 = HEADER_ROW + 1)

function doGet(e) {
  const action = e.parameter.action || 'list';

  try {
    if (action === 'list')   return respond(listItems());
    if (action === 'qty')    return respond(updateQty(e.parameter));
    return respond({ error: '알 수 없는 action: ' + action });
  } catch (err) {
    return respond({ error: err.message });
  }
}

// ── 목록 가져오기 ─────────────────────────────────────────────
function listItems() {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow <= HEADER_ROW) return { items: [] };

  const range = sheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, COL_QTY);
  const data  = range.getValues();

  const items = data
    .map((row, i) => {
      const name = String(row[COL_NAME - 1] || '').trim();
      if (!name) return null;

      const qtyRaw = row[COL_QTY - 1];
      const qty    = (qtyRaw !== '' && qtyRaw !== null && qtyRaw !== undefined)
        ? Number(qtyRaw)
        : null;

      return {
        id:   'sheet_' + (HEADER_ROW + 1 + i),
        name: name,
        qty:  isNaN(qty) ? null : qty,
        // 유통기한 열 없음 — 앱 표시용 더미 날짜 (먼 미래)
        date: '2099-12-31',
        row:  HEADER_ROW + 1 + i
      };
    })
    .filter(Boolean);

  return { items: items };
}

// ── 재고수량 업데이트 ─────────────────────────────────────────
function updateQty(params) {
  const id  = params.id;
  const qty = Number(params.qty);

  if (!id)        throw new Error('id 파라미터 없음');
  if (isNaN(qty)) throw new Error('qty 파라미터가 숫자가 아님');

  // id 형식: "sheet_<행번호>"
  const rowNum = parseInt(id.replace('sheet_', ''), 10);
  if (!rowNum)  throw new Error('row 번호를 파싱할 수 없음: ' + id);

  const sheet = getSheet();
  sheet.getRange(rowNum, COL_QTY).setValue(qty);

  return { ok: true, id: id, qty: qty };
}

// ── 헬퍼 ─────────────────────────────────────────────────────
function getSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  if (!sheet) throw new Error('시트를 찾을 수 없습니다: ' + SHEET_NAME);
  return sheet;
}

function respond(data) {
  const output = ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}
