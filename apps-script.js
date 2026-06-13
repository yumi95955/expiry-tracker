/**
 * 제로초이스 무인 편의점 - Google Apps Script
 *
 * 시트 탭 구조:
 *   [무잉 데이터]         1~4행 헤더, 5행~데이터
 *                         C=상품이름  D=바코드  F=재고수량  G=입고가  H=판매가
 *   [재고수량및유통기한]  1행 헤더, 2행~데이터
 *                         B=상품이름  C=바코드  H=유통기한
 *   [체크리스트]          A=key  B=checked  C=savedDate
 *   [요거트기록]          A=filled  B=remove
 *   [에이전트기록]        A=agentIdx  B=role  C=content  D=timestamp
 *
 * 배포 설정: 실행 계정 = 나, 액세스 권한 = 모든 사용자
 */

const SHEET_MUING     = '무잉 데이터';
const SHEET_EXPIRY    = '재고수량및유통기한';
const SHEET_CHECKLIST = '체크리스트';
const SHEET_YOGURT    = '요거트기록';
const SHEET_AGENT     = '에이전트기록';

// 무잉 데이터 탭
const MUING_HEADER_ROWS = 4;   // 1~4행 헤더, 5행부터 데이터
const COL_MUING_NAME    = 3;   // C열: 상품이름
const COL_MUING_BARCODE = 4;   // D열: 바코드
const COL_MUING_QTY     = 6;   // F열: 재고수량

// 재고수량및유통기한 탭
const COL_EXP_BARCODE = 3;     // C열: 바코드 (조인 키)
const COL_EXP_DATE    = 8;     // H열: 유통기한

// ── GET: 읽기 + 쓰기 모두 처리 ───────────────────────────────
// 브라우저 fetch의 POST는 Apps Script 리다이렉트에서 GET으로 변환되므로
// 쓰기 데이터도 ?data=<JSON> 파라미터로 GET 전송한다.
function doGet(e) {
  try {
    // 쓰기 요청: data 파라미터에 JSON 인코딩
    if (e.parameter.data) {
      const params = JSON.parse(e.parameter.data);
      const action = params.action;
      if (action === 'checklist_save') return respond(saveChecklist(params));
      if (action === 'yogurt_save')    return respond(saveYogurt(params));
      if (action === 'agent_append')   return respond(appendAgent(params));
      if (action === 'agent_clear')    return respond(clearAgent(params));
      return respond({ error: '알 수 없는 쓰기 action: ' + action });
    }

    // 읽기 요청
    const action = e.parameter.action || 'list';
    if (action === 'list')           return respond(listItems());
    if (action === 'qty')            return respond(updateQty(e.parameter));
    if (action === 'checklist_load') return respond(loadChecklist());
    if (action === 'yogurt_load')    return respond(loadYogurt());
    if (action === 'agent_load')     return respond(loadAgent());
    return respond({ error: '알 수 없는 action: ' + action });
  } catch (err) {
    return respond({ error: err.message });
  }
}

// doPost는 스크립트 간 직접 호출용으로 유지
function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    const action = params.action;
    if (action === 'checklist_save') return respond(saveChecklist(params));
    if (action === 'yogurt_save')    return respond(saveYogurt(params));
    if (action === 'agent_append')   return respond(appendAgent(params));
    if (action === 'agent_clear')    return respond(clearAgent(params));
    return respond({ error: '알 수 없는 action: ' + action });
  } catch (err) {
    return respond({ error: err.message });
  }
}

// ── 상품 목록 (무잉 데이터 + 재고수량및유통기한 조인) ──────────
function listItems() {
  // 1. 재고수량및유통기한에서 바코드→유통기한 맵 생성
  const dateMap = buildDateMap();

  // 2. 무잉 데이터에서 상품 읽기
  const sheet   = getSheet(SHEET_MUING);
  const lastRow = sheet.getLastRow();
  const startRow = MUING_HEADER_ROWS + 1;  // 5행부터
  if (lastRow < startRow) return { items: [] };

  const numRows  = lastRow - MUING_HEADER_ROWS;
  const nameData = sheet.getRange(startRow, COL_MUING_NAME,    numRows, 1).getValues();
  const bcData   = sheet.getRange(startRow, COL_MUING_BARCODE, numRows, 1).getValues();
  const qtyData  = sheet.getRange(startRow, COL_MUING_QTY,     numRows, 1).getValues();

  const items = nameData.map((nameRow, i) => {
    const name = String(nameRow[0] || '').trim();
    if (!name) return null;

    const barcode = String(bcData[i][0] || '').trim();
    const qtyRaw  = qtyData[i][0];
    const qty     = (qtyRaw !== '' && qtyRaw !== null && qtyRaw !== undefined)
      ? Number(qtyRaw) : null;

    // 바코드로 유통기한 조회
    const date = (barcode && dateMap[barcode]) ? dateMap[barcode] : '2099-12-31';

    return {
      id:      'sheet_' + (startRow + i),
      name:    name,
      barcode: barcode,
      qty:     isNaN(qty) ? null : qty,
      date:    date,
      row:     startRow + i
    };
  }).filter(Boolean);

  return { items };
}

// 재고수량및유통기한 탭에서 바코드→유통기한 맵 반환
function buildDateMap() {
  const map = {};
  try {
    const sheet   = getSheet(SHEET_EXPIRY);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return map;

    const numRows  = lastRow - 1;
    const bcData   = sheet.getRange(2, COL_EXP_BARCODE, numRows, 1).getValues();
    const dateData = sheet.getRange(2, COL_EXP_DATE,    numRows, 1).getValues();

    bcData.forEach((row, i) => {
      const bc = String(row[0] || '').trim();
      if (!bc) return;
      const raw = dateData[i][0];
      if (!raw) return;
      let dateStr;
      if (raw instanceof Date) {
        const y = raw.getFullYear();
        const m = String(raw.getMonth() + 1).padStart(2, '0');
        const d = String(raw.getDate()).padStart(2, '0');
        dateStr = y + '-' + m + '-' + d;
      } else {
        dateStr = String(raw).trim();
      }
      if (dateStr) map[bc] = dateStr;
    });
  } catch(e) {}
  return map;
}

function updateQty(params) {
  const id  = params.id;
  const qty = Number(params.qty);
  if (!id)        throw new Error('id 파라미터 없음');
  if (isNaN(qty)) throw new Error('qty 파라미터가 숫자가 아님');
  const rowNum = parseInt(id.replace('sheet_', ''), 10);
  if (!rowNum)    throw new Error('row 번호를 파싱할 수 없음: ' + id);
  getSheet(SHEET_MUING).getRange(rowNum, COL_MUING_QTY).setValue(qty);
  return { ok: true, id, qty };
}

// ── 체크리스트 ────────────────────────────────────────────────
function loadChecklist() {
  const sheet = getSheet(SHEET_CHECKLIST);
  const last  = sheet.getLastRow();
  if (last < 2) return { state: {}, savedDate: '' };

  const data  = sheet.getRange(2, 1, last - 1, 3).getValues();
  const state = {};
  let savedDate = '';
  data.forEach(row => {
    const key = String(row[0] || '').trim();
    if (!key) return;
    state[key] = row[1] === true || row[1] === 'TRUE';
    if (!savedDate && row[2]) {
      // Sheets가 날짜 문자열을 Date 객체로 자동변환하므로 YYYY-MM-DD로 정규화
      const raw = row[2];
      if (raw instanceof Date) {
        const y = raw.getFullYear();
        const m = String(raw.getMonth() + 1).padStart(2, '0');
        const d = String(raw.getDate()).padStart(2, '0');
        savedDate = y + '-' + m + '-' + d;
      } else {
        savedDate = String(raw);
      }
    }
  });
  return { state, savedDate };
}

function saveChecklist(params) {
  const state     = params.state     || {};
  const savedDate = params.savedDate || '';
  const sheet     = getSheet(SHEET_CHECKLIST);

  ensureHeader(sheet, ['key', 'checked', 'savedDate']);

  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).clearContent();
  }

  const entries = Object.entries(state);
  if (!entries.length) return { ok: true };

  sheet.getRange(2, 1, entries.length, 3)
    .setValues(entries.map(([k, v]) => [k, v, savedDate]));
  return { ok: true };
}

// ── 요거트 기록 ───────────────────────────────────────────────
function loadYogurt() {
  const sheet = getSheet(SHEET_YOGURT);
  const last  = sheet.getLastRow();
  if (last < 2) return { logs: [] };

  const data = sheet.getRange(2, 1, last - 1, 2).getValues();
  const logs = data
    .filter(row => row[0])
    .map(row => ({ filled: String(row[0]), remove: String(row[1]) }));
  return { logs };
}

function saveYogurt(params) {
  const logs  = params.logs || [];
  const sheet = getSheet(SHEET_YOGURT);

  ensureHeader(sheet, ['filled', 'remove']);

  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).clearContent();
  }
  if (!logs.length) return { ok: true };

  sheet.getRange(2, 1, logs.length, 2)
    .setValues(logs.map(l => [l.filled, l.remove]));
  return { ok: true };
}

// ── 에이전트 기록 ─────────────────────────────────────────────
function loadAgent() {
  const sheet = getSheet(SHEET_AGENT);
  const last  = sheet.getLastRow();
  const histories = [[], [], []];
  if (last < 2) return { histories };

  const data = sheet.getRange(2, 1, last - 1, 3).getValues();
  data.forEach(row => {
    const idx     = Number(row[0]);
    const role    = String(row[1] || '');
    const content = String(row[2] || '');
    if (idx >= 0 && idx <= 2 && role && content) {
      histories[idx].push({ role, content });
    }
  });
  return { histories };
}

function appendAgent(params) {
  const { agentIdx, role, content } = params;
  if (agentIdx === undefined || !role || !content) {
    throw new Error('agentIdx / role / content 필요');
  }
  const sheet = getSheet(SHEET_AGENT);
  ensureHeader(sheet, ['agentIdx', 'role', 'content', 'timestamp']);
  sheet.appendRow([agentIdx, role, content, new Date().toISOString()]);
  return { ok: true };
}

function clearAgent(params) {
  const { agentIdx } = params;
  const sheet = getSheet(SHEET_AGENT);
  const last  = sheet.getLastRow();
  if (last < 2) return { ok: true };

  const col = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (let i = col.length - 1; i >= 0; i--) {
    if (Number(col[i][0]) === Number(agentIdx)) sheet.deleteRow(i + 2);
  }
  return { ok: true };
}

// ── 헬퍼 ─────────────────────────────────────────────────────
function getSheet(name) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('시트를 찾을 수 없습니다: ' + name);
  return sheet;
}

function ensureHeader(sheet, headers) {
  if (sheet.getLastRow() < 1) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
