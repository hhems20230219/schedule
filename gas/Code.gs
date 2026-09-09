/**
 * =========================================================
 * 火警出動人員車輛分配表 - Google Apps Script API
 * =========================================================
 * 架構：
 * GitHub Pages -> fetch() -> Apps Script Web App -> Google Sheet
 *
 * Google Sheet 只使用一張工作表：BoardData
 * 不需要另外建立 Personnel / DutySchedule / BoardState 等工作表。
 */

/* =========================================================
 * 01. 固定設定
 * ---------------------------------------------------------
 * 本 Apps Script 建議直接綁定在要使用的 Google Sheet 上。
 * 若 BoardData 不存在，第一次讀寫時會自動建立。
 * ========================================================= */
const SHEET_NAME = 'BoardData';
const TIME_ZONE = 'Asia/Taipei';

/* =========================================================
 * 02. GET API
 * ---------------------------------------------------------
 * GitHub Pages 開啟時，以及之後定時更新時使用。
 * 範例：WEB_APP_URL?action=get
 * ========================================================= */
function doGet(e) {
  try {
    const action = String((e && e.parameter && e.parameter.action) || 'get').trim();

    if (action !== 'get') {
      return jsonResponse_({
        ok: false,
        message: '不支援的 GET action：' + action
      });
    }

    return jsonResponse_(readBoardState_());
  } catch (error) {
    return jsonResponse_({
      ok: false,
      message: error.message || String(error)
    });
  }
}

/* =========================================================
 * 03. POST API
 * ---------------------------------------------------------
 * 值班台匯入 Excel 或完成拖曳後，由 GitHub Pages 自動呼叫。
 * 前端送出的 body：{ action: "save", state: {...} }
 * ========================================================= */
function doPost(e) {
  try {
    const payload = parseRequestBody_(e);
    const action = String(payload.action || '').trim();

    if (action !== 'save') {
      return jsonResponse_({
        ok: false,
        message: '不支援的 POST action：' + action
      });
    }

    return jsonResponse_(saveBoardState_(payload.state));
  } catch (error) {
    return jsonResponse_({
      ok: false,
      message: error.message || String(error)
    });
  }
}

/* =========================================================
 * 04. 取得唯一資料工作表
 * ---------------------------------------------------------
 * BoardData 第 1 列固定為欄名：key / value。
 * 第一次沒有工作表時會自動建立，因此不需要人工建欄位。
 * ========================================================= */
function getDataSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error('找不到綁定的 Google Sheet，請從試算表的「擴充功能 → Apps Script」建立此專案。');
  }

  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange('A1:B1').setValues([['key', 'value']]);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

/* =========================================================
 * 05. 讀取目前完整看板
 * ---------------------------------------------------------
 * boardState 是值班台最後一次自動同步的完整 JSON。
 * 完全沒有 boardState 時回傳 state:null，前端就顯示全白。
 * ========================================================= */
function readBoardState_() {
  const sheet = getDataSheet_();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return {
      ok: true,
      state: null,
      updatedAt: ''
    };
  }

  const rows = sheet.getRange(2, 1, lastRow - 1, 2).getDisplayValues();
  const data = {};

  rows.forEach(function(row) {
    const key = String(row[0] || '').trim();
    if (key) data[key] = row[1];
  });

  if (!data.boardState) {
    return {
      ok: true,
      state: null,
      updatedAt: data.updatedAt || ''
    };
  }

  return {
    ok: true,
    state: JSON.parse(String(data.boardState)),
    updatedAt: data.updatedAt || ''
  };
}

/* =========================================================
 * 06. 儲存目前完整看板
 * ---------------------------------------------------------
 * 不論是「剛匯入但沒人工修改」或「人工拖曳後」，目前畫面
 * 都直接覆蓋成最新 boardState，不存在草稿 / 發布兩套資料。
 * ========================================================= */
function saveBoardState_(state) {
  if (!state || typeof state !== 'object' || Array.isArray(state)) {
    throw new Error('沒有可儲存的看板資料');
  }

  const json = JSON.stringify(state);

  // Google Sheets 單一儲存格上限約 50,000 字元，先主動檢查避免寫入失敗。
  if (json.length > 48000) {
    throw new Error('boardState 資料過大，已接近 Google Sheet 單一儲存格限制');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheet = getDataSheet_();
    const updatedAt = Utilities.formatDate(new Date(), TIME_ZONE, 'yyyy-MM-dd HH:mm:ss');

    // 永遠只有一份最新狀態；直接覆蓋 A1:B3，舊資料不累積。
    sheet.clearContents();
    sheet.getRange('A1:B3').setValues([
      ['key', 'value'],
      ['boardState', json],
      ['updatedAt', updatedAt]
    ]);
    sheet.setFrozenRows(1);

    return {
      ok: true,
      updatedAt: updatedAt
    };
  } finally {
    lock.releaseLock();
  }
}

/* =========================================================
 * 07. 解析 POST Body
 * ---------------------------------------------------------
 * 前端用 text/plain 傳 JSON，這裡統一解析並檢查格式。
 * ========================================================= */
function parseRequestBody_(e) {
  const text = e && e.postData ? String(e.postData.contents || '') : '';

  if (!text) {
    throw new Error('POST Body 為空');
  }

  try {
    return JSON.parse(text);
  } catch (error) {
    throw new Error('POST JSON 格式錯誤：' + error.message);
  }
}

/* =========================================================
 * 08. 統一 JSON 回應
 * ---------------------------------------------------------
 * Apps Script Web App 只負責 API，不輸出任何 HTML。
 * ========================================================= */
function jsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
