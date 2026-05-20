/**
 * 雙蓮國小社團報名系統 — Google Apps Script 後端
 *
 * 部署方式（請見下方說明）：
 *   1. 開啟 https://script.google.com，建立新專案
 *   2. 將此檔案全部內容貼入 Code.gs
 *   3. 點選「部署 > 新增部署」
 *      - 類型：網頁應用程式
 *      - 執行身分：我（Me）
 *      - 誰可以存取：所有人（Anyone）
 *   4. 複製「網頁應用程式 URL」貼回 index.html 的 API_URL 常數
 *
 * 資料會存在與此 Apps Script 同一個 Google 試算表的「報名資料」工作表。
 */

// ── 設定 ───────────────────────────────────────────────
const SHEET_NAME  = "報名資料";
const COLUMNS     = [
  "id", "category", "clubId", "clubName",
  "classNo", "seatNo", "studentName",
  "contactName", "contactPhone", "email",
  "totalFee", "agreedTerms", "timestamp", "status"
];

// ── 工具函式 ────────────────────────────────────────────
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(COLUMNS);          // 第一列標題
    sheet.setFrozenRows(1);            // 凍結標題列
    // 美化標題
    const headerRange = sheet.getRange(1, 1, 1, COLUMNS.length);
    headerRange.setBackground("#1a5fa8");
    headerRange.setFontColor("#ffffff");
    headerRange.setFontWeight("bold");
  }
  return sheet;
}

function jsonOut(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];          // 只有標題，沒資料
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

// ── GET：回傳所有報名資料 ───────────────────────────────
function doGet(e) {
  try {
    const sheet = getSheet();
    const regs   = sheetToObjects(sheet);
    return jsonOut(regs);
  } catch (err) {
    return jsonOut({ ok: false, error: err.message });
  }
}

// ── POST：依 action 分派處理 ────────────────────────────
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

    if (body.action === "delete") {
      return handleDelete(body.id);
    } else if (body.action === "updateStatus") {
      return handleUpdateStatus(body.id, body.status);
    } else if (body.action === "exportPayment") {
      return handleExportPayment(body);
    } else if (body.action === "clearAll") {
      return handleClearAll();
    } else if (body.action === "bulkSave") {
      return handleBulkSave(body.regs);
    } else {
      return handleSave(body);      // 新增報名（預設）
    }
  } catch (err) {
    return jsonOut({ ok: false, error: err.message });
  }
}

// ── 新增報名 ────────────────────────────────────────────
function handleSave(reg) {
  const sheet = getSheet();

  // 檢查 id 是否已存在（防止重複送出）
  const existing = sheetToObjects(sheet);
  if (existing.some(r => r.id === reg.id)) {
    return jsonOut({ ok: false, error: "此報名編號已存在，請勿重複送出。" });
  }

  const row = COLUMNS.map(col => {
    const val = reg[col];
    if (val === undefined || val === null) return "";
    // category 強制存字串，避免數字型態造成後續篩選失敗
    if (col === "category") return String(val);
    return val;
  });
  sheet.appendRow(row);
  return jsonOut({ ok: true });
}

// ── 刪除報名 ────────────────────────────────────────────
function handleDelete(id) {
  const sheet = getSheet();
  const data   = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol   = headers.indexOf("id");

  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][idCol]) === String(id)) {
      sheet.deleteRow(i + 1);
      return jsonOut({ ok: true });
    }
  }
  return jsonOut({ ok: false, error: `找不到 id = ${id} 的報名資料。` });
}

// ── 清空所有報名資料（保留標題列）───────────────────────
function handleClearAll() {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }
  return jsonOut({ ok: true });
}

// ── 批次寫入（清空後整批重寫，供匯入備份 / 示範資料使用）──
function handleBulkSave(regs) {
  const sheet = getSheet();
  // 清空現有資料（保留標題列）
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }
  // 整批寫入
  if (regs && regs.length > 0) {
    const rows = regs.map(function(reg) {
      return COLUMNS.map(function(col) {
        const val = reg[col];
        if (val === undefined || val === null) return "";
        if (col === "category") return String(val);
        return val;
      });
    });
    sheet.getRange(2, 1, rows.length, COLUMNS.length).setValues(rows);
  }
  return jsonOut({ ok: true, count: regs ? regs.length : 0 });
}

// ── 更新繳費狀態（admin 後台用，前端可擴充呼叫） ─────────
function handleUpdateStatus(id, status) {
  const sheet   = getSheet();
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol   = headers.indexOf("id");
  const stCol   = headers.indexOf("status");

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(id)) {
      sheet.getRange(i + 1, stCol + 1).setValue(status);
      return jsonOut({ ok: true });
    }
  }
  return jsonOut({ ok: false, error: `找不到 id = ${id} 的報名資料。` });
}

// ── 匯出繳費清單（回傳 JSON 資料，由瀏覽器端 xlsx.js 產生 Excel） ─
function handleExportPayment(body) {
  const sheet    = getSheet();
  const allRegs  = sheetToObjects(sheet);
  const category = body.category || "all";
  const title    = body.title    || "社團繳費清單";

  // 篩選類別（用 String() 轉型，避免試算表把數字型 category 存成 115 而非 "115"）
  const filtered = category === "all"
    ? allRegs
    : allRegs.filter(function(r){ return String(r.category) === String(category); });

  if (filtered.length === 0) {
    return jsonOut({ ok: false, error: "此類別目前沒有報名資料。" });
  }

  // 產生建議檔名
  const filename = "雙蓮國小_" + title + "_"
    + Utilities.formatDate(new Date(), "Asia/Taipei", "yyyyMMdd") + ".xlsx";

  // 回傳純資料給瀏覽器，由前端 xlsx.js 產生 Excel（不在 GAS 建立暫存試算表）
  return jsonOut({ ok: true, regs: filtered, filename: filename });
}
