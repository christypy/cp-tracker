/**
 * 劃算度帳本 - Google Apps Script 後端
 *
 * 使用方式：
 * 1. 打開你的 Google 試算表
 * 2. 上方選單「擴充功能」→「Apps Script」
 * 3. 把這個檔案的內容整個貼進去（取代預設的 myFunction 內容）
 * 4. 把下面 SHEET_NAME 改成你試算表分頁的名稱（分頁下方的標籤名稱）
 * 5. 儲存後，點右上角「部署」→「新增部署作業」
 *    - 類型選「網頁應用程式」
 *    - 「執行身份」選「我」
 *    - 「誰可以存取」選「所有人」
 * 6. 部署後會拿到一組網址，例如：
 *    https://script.google.com/macros/s/xxxxxxxxxxxx/exec
 *    把這組網址貼到網站的「設定同步網址」欄位
 *
 * 試算表第一列（標題列）請依序填入：
 * 品項名稱 | 類別 | 子類別 | 克數 | 價格 | 數量 | 單罐價格 | CP值 | 日期 | 地點 | 新增時間
 *
 * 「類別」是主類別（必填，例如「零食」）；「子類別」是選填的細分類，可以複選
 * （例如「甜的,鹹的」，多個子類別用半形逗號分隔），沒有子類別時這一格留空即可。
 * 「日期」是使用者自己填的購買日期；「地點」是購買地點（例如全聯、寶雅…）；
 * 「新增時間」是系統自動寫入的紀錄時間。
 *
 * ⚠ 如果你是從舊版（沒有「子類別」欄位）升級上來：
 *    請在試算表裡「類別」欄位右邊手動插入一個新欄，標題列填「子類別」，
 *    這樣舊資料的「地點」「新增時間」等欄位才不會被錯位覆蓋。
 *    舊資料本來就沒有子類別，留空即可，之後編輯該筆紀錄時再補上就會存進新欄位。
 *
 * ℹ 這個版本加了「比價紀錄」功能：在比價彈窗按「記錄這次比價」，不管划不划算都會存一筆進
 *    一個叫「比價紀錄」的新分頁（跟「類別排序」一樣，第一次用到時會自動建立，不用手動新增），
 *    下次打開同一個品項的比價視窗就能看到之前查過哪些地點／價格，不用重複比價。
 *
 * ℹ 這個版本也把「手動先建立、還沒有任何紀錄的空類別」改成會同步存進一個叫「手動類別」的
 *    新分頁（一樣第一次用到時自動建立，不用手動新增）。這樣在網頁上手動新增的主／子類別
 *    就會跟其他裝置、其他瀏覽器保持一致，不會有「網頁上看得到、試算表裡找不到」的落差。
 */

const SHEET_NAME = '工作表1'; // 依實際分頁名稱調整

// 存放「主類別／子類別排序」的分頁，會在第一次用到時自動建立，不需要手動新增
const CATEGORY_ORDER_SHEET_NAME = '類別排序';

// ⚠ 這個版本加了「一鍵修復」功能：如果你直接在試算表裡手動改「子類別」欄位、或手動刪除／
//    搬移資料列，「類別排序」「比價紀錄」這兩個輔助分頁不會自動知道（因為手動編輯不會觸發
//    Apps Script），久了就會跟主資料兜不起來。打開這份試算表後，選單列會多一個「🧹 帳本維護」，
//    點裡面的「修復子類別排序／比價紀錄同步」就會：
//    1. 清掉「類別排序」分頁裡已經不存在於目前資料的主／子類別殘留紀錄
//    2. 依「品項名稱」重新比對「比價紀錄」分頁的「對應列號」（只有名稱目前唯一對應到一列時才會更新，
//       避免同名品項比對錯筆；完全找不到對應品項的紀錄會保留原樣，不會被清空或亂改）
//    這個動作不會刪除或更動品項本身的資料，也不會動到「還沒有任何紀錄的空子類別」
//    （那些只存在瀏覽器本機，不在試算表裡，這裡管不到也不會影響它們）。

// 存放「手動類別」的分頁：使用者在網頁上「還沒有任何紀錄也可以先建立」的主／子類別清單，
// 會在第一次用到時自動建立。這樣手動先建立的空類別也會寫進試算表，不同裝置、清瀏覽器資料
// 都不會跑掉，也不會有「網頁上有、試算表沒有」的落差。
const MANUAL_CATEGORY_SHEET_NAME = '手動類別';

// 取得（或自動建立）存放手動類別的分頁
// 欄位：主類別 | 子類別（子類別留空代表「這個主類別本身還沒有任何子類別」的存在紀錄）
function getManualCategorySheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(MANUAL_CATEGORY_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(MANUAL_CATEGORY_SHEET_NAME);
    sheet.appendRow(['主類別', '子類別']);
  }
  return sheet;
}

// 讀出目前存在試算表裡的手動類別，整理成 { 主類別: [子類別, ...] }
function readManualCategories_() {
  const sheet = getManualCategorySheet_();
  const values = sheet.getDataRange().getValues();
  const result = {};
  if (values.length < 2) return result;

  values.shift(); // 標題列
  values.forEach(row => {
    const main = String(row[0] || '').trim();
    const sub = String(row[1] || '').trim();
    if (!main) return;
    if (!result[main]) result[main] = [];
    if (sub && !result[main].includes(sub)) result[main].push(sub);
  });
  return result;
}

// 把前端傳來的完整手動類別清單（{ 主類別: [子類別, ...] }）整批覆寫回試算表
function writeManualCategories_(categories) {
  const sheet = getManualCategorySheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 2).clearContent();
  }

  const rows = [];
  const map = (categories && typeof categories === 'object') ? categories : {};
  Object.keys(map).forEach(main => {
    const m = String(main || '').trim();
    if (!m) return;
    const subs = Array.isArray(map[main]) ? map[main] : [];
    const cleanSubs = subs.map(s => String(s).trim()).filter(Boolean);
    if (cleanSubs.length === 0) {
      // 還沒有任何子類別的主類別，也要存一列，才不會存完又消失
      rows.push([m, '']);
    } else {
      cleanSubs.forEach(s => rows.push([m, s]));
    }
  });

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  }
}

// 存放「比價紀錄」的分頁，同樣會在第一次用到時自動建立。
// 不管這次查價划不划算都可以存一筆進來，用來記住「這個品項在哪個地點查過多少錢、結果如何」，
// 避免下次逛街又重新比一次同一間店。跟品項本身的正式紀錄是分開的兩張表。
const COMPARISON_LOG_SHEET_NAME = '比價紀錄';
const COMPARISON_LOG_FIELDS = ['時間', '品項名稱', '對應列號', '容量', '數量', '售價', '換算CP值', '基準CP值', '結果', '地點'];

// 欄位固定順序（對應試算表由左到右的欄）
const FIELD_ORDER = ['品項名稱', '類別', '子類別', '克數', '價格', '數量', '單罐價格', 'CP值', '日期', '地點'];

function getSheet_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error('找不到分頁「' + SHEET_NAME + '」，請確認 SHEET_NAME 設定是否正確');
  }
  return sheet;
}

// 取得（或自動建立）存放類別排序的分頁
// 欄位：層級（主／子）| 主類別 | 子類別（主類別這一列留空）| 排序（數字，愈小排愈前面）
function getCategoryOrderSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CATEGORY_ORDER_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CATEGORY_ORDER_SHEET_NAME);
    sheet.appendRow(['層級', '主類別', '子類別', '排序']);
  }
  return sheet;
}

// 讀出目前存在試算表裡的類別排序，整理成 { main: [主類別...], sub: { 主類別: [子類別...] } }
function readCategoryOrder_() {
  const sheet = getCategoryOrderSheet_();
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return { main: [], sub: {} };

  values.shift(); // 標題列
  const mainEntries = []; // [{name, order}]
  const subEntries = {}; // { 主類別: [{name, order}] }

  values.forEach(row => {
    const level = String(row[0] || '').trim();
    const main = String(row[1] || '').trim();
    const sub = String(row[2] || '').trim();
    const order = Number(row[3]);
    if (!main) return;
    if (level === '子' && sub) {
      if (!subEntries[main]) subEntries[main] = [];
      subEntries[main].push({ name: sub, order: isNaN(order) ? 0 : order });
    } else if (!sub) {
      mainEntries.push({ name: main, order: isNaN(order) ? 0 : order });
    }
  });

  const main = mainEntries.sort((a, b) => a.order - b.order).map(e => e.name);
  const sub = {};
  Object.keys(subEntries).forEach(parent => {
    sub[parent] = subEntries[parent].sort((a, b) => a.order - b.order).map(e => e.name);
  });
  return { main, sub };
}

// 把前端傳來的完整排序（{ main: [...], sub: { 主類別: [...] } }）整批覆寫回試算表
function writeCategoryOrder_(order) {
  const sheet = getCategoryOrderSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 4).clearContent();
  }

  const rows = [];
  const mainList = Array.isArray(order && order.main) ? order.main : [];
  mainList.forEach((name, idx) => {
    rows.push(['主', String(name), '', idx]);
  });

  const subMap = (order && order.sub && typeof order.sub === 'object') ? order.sub : {};
  Object.keys(subMap).forEach(parent => {
    const subList = Array.isArray(subMap[parent]) ? subMap[parent] : [];
    subList.forEach((name, idx) => {
      rows.push(['子', String(parent), String(name), idx]);
    });
  });

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 4).setValues(rows);
  }
}

// 取得（或自動建立）存放比價紀錄的分頁
function getComparisonLogSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(COMPARISON_LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(COMPARISON_LOG_SHEET_NAME);
    sheet.appendRow(COMPARISON_LOG_FIELDS);
  }
  return sheet;
}

// 讀出所有比價紀錄，給前端在比價視窗顯示「之前查過哪些地點／價格」
function readComparisonLog_() {
  const sheet = getComparisonLogSheet_();
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values.shift();
  const logs = [];
  values.forEach((row, i) => {
    if (!row[1]) return; // 沒有品項名稱的空白列跳過
    const obj = {};
    headers.forEach((h, idx) => {
      let val = row[idx];
      if (h === '時間' && val instanceof Date) {
        val = Utilities.formatDate(val, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
      }
      obj[h] = val;
    });
    obj._row = i + 2; // 對應到試算表的實際列號（用於刪除）
    logs.push(obj);
  });
  return logs;
}

// 新增一筆比價紀錄。不管這次划不划算都可以呼叫，單純記錄「查過」這件事。
function appendComparisonLog_(body) {
  const name = String(body.name || '').trim();
  if (!name) throw new Error('缺少品項名稱，無法記錄比價');

  const sheet = getComparisonLogSheet_();
  const rowValues = [
    new Date(),
    name,
    (body.row !== undefined && body.row !== null && body.row !== '' && !isNaN(Number(body.row))) ? Number(body.row) : '',
    Number(body.grams) || '',
    Number(body.count) || 1,
    Number(body.price) || '',
    Number(body.cp) || '',
    Number(body.baseCp) || '',
    String(body.verdict || '').trim(),
    String(body.location || '').trim()
  ];
  sheet.appendRow(rowValues);
  return sheet.getLastRow();
}

// ---------- 一鍵修復：讓「類別排序」「比價紀錄」跟主資料（工作表1）重新對齊 ----------
// 適用情境：直接在試算表手動改子類別文字、或手動刪除／搬移資料列之後，
// 這兩個輔助分頁不會自動更新，跑這個函式把它們「照目前的主資料」重新校正一次。
function repairAuxSheets_() {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  const result = { removedOrderRows: 0, fixedLogRows: 0, unresolvedLogRows: 0 };
  if (values.length < 2) return result;

  const headers = values.shift();
  const nameIdx = headers.indexOf('品項名稱');
  const catIdx = headers.indexOf('類別');
  const subIdx = headers.indexOf('子類別');
  if (nameIdx === -1 || catIdx === -1 || subIdx === -1) {
    throw new Error('工作表1的標題列缺少「品項名稱」「類別」或「子類別」欄位，請確認欄位標題');
  }

  // 依目前實際資料，整理出「主類別 -> 子類別集合」，以及「品項名稱 -> 目前所在列號」
  const catMap = {}; // { 主類別: Set(子類別) }
  const nameToRows = {}; // { 品項名稱: [列號, ...] }
  values.forEach((row, i) => {
    const name = String(row[nameIdx] || '').trim();
    const cat = String(row[catIdx] || '').trim();
    const subRaw = String(row[subIdx] || '').trim();
    if (!name || !cat) return;
    if (!catMap[cat]) catMap[cat] = new Set();
    subRaw.split(',').map(s => s.trim()).filter(Boolean).forEach(s => catMap[cat].add(s));
    if (!nameToRows[name]) nameToRows[name] = [];
    nameToRows[name].push(i + 2);
  });

  // --- 修復「類別排序」：只保留目前資料裡還存在的主／子類別，捨棄找不到對應資料的舊排序紀錄 ---
  const orderSheet = getCategoryOrderSheet_();
  const orderLastRow = orderSheet.getLastRow();
  if (orderLastRow > 1) {
    const orderValues = orderSheet.getRange(2, 1, orderLastRow - 1, 4).getValues();
    const keptRows = [];
    orderValues.forEach(row => {
      const level = String(row[0] || '').trim();
      const main = String(row[1] || '').trim();
      const sub = String(row[2] || '').trim();
      if (!main) return;
      const stillValid = (level === '子')
        ? !!(catMap[main] && catMap[main].has(sub))
        : !!catMap[main];
      if (stillValid) {
        keptRows.push(row);
      } else {
        result.removedOrderRows++;
      }
    });
    orderSheet.getRange(2, 1, orderLastRow - 1, 4).clearContent();
    if (keptRows.length > 0) {
      orderSheet.getRange(2, 1, keptRows.length, 4).setValues(keptRows);
    }
  }

  // --- 修復「比價紀錄」的對應列號：依品項名稱重新對到目前實際的列號 ---
  // 只有「這個名稱目前唯一對應到一列」時才更新，避免同名品項比對到錯的一筆；
  // 完全找不到對應品項（可能改名或刪除過）的紀錄則保留原狀，不清空、不亂猜。
  const logSheet = getComparisonLogSheet_();
  const logLastRow = logSheet.getLastRow();
  if (logLastRow > 1) {
    const logHeaders = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0];
    const nameIdxLog = logHeaders.indexOf('品項名稱');
    const rowIdxLog = logHeaders.indexOf('對應列號');
    if (nameIdxLog !== -1 && rowIdxLog !== -1) {
      const logValues = logSheet.getRange(2, 1, logLastRow - 1, logSheet.getLastColumn()).getValues();
      logValues.forEach((row, i) => {
        const name = String(row[nameIdxLog] || '').trim();
        if (!name) return;
        const matches = nameToRows[name];
        if (matches && matches.length === 1) {
          if (row[rowIdxLog] !== matches[0]) {
            logSheet.getRange(i + 2, rowIdxLog + 1).setValue(matches[0]);
            result.fixedLogRows++;
          }
        } else {
          result.unresolvedLogRows++;
        }
      });
    }
  }

  return result;
}

// 試算表選單：打開試算表時自動加一個「🧹 帳本維護」選單，不用進 Apps Script 編輯器也能跑修復
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🧹 帳本維護')
    .addItem('修復子類別排序／比價紀錄同步', 'runRepairAuxSheets')
    .addToUi();
}

// 給選單按鈕呼叫：跑修復，並跳出結果摘要
function runRepairAuxSheets() {
  try {
    const r = repairAuxSheets_();
    let msg = '✅ 修復完成\n\n';
    msg += '「類別排序」分頁：清掉 ' + r.removedOrderRows + ' 筆已經不存在的舊主／子類別排序紀錄\n';
    msg += '「比價紀錄」分頁：修正 ' + r.fixedLogRows + ' 筆對應列號';
    if (r.unresolvedLogRows > 0) {
      msg += '\n（另有 ' + r.unresolvedLogRows + ' 筆因品項可能已改名或刪除，找不到目前對應的列，暫時保留原狀）';
    }
    SpreadsheetApp.getUi().alert(msg);
  } catch (err) {
    SpreadsheetApp.getUi().alert('❌ 修復失敗：' + String(err));
  }
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// 讀取所有品項
function doGet(e) {
  try {
    const sheet = getSheet_();
    const range = sheet.getDataRange();
    const values = range.getValues();
    if (values.length < 1) return jsonOut_({ status: 'ok', items: [] });

    const headers = values.shift();
    const items = [];
    values.forEach((row, i) => {
      if (!row[0]) return; // 跳過空白列
      const obj = {};
      headers.forEach((h, idx) => {
        let val = row[idx];
        // 「日期」欄位如果被試算表存成日期格式，讀出來會是 Date 物件，
        // 直接塞進 JSON 會被轉成完整的日期時間字串（帶時區），這裡統一只留日期部分，
        // 並用試算表的時區換算，避免因為時區問題差了一天。
        if (h === '日期' && val instanceof Date) {
          val = Utilities.formatDate(val, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
        }
        obj[h] = val;
      });
      obj._row = i + 2; // 對應到試算表的實際列號（用於刪除／編輯）
      items.push(obj);
    });
    const categoryOrder = readCategoryOrder_();
    const comparisonLog = readComparisonLog_();
    const manualCategories = readManualCategories_();
    return jsonOut_({ status: 'ok', items, categoryOrder, comparisonLog, manualCategories });
  } catch (err) {
    return jsonOut_({ status: 'error', message: String(err) });
  }
}

// 從 request body 整理出一筆品項的欄位，並做基本驗證
function parseItemFields_(body) {
  const name = String(body.name || '').trim();
  const category = String(body.category || '').trim(); // 主類別（必填）
  // 子類別（選填，可複選）：前端可能傳陣列（多選）或字串，統一轉成用逗號分隔的字串存進試算表
  const subCategory = Array.isArray(body.subCategory)
    ? body.subCategory.map(s => String(s).trim()).filter(Boolean).join(',')
    : String(body.subCategory || '').trim();
  const grams = Number(body.grams);
  const price = Number(body.price);
  const count = Number(body.count) || 1;
  const date = String(body.date || '').trim();
  const location = String(body.location || '').trim();

  if (!name || !category || !(grams > 0) || isNaN(price) || count < 1) {
    throw new Error('欄位不完整或格式錯誤（品項名稱／主類別／克數／價格為必填，子類別可留空）');
  }

  const unitPrice = +(price / count).toFixed(2);
  const cp = +(price / (grams * count)).toFixed(4);

  return { name, category, subCategory, grams, price, count, unitPrice, cp, date, location };
}

// 新增 / 刪除 / 編輯品項
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

    // 儲存主／子類別的自訂排序（跟品項資料是不同分頁，這裡直接處理、不用碰到品項工作表）
    if (body.action === 'saveCategoryOrder') {
      writeCategoryOrder_(body.order || {});
      return jsonOut_({ status: 'ok' });
    }

    // 一鍵修復：讓「類別排序」「比價紀錄」重新對齊目前的主資料（跟選單裡的按鈕是同一個函式）
    if (body.action === 'repairAuxSheets') {
      const result = repairAuxSheets_();
      return jsonOut_({ status: 'ok', result });
    }

    // 儲存手動建立的主／子類別清單（跟品項資料是不同分頁，這裡直接處理、不用碰到品項工作表）
    if (body.action === 'saveManualCategories') {
      writeManualCategories_(body.categories || {});
      return jsonOut_({ status: 'ok' });
    }

    // 記錄一筆比價結果（跟品項資料是不同分頁，不管這次划不划算都可以記錄，方便下次不用重比）
    if (body.action === 'logComparison') {
      const row = appendComparisonLog_(body);
      return jsonOut_({ status: 'ok', row });
    }

    // 刪除一筆比價紀錄（例如記錯了，或想清掉太舊的查價紀錄）
    if (body.action === 'deleteComparisonLog') {
      const row = Number(body.row);
      if (row < 2) throw new Error('無效的列號');
      getComparisonLogSheet_().deleteRow(row);
      return jsonOut_({ status: 'ok' });
    }

    const sheet = getSheet_();

    if (body.action === 'delete') {
      const row = Number(body.row);
      if (row < 2) throw new Error('無效的列號');
      sheet.deleteRow(row);
      return jsonOut_({ status: 'ok' });
    }

    if (body.action === 'update') {
      const row = Number(body.row);
      if (row < 2) throw new Error('無效的列號');
      const f = parseItemFields_(body);

      // 更新前 10 欄（品項名稱～地點），第 11 欄「新增時間」維護原本紀錄不變
      sheet.getRange(row, 1, 1, 10).setValues([[
        f.name, f.category, f.subCategory, f.grams, f.price, f.count, f.unitPrice, f.cp, f.date, f.location
      ]]);
      return jsonOut_({ status: 'ok', row, cp: f.cp, unitPrice: f.unitPrice });
    }

    // 預設為新增品項
    const f = parseItemFields_(body);
    sheet.appendRow([
      f.name, f.category, f.subCategory, f.grams, f.price, f.count, f.unitPrice, f.cp, f.date, f.location, new Date()
    ]);

    // 直接回傳這筆新資料實際寫入的列號，前端就不需要再整張表重新讀取一次
    const row = sheet.getLastRow();
    return jsonOut_({ status: 'ok', row, cp: f.cp, unitPrice: f.unitPrice });
  } catch (err) {
    return jsonOut_({ status: 'error', message: String(err) });
  }
}
