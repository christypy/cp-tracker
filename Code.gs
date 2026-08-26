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
 */

const SHEET_NAME = '工作表1'; // 依實際分頁名稱調整

// 欄位固定順序（對應試算表由左到右的欄）
const FIELD_ORDER = ['品項名稱', '類別', '子類別', '克數', '價格', '數量', '單罐價格', 'CP值', '日期', '地點'];

function getSheet_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error('找不到分頁「' + SHEET_NAME + '」，請確認 SHEET_NAME 設定是否正確');
  }
  return sheet;
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
      headers.forEach((h, idx) => { obj[h] = row[idx]; });
      obj._row = i + 2; // 對應到試算表的實際列號（用於刪除／編輯）
      items.push(obj);
    });
    return jsonOut_({ status: 'ok', items });
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
