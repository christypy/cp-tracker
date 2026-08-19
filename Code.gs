/**
 * 划算度帳本 - Google Apps Script 後端
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
 * 品項名稱 | 類別 | 克數 | 價格 | CP值 | 日期 | 地點 | 新增時間
 *
 * 「日期」是使用者自己填的購買日期；「地點」是購買地點（例如全聯、寶雅…）；
 * 「新增時間」是系統自動寫入的紀錄時間，三者不同，都要保留。
 */

const SHEET_NAME = '工作表1'; // 依實際分頁名稱調整

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
      obj._row = i + 2; // 對應到試算表的實際列號（用於刪除）
      items.push(obj);
    });
    return jsonOut_({ status: 'ok', items });
  } catch (err) {
    return jsonOut_({ status: 'error', message: String(err) });
  }
}

// 新增 / 刪除品項
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
      const name = String(body.name || '').trim();
      const category = String(body.category || '').trim();
      const grams = Number(body.grams);
      const price = Number(body.price);
      const date = String(body.date || '').trim();
      const location = String(body.location || '').trim();

      if (!name || !category || !(grams > 0) || isNaN(price)) {
        throw new Error('欄位不完整或格式錯誤');
      }

      const cp = +(price / grams).toFixed(4);
      // 只更新前 7 欄（品項名稱～地點），第 8 欄「新增時間」維持原本紀錄不變
      sheet.getRange(row, 1, 1, 7).setValues([[name, category, grams, price, cp, date, location]]);
      return jsonOut_({ status: 'ok', cp });
    }

    // 預設為新增品項
    const name = String(body.name || '').trim();
    const category = String(body.category || '').trim();
    const grams = Number(body.grams);
    const price = Number(body.price);
    const date = String(body.date || '').trim();
    const location = String(body.location || '').trim();

    if (!name || !category || !(grams > 0) || isNaN(price)) {
      throw new Error('欄位不完整或格式錯誤');
    }

    const cp = +(price / grams).toFixed(4);
    sheet.appendRow([name, category, grams, price, cp, date, location, new Date()]);

    return jsonOut_({ status: 'ok', cp });
  } catch (err) {
    return jsonOut_({ status: 'error', message: String(err) });
  }
}