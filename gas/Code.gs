/**
 * 煎好 GOOD CRISP - 結帳訂單寫入 Google Sheet
 *
 * 部署步驟請見同資料夾的 README.md。
 * 這個檔案要貼進「綁定此試算表」的 Apps Script 專案（從 Google Sheet 的
 * 「擴充功能 → Apps Script」打開），這樣 SpreadsheetApp.getActiveSpreadsheet()
 * 才能直接抓到這張表，不用自己填 Sheet ID。
 */

// 跟 index.html 裡的 ORDER_SECRET 設成一樣的字串（留空則不驗證，任何人都可寫入，不建議）。
const SHARED_SECRET = '';

const SHEET_NAME = '訂單';

const HEADERS = ['時間', '訂單編號', '姓名', '電話', 'Email', '取餐方式', '配送地址', '備註', '品項', '件數', '金額'];

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ ok: false, error: 'empty request' });
    }
    const data = JSON.parse(e.postData.contents);

    if (SHARED_SECRET && data.secret !== SHARED_SECRET) {
      return jsonResponse({ ok: false, error: 'unauthorized' });
    }

    const sheet = getOrderSheet();
    const itemsText = (data.items || [])
      .map(function (it) { return it.name + ' x' + it.qty; })
      .join('; ');

    sheet.appendRow([
      new Date(),
      data.orderNumber || '',
      data.customerName || '',
      data.phone || '',
      data.email || '',
      fulfillmentLabel(data.fulfillment),
      data.address || '',
      data.note || '',
      itemsText,
      data.itemCount || 0,
      data.totalAmount || 0
    ]);

    return jsonResponse({ ok: true, orderNumber: data.orderNumber || '' });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

// 方便你直接用瀏覽器打開部署網址測試有沒有連通（會回傳一個小 JSON，不會寫入任何資料）。
function doGet(e) {
  return jsonResponse({ ok: true, message: '煎好訂單 API 運作中。請用 POST 送出訂單資料。' });
}

function getOrderSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function fulfillmentLabel(v) {
  const map = { takeout: '外帶', dinein: '內用', delivery: '宅配' };
  return map[v] || v || '';
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
