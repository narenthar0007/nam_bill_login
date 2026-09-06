/**
 * NAM bill — Google Sheet write API
 *
 * HOW TO DEPLOY (important — fixes "fetch failed"):
 * 1. In the spreadsheet: Extensions → Apps Script
 * 2. Delete old code, paste THIS entire file, Save (Ctrl+S)
 * 3. Deploy → New deployment → Type: Web app
 * 4. Description: nam-bill
 * 5. Execute as: Me
 * 6. Who has access: Anyone   ← MUST be Anyone (not "Only myself")
 * 7. Deploy → Authorize → copy the Web app URL (.../exec)
 * 8. Paste that URL in admin Settings → Google Apps Script URL → Save
 *
 * If you change code later: Deploy → Manage deployments → Edit (pencil)
 * → Version: New version → Deploy
 *
 * Sheet:
 * https://docs.google.com/spreadsheets/d/1Gr7vHrtQZ_wRrFsv3y_mCJ6qdNeSB7x_TkQhUCY3ato/edit?gid=0#gid=0
 *
 * Headers (row 1):
 * Name, Mobile Number, Email ID, Date of payment, Date of login, Expire date,
 * Payment mode, Amount, Login key, Shop name, User Code, Is Login,
 * Status, Login session, Remarks
 */

var REQUIRED_HEADERS = [
  'Name',
  'Mobile Number',
  'Email ID',
  'Date of payment',
  'Date of login',
  'Expire date',
  'Payment mode',
  'Amount',
  'Login key',
  'Shop name',
  'User Code',
  'Is Login',
  'Status',
  'Login session',
  'Remarks',
];

function ensureHeaders_(sheet) {
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {
    return String(h || '').trim();
  });
  var changed = false;
  REQUIRED_HEADERS.forEach(function (h) {
    if (headers.indexOf(h) === -1) {
      headers.push(h);
      changed = true;
    }
  });
  if (changed) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return headers;
}

function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var headers = ensureHeaders_(sheet);
  var action = String((e && e.parameter && e.parameter.action) || 'list').toLowerCase();
  var values = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var row = {};
    headers.forEach(function (h, idx) {
      row[String(h)] = values[i][idx];
    });
    rows.push(row);
  }

  if (action === 'check') {
    var shop = String((e.parameter && e.parameter.shop) || '').trim().toLowerCase();
    var key = String((e.parameter && e.parameter.key) || '').trim().toUpperCase();
    var found = rows.filter(function (r) {
      var rowShop = String(r['Shop name'] || '').trim().toLowerCase();
      var rowKey = String(r['Login key'] || r['Login Key'] || '').trim().toUpperCase();
      if (shop && key) return rowShop === shop && rowKey === key;
      if (shop) return rowShop === shop;
      if (key) return rowKey === key;
      return false;
    });
    return json_({ rows: found, ok: true });
  }

  return json_({ rows: rows, ok: true });
}

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    var headers = ensureHeaders_(sheet);
    var raw = (e && e.postData && e.postData.contents) || '{}';
    var payload = JSON.parse(raw);
    var data = payload.data || payload;
    var action = String(payload.action || '').toLowerCase();
    if (action === 'loginstate') {
      return json_(patchLoginState_(sheet, headers, data));
    }
    return json_(upsertRow_(sheet, headers, data));
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function upsertRow_(sheet, headers, data) {
  var values = sheet.getDataRange().getValues();
  var shopName = String(data['Shop name'] || '').trim();
  var mobile = String(data['Mobile Number'] || '').trim();
  var loginKey = String(data['Login key'] || data['Login Key'] || '').trim().toUpperCase();
  var rowIndex = -1;
  var shopIdx = headers.indexOf('Shop name');
  var mobileIdx = headers.indexOf('Mobile Number');
  var keyIdx = Math.max(headers.indexOf('Login key'), headers.indexOf('Login Key'));
  for (var i = 1; i < values.length; i++) {
    var rowShop = shopIdx >= 0 ? String(values[i][shopIdx] || '').trim() : '';
    var rowMobile = mobileIdx >= 0 ? String(values[i][mobileIdx] || '').trim() : '';
    var rowKey = keyIdx >= 0 ? String(values[i][keyIdx] || '').trim().toUpperCase() : '';
    if ((shopName && rowShop === shopName) || (mobile && rowMobile === mobile) || (loginKey && rowKey === loginKey)) {
      rowIndex = i + 1;
      break;
    }
  }
  var line = headers.map(function (h, idx) {
    if (Object.prototype.hasOwnProperty.call(data, h)) return data[h];
    return rowIndex > 0 ? values[rowIndex - 1][idx] : '';
  });
  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 1, 1, headers.length).setValues([line]);
  } else {
    sheet.appendRow(line);
  }
  return { ok: true };
}

function patchLoginState_(sheet, headers, data) {
  var values = sheet.getDataRange().getValues();
  var shopName = String(data['Shop name'] || '').trim().toLowerCase();
  var loginKey = String(data['Login key'] || data['Login Key'] || '').trim().toUpperCase();
  var shopIdx = headers.indexOf('Shop name');
  var keyIdx = Math.max(headers.indexOf('Login key'), headers.indexOf('Login Key'));
  var rowIndex = -1;
  for (var i = 1; i < values.length; i++) {
    var rowShop = shopIdx >= 0 ? String(values[i][shopIdx] || '').trim().toLowerCase() : '';
    var rowKey = keyIdx >= 0 ? String(values[i][keyIdx] || '').trim().toUpperCase() : '';
    if (
      (shopName && loginKey && rowShop === shopName && rowKey === loginKey) ||
      (shopName && !loginKey && rowShop === shopName) ||
      (loginKey && !shopName && rowKey === loginKey)
    ) {
      rowIndex = i + 1;
      break;
    }
  }
  if (rowIndex < 0) return { ok: false, error: 'Row not found' };
  var line = values[rowIndex - 1].slice();
  while (line.length < headers.length) line.push('');
  headers.forEach(function (h, idx) {
    if (Object.prototype.hasOwnProperty.call(data, h)) line[idx] = data[h];
  });
  sheet.getRange(rowIndex, 1, 1, headers.length).setValues([line]);
  return { ok: true };
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
