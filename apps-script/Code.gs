function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const action = String((e && e.parameter && e.parameter.action) || 'list').toLowerCase();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = sheet.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const row = {};
    headers.forEach((h, idx) => {
      row[String(h)] = values[i][idx];
    });
    rows.push(row);
  }

  if (action === 'check') {
    const shop = String((e.parameter && e.parameter.shop) || '').trim().toLowerCase();
    const found = rows.filter((r) => String(r['Shop name'] || '').trim().toLowerCase() === shop);
    return json_({ rows: found });
  }

  return json_({ rows: rows });
}

function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
  const payload = JSON.parse((e.postData && e.postData.contents) || '{}');
  const data = payload.data || payload;
  return json_(upsertRow_(sheet, headers, data));
}

function upsertRow_(sheet, headers, data) {
  const values = sheet.getDataRange().getValues();
  const shopName = String(data['Shop name'] || '').trim();
  const mobile = String(data['Mobile Number'] || '').trim();
  let rowIndex = -1;
  const shopIdx = headers.indexOf('Shop name');
  const mobileIdx = headers.indexOf('Mobile Number');
  for (let i = 1; i < values.length; i++) {
    const rowShop = shopIdx >= 0 ? String(values[i][shopIdx] || '').trim() : '';
    const rowMobile = mobileIdx >= 0 ? String(values[i][mobileIdx] || '').trim() : '';
    if ((shopName && rowShop === shopName) || (mobile && rowMobile === mobile)) {
      rowIndex = i + 1;
      break;
    }
  }
  const line = headers.map((h, idx) => {
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

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
