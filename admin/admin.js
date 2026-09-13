(() => {
  const Secure = window.NamSecureCrypto;
  const AUTH_USER_HASH = '8c6976e5b5410415bde908bd4dee15dfb167a9c873fc4bb8a81f6f2ab448a918';
  const AUTH_PASS_HASH = 'ac9689e2272427085e35b9d3e3e8bed88cb3434828b43b86fc0596cad4c6e270';
  const AUTH_SESSION = Secure.sha256Hex(`${Secure.DATA_ENCRYPTION_SECRET}:admin-session`);
  const LICENSE_SECRET = 'RIYE_POS_SECRET_2026';
  const APP_LICENSE_ID = 'com.riye.posbilling';
  const CONTROL_SHOP = '__NAM_BILL_CONTROL__';
  const DEFAULT_SHEET_ID = '1Gr7vHrtQZ_wRrFsv3y_mCJ6qdNeSB7x_TkQhUCY3ato';
  const DEFAULT_SHEET_GID = '0';
  const DEFAULT_SHEET_URL = `https://docs.google.com/spreadsheets/d/${DEFAULT_SHEET_ID}/edit?gid=${DEFAULT_SHEET_GID}#gid=${DEFAULT_SHEET_GID}`;
  const DEFAULT_APPS_SCRIPT_URL =
    'https://script.google.com/macros/s/AKfycbxjLrd4XB11GwxmntVLpOAWPENBDmmaq9L01OH7ek28yd-sbM8eP1-asZ3PdcgDhQuQeQ/exec';
  const ISSUE_LEN = 4;
  const EXPIRY_LEN = 4;
  const SIG_LEN = 12;
  const CHARS = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  const ALPHA_NUM = CHARS;
  const DIGITS = '0123456789';
  const RENEWAL_LEN = 10;
  const RENEWAL_MS = 45 * 60 * 1000;

  // Exact header order matching the live Google Sheet, plus Status / Login session / Remarks.
  const SHEET_HEADER_ORDER = [
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

  const COLS = {
    name: 'Name',
    mobile: 'Mobile Number',
    email: 'Email ID',
    paymentDate: 'Date of payment',
    loginDate: 'Date of login',
    expireDate: 'Expire date',
    paymentMode: 'Payment mode',
    amount: 'Amount',
    loginKey: 'Login key',
    shop: 'Shop name',
    userCode: 'User Code',
    isLogin: 'Is Login',
    status: 'Status',
    loginSession: 'Login session',
    remarks: 'Remarks',
  };

  const APPS_SCRIPT = "/**\n * NAM bill — Google Sheet write API\n *\n * Deploy as WEB APP (not Library):\n * Deploy → New deployment → Web app\n * Execute as: Me\n * Who has access: Anyone\n * Copy URL ending with /exec\n *\n * Upsert rules:\n * 1) Same Shop name (case-insensitive) → update that row\n * 2) Else → append a NEW row\n * Never match by Login key alone — keys repeat for the same day/period.\n */\n\nvar REQUIRED_HEADERS = [\n  'Name',\n  'Mobile Number',\n  'Email ID',\n  'Date of payment',\n  'Date of login',\n  'Expire date',\n  'Payment mode',\n  'Amount',\n  'Login key',\n  'Shop name',\n  'User Code',\n  'Is Login',\n  'Status',\n  'Login session',\n  'Remarks',\n];\n\nfunction headerIndex_(headers, names) {\n  for (var i = 0; i < names.length; i++) {\n    var idx = headers.indexOf(names[i]);\n    if (idx >= 0) return idx;\n  }\n  return -1;\n}\n\nfunction ensureHeaders_(sheet) {\n  var lastCol = Math.max(sheet.getLastColumn(), 1);\n  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {\n    return String(h || '').trim();\n  });\n  var changed = false;\n  REQUIRED_HEADERS.forEach(function (h) {\n    if (headers.indexOf(h) === -1) {\n      headers.push(h);\n      changed = true;\n    }\n  });\n  if (changed) {\n    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);\n  }\n  return headers;\n}\n\nfunction doGet(e) {\n  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];\n  var headers = ensureHeaders_(sheet);\n  var action = String((e && e.parameter && e.parameter.action) || 'list').toLowerCase();\n  var values = sheet.getDataRange().getValues();\n  var rows = [];\n  for (var i = 1; i < values.length; i++) {\n    var row = {};\n    headers.forEach(function (h, idx) {\n      row[String(h)] = values[i][idx];\n    });\n    // Skip fully empty rows\n    var hasData = headers.some(function (h, idx) {\n      return String(values[i][idx] == null ? '' : values[i][idx]).trim() !== '';\n    });\n    if (hasData) rows.push(row);\n  }\n\n  if (action === 'check') {\n    var shop = String((e.parameter && e.parameter.shop) || '').trim().toLowerCase();\n    var key = String((e.parameter && e.parameter.key) || '').trim().toUpperCase();\n    var found = rows.filter(function (r) {\n      var rowShop = String(r['Shop name'] || '').trim().toLowerCase();\n      var rowKey = String(r['Login key'] || r['Login Key'] || '').trim().toUpperCase();\n      if (shop && key) return rowShop === shop && rowKey === key;\n      if (shop) return rowShop === shop;\n      if (key) return rowKey === key;\n      return false;\n    });\n    return json_({ rows: found, ok: true });\n  }\n\n  return json_({ rows: rows, ok: true });\n}\n\nfunction doPost(e) {\n  try {\n    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];\n    var headers = ensureHeaders_(sheet);\n    var raw = (e && e.postData && e.postData.contents) || '{}';\n    var payload = JSON.parse(raw);\n    var data = payload.data || payload;\n    var action = String(payload.action || '').toLowerCase();\n    if (action === 'loginstate') {\n      return json_(patchLoginState_(sheet, headers, data));\n    }\n    return json_(upsertRow_(sheet, headers, data, action === 'create'));\n  } catch (err) {\n    return json_({ ok: false, error: String(err && err.message ? err.message : err) });\n  }\n}\n\nfunction findRowIndex_(values, headers, data, forceCreate) {\n  if (forceCreate) return -1;\n\n  // IMPORTANT: License keys are derived from issue/expiry day only, so many\n  // shops can share the same key on the same day. Never match by Login key alone\n  // or new shops overwrite an existing row.\n  var shopName = String(data['Shop name'] || '').trim().toLowerCase();\n  var shopIdx = headerIndex_(headers, ['Shop name']);\n  if (!shopName || shopIdx < 0) return -1;\n\n  for (var i = 1; i < values.length; i++) {\n    var rowShop = String(values[i][shopIdx] || '').trim().toLowerCase();\n    if (rowShop && rowShop === shopName) return i + 1;\n  }\n  return -1;\n}\n\nfunction upsertRow_(sheet, headers, data, forceCreate) {\n  var values = sheet.getDataRange().getValues();\n  var rowIndex = findRowIndex_(values, headers, data, forceCreate);\n\n  var line = headers.map(function (h, idx) {\n    if (Object.prototype.hasOwnProperty.call(data, h)) return data[h];\n    return rowIndex > 0 ? values[rowIndex - 1][idx] : '';\n  });\n\n  if (rowIndex > 0) {\n    sheet.getRange(rowIndex, 1, 1, headers.length).setValues([line]);\n    return { ok: true, updated: true, created: false, row: rowIndex };\n  }\n\n  // Explicit next-row write (more reliable than appendRow with sparse sheets)\n  var nextRow = Math.max(sheet.getLastRow() + 1, 2);\n  sheet.getRange(nextRow, 1, 1, headers.length).setValues([line]);\n  return { ok: true, updated: false, created: true, row: nextRow };\n}\n\nfunction patchLoginState_(sheet, headers, data) {\n  var values = sheet.getDataRange().getValues();\n  var shopName = String(data['Shop name'] || '').trim().toLowerCase();\n  var loginKey = String(data['Login key'] || data['Login Key'] || '').trim().toUpperCase();\n  var shopIdx = headerIndex_(headers, ['Shop name']);\n  var keyIdx = headerIndex_(headers, ['Login key', 'Login Key']);\n  var rowIndex = -1;\n\n  for (var i = 1; i < values.length; i++) {\n    var rowShop = shopIdx >= 0 ? String(values[i][shopIdx] || '').trim().toLowerCase() : '';\n    var rowKey = keyIdx >= 0 ? String(values[i][keyIdx] || '').trim().toUpperCase() : '';\n    if (shopName && loginKey && rowShop === shopName && rowKey === loginKey) {\n      rowIndex = i + 1;\n      break;\n    }\n    if (shopName && rowShop === shopName) {\n      rowIndex = i + 1;\n      break;\n    }\n    if (loginKey && rowKey === loginKey) {\n      rowIndex = i + 1;\n      break;\n    }\n  }\n\n  if (rowIndex < 0) return { ok: false, error: 'Row not found' };\n  var line = values[rowIndex - 1].slice();\n  while (line.length < headers.length) line.push('');\n  headers.forEach(function (h, idx) {\n    if (Object.prototype.hasOwnProperty.call(data, h)) line[idx] = data[h];\n  });\n  sheet.getRange(rowIndex, 1, 1, headers.length).setValues([line]);\n  return { ok: true, row: rowIndex };\n}\n\nfunction json_(obj) {\n  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);\n}\n";
  let customers = [];
  let selectedDays = 90;
  let useCustom = false;
  let existingRecord = null;

  const $ = (id) => document.getElementById(id);

  function loadSettings() {
    try {
      const raw = localStorage.getItem('namAdminSettings') || '{}';
      const json = Secure.isEncryptedValue(raw) ? Secure.decryptText(raw) : raw;
      return JSON.parse(json || '{}');
    } catch {
      return {};
    }
  }

  function saveSettings(partial) {
    const next = { ...defaultSettings(), ...normalizedSettings(loadSettings()), ...partial };
    localStorage.setItem('namAdminSettings', Secure.encryptText(JSON.stringify(next)));
    return next;
  }

  function defaultSettings() {
    return {
      googleSheetUrl: DEFAULT_SHEET_URL,
      appsScriptUrl: DEFAULT_APPS_SCRIPT_URL,
      minAppVersion: '1.0.0',
      latestAppVersion: '1.0.0',
      downloadUrl: 'https://github.com/narenthar0007/mahi/releases',
      updateMessage: 'A new version of NAM bill is available. Please download the latest app to continue.',
      forceUpdate: false,
      requireSheetRecord: true,
    };
  }

  function isValidAppsScriptUrl(url) {
    const value = String(url || '').trim();
    return /^https:\/\/script\.google\.com\/macros\/s\/[A-Za-z0-9_-]+\/exec\/?$/.test(value);
  }

  function normalizedSettings(raw) {
    const base = { ...defaultSettings(), ...(raw || {}) };
    const apps = String(base.appsScriptUrl || '').trim();
    if (!isValidAppsScriptUrl(apps)) base.appsScriptUrl = DEFAULT_APPS_SCRIPT_URL;
    base.sheetdbUrl = '';
    const sheet = String(base.googleSheetUrl || '').trim();
    if (!extractSheetId(sheet)) base.googleSheetUrl = DEFAULT_SHEET_URL;
    return base;
  }

  function currentSettings() {
    return normalizedSettings(loadSettings());
  }

  function extractSheetId(urlOrId) {
    const raw = String(urlOrId || '').trim();
    const m = raw.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (m) return m[1];
    if (/^[a-zA-Z0-9-_]{30,}$/.test(raw)) return raw;
    return '';
  }

  function extractSheetGid(url) {
    const raw = String(url || '');
    const m = raw.match(/[?&#]gid=([0-9]+)/);
    return m ? m[1] : DEFAULT_SHEET_GID;
  }

  function connectDefaultSheet() {
    const current = currentSettings();
    saveSettings({
      googleSheetUrl: current.googleSheetUrl || DEFAULT_SHEET_URL,
      appsScriptUrl: isValidAppsScriptUrl(current.appsScriptUrl)
        ? current.appsScriptUrl
        : DEFAULT_APPS_SCRIPT_URL,
      sheetdbUrl: '',
    });
  }

  function toast(message) {
    const el = $('toast');
    el.textContent = message;
    el.classList.add('show');
    setTimeout(() => el.classList.remove('show'), 2800);
  }

  function setSheetStatus(message, type) {
    const el = $('dashSheetStatus');
    const url = currentSettings().googleSheetUrl || DEFAULT_SHEET_URL;
    el.className = `sheet-status ${type || ''}`.trim();
    el.innerHTML = `${escapeHtml(message)}<br><a href="${escapeHtml(url)}" target="_blank" rel="noopener">${escapeHtml(url)}</a>`;
  }

  function todayStr() {
    return formatDateInput(new Date());
  }

  function formatDateInput(date) {
    const d = date instanceof Date ? date : new Date(date);
    if (Number.isNaN(d.getTime())) return '';
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  }

  function parseDate(value) {
    const text = String(value || '').trim();
    if (!text) return null;
    const iso = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (iso) return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]), 23, 59, 59, 999);
    const dmy = text.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})/);
    if (dmy) return new Date(Number(dmy[3]), Number(dmy[2]) - 1, Number(dmy[1]), 23, 59, 59, 999);
    const d = new Date(text);
    return Number.isNaN(d.getTime()) ? null : d;
  }

  function isExpired(row) {
    const d = parseDate(field(row, COLS.expireDate));
    return d ? d.getTime() < Date.now() : false;
  }

  function field(row, ...keys) {
    if (!row) return '';
    for (const key of keys) {
      if (row[key] != null && String(row[key]).trim() !== '') return String(row[key]).trim();
    }
    const lower = {};
    Object.keys(row).forEach((k) => { lower[k.trim().toLowerCase()] = row[k]; });
    for (const key of keys) {
      const found = lower[String(key).toLowerCase()];
      if (found != null && String(found).trim() !== '') return String(found).trim();
    }
    return '';
  }

  function isControl(row) {
    return field(row, COLS.shop) === CONTROL_SHOP;
  }

  function computeSignature(input, sigLen) {
    let state = 0;
    for (let i = 0; i < input.length; i += 1) state = (state * 31 + input.charCodeAt(i)) >>> 0;
    let sig = '';
    for (let i = 0; i < sigLen; i += 1) {
      state = (state * 1103515245 + 12345 + input.charCodeAt(i % input.length)) >>> 0;
      sig += CHARS[state % 36];
    }
    return sig;
  }

  function generateLicenseKey(expiryDate) {
    const issuedDay = Math.floor(Date.now() / 86400000);
    const expiryDay = Math.floor(expiryDate.getTime() / 86400000);
    const issuePart = issuedDay.toString(36).toUpperCase().padStart(ISSUE_LEN, '0').slice(-ISSUE_LEN);
    const expiryPart = expiryDay.toString(36).toUpperCase().padStart(EXPIRY_LEN, '0').slice(-EXPIRY_LEN);
    const sig = computeSignature(`${LICENSE_SECRET}:${APP_LICENSE_ID}:${issuedDay}:${expiryDay}`, SIG_LEN);
    return (issuePart + expiryPart + sig).slice(0, 20);
  }

  function computeCharsetSignature(input, charset, length) {
    let state = 0;
    for (let i = 0; i < input.length; i += 1) state = (state * 31 + input.charCodeAt(i)) >>> 0;
    let out = '';
    for (let i = 0; i < length; i += 1) {
      state = (state * 1103515245 + 12345 + input.charCodeAt(i % input.length)) >>> 0;
      out += charset[state % charset.length];
    }
    return out;
  }

  function computeRenewalChecksum(dateTimePart, shopName) {
    const input = `${LICENSE_SECRET}:${APP_LICENSE_ID}:${shopName.trim().toLowerCase()}:${dateTimePart}`;
    let state = 0;
    for (let i = 0; i < input.length; i += 1) state = (state * 31 + input.charCodeAt(i)) >>> 0;
    return String(state % 100).padStart(2, '0');
  }

  function resolveDateTimePart(dateTimePart, referenceMs = Date.now()) {
    const mm = parseInt(dateTimePart.slice(0, 2), 10);
    const dd = parseInt(dateTimePart.slice(2, 4), 10);
    const hh = parseInt(dateTimePart.slice(4, 6), 10);
    const mi = parseInt(dateTimePart.slice(6, 8), 10);
    if ([mm, dd, hh, mi].some((n) => Number.isNaN(n)) || mm < 1 || mm > 12 || dd < 1 || dd > 31 || hh > 23 || mi > 59) {
      return null;
    }
    const year = new Date(referenceMs).getFullYear();
    const candidates = [year - 1, year, year + 1]
      .map((y) => new Date(y, mm - 1, dd, hh, mi, 0, 0))
      .filter((d) => d.getMonth() === mm - 1 && d.getDate() === dd);
    if (!candidates.length) return null;
    candidates.sort((a, b) => Math.abs(referenceMs - a.getTime()) - Math.abs(referenceMs - b.getTime()));
    return candidates[0];
  }

  function validateRenewalCode(code, shopName) {
    const shop = String(shopName || '').trim();
    if (!shop) return { valid: false, error: 'Enter the shop name that matches the app login.' };
    const normalized = String(code).replace(/\D/g, '');
    if (normalized.length !== RENEWAL_LEN) return { valid: false, error: 'Renewal code must be 10 digits.' };
    const dateTimePart = normalized.slice(0, 8);
    const checksum = normalized.slice(8);
    if (checksum !== computeRenewalChecksum(dateTimePart, shop)) {
      return { valid: false, error: 'Invalid renewal code for this shop.' };
    }
    const issuedAt = resolveDateTimePart(dateTimePart);
    if (!issuedAt) return { valid: false, error: 'The date/time in this code is not valid.' };
    const age = Date.now() - issuedAt.getTime();
    if (age < 0) return { valid: false, error: 'This renewal code is not active yet.' };
    if (age > RENEWAL_MS) return { valid: false, error: 'This renewal code has expired (45 minute limit).' };
    return { valid: true };
  }

  function endOfDay(date) {
    const d = new Date(date);
    d.setHours(23, 59, 59, 999);
    return d;
  }

  function addDays(days) {
    const d = new Date();
    d.setDate(d.getDate() + days);
    return endOfDay(d);
  }

  function parseGviz(resp) {
    const cols = (resp.table.cols || []).map((c) => c.label || c.id || '');
    return (resp.table.rows || []).map((row) => {
      const obj = {};
      (row.c || []).forEach((cell, i) => {
        obj[cols[i]] = cell ? String(cell.f ?? cell.v ?? '') : '';
      });
      return obj;
    }).filter((row) => Object.values(row).some((v) => String(v).trim()));
  }

  function loadGviz(sheetId) {
    return new Promise((resolve, reject) => {
      const timeout = setTimeout(() => reject(new Error('Sheet read timed out. Share it as Anyone with the link.')), 15000);
      window.google = window.google || {};
      window.google.visualization = window.google.visualization || {};
      window.google.visualization.Query = window.google.visualization.Query || {};
      window.google.visualization.Query.setResponse = (resp) => {
        clearTimeout(timeout);
        try {
          resolve(parseGviz(resp));
        } catch (err) {
          reject(err);
        }
      };
      const script = document.createElement('script');
      script.src = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:json&gid=${extractSheetGid(currentSettings().googleSheetUrl)}`;
      script.onerror = () => {
        clearTimeout(timeout);
        reject(new Error('Could not read Google Sheet. Share as Anyone with the link can view.'));
      };
      document.head.appendChild(script);
    });
  }

  async function fetchJson(url, options) {
    const res = await fetch(url, options);
    if (!res.ok) {
      let detail = '';
      try {
        const body = await res.text();
        detail = body ? `: ${body.slice(0, 180)}` : '';
      } catch {
        // ignore
      }
      if (res.status === 429) {
        throw new Error('Too many requests (429). Wait a minute and try again.');
      }
      throw new Error(`Request failed (${res.status})${detail}`);
    }
    const text = await res.text();
    if (!text) return {};
    try {
      return JSON.parse(text);
    } catch {
      return { ok: true, raw: text };
    }
  }

  function httpsUrl(url) {
    return Secure.assertHttpsUrl(url);
  }

  async function loadCustomers() {
    const settings = currentSettings();
    const sheetId = extractSheetId(settings.googleSheetUrl) || DEFAULT_SHEET_ID;
    const errors = [];

    // Prefer Apps Script (fresh writes). Fall back to public sheet read.
    if (settings.appsScriptUrl) {
      try {
        const base = httpsUrl(settings.appsScriptUrl);
        const joiner = base.includes('?') ? '&' : '?';
        const data = await fetchJson(`${base}${joiner}action=list`);
        const rows = Array.isArray(data?.rows) ? data.rows : Array.isArray(data) ? data : [];
        if (rows.length) return rows.map((row) => Secure.decryptCustomerRow(row));
      } catch (err) {
        errors.push(err.message);
      }
    }

    try {
      const rows = await loadGviz(sheetId);
      if (rows.length) return rows.map((row) => Secure.decryptCustomerRow(row));
    } catch (err) {
      errors.push(err.message);
    }

    throw new Error(errors.filter(Boolean).join(' ') || 'Could not load Google Sheet. Share it as Anyone with the link can view.');
  }

  function customerPayloadFromForm(loginKey, expireDate) {
    return {
      [COLS.name]: $('custName').value.trim(),
      [COLS.mobile]: $('custMobile').value.trim(),
      [COLS.email]: $('custEmail').value.trim(),
      [COLS.shop]: $('custShop').value.trim(),
      [COLS.userCode]: $('custUserCode').value.replace(/\D/g, '').slice(0, 10),
      [COLS.paymentDate]: $('custPaymentDate').value,
      [COLS.loginDate]: $('custLoginDate').value,
      [COLS.expireDate]: expireDate || $('custExpireDate').value,
      [COLS.paymentMode]: $('custPaymentMode').value,
      [COLS.amount]: $('custAmount').value.trim(),
      [COLS.loginKey]: loginKey || $('custLoginKey').value.trim(),
      [COLS.status]: $('custStatus').value,
      [COLS.remarks]: $('custRemarks').value.trim(),
      [COLS.isLogin]: 'FALSE',
      [COLS.loginSession]: '',
    };
  }

  function missingRequired(payload) {
    const needed = [COLS.name, COLS.mobile, COLS.shop, COLS.paymentDate, COLS.loginDate, COLS.expireDate, COLS.paymentMode, COLS.amount, COLS.loginKey, COLS.status];
    return needed.filter((key) => !String(payload[key] || '').trim());
  }

  function sheetRowPayload(payload) {
    const base = {};
    SHEET_HEADER_ORDER.forEach((key) => {
      if (payload[key] != null && String(payload[key]).trim() !== '') {
        base[key] = payload[key];
      } else if (key === COLS.isLogin) {
        base[key] = payload[key] != null ? payload[key] : 'FALSE';
      } else if (key === COLS.status) {
        base[key] = payload[key] || 'Active';
      } else if (key === COLS.loginSession) {
        base[key] = payload[key] != null ? payload[key] : '';
      } else if (payload[key] != null) {
        base[key] = payload[key];
      } else {
        base[key] = '';
      }
    });
    return Secure.encryptCustomerRow(base);
  }

  async function saveViaAppsScript(appsScriptUrl, plain, action) {
    const url = httpsUrl(appsScriptUrl);
    const body = JSON.stringify({ action: action || 'upsert', data: plain });
    let text = '';
    let res;

    try {
      res = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body,
        redirect: 'follow',
      });
      text = await res.text();
    } catch (err) {
      // Browser CORS / opaque redirect — still try fire-and-verify via GET.
      try {
        await fetch(url, {
          method: 'POST',
          mode: 'no-cors',
          headers: { 'Content-Type': 'text/plain;charset=utf-8' },
          body,
        });
        return { opaque: true };
      } catch {
        throw new Error(
          `Fetch failed: ${err.message || err}. Open Settings, paste the /exec Apps Script URL, Save and sync, then try again.`
        );
      }
    }

    if (/accounts\.google\.com|Sign in|signin/i.test(text) || res.status === 401) {
      throw new Error(
        'Apps Script is locked (Google login page). Deploy → Manage deployments → Edit → Who has access: Anyone → New version → Deploy.'
      );
    }
    if (!res.ok) throw new Error(`Apps Script request failed (${res.status}): ${text.slice(0, 120)}`);
    try {
      const json = JSON.parse(text);
      if (json && json.ok === false) throw new Error(json.error || 'Apps Script returned ok:false');
      return json || { ok: true };
    } catch (err) {
      if (/Apps Script|ok:false/.test(String(err.message || ''))) throw err;
      return { ok: true };
    }
  }

  async function verifySheetWrite(shop, loginKey) {
    const settings = currentSettings();
    const base = httpsUrl(settings.appsScriptUrl);
    const joiner = base.includes('?') ? '&' : '?';
    const url = `${base}${joiner}action=check&shop=${encodeURIComponent(shop)}`;
    const res = await fetch(url, { redirect: 'follow', cache: 'no-store' });
    const text = await res.text();
    if (/accounts\.google\.com|Sign in/i.test(text)) {
      throw new Error('Could not verify save: Apps Script access is not set to Anyone.');
    }
    let data;
    try {
      data = JSON.parse(text);
    } catch {
      throw new Error('Could not verify save: Apps Script did not return JSON.');
    }
    const rows = Array.isArray(data?.rows) ? data.rows : [];
    const key = String(loginKey || '').trim().toUpperCase();
    const found = rows.some((row) => {
      const rowShop = field(row, COLS.shop).toLowerCase();
      const rowKey = field(row, COLS.loginKey, 'Login Key').toUpperCase();
      return rowShop === String(shop || '').trim().toLowerCase() && rowKey === key;
    });
    if (!found) {
      throw new Error(
        'Save did not appear in Google Sheet. Check Apps Script is bound to this spreadsheet and redeploy as Web app (Anyone).'
      );
    }
  }

  async function saveRow(payload) {
    const settings = currentSettings();
    const plain = sheetRowPayload(payload);
    const appsScriptUrl = String(settings.appsScriptUrl || '').trim();
    if (!isValidAppsScriptUrl(appsScriptUrl)) {
      throw new Error(
        'Paste a valid Google Apps Script /exec URL in Settings (Deploy → Web app → Anyone).'
      );
    }
    const shop = String(payload[COLS.shop] || '').trim();
    const key = payload[COLS.loginKey];
    if (!shop) throw new Error('Shop name is required.');

    // Ask the sheet if this shop already exists (do not trust login-key uniqueness).
    let shopExists = false;
    try {
      const base = httpsUrl(appsScriptUrl);
      const joiner = base.includes('?') ? '&' : '?';
      const check = await fetchJson(`${base}${joiner}action=check&shop=${encodeURIComponent(shop)}`);
      const rows = Array.isArray(check?.rows) ? check.rows : [];
      shopExists = rows.some((row) => field(row, COLS.shop).toLowerCase() === shop.toLowerCase());
    } catch {
      shopExists = customers.some((row) => field(row, COLS.shop).toLowerCase() === shop.toLowerCase());
    }

    const action = shopExists ? 'upsert' : 'create';
    const result = await saveViaAppsScript(appsScriptUrl, plain, action);
    if (shop && key) {
      await new Promise((r) => setTimeout(r, 800));
      await verifySheetWrite(shop, key);
    }
    return { ...(result || {}), action, shopExists };
  }

  function showPage(name) {
    document.querySelectorAll('.page').forEach((p) => p.classList.remove('active'));
    const page = $(name);
    if (page) page.classList.add('active');
    document.querySelectorAll('.menu button').forEach((b) => {
      b.classList.toggle('active', b.dataset.page === name);
    });
    const titles = {
      dashboard: 'Dashboard',
      generate: 'Generate Key',
      api: 'API Keys',
      reports: 'Reports',
      settings: 'Settings',
    };
    $('pageTitle').textContent = titles[name] || name;
    if (name === 'dashboard') updateStats();
  }

  function statusBadge(row) {
    if (isExpired(row) && field(row, COLS.status) !== 'Disabled') {
      return '<span class="badge badge-orange">Expired</span>';
    }
    const status = field(row, COLS.status) || 'Active';
    if (status === 'Disabled') return '<span class="badge badge-red">Disabled</span>';
    if (status === 'Inactive') return '<span class="badge badge-gray">Inactive</span>';
    return '<span class="badge badge-green">Active</span>';
  }

  function visibleCustomers() {
    return customers.filter((row) => !isControl(row));
  }

  function updateStats() {
    const list = visibleCustomers();
    const expired = list.filter((row) => isExpired(row)).length;
    const disabled = list.filter((row) => field(row, COLS.status) === 'Disabled').length;
    const active = list.filter((row) => !isExpired(row) && (field(row, COLS.status) || 'Active') === 'Active').length;
    $('statTotal').textContent = String(list.length);
    $('statActive').textContent = String(active);
    $('statExpired').textContent = String(expired);
    $('statDisabled').textContent = String(disabled);
  }

  function renderCustomers() {
    // Customers page removed — keep as no-op so sync/stats stay safe.
  }

  function escapeHtml(value) {
    return String(value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function fillSettingsForm() {
    const s = currentSettings();
    $('googleSheetUrl').value = s.googleSheetUrl || '';
    $('appsScriptUrl').value = s.appsScriptUrl || '';
    $('minAppVersion').value = s.minAppVersion || '1.0.0';
    $('latestAppVersion').value = s.latestAppVersion || '1.0.0';
    $('downloadUrl').value = s.downloadUrl || '';
    $('updateMessage').value = s.updateMessage || '';
    $('forceUpdate').value = s.forceUpdate ? 'true' : 'false';
    $('requireSheetRecord').value = s.requireSheetRecord === false ? 'false' : 'true';
    $('appsScriptSource').textContent = APPS_SCRIPT;
  }

  function applyControlRow(row) {
    if (!row) return;
    $('minAppVersion').value = field(row, COLS.loginKey) || $('minAppVersion').value;
    $('latestAppVersion').value = field(row, COLS.amount) || $('latestAppVersion').value;
    $('downloadUrl').value = field(row, COLS.email) || $('downloadUrl').value;
    $('updateMessage').value = field(row, COLS.remarks) || $('updateMessage').value;
    const status = field(row, COLS.status);
    $('forceUpdate').value = status === 'ForceUpdate' || status === 'Disabled' ? 'true' : 'false';
    const mode = field(row, COLS.paymentMode);
    $('requireSheetRecord').value = /open|false|no/i.test(mode) ? 'false' : 'true';
  }

  async function syncSheet() {
    setSheetStatus('Syncing Google Sheet…');
    try {
      customers = await loadCustomers();
      const control = customers.find(isControl);
      applyControlRow(control);
      renderCustomers();
      updateStats();
      const count = visibleCustomers().length;
      setSheetStatus(`Connected to your Google Sheet. Loaded ${count} customer key${count === 1 ? '' : 's'}.`);
      toast('Google Sheet synced');
    } catch (err) {
      setSheetStatus(err.message || 'Could not sync Google Sheet.', 'error');
    }
  }

  function fillGenerateDates() {
    $('custPaymentDate').value = todayStr();
    $('custLoginDate').value = todayStr();
    $('custExpireDate').value = formatDateInput(addDays(90));
    $('customDate').value = formatDateInput(addDays(90));
  }

  function openCustomerModal() {
    // Customers page removed.
  }

  function controlPayload() {
    const s = currentSettings();
    return {
      [COLS.name]: 'NAM bill app control',
      [COLS.mobile]: '0000000000',
      [COLS.email]: $('downloadUrl').value.trim() || s.downloadUrl,
      [COLS.shop]: CONTROL_SHOP,
      [COLS.userCode]: '',
      [COLS.paymentDate]: todayStr(),
      [COLS.loginDate]: todayStr(),
      [COLS.expireDate]: '2099-12-31',
      [COLS.paymentMode]: $('requireSheetRecord').value === 'false' ? 'Open' : 'Strict',
      [COLS.amount]: $('latestAppVersion').value.trim() || '1.0.0',
      [COLS.loginKey]: $('minAppVersion').value.trim() || '1.0.0',
      [COLS.status]: $('forceUpdate').value === 'true' ? 'ForceUpdate' : 'Active',
      [COLS.remarks]: $('updateMessage').value.trim(),
      [COLS.isLogin]: 'FALSE',
      [COLS.loginSession]: '',
    };
  }

  function downloadJson(filename, data) {
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
  }

  function downloadCsv(rows, filename) {
    if (!rows.length) {
      toast('No data to export');
      return;
    }
    const headers = SHEET_HEADER_ORDER;
    const lines = [headers.join(',')];
    rows.forEach((row) => {
      lines.push(headers.map((h) => `"${String(field(row, h)).replace(/"/g, '""')}"`).join(','));
    });
    const blob = new Blob([lines.join('\n')], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
  }

  function unlockApp() {
    $('loginScreen').classList.add('hidden');
    $('appShell').classList.remove('hidden');
    connectDefaultSheet();
    fillSettingsForm();
    fillGenerateDates();
    syncSheet();
  }

  $('adminLoginForm').addEventListener('submit', (e) => {
    e.preventDefault();
    const userOk = Secure.hashesMatch(Secure.sha256Hex($('authUser').value.trim()), AUTH_USER_HASH);
    const passOk = Secure.hashesMatch(Secure.sha256Hex($('authPass').value), AUTH_PASS_HASH);
    if (userOk && passOk) {
      sessionStorage.setItem('namAdminAuth', AUTH_SESSION);
      unlockApp();
    } else {
      $('authError').classList.add('show');
    }
  });

  document.querySelectorAll('.menu button, .quick-action[data-page]').forEach((btn) => {
    btn.addEventListener('click', () => showPage(btn.dataset.page));
  });

  $('syncNowBtn').addEventListener('click', syncSheet);
  $('dashSyncBtn').addEventListener('click', syncSheet);

  document.querySelectorAll('[data-close]').forEach((btn) => {
    btn.addEventListener('click', () => $(btn.dataset.close).classList.remove('show'));
  });

  $('periodGrid').addEventListener('click', (e) => {
    const btn = e.target.closest('.period-btn');
    if (!btn) return;
    document.querySelectorAll('.period-btn').forEach((b) => b.classList.remove('active'));
    btn.classList.add('active');
    if (btn.dataset.custom) {
      useCustom = true;
      $('customDateWrap').style.display = 'block';
    } else {
      useCustom = false;
      $('customDateWrap').style.display = 'none';
      selectedDays = Number(btn.dataset.days);
      $('custExpireDate').value = formatDateInput(addDays(selectedDays));
    }
  });

  $('customDate').addEventListener('change', () => {
    if ($('customDate').value) $('custExpireDate').value = $('customDate').value;
  });

  $('generateBtn').addEventListener('click', async () => {
    const err = $('generateError');
    const btn = $('generateBtn');
    err.classList.remove('show');
    const shop = $('custShop').value.trim();
    const renewalInput = ($('renewalCode').value || '').replace(/\D/g, '');
    // Renewal is only required for renewals; new customers can generate without it.
    if (renewalInput) {
      const renewal = validateRenewalCode(renewalInput, shop);
      if (!renewal.valid) {
        err.textContent = renewal.error;
        err.classList.add('show');
        return;
      }
      if (!$('custUserCode').value.trim()) $('custUserCode').value = renewalInput;
    }
    let expiryDate;
    if (useCustom) {
      if (!$('customDate').value) {
        err.textContent = 'Select a custom expiry date.';
        err.classList.add('show');
        return;
      }
      expiryDate = endOfDay(new Date($('customDate').value));
    } else {
      expiryDate = addDays(selectedDays);
    }
    const key = generateLicenseKey(expiryDate);
    $('custLoginKey').value = key;
    $('custExpireDate').value = formatDateInput(expiryDate);
    const payload = customerPayloadFromForm(key, formatDateInput(expiryDate));
    const missing = missingRequired(payload);
    if (missing.length) {
      err.textContent = `Fill required fields: ${missing.join(', ')}`;
      err.classList.add('show');
      return;
    }
    if (!/^[0-9]{10}$/.test(payload[COLS.mobile])) {
      err.textContent = 'Enter a valid 10-digit mobile number.';
      err.classList.add('show');
      return;
    }
    btn.disabled = true;
    btn.textContent = 'Saving to Google Sheet…';
    try {
      existingRecord = customers.find((row) => field(row, COLS.shop).toLowerCase() === shop.toLowerCase())
        || customers.find((row) => field(row, COLS.mobile) === payload[COLS.mobile])
        || null;
      const saveInfo = await saveRow(payload);
      if (existingRecord) {
        const idx = customers.indexOf(existingRecord);
        if (idx >= 0) customers[idx] = { ...existingRecord, ...payload };
        else customers.push(payload);
      } else {
        customers.push(payload);
      }
      updateStats();
      const result = $('generateResult');
      result.style.display = 'flex';
      result.className = 'sheet-status';
      const mode = saveInfo?.created || saveInfo?.action === 'create'
        ? `New row created${saveInfo?.row ? ` (#${saveInfo.row})` : ''}`
        : `Updated existing shop${saveInfo?.row ? ` (row #${saveInfo.row})` : ''}`;
      result.innerHTML = `<div><strong>${escapeHtml(mode)}</strong><br><span class="mono">${escapeHtml(key)}</span><br>Shop: ${escapeHtml(shop)} · Valid until ${expiryDate.toLocaleDateString('en-IN')}<br><a href="${escapeHtml(currentSettings().googleSheetUrl)}" target="_blank" rel="noopener">Open sheet</a> and scroll to the bottom if it is a new shop.</div>`;
      toast(mode);
      syncSheet().catch(() => {});
    } catch (ex) {
      err.textContent = ex.message || 'Could not save key to Google Sheet.';
      err.classList.add('show');
      const result = $('generateResult');
      result.style.display = 'none';
    } finally {
      btn.disabled = false;
      btn.textContent = 'Generate and save to sheet';
    }
  });

  $('generateApiBtn').addEventListener('click', () => {
    const unique = $('apiUniqueKey').value.trim().toUpperCase().replace(/[^0-9A-Z]/g, '').slice(0, 10);
    $('apiError').classList.remove('show');
    if (unique.length !== 10) {
      $('apiError').textContent = 'App Unique Key must be 10 alphanumeric characters.';
      $('apiError').classList.add('show');
      return;
    }
    const base = `${LICENSE_SECRET}:${APP_LICENSE_ID}:${unique}`;
    $('outUniqueKey').value = unique;
    $('outAppKey').value = computeCharsetSignature(`${base}:APP_KEY`, ALPHA_NUM, 20);
    $('outSecretToken').value = computeCharsetSignature(`${base}:SECRET`, ALPHA_NUM, 25);
    $('outClientToken').value = computeCharsetSignature(`${base}:CLIENT`, DIGITS, 10);
    $('apiResult').style.display = 'block';
  });

  $('copyApiBtn').addEventListener('click', async () => {
    const text = [
      'NAM bill API credentials',
      `App Unique Key: ${$('outUniqueKey').value}`,
      `App Key: ${$('outAppKey').value}`,
      `Secret Token: ${$('outSecretToken').value}`,
      `Client Token: ${$('outClientToken').value}`,
    ].join('\n');
    try {
      await navigator.clipboard.writeText(text);
      toast('Copied API credentials');
    } catch {
      prompt('Copy credentials:', text);
    }
  });

  $('exportCsvBtn').addEventListener('click', () => downloadCsv(visibleCustomers(), 'nam-bill-customers.csv'));
  $('exportControlBtn').addEventListener('click', () => {
    const s = currentSettings();
    downloadJson('license-control.json', {
      minAppVersion: $('minAppVersion').value.trim() || s.minAppVersion,
      latestAppVersion: $('latestAppVersion').value.trim() || s.latestAppVersion,
      forceUpdate: $('forceUpdate').value === 'true',
      downloadUrl: $('downloadUrl').value.trim(),
      updateMessage: $('updateMessage').value.trim(),
      googleSheetId: extractSheetId($('googleSheetUrl').value) || DEFAULT_SHEET_ID,
      googleSheetUrl: $('googleSheetUrl').value.trim(),
      appsScriptUrl: $('appsScriptUrl').value.trim() || s.appsScriptUrl,
      requireSheetRecord: $('requireSheetRecord').value !== 'false',
      blockDisabledAccounts: true,
    });
  });

  $('saveSheetSettingsBtn').addEventListener('click', async () => {
    const url = $('googleSheetUrl').value.trim();
    const apps = $('appsScriptUrl').value.trim() || DEFAULT_APPS_SCRIPT_URL;
    if (!extractSheetId(url)) {
      toast('Paste a valid Google Sheet URL');
      return;
    }
    if (!isValidAppsScriptUrl(apps)) {
      toast('Apps Script URL must end with /exec (Web app, not Library)');
      return;
    }
    try {
      saveSettings({
        googleSheetUrl: url.startsWith('http') ? Secure.assertHttpsUrl(url) : url,
        appsScriptUrl: Secure.assertHttpsUrl(apps),
        sheetdbUrl: '',
      });
    } catch (err) {
      toast(err.message);
      return;
    }
    toast('Settings saved');
    await syncSheet();
  });

  $('saveControlBtn').addEventListener('click', async () => {
    saveSettings({
      minAppVersion: $('minAppVersion').value.trim(),
      latestAppVersion: $('latestAppVersion').value.trim(),
      downloadUrl: $('downloadUrl').value.trim(),
      updateMessage: $('updateMessage').value.trim(),
      forceUpdate: $('forceUpdate').value === 'true',
      requireSheetRecord: $('requireSheetRecord').value !== 'false',
    });
    try {
      existingRecord = customers.find(isControl) || null;
      await saveRow(controlPayload());
      toast('Control row saved. App logins will use it.');
      await syncSheet();
    } catch (err) {
      toast(err.message);
    }
  });

  $('copyScriptBtn').addEventListener('click', async () => {
    try {
      await navigator.clipboard.writeText(APPS_SCRIPT);
      toast('Apps Script copied');
    } catch {
      prompt('Copy script:', APPS_SCRIPT);
    }
  });

  if (sessionStorage.getItem('namAdminAuth') === AUTH_SESSION) unlockApp();
})();
