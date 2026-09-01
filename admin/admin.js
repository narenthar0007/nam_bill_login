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
  const DEFAULT_SHEETDB = 'https://sheetdb.io/api/v1/y019wjypdybot';
  const DEFAULT_SHEET_URL = `https://docs.google.com/spreadsheets/d/${DEFAULT_SHEET_ID}/edit?gid=${DEFAULT_SHEET_GID}#gid=${DEFAULT_SHEET_GID}`;
  const ISSUE_LEN = 4;
  const EXPIRY_LEN = 4;
  const SIG_LEN = 12;
  const CHARS = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  const ALPHA_NUM = CHARS;
  const DIGITS = '0123456789';
  const RENEWAL_LEN = 10;
  const RENEWAL_MS = 45 * 60 * 1000;

  const COLS = {
    name: 'Name',
    mobile: 'Mobile Number',
    email: 'Email ID',
    shop: 'Shop name',
    userCode: 'User Code',
    paymentDate: 'Date of payment',
    loginDate: 'Date of login',
    expireDate: 'Expire date',
    paymentMode: 'Payment mode',
    amount: 'Amount',
    loginKey: 'Login key',
    status: 'Status',
    remarks: 'Remarks',
    isLogin: 'Is Login',
    loginSession: 'Login session',
  };

  const APPS_SCRIPT = `function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const action = (e.parameter.action || 'list').toLowerCase();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = sheet.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const row = {};
    headers.forEach((h, idx) => { row[String(h)] = values[i][idx]; });
    rows.push(row);
  }
  if (action === 'check') {
    const shop = String(e.parameter.shop || '').trim().toLowerCase();
    const found = rows.filter((r) => String(r['Shop name'] || '').trim().toLowerCase() === shop);
    return json({ rows: found });
  }
  return json({ rows });
}

function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
  const payload = JSON.parse(e.postData.contents || '{}');
  const data = payload.data || payload;
  return json(upsertRow_(sheet, headers, data));
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
  const line = headers.map((h) => data[h] != null ? data[h] : (rowIndex > 0 ? values[rowIndex - 1][headers.indexOf(h)] : ''));
  if (rowIndex > 0) sheet.getRange(rowIndex, 1, 1, headers.length).setValues([line]);
  else sheet.appendRow(line);
  return { ok: true };
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}`;

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
    const next = { ...defaultSettings(), ...loadSettings(), ...partial };
    localStorage.setItem('namAdminSettings', Secure.encryptText(JSON.stringify(next)));
    return next;
  }

  function defaultSettings() {
    return {
      googleSheetUrl: DEFAULT_SHEET_URL,
      sheetdbUrl: DEFAULT_SHEETDB,
      appsScriptUrl: '',
      minAppVersion: '1.0.0',
      latestAppVersion: '1.0.0',
      downloadUrl: 'https://github.com/narenthar0007/mahi/releases',
      updateMessage: 'A new version of NAM bill is available. Please download the latest app to continue.',
      forceUpdate: false,
      requireSheetRecord: true,
    };
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
    const current = loadSettings();
    const id = extractSheetId(current.googleSheetUrl);
    if (!id || id === DEFAULT_SHEET_ID) {
      saveSettings({
        googleSheetUrl: DEFAULT_SHEET_URL,
        sheetdbUrl: current.sheetdbUrl || DEFAULT_SHEETDB,
      });
    }
  }

  function toast(message) {
    const el = $('toast');
    el.textContent = message;
    el.classList.add('show');
    setTimeout(() => el.classList.remove('show'), 2800);
  }

  function setSheetStatus(message, type) {
    const el = $('dashSheetStatus');
    const url = ({ ...defaultSettings(), ...loadSettings() }).googleSheetUrl || DEFAULT_SHEET_URL;
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
      script.src = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:json&gid=${extractSheetGid(loadSettings().googleSheetUrl)}`;
      script.onerror = () => {
        clearTimeout(timeout);
        reject(new Error('Could not read Google Sheet. Share as Anyone with the link can view.'));
      };
      document.head.appendChild(script);
    });
  }

  async function fetchJson(url, options) {
    const res = await fetch(url, options);
    if (!res.ok) throw new Error(`Request failed (${res.status})`);
    return res.json();
  }

  function httpsUrl(url) {
    return Secure.assertHttpsUrl(url);
  }

  async function loadCustomers() {
    const settings = { ...defaultSettings(), ...loadSettings() };
    const sheetId = extractSheetId(settings.googleSheetUrl) || DEFAULT_SHEET_ID;
    const errors = [];

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

    if (settings.sheetdbUrl) {
      try {
        const data = await fetchJson(httpsUrl(settings.sheetdbUrl.replace(/\/$/, '')));
        if (Array.isArray(data)) return data.map((row) => Secure.decryptCustomerRow(row));
      } catch (err) {
        errors.push(err.message);
      }
    }

    try {
      const rows = await loadGviz(sheetId);
      return rows.map((row) => Secure.decryptCustomerRow(row));
    } catch (err) {
      errors.push(err.message);
      throw new Error(errors.filter(Boolean).join(' ') || 'Could not load Google Sheet.');
    }
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

  async function saveRow(payload) {
    const settings = { ...defaultSettings(), ...loadSettings() };
    const plain = Secure.encryptCustomerRow(payload);
    const shop = payload[COLS.shop];
    const mobile = payload[COLS.mobile];
    const loginKey = payload[COLS.loginKey];

    if (settings.sheetdbUrl) {
      const base = httpsUrl(settings.sheetdbUrl.replace(/\/$/, ''));
      const existing = existingRecord || customers.find((row) => {
        const sameShop = shop && field(row, COLS.shop).toLowerCase() === shop.toLowerCase();
        const sameMobile = mobile && field(row, COLS.mobile) === mobile;
        return sameShop || sameMobile;
      });
      if (existing) {
        const path = shop
          ? `${encodeURIComponent(COLS.shop)}/${encodeURIComponent(shop)}`
          : mobile
            ? `${encodeURIComponent(COLS.mobile)}/${encodeURIComponent(mobile)}`
            : `${encodeURIComponent(COLS.loginKey)}/${encodeURIComponent(loginKey)}`;
        await fetchJson(`${base}/${path}`, {
          method: 'PATCH',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ data: plain }),
        });
        return;
      }
      await fetchJson(base, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ data: plain }),
      });
      return;
    }

    if (settings.appsScriptUrl) {
      const res = await fetch(httpsUrl(settings.appsScriptUrl), {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({ data: plain }),
      });
      if (!res.ok) throw new Error(`Request failed (${res.status})`);
      return;
    }

    throw new Error('Add a SheetDB API URL or Apps Script URL in Settings so rows can be saved.');
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
      customers: 'Customers',
      generate: 'Generate Key',
      api: 'API Keys',
      reports: 'Reports',
      settings: 'Settings',
    };
    $('pageTitle').textContent = titles[name] || name;
    if (name === 'customers') renderCustomers();
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
    const search = ($('customerSearch').value || '').toLowerCase();
    const status = $('customerStatusFilter').value;
    const tbody = $('customerTable');
    const filtered = visibleCustomers().filter((row) => {
      const blob = [field(row, COLS.name), field(row, COLS.shop), field(row, COLS.mobile), field(row, COLS.loginKey)].join(' ').toLowerCase();
      const matchesSearch = !search || blob.includes(search);
      let matchesStatus = true;
      if (status === 'Expired') matchesStatus = isExpired(row);
      else if (status) matchesStatus = (field(row, COLS.status) || 'Active') === status;
      return matchesSearch && matchesStatus;
    });

    if (!filtered.length) {
      tbody.innerHTML = '<tr><td colspan="11" class="empty">No customers yet. Paste your Google Sheet URL in Settings and sync, or generate a key.</td></tr>';
      updateStats();
      return;
    }

    tbody.innerHTML = filtered.map((row) => {
      const index = customers.indexOf(row);
      const status = field(row, COLS.status) || 'Active';
      const toggleLabel = status === 'Disabled' ? 'Enable' : 'Disable';
      const loggedIn = /^(true|yes|1|y)$/i.test(field(row, COLS.isLogin) || field(row, 'Is Login id'));
      return `<tr>
        <td>${escapeHtml(field(row, COLS.name) || '-')}</td>
        <td>${escapeHtml(field(row, COLS.shop) || '-')}</td>
        <td>${escapeHtml(field(row, COLS.mobile) || '-')}</td>
        <td>${escapeHtml(field(row, COLS.email) || '-')}</td>
        <td class="mono">${escapeHtml(field(row, COLS.loginKey) || '-')}</td>
        <td>${escapeHtml(field(row, COLS.expireDate) || '-')}</td>
        <td>${escapeHtml(field(row, COLS.paymentMode) || '-')}</td>
        <td>${escapeHtml(field(row, COLS.amount) || '-')}</td>
        <td>${statusBadge(row)}</td>
        <td>${loggedIn ? '<span class="badge badge-green">TRUE</span>' : '<span class="badge badge-gray">FALSE</span>'}</td>
        <td>
          <button class="btn btn-blue" type="button" data-edit="${index}">Edit</button>
          <button class="btn btn-danger" type="button" data-toggle="${index}">${toggleLabel}</button>
        </td>
      </tr>`;
    }).join('');
    updateStats();
  }

  function escapeHtml(value) {
    return String(value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function fillSettingsForm() {
    const s = { ...defaultSettings(), ...loadSettings() };
    $('googleSheetUrl').value = s.googleSheetUrl || '';
    $('sheetdbUrl').value = s.sheetdbUrl || '';
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

  function openCustomerModal(index) {
    const row = index == null ? null : customers[index];
    $('editIndex').value = index == null ? '' : String(index);
    $('customerModalTitle').textContent = row ? 'Edit customer' : 'Add customer';
    $('editName').value = row ? field(row, COLS.name) : '';
    $('editMobile').value = row ? field(row, COLS.mobile) : '';
    $('editEmail').value = row ? field(row, COLS.email) : '';
    $('editShop').value = row ? field(row, COLS.shop) : '';
    $('editExpire').value = row ? formatDateInput(parseDate(field(row, COLS.expireDate)) || new Date()) : todayStr();
    $('editStatus').value = row ? (field(row, COLS.status) || 'Active') : 'Active';
    $('editKey').value = row ? field(row, COLS.loginKey) : '';
    $('editAmount').value = row ? field(row, COLS.amount) : '';
    $('editRemarks').value = row ? field(row, COLS.remarks) : '';
    $('customerModal').classList.add('show');
  }

  function controlPayload() {
    const s = { ...defaultSettings(), ...loadSettings() };
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
    const headers = Object.keys(COLS).map((k) => COLS[k]);
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
  $('customerSearch').addEventListener('input', renderCustomers);
  $('customerStatusFilter').addEventListener('change', renderCustomers);
  $('addCustomerBtn').addEventListener('click', () => openCustomerModal(null));

  $('customerTable').addEventListener('click', async (e) => {
    const edit = e.target.closest('[data-edit]');
    const toggle = e.target.closest('[data-toggle]');
    if (edit) openCustomerModal(Number(edit.dataset.edit));
    if (toggle) {
      const index = Number(toggle.dataset.toggle);
      const row = customers[index];
      if (!row) return;
      const current = field(row, COLS.status) || 'Active';
      const next = current === 'Disabled' ? 'Active' : 'Disabled';
      const payload = { ...row, [COLS.status]: next };
      try {
        existingRecord = row;
        await saveRow(payload);
        customers[index] = payload;
        renderCustomers();
        toast(next === 'Disabled' ? 'Account disabled' : 'Account enabled');
      } catch (err) {
        toast(err.message);
      }
    }
  });

  document.querySelectorAll('[data-close]').forEach((btn) => {
    btn.addEventListener('click', () => $(btn.dataset.close).classList.remove('show'));
  });

  $('saveCustomerBtn').addEventListener('click', async () => {
    const index = $('editIndex').value;
    const current = index === '' ? {} : { ...customers[Number(index)] };
    const payload = {
      ...current,
      [COLS.name]: $('editName').value.trim(),
      [COLS.mobile]: $('editMobile').value.trim(),
      [COLS.email]: $('editEmail').value.trim(),
      [COLS.shop]: $('editShop').value.trim(),
      [COLS.expireDate]: $('editExpire').value,
      [COLS.status]: $('editStatus').value,
      [COLS.loginKey]: $('editKey').value.trim(),
      [COLS.amount]: $('editAmount').value.trim(),
      [COLS.remarks]: $('editRemarks').value.trim(),
      [COLS.paymentDate]: current[COLS.paymentDate] || todayStr(),
      [COLS.loginDate]: current[COLS.loginDate] || todayStr(),
      [COLS.paymentMode]: current[COLS.paymentMode] || 'UPI',
    };
    if (!payload[COLS.name] || !payload[COLS.mobile] || !payload[COLS.shop]) {
      toast('Name, mobile, and shop are required');
      return;
    }
    try {
      existingRecord = index === '' ? null : customers[Number(index)];
      await saveRow(payload);
      if (index === '') customers.push(payload);
      else customers[Number(index)] = payload;
      $('customerModal').classList.remove('show');
      renderCustomers();
      toast('Customer saved');
    } catch (err) {
      toast(err.message);
    }
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
    err.classList.remove('show');
    const shop = $('custShop').value.trim();
    const renewal = validateRenewalCode($('renewalCode').value, shop);
    if (!renewal.valid) {
      err.textContent = renewal.error;
      err.classList.add('show');
      return;
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
    try {
      existingRecord = customers.find((row) => field(row, COLS.shop).toLowerCase() === shop.toLowerCase()) || null;
      await saveRow(payload);
      const result = $('generateResult');
      result.style.display = 'flex';
      result.className = 'sheet-status';
      result.innerHTML = `<div><strong>Key saved to Google Sheet</strong><br><span class="mono">${key}</span><br>Valid until ${expiryDate.toLocaleDateString('en-IN')}</div>`;
      await syncSheet();
    } catch (ex) {
      err.textContent = ex.message;
      err.classList.add('show');
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
    const s = { ...defaultSettings(), ...loadSettings() };
    downloadJson('license-control.json', {
      minAppVersion: $('minAppVersion').value.trim() || s.minAppVersion,
      latestAppVersion: $('latestAppVersion').value.trim() || s.latestAppVersion,
      forceUpdate: $('forceUpdate').value === 'true',
      downloadUrl: $('downloadUrl').value.trim(),
      updateMessage: $('updateMessage').value.trim(),
      googleSheetId: extractSheetId($('googleSheetUrl').value) || DEFAULT_SHEET_ID,
      googleSheetUrl: $('googleSheetUrl').value.trim(),
      sheetdbUrl: $('sheetdbUrl').value.trim(),
      appsScriptUrl: $('appsScriptUrl').value.trim(),
      requireSheetRecord: $('requireSheetRecord').value !== 'false',
      blockDisabledAccounts: true,
    });
  });

  $('saveSheetSettingsBtn').addEventListener('click', async () => {
    const url = $('googleSheetUrl').value.trim();
    if (!extractSheetId(url)) {
      toast('Paste a valid Google Sheet URL');
      return;
    }
    try {
      saveSettings({
        googleSheetUrl: url.startsWith('http') ? Secure.assertHttpsUrl(url) : url,
        sheetdbUrl: $('sheetdbUrl').value.trim() ? Secure.assertHttpsUrl($('sheetdbUrl').value.trim()) : '',
        appsScriptUrl: $('appsScriptUrl').value.trim() ? Secure.assertHttpsUrl($('appsScriptUrl').value.trim()) : '',
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
