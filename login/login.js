(() => {
  const CONTROL_JSON = '../license-control.json';
  const CONTROL_SHOP = '__NAM_BILL_CONTROL__';
  const DEFAULT_SHEET_ID = '1Gr7vHrtQZ_wRrFsv3y_mCJ6qdNeSB7x_TkQhUCY3ato';

  const $ = (id) => document.getElementById(id);

  function field(row, ...keys) {
    if (!row) return '';
    for (const key of keys) {
      if (row[key] != null && String(row[key]).trim() !== '') return String(row[key]).trim();
    }
    return '';
  }

  function parseDate(value) {
    const text = String(value || '').trim();
    const iso = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (iso) return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]), 23, 59, 59, 999);
    const dmy = text.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})/);
    if (dmy) return new Date(Number(dmy[3]), Number(dmy[2]) - 1, Number(dmy[1]), 23, 59, 59, 999);
    const d = new Date(text);
    return Number.isNaN(d.getTime()) ? null : d;
  }

  function compareVersions(a, b) {
    const pa = String(a || '0').split('.').map((n) => parseInt(n, 10) || 0);
    const pb = String(b || '0').split('.').map((n) => parseInt(n, 10) || 0);
    for (let i = 0; i < 3; i += 1) {
      if ((pa[i] || 0) > (pb[i] || 0)) return 1;
      if ((pa[i] || 0) < (pb[i] || 0)) return -1;
    }
    return 0;
  }

  function parseCsv(text) {
    const rows = [];
    let row = [];
    let cell = '';
    let inQuotes = false;
    const input = String(text || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    for (let i = 0; i < input.length; i += 1) {
      const ch = input[i];
      if (inQuotes) {
        if (ch === '"') {
          if (input[i + 1] === '"') {
            cell += '"';
            i += 1;
          } else inQuotes = false;
        } else cell += ch;
      } else if (ch === '"') inQuotes = true;
      else if (ch === ',') {
        row.push(cell);
        cell = '';
      } else if (ch === '\n') {
        row.push(cell);
        rows.push(row);
        row = [];
        cell = '';
      } else cell += ch;
    }
    if (cell.length || row.length) {
      row.push(cell);
      rows.push(row);
    }
    if (rows.length < 2) return [];
    const headers = rows[0].map((h) => String(h || '').trim());
    return rows.slice(1).map((cells) => {
      const obj = {};
      headers.forEach((h, idx) => {
        obj[h] = cells[idx] != null ? String(cells[idx]) : '';
      });
      return obj;
    }).filter((r) => Object.values(r).some((v) => String(v).trim()));
  }

  async function loadControl() {
    try {
      const res = await fetch(CONTROL_JSON, { cache: 'no-store' });
      if (!res.ok) throw new Error('control missing');
      return await res.json();
    } catch {
      return {
        googleSheetId: DEFAULT_SHEET_ID,
        requireSheetRecord: true,
        forceUpdate: false,
        downloadUrl: 'https://github.com/narenthar0007/mahi/releases',
        updateMessage: 'Please download the latest NAM bill app.',
        minAppVersion: '1.0.0',
        appsScriptUrl: '',
      };
    }
  }

  async function loadRows(control, shop) {
    const Secure = window.NamSecureCrypto;
    const decryptAll = (rows) => (Array.isArray(rows) ? rows : []).map((row) => Secure.decryptCustomerRow(row));
    const sheetId = String(control.googleSheetId || DEFAULT_SHEET_ID).trim();

    try {
      const res = await fetch(`https://docs.google.com/spreadsheets/d/${sheetId}/export?format=csv&gid=0`, {
        cache: 'no-store',
      });
      if (res.ok) {
        const rows = parseCsv(await res.text());
        if (rows.length) {
          const decrypted = decryptAll(rows);
          if (!shop) return decrypted;
          const needle = shop.trim().toLowerCase();
          return decrypted.filter((r) => {
            const name = field(r, 'Shop name').toLowerCase();
            return name === needle || name === CONTROL_SHOP.toLowerCase();
          });
        }
      }
    } catch {
      // fall through to Apps Script
    }

    if (control.appsScriptUrl) {
      const base = Secure.assertHttpsUrl(control.appsScriptUrl);
      const joiner = base.includes('?') ? '&' : '?';
      const lookup = shop ? `&shop=${encodeURIComponent(shop)}` : '';
      const res = await fetch(`${base}${joiner}action=${shop ? 'check' : 'list'}${lookup}`);
      const data = await res.json();
      if (Array.isArray(data?.rows)) return decryptAll(data.rows);
      if (Array.isArray(data)) return decryptAll(data);
    }

    throw new Error('Could not read Google Sheet. Share it as Anyone with the link can view.');
  }

  $('licenseKey').addEventListener('input', (e) => {
    e.target.value = e.target.value.replace(/[^a-zA-Z0-9]/g, '').toUpperCase().slice(0, 20);
  });

  $('loginForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    ['errorBox', 'okBox', 'warnBox'].forEach((id) => $(id).classList.remove('show'));
    $('downloadBtn').style.display = 'none';

    const shop = $('shopName').value.trim();
    const key = $('licenseKey').value.trim().toUpperCase();
    if (key.length !== 20) {
      $('errorBox').textContent = 'License key must be exactly 20 characters.';
      $('errorBox').classList.add('show');
      return;
    }

    try {
      const control = await loadControl();
      const rows = await loadRows(control, shop);
      const controlRow = rows.find((r) => field(r, 'Shop name') === CONTROL_SHOP);
      const minVersion = field(controlRow, 'Login key') || control.minAppVersion || '1.0.0';
      const forceUpdate = /forceupdate|disabled/i.test(field(controlRow, 'Status') || '') || control.forceUpdate;
      const downloadUrl = field(controlRow, 'Email ID') || control.downloadUrl || '';
      const updateMessage = field(controlRow, 'Remarks') || control.updateMessage || '';

      if (forceUpdate && compareVersions(minVersion, '1.0.0') > 0) {
        $('errorBox').textContent = updateMessage || 'Please update the app.';
        $('errorBox').classList.add('show');
        if (downloadUrl) {
          $('downloadBtn').style.display = 'inline-flex';
          $('downloadBtn').onclick = () => window.open(downloadUrl, '_blank');
        }
        return;
      }

      const match = rows.find(
        (r) =>
          field(r, 'Shop name').toLowerCase() === shop.toLowerCase() &&
          field(r, 'Login key', 'Login Key').toUpperCase() === key
      );

      if (!match) {
        $('errorBox').textContent =
          'Shop name and license key do not match the Google Sheet. Check both carefully.';
        $('errorBox').classList.add('show');
        return;
      }

      const status = field(match, 'Status') || 'Active';
      if (/disabled|inactive/i.test(status)) {
        $('errorBox').textContent = `This account is ${status}. Contact your provider.`;
        $('errorBox').classList.add('show');
        return;
      }

      const expiry = parseDate(field(match, 'Expire date'));
      if (expiry && expiry.getTime() < Date.now()) {
        $('errorBox').textContent = 'This license has expired. Generate a 10-digit renewal code in the app and send it to your provider.';
        $('errorBox').classList.add('show');
        return;
      }

      $('okBox').textContent = `License valid until ${field(match, 'Expire date') || '—'}. You can open the NAM bill app.`;
      $('okBox').classList.add('show');
    } catch (err) {
      $('errorBox').textContent = err.message || 'Login check failed.';
      $('errorBox').classList.add('show');
    }
  });
})();
