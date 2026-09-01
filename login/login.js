(() => {
  const CONTROL_JSON = '../license-control.json';
  const CONTROL_SHOP = '__NAM_BILL_CONTROL__';
  const DEFAULT_SHEETDB = 'https://sheetdb.io/api/v1/y019wjypdybot';

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

  async function loadControl() {
    try {
      const res = await fetch(CONTROL_JSON, { cache: 'no-store' });
      if (!res.ok) throw new Error('control missing');
      return await res.json();
    } catch {
      return {
        sheetdbUrl: DEFAULT_SHEETDB,
        requireSheetRecord: true,
        forceUpdate: false,
        downloadUrl: 'https://github.com/narenthar0007/mahi/releases',
        updateMessage: 'Please download the latest NAM bill app.',
        minAppVersion: '1.0.0',
      };
    }
  }

  async function loadRows(control, shop) {
    const Secure = window.NamSecureCrypto;
    const decryptAll = (rows) => (Array.isArray(rows) ? rows : []).map((row) => Secure.decryptCustomerRow(row));

    if (control.appsScriptUrl) {
      const base = Secure.assertHttpsUrl(control.appsScriptUrl);
      const joiner = base.includes('?') ? '&' : '?';
      const lookup = shop ? `&shop=${encodeURIComponent(shop)}` : '';
      const res = await fetch(`${base}${joiner}action=${shop ? 'check' : 'list'}${lookup}`);
      const data = await res.json();
      if (Array.isArray(data?.rows)) return decryptAll(data.rows);
      if (Array.isArray(data)) return decryptAll(data);
    }
    const base = Secure.assertHttpsUrl(String(control.sheetdbUrl || DEFAULT_SHEETDB).replace(/\/$/, ''));
    if (shop) {
      const [legacy, controlRows] = await Promise.all([
        fetch(`${base}/search?${encodeURIComponent('Shop name')}=${encodeURIComponent(shop)}`).then((r) => r.json()),
        fetch(`${base}/search?${encodeURIComponent('Shop name')}=${encodeURIComponent(CONTROL_SHOP)}`).then((r) => r.json()),
      ]);
      return decryptAll([
        ...(Array.isArray(legacy) ? legacy : []),
        ...(Array.isArray(controlRows) ? controlRows : []),
      ]);
    }
    const res = await fetch(base);
    const data = await res.json();
    return decryptAll(data);
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
      const controlRow = rows.find((row) => field(row, 'Shop name') === CONTROL_SHOP);
      const minVersion = field(controlRow, 'Login key') || control.minAppVersion || '1.0.0';
      const forceUpdate = control.forceUpdate || field(controlRow, 'Status') === 'ForceUpdate';
      const downloadUrl = field(controlRow, 'Email ID') || control.downloadUrl;
      const message = field(controlRow, 'Remarks') || control.updateMessage;
      if (downloadUrl) $('downloadBtn').href = downloadUrl;

      if (forceUpdate && compareVersions(minVersion, '1.0.0') > 0) {
        $('warnBox').textContent = message || 'Please download the latest NAM bill app.';
        $('warnBox').classList.add('show');
        $('downloadBtn').style.display = 'block';
      }

      const match = rows.find((row) => {
        if (field(row, 'Shop name') === CONTROL_SHOP) return false;
        return field(row, 'Shop name').toLowerCase() === shop.toLowerCase()
          && field(row, 'Login key', 'Login Key').toUpperCase() === key;
      });

      if (!match) {
        $('errorBox').textContent = 'This shop/key is not registered on the license sheet. Contact your provider.';
        $('errorBox').classList.add('show');
        return;
      }

      const status = field(match, 'Status') || 'Active';
      if (status === 'Disabled' || status === 'Inactive') {
        $('errorBox').textContent = status === 'Disabled'
          ? 'This account has been disabled. Contact your NAM bill provider.'
          : 'This account is inactive.';
        $('errorBox').classList.add('show');
        return;
      }

      const expiry = parseDate(field(match, 'Expire date'));
      if (expiry && expiry.getTime() < Date.now()) {
        $('errorBox').textContent = 'This license has expired. Generate a 10-digit renewal code in the app and send it to your provider.';
        $('errorBox').classList.add('show');
        return;
      }

      $('okBox').innerHTML = `<strong>License active for ${shop}.</strong><br>Use this same shop name and key in the NAM bill app. Expires ${expiry ? expiry.toLocaleDateString('en-IN') : 'on the key date'}.`;
      $('okBox').classList.add('show');
      if (downloadUrl) $('downloadBtn').style.display = 'block';
    } catch (err) {
      $('errorBox').textContent = 'Could not reach the license sheet. Check your internet and try again.';
      $('errorBox').classList.add('show');
    }
  });
})();
