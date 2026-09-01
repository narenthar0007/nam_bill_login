(function (global) {
const ENC_PREFIX = 'enc:v1:';
const DATA_ENCRYPTION_SECRET = 'NAM_BILL_DATA_AES256_2026';

const ENCRYPTED_SHEET_FIELDS = [
  'Name',
  'Mobile Number',
  'Email ID',
  'Shop name',
  'User Code',
  'Login key',
  'Remarks',
  'Amount',
];

const LOOKUP_COLUMNS = {
  shop: 'Shop lookup',
  key: 'Key lookup',
  mobile: 'Mobile lookup',
};

const CONTROL_SHOP = '__NAM_BILL_CONTROL__';

const SHA_K = new Uint32Array([
  0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5,
  0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174,
  0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da,
  0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967,
  0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85,
  0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070,
  0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3,
  0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2,
]);

const AES_SBOX = new Uint8Array([
  99, 124, 119, 123, 242, 107, 111, 197, 48, 1, 103, 43, 254, 215, 171, 118, 202, 130, 201, 125, 250, 89, 71, 240, 173, 212, 162, 175, 156, 164, 114, 192, 183, 253, 147, 38, 54, 63, 247, 204, 52, 165, 229, 241, 113, 216, 49, 21, 4, 199, 35, 195, 24, 150, 5, 154, 7, 18, 128, 226, 235, 39, 178, 117, 9, 131, 44, 26, 27, 110, 90, 160, 82, 59, 214, 179, 41, 227, 47, 132, 83, 209, 0, 237, 32, 252, 177, 91, 106, 203, 190, 57, 74, 76, 88, 207, 208, 239, 170, 251, 67, 77, 51, 133, 69, 249, 2, 127, 80, 60, 159, 168, 81, 163, 64, 143, 146, 157, 56, 245, 188, 182, 218, 33, 16, 255, 243, 210, 205, 12, 19, 236, 95, 151, 68, 23, 196, 167, 126, 61, 100, 93, 25, 115, 96, 129, 79, 220, 34, 42, 144, 136, 70, 238, 184, 20, 222, 94, 11, 219, 224, 50, 58, 10, 73, 6, 36, 92, 194, 211, 172, 98, 145, 149, 228, 121, 231, 200, 55, 109, 141, 213, 78, 169, 108, 86, 244, 234, 101, 122, 174, 8, 186, 120, 37, 46, 28, 166, 180, 198, 232, 221, 116, 31, 75, 189, 139, 138, 112, 62, 181, 102, 72, 3, 246, 14, 97, 53, 87, 185, 134, 193, 29, 158, 225, 248, 152, 17, 105, 217, 142, 148, 155, 30, 135, 233, 206, 85, 40, 223, 140, 161, 137, 13, 191, 230, 66, 104, 65, 153, 45, 15, 176, 84, 187, 22,
]);

const AES_RCON = new Uint8Array([0x00, 0x01, 0x02, 0x04, 0x08, 0x10, 0x20, 0x40, 0x80, 0x1b, 0x36]);

function rotr(x, n) {
  return (x >>> n) | (x << (32 - n));
}

function utf8Bytes(text) {
  if (typeof TextEncoder !== 'undefined') return new TextEncoder().encode(String(text));
  const s = String(text);
  const out = [];
  for (let i = 0; i < s.length; i += 1) {
    let c = s.charCodeAt(i);
    if (c < 0x80) out.push(c);
    else if (c < 0x800) out.push(0xc0 | (c >> 6), 0x80 | (c & 63));
    else if (c >= 0xd800 && c <= 0xdbff && i + 1 < s.length) {
      const c2 = s.charCodeAt(i + 1);
      c = 0x10000 + ((c & 0x3ff) << 10) + (c2 & 0x3ff);
      i += 1;
      out.push(0xf0 | (c >> 18), 0x80 | ((c >> 12) & 63), 0x80 | ((c >> 6) & 63), 0x80 | (c & 63));
    } else out.push(0xe0 | (c >> 12), 0x80 | ((c >> 6) & 63), 0x80 | (c & 63));
  }
  return Uint8Array.from(out);
}

function bytesToUtf8(bytes) {
  if (typeof TextDecoder !== 'undefined') return new TextDecoder().decode(bytes);
  let out = '';
  for (let i = 0; i < bytes.length; ) {
    const b = bytes[i];
    if (b < 0x80) {
      out += String.fromCharCode(b);
      i += 1;
    } else if (b < 0xe0) {
      out += String.fromCharCode(((b & 31) << 6) | (bytes[i + 1] & 63));
      i += 2;
    } else if (b < 0xf0) {
      out += String.fromCharCode(((b & 15) << 12) | ((bytes[i + 1] & 63) << 6) | (bytes[i + 2] & 63));
      i += 3;
    } else {
      const cp =
        ((b & 7) << 18) | ((bytes[i + 1] & 63) << 12) | ((bytes[i + 2] & 63) << 6) | (bytes[i + 3] & 63);
      const c = cp - 0x10000;
      out += String.fromCharCode(0xd800 + (c >> 10), 0xdc00 + (c & 0x3ff));
      i += 4;
    }
  }
  return out;
}

function bytesToHex(bytes) {
  return Array.from(bytes, (b) => b.toString(16).padStart(2, '0')).join('');
}

function sha256Bytes(data) {
  const msg = data instanceof Uint8Array ? data : utf8Bytes(data);
  const bitLen = msg.length * 8;
  const padded = new Uint8Array(((msg.length + 9 + 63) >> 6) << 6);
  padded.set(msg);
  padded[msg.length] = 0x80;
  const view = new DataView(padded.buffer);
  view.setUint32(padded.length - 4, bitLen >>> 0);

  let h0 = 0x6a09e667;
  let h1 = 0xbb67ae85;
  let h2 = 0x3c6ef372;
  let h3 = 0xa54ff53a;
  let h4 = 0x510e527f;
  let h5 = 0x9b05688c;
  let h6 = 0x1f83d9ab;
  let h7 = 0x5be0cd19;
  const w = new Uint32Array(64);

  for (let offset = 0; offset < padded.length; offset += 64) {
    for (let i = 0; i < 16; i += 1) w[i] = view.getUint32(offset + i * 4);
    for (let i = 16; i < 64; i += 1) {
      const s0 = rotr(w[i - 15], 7) ^ rotr(w[i - 15], 18) ^ (w[i - 15] >>> 3);
      const s1 = rotr(w[i - 2], 17) ^ rotr(w[i - 2], 19) ^ (w[i - 2] >>> 10);
      w[i] = (w[i - 16] + s0 + w[i - 7] + s1) >>> 0;
    }
    let a = h0;
    let b = h1;
    let c = h2;
    let d = h3;
    let e = h4;
    let f = h5;
    let g = h6;
    let h = h7;
    for (let i = 0; i < 64; i += 1) {
      const S1 = rotr(e, 6) ^ rotr(e, 11) ^ rotr(e, 25);
      const ch = (e & f) ^ (~e & g);
      const temp1 = (h + S1 + ch + SHA_K[i] + w[i]) >>> 0;
      const S0 = rotr(a, 2) ^ rotr(a, 13) ^ rotr(a, 22);
      const maj = (a & b) ^ (a & c) ^ (b & c);
      const temp2 = (S0 + maj) >>> 0;
      h = g;
      g = f;
      f = e;
      e = (d + temp1) >>> 0;
      d = c;
      c = b;
      b = a;
      a = (temp1 + temp2) >>> 0;
    }
    h0 = (h0 + a) >>> 0;
    h1 = (h1 + b) >>> 0;
    h2 = (h2 + c) >>> 0;
    h3 = (h3 + d) >>> 0;
    h4 = (h4 + e) >>> 0;
    h5 = (h5 + f) >>> 0;
    h6 = (h6 + g) >>> 0;
    h7 = (h7 + h) >>> 0;
  }

  const out = new Uint8Array(32);
  const outView = new DataView(out.buffer);
  outView.setUint32(0, h0);
  outView.setUint32(4, h1);
  outView.setUint32(8, h2);
  outView.setUint32(12, h3);
  outView.setUint32(16, h4);
  outView.setUint32(20, h5);
  outView.setUint32(24, h6);
  outView.setUint32(28, h7);
  return out;
}

function sha256Hex(text) {
  return bytesToHex(sha256Bytes(utf8Bytes(text)));
}

function hmacSha256(keyBytes, dataBytes) {
  const block = 64;
  let key = keyBytes instanceof Uint8Array ? keyBytes : utf8Bytes(keyBytes);
  if (key.length > block) key = sha256Bytes(key);
  const oKey = new Uint8Array(block);
  const iKey = new Uint8Array(block);
  oKey.fill(0x5c);
  iKey.fill(0x36);
  for (let i = 0; i < key.length; i += 1) {
    oKey[i] ^= key[i];
    iKey[i] ^= key[i];
  }
  const data = dataBytes instanceof Uint8Array ? dataBytes : utf8Bytes(dataBytes);
  const inner = new Uint8Array(iKey.length + data.length);
  inner.set(iKey);
  inner.set(data, iKey.length);
  const innerHash = sha256Bytes(inner);
  const outer = new Uint8Array(oKey.length + innerHash.length);
  outer.set(oKey);
  outer.set(innerHash, oKey.length);
  return sha256Bytes(outer);
}

function xorBytes(a, b) {
  const out = new Uint8Array(a.length);
  for (let i = 0; i < a.length; i += 1) out[i] = a[i] ^ b[i];
  return out;
}

function subWord(word) {
  return (
    (AES_SBOX[(word >>> 24) & 255] << 24) |
    (AES_SBOX[(word >>> 16) & 255] << 16) |
    (AES_SBOX[(word >>> 8) & 255] << 8) |
    AES_SBOX[word & 255]
  ) >>> 0;
}

function rotWord(word) {
  return ((word << 8) | (word >>> 24)) >>> 0;
}

function expandKey(key) {
  const Nk = 8;
  const Nr = 14;
  const W = new Uint32Array(4 * (Nr + 1));
  const view = new DataView(key.buffer, key.byteOffset, key.byteLength);
  for (let i = 0; i < Nk; i += 1) W[i] = view.getUint32(i * 4);
  for (let i = Nk; i < W.length; i += 1) {
    let temp = W[i - 1];
    if (i % Nk === 0) temp = subWord(rotWord(temp)) ^ (AES_RCON[i / Nk] << 24);
    else if (i % Nk === 4) temp = subWord(temp);
    W[i] = (W[i - Nk] ^ temp) >>> 0;
  }
  return W;
}

function gmul2(b) {
  return ((b << 1) & 0xff) ^ (b & 0x80 ? 0x1b : 0);
}

function mixColumn(col) {
  const a0 = col[0];
  const a1 = col[1];
  const a2 = col[2];
  const a3 = col[3];
  const t = a0 ^ a1 ^ a2 ^ a3;
  col[0] ^= t ^ gmul2(a0 ^ a1);
  col[1] ^= t ^ gmul2(a1 ^ a2);
  col[2] ^= t ^ gmul2(a2 ^ a3);
  col[3] ^= t ^ gmul2(a3 ^ a0);
}

function aesEncryptBlock(keySchedule, input) {
  const s = new Uint8Array(input);
  const Nr = 14;
  const add = (round) => {
    for (let c = 0; c < 4; c += 1) {
      const w = keySchedule[round * 4 + c];
      s[c * 4] ^= (w >>> 24) & 255;
      s[c * 4 + 1] ^= (w >>> 16) & 255;
      s[c * 4 + 2] ^= (w >>> 8) & 255;
      s[c * 4 + 3] ^= w & 255;
    }
  };
  add(0);
  for (let round = 1; round < Nr; round += 1) {
    for (let i = 0; i < 16; i += 1) s[i] = AES_SBOX[s[i]];
    const t = new Uint8Array(s);
    s[1] = t[5];
    s[5] = t[9];
    s[9] = t[13];
    s[13] = t[1];
    s[2] = t[10];
    s[6] = t[14];
    s[10] = t[2];
    s[14] = t[6];
    s[3] = t[15];
    s[7] = t[3];
    s[11] = t[7];
    s[15] = t[11];
    for (let c = 0; c < 4; c += 1) mixColumn(s.subarray(c * 4, c * 4 + 4));
    add(round);
  }
  for (let i = 0; i < 16; i += 1) s[i] = AES_SBOX[s[i]];
  const t = new Uint8Array(s);
  s[1] = t[5];
  s[5] = t[9];
  s[9] = t[13];
  s[13] = t[1];
  s[2] = t[10];
  s[6] = t[14];
  s[10] = t[2];
  s[14] = t[6];
  s[3] = t[15];
  s[7] = t[3];
  s[11] = t[7];
  s[15] = t[11];
  add(Nr);
  return s;
}

function inc32(block) {
  const out = new Uint8Array(block);
  for (let i = 15; i >= 12; i -= 1) {
    out[i] = (out[i] + 1) & 255;
    if (out[i]) break;
  }
  return out;
}

function ghashMul(x, y) {
  const z = new Uint8Array(16);
  const v = new Uint8Array(y);
  for (let i = 0; i < 128; i += 1) {
    if ((x[i >> 3] >>> (7 - (i & 7))) & 1) {
      for (let j = 0; j < 16; j += 1) z[j] ^= v[j];
    }
    const lsb = v[15] & 1;
    for (let j = 15; j > 0; j -= 1) v[j] = ((v[j] >>> 1) | ((v[j - 1] & 1) << 7)) & 255;
    v[0] >>>= 1;
    if (lsb) v[0] ^= 0xe1;
  }
  return z;
}

function ghash(h, data) {
  let y = new Uint8Array(16);
  for (let i = 0; i < data.length; i += 16) {
    const block = new Uint8Array(16);
    block.set(data.subarray(i, Math.min(i + 16, data.length)));
    y = ghashMul(xorBytes(y, block), h);
  }
  return y;
}

function gctr(schedule, icb, data) {
  if (!data.length) return new Uint8Array(0);
  const out = new Uint8Array(data.length);
  let cb = new Uint8Array(icb);
  for (let i = 0; i < data.length; i += 16) {
    const keystream = aesEncryptBlock(schedule, cb);
    const n = Math.min(16, data.length - i);
    for (let j = 0; j < n; j += 1) out[i + j] = data[i + j] ^ keystream[j];
    cb = inc32(cb);
  }
  return out;
}

function randomBytes(n) {
  const out = new Uint8Array(n);
  if (globalThis.crypto?.getRandomValues) {
    globalThis.crypto.getRandomValues(out);
    return out;
  }
  for (let i = 0; i < n; i += 1) out[i] = Math.floor(Math.random() * 256);
  return out;
}

const B64 = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';

function bytesToBase64(bytes) {
  let out = '';
  for (let i = 0; i < bytes.length; i += 3) {
    const a = bytes[i];
    const b = i + 1 < bytes.length ? bytes[i + 1] : 0;
    const c = i + 2 < bytes.length ? bytes[i + 2] : 0;
    out += B64[a >> 2];
    out += B64[((a & 3) << 4) | (b >> 4)];
    out += i + 1 < bytes.length ? B64[((b & 15) << 2) | (c >> 6)] : '=';
    out += i + 2 < bytes.length ? B64[c & 63] : '=';
  }
  return out;
}

function base64ToBytes(str) {
  const clean = String(str).replace(/[^A-Za-z0-9+/=]/g, '');
  const out = [];
  for (let i = 0; i < clean.length; i += 4) {
    const a = B64.indexOf(clean[i]);
    const b = B64.indexOf(clean[i + 1]);
    const c = B64.indexOf(clean[i + 2]);
    const d = B64.indexOf(clean[i + 3]);
    out.push(((a << 2) | (b >> 4)) & 255);
    if (clean[i + 2] !== '=') out.push((((b & 15) << 4) | (c >> 2)) & 255);
    if (clean[i + 3] !== '=') out.push((((c & 3) << 6) | d) & 255);
  }
  return Uint8Array.from(out);
}

function deriveKey(secret) {
  return sha256Bytes(utf8Bytes(secret || DATA_ENCRYPTION_SECRET));
}

function encryptBytes(plainBytes, secret = DATA_ENCRYPTION_SECRET) {
  const key = deriveKey(secret);
  const schedule = expandKey(key);
  const iv = randomBytes(12);
  const h = aesEncryptBlock(schedule, new Uint8Array(16));
  const j0 = new Uint8Array(16);
  j0.set(iv);
  j0[15] = 1;
  const ciphertext = gctr(schedule, inc32(j0), plainBytes);
  const u = (16 - (ciphertext.length % 16)) % 16;
  const ghashIn = new Uint8Array(ciphertext.length + u + 16);
  ghashIn.set(ciphertext);
  const lenView = new DataView(ghashIn.buffer, ghashIn.byteOffset + ghashIn.length - 16);
  lenView.setUint32(8, 0);
  lenView.setUint32(12, ciphertext.length * 8);
  const s = ghash(h, ghashIn);
  const tag = xorBytes(aesEncryptBlock(schedule, j0), s);
  const packed = new Uint8Array(12 + ciphertext.length + 16);
  packed.set(iv, 0);
  packed.set(ciphertext, 12);
  packed.set(tag, 12 + ciphertext.length);
  return packed;
}

function decryptBytes(packed, secret = DATA_ENCRYPTION_SECRET) {
  if (!packed || packed.length < 28) throw new Error('invalid_cipher');
  const key = deriveKey(secret);
  const schedule = expandKey(key);
  const iv = packed.subarray(0, 12);
  const tag = packed.subarray(packed.length - 16);
  const ciphertext = packed.subarray(12, packed.length - 16);
  const h = aesEncryptBlock(schedule, new Uint8Array(16));
  const j0 = new Uint8Array(16);
  j0.set(iv);
  j0[15] = 1;
  const u = (16 - (ciphertext.length % 16)) % 16;
  const ghashIn = new Uint8Array(ciphertext.length + u + 16);
  ghashIn.set(ciphertext);
  const lenView = new DataView(ghashIn.buffer, ghashIn.byteOffset + ghashIn.length - 16);
  lenView.setUint32(12, ciphertext.length * 8);
  const s = ghash(h, ghashIn);
  const expected = xorBytes(aesEncryptBlock(schedule, j0), s);
  let diff = 0;
  for (let i = 0; i < 16; i += 1) diff |= tag[i] ^ expected[i];
  if (diff !== 0) throw new Error('auth_failed');
  return gctr(schedule, inc32(j0), ciphertext);
}

function isEncryptedValue(value) {
  return String(value || '').startsWith(ENC_PREFIX);
}

function encryptText(text, secret = DATA_ENCRYPTION_SECRET) {
  if (text == null || text === '') return '';
  const str = String(text);
  if (isEncryptedValue(str)) return str;
  return ENC_PREFIX + bytesToBase64(encryptBytes(utf8Bytes(str), secret));
}

function decryptText(text, secret = DATA_ENCRYPTION_SECRET) {
  const str = String(text || '');
  if (!str) return '';
  if (!isEncryptedValue(str)) return str;
  const packed = base64ToBytes(str.slice(ENC_PREFIX.length));
  return bytesToUtf8(decryptBytes(packed, secret));
}

function lookupToken(kind, value) {
  const normalized = String(value || '').trim().toLowerCase();
  if (!normalized) return '';
  return bytesToHex(hmacSha256(deriveKey(DATA_ENCRYPTION_SECRET), utf8Bytes(`${kind}:${normalized}`))).slice(0, 40);
}

function hashesMatch(a, b) {
  const left = String(a || '');
  const right = String(b || '');
  if (left.length !== right.length || !left.length) return false;
  let diff = 0;
  for (let i = 0; i < left.length; i += 1) diff |= left.charCodeAt(i) ^ right.charCodeAt(i);
  return diff === 0;
}

function assertHttpsUrl(url) {
  const value = String(url || '').trim();
  if (!value) return '';
  if (value.startsWith('https://')) return value;
  if (/^http:\/\/(localhost|127\.0\.0\.1)(:\d+)?(\/|$)/i.test(value)) return value;
  throw new Error('Only HTTPS URLs are allowed for license data.');
}

function encryptCustomerRow(row, options) {
  const next = decryptCustomerRow(row, options);
  next[LOOKUP_COLUMNS.shop] = '';
  next[LOOKUP_COLUMNS.mobile] = '';
  next[LOOKUP_COLUMNS.key] = '';
  return next;
}

function decryptCustomerRow(row, { controlShop = CONTROL_SHOP } = {}) {
  if (!row || typeof row !== 'object') return row;
  const next = { ...row };
  ENCRYPTED_SHEET_FIELDS.forEach((field) => {
    if (next[field] == null || next[field] === '') return;
    if (field === 'Shop name' && String(next[field]).trim() === controlShop) return;
    try {
      next[field] = decryptText(String(next[field]));
    } catch {
      // leave ciphertext if this device cannot decrypt it
    }
  });
  if (next['Login Key'] && isEncryptedValue(next['Login Key'])) {
    try {
      next['Login Key'] = decryptText(String(next['Login Key']));
    } catch {
      // keep original
    }
  }
  return next;
}

function createEncryptedAsyncStorage(baseStorage) {
  return {
    getItem: async (name) => {
      const raw = await baseStorage.getItem(name);
      if (raw == null) return null;
      try {
        return decryptText(raw);
      } catch {
        return raw;
      }
    },
    setItem: async (name, value) => {
      await baseStorage.setItem(name, encryptText(value));
    },
    removeItem: (name) => baseStorage.removeItem(name),
  };
}

function encryptBackupPayload(payload) {
  return encryptText(JSON.stringify(payload));
}

function decryptBackupText(text) {
  const trimmed = String(text || '').trim();
  if (!trimmed) return trimmed;
  if (isEncryptedValue(trimmed)) return decryptText(trimmed);
  return trimmed;
}

function wrapSyncMessage(message, token) {
  return {
    _enc: 1,
    d: encryptText(JSON.stringify(message), `NAM_BILL_SYNC:${token}`),
  };
}

function unwrapSyncMessage(message, token) {
  if (!message || typeof message !== 'object') return message;
  if (message._enc === 1 && message.d) {
    try {
      return JSON.parse(decryptText(message.d, `NAM_BILL_SYNC:${token}`));
    } catch {
      return null;
    }
  }
  return message;
}

  global.NamSecureCrypto = {
    DATA_ENCRYPTION_SECRET,
    ENCRYPTED_SHEET_FIELDS,
    LOOKUP_COLUMNS,
    sha256Hex,
    hashesMatch,
    assertHttpsUrl,
    lookupToken,
    isEncryptedValue,
    encryptText,
    decryptText,
    encryptCustomerRow,
    decryptCustomerRow,
    encryptBackupPayload,
    decryptBackupText,
    wrapSyncMessage,
    unwrapSyncMessage
  };
})(typeof globalThis !== 'undefined' ? globalThis : window);
