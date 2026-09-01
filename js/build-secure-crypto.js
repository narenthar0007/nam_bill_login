const fs = require('fs');
const path = require('path');
const src = fs
  .readFileSync(path.join(__dirname, '../../utils/secureCrypto.js'), 'utf8')
  .replace(/^export /gm, '');
const out = `(function (global) {
${src}
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
`;
fs.writeFileSync(path.join(__dirname, 'secureCrypto.js'), out);
