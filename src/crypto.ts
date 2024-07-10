let WebCryptoImpl = globalThis.crypto;

if (typeof WebCryptoImpl === 'undefined') {
    const { webcrypto } = require('node:crypto');
    WebCryptoImpl = webcrypto;
}

export { WebCryptoImpl as crypto };
