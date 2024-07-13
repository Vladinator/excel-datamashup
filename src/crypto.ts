let WebCryptoImpl = globalThis.crypto;

/**
 * If the crypto is already available, it means we're running in the browser, so we simply
 * keep it. Otherwise, we load the Node crypto module.
 * This file will exclusively be used with decryption/encryption of permission bindings.
 */
if (typeof WebCryptoImpl === 'undefined') {
    const { webcrypto } = require('node:crypto');
    WebCryptoImpl = webcrypto;
}

export { WebCryptoImpl as crypto };
