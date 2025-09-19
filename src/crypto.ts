const crypto = globalThis.crypto || eval('require')('node:crypto').webcrypto;

export { crypto };
