#!/usr/bin/env node
'use strict';
const crypto = require('node:crypto');
const { StringDecoder } = require('node:string_decoder');
const ITERATIONS = 600000;
const SALT_BYTES = 16;
const KEY_BYTES = 32;
const TEST_PASSWORD = 'wishlist-test-password-v1';
const TEST_SALT_HEX = '000102030405060708090a0b0c0d0e0f';

function toHex(bytes) { return Buffer.from(bytes).toString('hex'); }
function makeRecord(derivedBytes, salt) {
  const verifier = crypto.createHash('sha256').update(derivedBytes).digest();
  return 'pbkdf2-sha256-client$' + ITERATIONS + '$' + salt.toString('hex') + '$' + toHex(verifier);
}
function deriveSync(password, salt, iterations) {
  return crypto.pbkdf2Sync(Buffer.from(password, 'utf8'), salt, iterations, KEY_BYTES, 'sha256');
}
async function selfTest() {
  const rfc = crypto.pbkdf2Sync('password', 'salt', 1, 32, 'sha256').toString('hex');
  const expected = '120fb6cffcf8b32c43e7225256c4f837a86548c92ccc35480805987cb70be17b';
  if (rfc !== expected) throw new Error('RFC PBKDF2 vector failed.');
  const salt = Buffer.from(TEST_SALT_HEX, 'hex');
  const syncKey = deriveSync(TEST_PASSWORD, salt, ITERATIONS);
  const webKey = await crypto.webcrypto.subtle.importKey('raw', Buffer.from(TEST_PASSWORD, 'utf8'), 'PBKDF2', false, ['deriveBits']);
  const webBits = await crypto.webcrypto.subtle.deriveBits({ name: 'PBKDF2', salt, iterations: ITERATIONS, hash: 'SHA-256' }, webKey, 256);
  if (!crypto.timingSafeEqual(syncKey, Buffer.from(webBits))) throw new Error('Node WebCrypto and pbkdf2Sync results differ.');
  const wrongKey = deriveSync('incorrect-test-password', salt, ITERATIONS);
  if (crypto.timingSafeEqual(syncKey, wrongKey)) throw new Error('Wrong-password test failed.');
  console.log('RFC PBKDF2 vector: PASS');
  console.log('Node WebCrypto matches pbkdf2Sync at 600000 iterations: PASS');
  console.log('Incorrect password derives a different key: PASS');
}
function makeTestVector() {
  const salt = Buffer.from(TEST_SALT_HEX, 'hex');
  const derived = deriveSync(TEST_PASSWORD, salt, ITERATIONS);
  console.log('PUBLIC TEST DATA ONLY — not a real credential');
  console.log('verifierRecord=' + makeRecord(derived, salt));
  console.log('derivedKeyHex=' + toHex(derived));
}
function hiddenPrompt(prompt) {
  return new Promise((resolve, reject) => {
    const input = process.stdin; const output = process.stdout; const decoder = new StringDecoder('utf8'); let value = '';
    if (!input.isTTY || typeof input.setRawMode !== 'function') return reject(new Error('Run in an interactive terminal.'));
    output.write(prompt); input.setRawMode(true); input.resume();
    function finish(error) { input.removeListener('data', onData); input.setRawMode(false); input.pause(); output.write('\n'); error ? reject(error) : resolve(value); }
    function onData(buffer) {
      for (const character of decoder.write(buffer)) {
        if (character === '\u0003') return finish(new Error('Cancelled.'));
        if (character === '\r' || character === '\n') return finish();
        if (character === '\u007f' || character === '\b') value = Array.from(value).slice(0, -1).join('');
        else if (character >= ' ') value += character;
      }
    }
    input.on('data', onData);
  });
}
(async function main() {
  try {
    if (process.argv.includes('--self-test')) return await selfTest();
    if (process.argv.includes('--test-vector')) return makeTestVector();
    const password = await hiddenPrompt('Enter your chosen password (input hidden): ');
    if (!password) throw new Error('Password cannot be empty.');
    if (Array.from(password).length > 256) throw new Error('Password must be no more than 256 characters.');
    const salt = crypto.randomBytes(SALT_BYTES);
    const derived = deriveSync(password, salt, ITERATIONS);
    console.log(makeRecord(derived, salt));
  } catch (error) {
    console.error(error.message);
    process.exitCode = 1;
  }
})();
