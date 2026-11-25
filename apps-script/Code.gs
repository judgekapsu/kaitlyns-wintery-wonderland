const SESSION_TTL_SECONDS = 2 * 60 * 60;
const LOGIN_WINDOW_SECONDS = 15 * 60;
const MAX_FAILED_LOGINS = 8;
const PASSWORD_ITERATIONS = 600000;
const MAX_CHALLENGE_AGE_MS = 5 * 60 * 1000;
const MAX_USERNAME_LENGTH = 128;
const MAX_PASSWORD_LENGTH = 256;

const SHEET_HEADERS = {
  People: ['PersonID', 'Name', 'Email', 'CreatedDate', 'Active'],
  WishlistItems: ['ItemID', 'PersonID', 'ItemName', 'ItemType', 'ProductURL', 'Description', 'Size', 'ColorPattern', 'PreferredPrice', 'PriceTrackingEnabled', 'Status', 'DateAdded', 'Category'],
  PriceSubscriptions: ['SubscriptionID', 'ItemID', 'SubscriberEmail', 'Active', 'DateSubscribed', 'LastAlertDate', 'UnsubscribeToken'],
  PriceOffers: ['OfferID', 'ItemID', 'Retailer', 'ProductURL', 'CurrentPrice', 'Currency', 'LastChecked', 'Availability', 'MatchConfidence', 'OfferStatus'],
  PriceHistory: ['HistoryID', 'OfferID', 'CheckedDate', 'Price', 'Currency', 'Availability'],
  PurchaseActions: ['ActionID', 'ItemID', 'BuyerEmail', 'ActionType', 'ActionDate', 'ActionToken', 'TokenUsed']
};

const LIMITS = { name: 100, email: 254, itemName: 160, productUrl: 2048, description: 2000, size: 100, colorPattern: 150 };

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle("Kaitlyn's Wintery Wonderland")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include_(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

function onOpen() {
  if (!isConfiguredOwner_()) return;
  SpreadsheetApp.getUi().createMenu('Wishlist admin')
    .addItem('Initialize database and connect existing items', 'runOwnerSetup')
    .addItem('Run compatibility checks', 'runOwnerCompatibilityTests')
    .addToUi();
}

function runOwnerSetup() {
  requireConfiguredOwner_();
  setupDatabase_();
  const result = connectOwnerSetupItems_();
  try {
    SpreadsheetApp.getUi().alert('Wishlist setup complete', result, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    Logger.log('Wishlist setup complete: ' + result);
  }
}

function runOwnerCompatibilityTests() {
  requireConfiguredOwner_();
  runCompatibilityTests_();
  try {
    SpreadsheetApp.getUi().alert('Compatibility checks passed', 'Verifier and product URL checks passed.', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    Logger.log('Compatibility checks passed.');
  }
}

function isConfiguredOwner_() {
  const expectedEmail = PropertiesService.getScriptProperties().getProperty('OWNER_EMAIL');
  if (!expectedEmail) return false;
  let activeEmail = '';
  try {
    activeEmail = Session.getActiveUser().getEmail();
  } catch (error) {
    return false;
  }
  return Boolean(activeEmail) && activeEmail.trim().toLowerCase() === expectedEmail.trim().toLowerCase();
}

function requireConfiguredOwner_() {
  if (!isConfiguredOwner_()) throw new Error('This owner-only action is unavailable for this account.');
}

/** Returns a signed, expiring challenge and the public PBKDF2 parameters. */
function beginLogin() {
  const record = parsePasswordRecord_();
  const payload = String(Date.now()) + '.' + Utilities.getUuid() + Utilities.getUuid();
  return {
    challenge: payload + '.' + signChallenge_(payload, record.verifierHex),
    saltHex: record.saltHex,
    iterations: PASSWORD_ITERATIONS
  };
}

/** derivedKeyHex is produced in the browser; plaintext password never reaches this server. */
function login(username, derivedKeyHex, challenge, proofCounter) {
  const user = typeof username === 'string' ? username : '';
  const derivedHex = typeof derivedKeyHex === 'string' ? derivedKeyHex : '';
  if (!user || user.length > MAX_USERNAME_LENGTH || !/^[0-9a-f]{64}$/i.test(derivedHex)) {
    throw new Error('Login failed. Check your credentials or try again later.');
  }

  const record = parsePasswordRecord_();
  if (record.username.length > MAX_USERNAME_LENGTH) throw new Error('Login is not configured.');
  const challengeData = validateChallenge_(challenge, record.verifierHex);
  if (!validProof_(challenge, proofCounter)) throw new Error('Login failed. Please try again.');

  const cache = CacheService.getScriptCache();
  const throttleKey = loginThrottleKey_();
  if (throttleKey && Number(cache.get(throttleKey) || 0) >= MAX_FAILED_LOGINS) {
    throw new Error('Too many failed attempts. Wait a few minutes and try again.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const replayKey = 'used-login:' + sha256Text_(challengeData.nonce);
    if (cache.get(replayKey)) throw new Error('Login failed. Please try again.');
    cache.put(replayKey, 'used', 600);
  } finally {
    lock.releaseLock();
  }

  const usernameMatches = constantTimeTextEqual_(user, record.username);
  const derivedBytes = hexToBytes_(derivedHex);
  const submittedVerifier = bytesToHex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, derivedBytes));
  const keyMatches = constantTimeTextEqual_(submittedVerifier, record.verifierHex);
  if (!usernameMatches || !keyMatches) {
    if (throttleKey) cache.put(throttleKey, String(Number(cache.get(throttleKey) || 0) + 1), LOGIN_WINDOW_SECONDS);
    throw new Error('Login failed. Check your credentials or try again later.');
  }

  if (throttleKey) cache.remove(throttleKey);
  const token = Utilities.getUuid() + Utilities.getUuid();
  cache.put(sessionKey_(token), 'valid', SESSION_TTL_SECONDS);
  return { sessionToken: token };
}

function getFamilyData(sessionToken) {
  requireSession_(sessionToken);
  const spreadsheet = getSpreadsheet_();
  const people = readRows_(spreadsheet.getSheetByName('People'));
  const items = readRows_(spreadsheet.getSheetByName('WishlistItems'));
  const visible = {};
  people.forEach(function (person) {
    if (isActive_(person.Active) && person.PersonID && person.Name) visible[String(person.PersonID)] = { name: String(person.Name), items: [] };
  });
  items.forEach(function (item) {
    const person = visible[String(item.PersonID)];
    if (!person || String(item.Status) !== 'Available') return;
    person.items.push({
      name: String(item.ItemName || ''), type: String(item.ItemType || ''),
      productUrl: safeHttpUrl_(item.ProductURL) ? String(item.ProductURL) : '',
      description: String(item.Description || ''), size: String(item.Size || ''),
      colorPattern: String(item.ColorPattern || ''),
      preferredPrice: item.PreferredPrice === '' || item.PreferredPrice == null ? '' : String(item.PreferredPrice)
    });
  });
  return Object.keys(visible).map(function (id) { return visible[id]; })
    .filter(function (person) { return person.items.length > 0; })
    .sort(function (a, b) { return a.name.localeCompare(b.name); });
}

function submitWishlistEntry(form, sessionToken) {
  requireSession_(sessionToken);
  const entry = validateEntry_(form);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    enforceEmailRateLimit_(entry.email);
    const spreadsheet = getSpreadsheet_();
    const peopleSheet = spreadsheet.getSheetByName('People');
    const itemsSheet = spreadsheet.getSheetByName('WishlistItems');
    if (!peopleSheet || !itemsSheet) throw new Error('The database is not initialized. Ask the site owner for help.');
    const matches = readRows_(peopleSheet).filter(function (row) { return String(row.Email || '').trim().toLowerCase() === entry.email; });
    if (matches.length > 1) throw new Error('This email has duplicate records. Please contact the site owner.');
    let person = matches[0];
    if (person && !isActive_(person.Active)) throw new Error('This email is associated with an inactive person record. Contact the site owner.');
    const now = new Date();
    if (!person) {
      person = { PersonID: Utilities.getUuid(), Name: entry.name, Email: entry.email, CreatedDate: now, Active: true };
      appendObject_(peopleSheet, SHEET_HEADERS.People, person);
    }
    appendObject_(itemsSheet, SHEET_HEADERS.WishlistItems, {
      ItemID: Utilities.getUuid(), PersonID: person.PersonID, ItemName: entry.itemName,
      ItemType: entry.itemType, ProductURL: entry.productUrl, Description: entry.description,
      Size: entry.size, ColorPattern: entry.colorPattern, PreferredPrice: entry.preferredPrice,
      PriceTrackingEnabled: false, Status: 'Available', DateAdded: now, Category: ''
    });
    return { success: true, message: 'Your wishlist item was added.' };
  } finally {
    lock.releaseLock();
  }
}

function logout(sessionToken) {
  if (typeof sessionToken === 'string' && sessionToken.length <= 200) CacheService.getScriptCache().remove(sessionKey_(sessionToken));
  return true;
}

function setupDatabase_() {
  const spreadsheet = getSpreadsheet_();
  Object.keys(SHEET_HEADERS).forEach(function (name) {
    let sheet = spreadsheet.getSheetByName(name);
    if (!sheet) sheet = spreadsheet.insertSheet(name);
    const expected = SHEET_HEADERS[name];
    const current = sheet.getRange(1, 1, 1, expected.length).getDisplayValues()[0];
    if (current.every(function (value) { return value === ''; })) {
      sheet.getRange(1, 1, 1, expected.length).setValues([expected]);
      sheet.setFrozenRows(1);
    } else if (expected.some(function (value, index) { return current[index] !== value; })) {
      throw new Error('Header mismatch in tab "' + name + '". No existing data was changed.');
    }
  });
}

function connectOwnerSetupItems_() {
  const props = PropertiesService.getScriptProperties();
  const ownerName = cleanText_(props.getProperty('OWNER_NAME'), 'OWNER_NAME', LIMITS.name, true);
  const ownerEmail = cleanText_(props.getProperty('OWNER_EMAIL'), 'OWNER_EMAIL', LIMITS.email, true).toLowerCase();
  if (!isValidEmail_(ownerEmail)) throw new Error('OWNER_EMAIL must be valid.');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const spreadsheet = getSpreadsheet_();
    const peopleSheet = spreadsheet.getSheetByName('People');
    const itemsSheet = spreadsheet.getSheetByName('WishlistItems');
    if (!peopleSheet || !itemsSheet) throw new Error('Initialize the database first.');
    const matches = readRows_(peopleSheet).filter(function (row) { return String(row.Email || '').trim().toLowerCase() === ownerEmail; });
    if (matches.length > 1) throw new Error('Duplicate owner email rows need manual review.');
    let owner = matches[0];
    if (owner && (!owner.PersonID || !isActive_(owner.Active))) throw new Error('The owner row is inactive or missing an ID. Review it manually.');
    if (!owner) {
      owner = { PersonID: Utilities.getUuid(), Name: ownerName, Email: ownerEmail, CreatedDate: new Date(), Active: true };
      appendObject_(peopleSheet, SHEET_HEADERS.People, owner);
    }
    const column = SHEET_HEADERS.WishlistItems.indexOf('PersonID') + 1;
    const lastRow = itemsSheet.getLastRow();
    if (lastRow < 2) return 'No wishlist items found.';
    const ids = itemsSheet.getRange(2, column, lastRow - 1, 1).getValues();
    let count = 0;
    ids.forEach(function (row, index) {
      if (String(row[0]) === 'OWNER-SETUP') { itemsSheet.getRange(index + 2, column).setValue(owner.PersonID); count++; }
    });
    return count + ' placeholder item(s) connected. Other rows and item data were not changed.';
  } finally {
    lock.releaseLock();
  }
}

function parsePasswordRecord_() {
  const props = PropertiesService.getScriptProperties();
  const username = props.getProperty('FAMILY_USERNAME') || '';
  const text = props.getProperty('FAMILY_PASSWORD_HASH') || '';
  const parts = text.split('$');
  if (!username || username.length > MAX_USERNAME_LENGTH || parts.length !== 4 || parts[0] !== 'pbkdf2-sha256-client' || Number(parts[1]) !== PASSWORD_ITERATIONS || !/^[0-9a-f]{32}$/i.test(parts[2]) || !/^[0-9a-f]{64}$/i.test(parts[3])) {
    throw new Error('Login is not configured.');
  }
  return { username: username, saltHex: parts[2].toLowerCase(), verifierHex: parts[3].toLowerCase() };
}

function requireSession_(token) {
  if (typeof token !== 'string' || token.length > 200) throw new Error('Your session expired. Please log in again.');
  const cache = CacheService.getScriptCache();
  const key = sessionKey_(token);
  if (cache.get(key) !== 'valid') throw new Error('Your session expired. Please log in again.');
  cache.put(key, 'valid', SESSION_TTL_SECONDS);
}

function validateChallenge_(challenge, verifierHex) {
  if (typeof challenge !== 'string' || challenge.length > 300) throw new Error('Login failed. Please try again.');
  const parts = challenge.split('.');
  if (parts.length !== 3) throw new Error('Login failed. Please try again.');
  const payload = parts[0] + '.' + parts[1];
  const time = Number(parts[0]);
  if (!Number.isFinite(time) || Math.abs(Date.now() - time) > MAX_CHALLENGE_AGE_MS) throw new Error('Login failed. Please try again.');
  if (!constantTimeTextEqual_(parts[2], signChallenge_(payload, verifierHex))) throw new Error('Login failed. Please try again.');
  return { nonce: parts[1] };
}

function signChallenge_(payload, verifierHex) {
  const key = hexToBytes_(verifierHex);
  return base64Url_(Utilities.computeHmacSha256Signature(payload, key));
}

function validProof_(challenge, counter) {
  const number = Number(counter);
  if (!Number.isSafeInteger(number) || number < 0 || number > 10000000) return false;
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, challenge + ':' + number, Utilities.Charset.UTF_8);
  return digest[0] === 0 && (digest[1] & 192) === 0;
}

function loginThrottleKey_() {
  try {
    const userKey = Session.getTemporaryActiveUserKey();
    return userKey ? 'login-fail:' + sha256Text_(userKey) : '';
  } catch (error) {
    return '';
  }
}

function runCompatibilityTests_() {
  // Deterministic verifier-function fixture: 32 zero bytes; SHA-256 digest below.
  // This checks the cheap server-side comparison, not the PBKDF2 derivation.
  const verifier = '66687aadf862bd776c8fc18b8e9f8e20089714856ee233b3902a591d0d5f2925';
  const validKey = '0'.repeat(64);
  const badKey = '1'.repeat(64);
  if (!verifyDerivedKey_(validKey, verifier)) throw new Error('Verifier acceptance test failed.');
  if (verifyDerivedKey_(badKey, verifier)) throw new Error('Verifier rejection test failed.');
  [
    ['https://shop.example/products/blue-scarf', true],
    ['https://shop.example/search?q=blue%20scarf&sort=price', true],
    ['https://shop.example:8443/products/item?color=blue', true],
    ['http://shop.example:80/item', true],
    ['ftp://shop.example/item', false],
    ['https://shop.example:99999/item', false],
    ['https:// bad.example/item', false]
  ].forEach(function (test) {
    if (safeHttpUrl_(test[0]) !== test[1]) throw new Error('URL validation test failed.');
  });
}

function verifyDerivedKey_(derivedHex, expectedVerifierHex) {
  if (typeof derivedHex !== 'string' || !/^[0-9a-f]{64}$/i.test(derivedHex) || !/^[0-9a-f]{64}$/i.test(expectedVerifierHex)) return false;
  const actual = bytesToHex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, hexToBytes_(derivedHex)));
  return constantTimeTextEqual_(actual.toLowerCase(), expectedVerifierHex.toLowerCase());
}

function safeHttpUrl_(value) {
  const text = String(value || '').trim();
  if (!text || /[\u0000-\u0020\u007f]/.test(text)) return false;
  const match = text.match(/^https?:\/\/(\[[0-9a-f:.]+\]|(?:[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)*[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?)(?::([0-9]{1,5}))?(?:[/?#][^\s]*)?$/i);
  return Boolean(match) && (!match[2] || (Number(match[2]) >= 1 && Number(match[2]) <= 65535));
}

function validateEntry_(form) {
  if (!form || typeof form !== 'object') throw new Error('Complete the form and try again.');
  const entry = {
    name: cleanText_(form.name, 'Name', LIMITS.name, true),
    email: cleanText_(form.email, 'Email', LIMITS.email, true).toLowerCase(),
    itemName: cleanText_(form.itemName, 'Item name', LIMITS.itemName, true),
    itemType: String(form.itemType || '').trim(),
    productUrl: cleanText_(form.productUrl, 'Product URL', LIMITS.productUrl, false),
    description: cleanText_(form.description, 'Description', LIMITS.description, false),
    size: cleanText_(form.size, 'Size', LIMITS.size, false),
    colorPattern: cleanText_(form.colorPattern, 'Color or pattern', LIMITS.colorPattern, false)
  };
  const price = String(form.preferredPrice == null ? '' : form.preferredPrice).trim();
  if (!isValidEmail_(entry.email)) throw new Error('Enter a valid email address.');
  if (entry.itemType !== 'Product' && entry.itemType !== 'General Idea') throw new Error('Choose Product or General Idea.');
  if (entry.productUrl && !safeHttpUrl_(entry.productUrl)) throw new Error('Product URL must be a valid HTTP or HTTPS URL.');
  if (price && (!/^\d{1,7}(\.\d{1,2})?$/.test(price) || Number(price) <= 0)) throw new Error('Preferred price must be positive with up to two decimals.');
  entry.preferredPrice = price ? Number(price) : '';
  return entry;
}

function cleanText_(value, label, maxLength, required) {
  const text = String(value == null ? '' : value).trim();
  if (required && !text) throw new Error(label + ' is required.');
  if (text.length > maxLength) throw new Error(label + ' is too long.');
  return text;
}

function isValidEmail_(email) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email || '')); }

function enforceEmailRateLimit_(email) {
  const cache = CacheService.getScriptCache();
  const key = 'submit:' + sha256Text_(email);
  const count = Number(cache.get(key) || 0);
  if (count >= 8) throw new Error('Too many submissions in a short time. Try again later.');
  cache.put(key, String(count + 1), 600);
}

function getSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('The site owner has not configured the spreadsheet.');
  return SpreadsheetApp.openById(id);
}

function readRows_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(String);
  return values.slice(1).filter(function (row) {
    return row.some(function (value) { return value !== ''; });
  }).map(function (row) {
    const object = {};
    headers.forEach(function (header, index) { object[header] = row[index]; });
    return object;
  });
}

function appendObject_(sheet, headers, object) {
  sheet.appendRow(headers.map(function (header) { return object[header] === undefined ? '' : object[header]; }));
}
function isActive_(value) { return value === true || String(value).toLowerCase() === 'true'; }
function sessionKey_(token) { return 'session:' + sha256Text_(token); }
function sha256Text_(text) { return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(text), Utilities.Charset.UTF_8)); }
function base64Url_(bytes) { return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/g, ''); }
function hexToBytes_(hex) {
  const bytes = [];
  for (let i = 0; i < hex.length; i += 2) {
    const value = parseInt(hex.substring(i, i + 2), 16);
    bytes.push(value > 127 ? value - 256 : value);
  }
  return bytes;
}
function bytesToHex_(bytes) {
  return bytes.map(function (value) { return ('0' + ((value + 256) % 256).toString(16)).slice(-2); }).join('');
}
function constantTimeBytesEqual_(left, right) {
  if (left.length !== right.length) return false;
  let diff = 0;
  for (let i = 0; i < left.length; i++) diff |= left[i] ^ right[i];
  return diff === 0;
}
function constantTimeTextEqual_(left, right) {
  return constantTimeBytesEqual_(Utilities.newBlob(left, 'text/plain').getBytes(), Utilities.newBlob(right, 'text/plain').getBytes());
}
