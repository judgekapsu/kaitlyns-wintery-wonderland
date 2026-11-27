const SESSION_TTL_SECONDS = 2 * 60 * 60;
const LOGIN_WINDOW_SECONDS = 15 * 60;
const MAX_FAILED_LOGINS = 8;
const PASSWORD_ITERATIONS = 600000;
const MAX_CHALLENGE_AGE_MS = 5 * 60 * 1000;
const MAX_USERNAME_LENGTH = 128;
const MAX_PASSWORD_LENGTH = 256;
const DEFAULT_PRICE_CHECK_DOMAINS = [
  'amazon.com', 'sephora.com', 'ulta.com', 'kiarasky.com', 'apresnail.com',
  'macys.com', 'dillards.com', 'basspro.com', 'academy.com', 'coach.com'
];

const SHEET_HEADERS = {
  People: ['PersonID', 'Name', 'Email', 'CreatedDate', 'Active'],
  WishlistItems: ['ItemID', 'PersonID', 'ItemName', 'ItemType', 'ProductURL', 'Description', 'Size', 'ColorPattern', 'PreferredPrice', 'PriceTrackingEnabled', 'Status', 'DateAdded', 'Category'],
  PriceSubscriptions: ['SubscriptionID', 'ItemID', 'SubscriberEmail', 'Active', 'DateSubscribed', 'LastAlertDate', 'UnsubscribeToken'],
  PriceOffers: ['OfferID', 'ItemID', 'Retailer', 'ProductURL', 'CurrentPrice', 'Currency', 'LastChecked', 'Availability', 'MatchConfidence', 'OfferStatus'],
  PriceHistory: ['HistoryID', 'OfferID', 'CheckedDate', 'Price', 'Currency', 'Availability'],
  PurchaseActions: ['ActionID', 'ItemID', 'BuyerEmail', 'ActionType', 'ActionDate', 'ActionToken', 'TokenUsed']
};

const LIMITS = { name: 100, email: 254, itemName: 160, productUrl: 2048, description: 2000, size: 100, colorPattern: 150 };

function doGet(e) {
  const unsubscribeToken = e && e.parameter ? String(e.parameter.unsubscribe || '') : '';
  if (unsubscribeToken) return unsubscribeFromAlertLink_(unsubscribeToken);
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
  const schedule = ensureDailyPriceCheckTrigger_();
  try {
    SpreadsheetApp.getUi().alert('Wishlist setup complete', result + '\n' + schedule, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    Logger.log('Wishlist setup complete: ' + result + ' ' + schedule);
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
  const subscriptions = readRows_(spreadsheet.getSheetByName('PriceSubscriptions'));
  const offers = readRows_(spreadsheet.getSheetByName('PriceOffers'));
  const counts = {};
  subscriptions.forEach(function (subscription) {
    if (!isActive_(subscription.Active) || !subscription.ItemID || !isValidEmail_(subscription.SubscriberEmail)) return;
    const itemId = String(subscription.ItemID);
    if (!counts[itemId]) counts[itemId] = {};
    counts[itemId][String(subscription.SubscriberEmail).trim().toLowerCase()] = true;
  });
  const subscriberCounts = {};
  Object.keys(counts).forEach(function (itemId) { subscriberCounts[itemId] = Object.keys(counts[itemId]).length; });
  const latestOffers = {};
  offers.forEach(function (offer) {
    const key = String(offer.ItemID || '') + '|' + String(offer.ProductURL || '');
    if (key.charAt(0) === '|') return;
    const current = latestOffers[key];
    if (!current || dateValue_(offer.LastChecked) >= dateValue_(current.LastChecked)) latestOffers[key] = offer;
  });
  const visible = {};
  people.forEach(function (person) {
    if (isActive_(person.Active) && person.PersonID && person.Name) visible[String(person.PersonID)] = { name: String(person.Name), items: [] };
  });
  items.forEach(function (item) {
    const person = visible[String(item.PersonID)];
    // Older wishlist rows may not have a Status value. Treat blank status as
    // available, while keeping explicitly unavailable items out of the list.
    const status = String(item.Status || '').trim().toLowerCase();
    if (!person || (status && status !== 'available')) return;
    const itemId = String(item.ItemID || '');
    const offer = latestOffers[itemId + '|' + String(item.ProductURL || '')] || {};
    person.items.push({
      name: String(item.ItemName || ''), type: String(item.ItemType || ''),
      productUrl: isApprovedRetailerUrl_(item.ProductURL) ? String(item.ProductURL) : '',
      unapprovedProductUrl: Boolean(item.ProductURL) && safeHttpUrl_(item.ProductURL) && !isApprovedRetailerUrl_(item.ProductURL),
      description: String(item.Description || ''), size: String(item.Size || ''),
      colorPattern: String(item.ColorPattern || ''),
      preferredPrice: item.PreferredPrice === '' || item.PreferredPrice == null ? '' : String(item.PreferredPrice),
      itemId: itemId,
      currentPrice: offer.CurrentPrice === '' || offer.CurrentPrice == null ? '' : String(offer.CurrentPrice),
      currency: String(offer.Currency || ''),
      lastChecked: formatTimestamp_(offer.LastChecked),
      priceCheckStatus: String(offer.OfferStatus || ''),
      subscriberCount: subscriberCounts[itemId] || 0
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

function subscribeToPriceAlerts(itemId, email, sessionToken) {
  requireSession_(sessionToken);
  const id = cleanText_(itemId, 'Wishlist item', 100, true);
  const subscriberEmail = cleanText_(email, 'Email', LIMITS.email, true).toLowerCase();
  if (!isValidEmail_(subscriberEmail)) throw new Error('Enter a valid email address.');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    enforcePriceSubscriptionRateLimit_(subscriberEmail);
    const spreadsheet = getSpreadsheet_();
    const itemsSheet = spreadsheet.getSheetByName('WishlistItems');
    const subscriptionsSheet = spreadsheet.getSheetByName('PriceSubscriptions');
    if (!itemsSheet || !subscriptionsSheet) throw new Error('The database is not initialized. Ask the site owner for help.');
    const item = readRows_(itemsSheet).filter(function (row) { return String(row.ItemID) === id; })[0];
    if (!item || (String(item.Status || '').trim() && String(item.Status).trim().toLowerCase() !== 'available')) {
      throw new Error('This wishlist item is no longer available.');
    }
    if (String(item.ItemType || '') !== 'Product' || !isApprovedRetailerUrl_(item.ProductURL)) {
      throw new Error('Price alerts require a product URL on an approved retailer domain.');
    }

    const subscriptions = readRows_(subscriptionsSheet);
    let existing;
    subscriptions.forEach(function (row, index) {
      if (String(row.ItemID) === id && String(row.SubscriberEmail || '').trim().toLowerCase() === subscriberEmail) {
        existing = { row: row, rowNumber: index + 2 };
      }
    });
    const trackingColumn = SHEET_HEADERS.WishlistItems.indexOf('PriceTrackingEnabled') + 1;
    const itemRows = readRows_(itemsSheet);
    const itemIndex = itemRows.findIndex(function (row) { return String(row.ItemID) === id; });
    if (existing && isActive_(existing.row.Active)) {
      itemsSheet.getRange(itemIndex + 2, trackingColumn).setValue(true);
      return { success: true, message: 'This email is already subscribed.', subscriberCount: countActiveSubscribers_(subscriptions, id) };
    }

    const subscription = {
      SubscriptionID: existing && existing.row.SubscriptionID ? existing.row.SubscriptionID : Utilities.getUuid(),
      ItemID: id,
      SubscriberEmail: subscriberEmail,
      Active: true,
      DateSubscribed: new Date(),
      LastAlertDate: '',
      UnsubscribeToken: Utilities.getUuid() + Utilities.getUuid()
    };
    if (existing) {
      subscriptionsSheet.getRange(existing.rowNumber, 1, 1, SHEET_HEADERS.PriceSubscriptions.length)
        .setValues([SHEET_HEADERS.PriceSubscriptions.map(function (header) { return subscription[header] === undefined ? '' : subscription[header]; })]);
    } else {
      appendObject_(subscriptionsSheet, SHEET_HEADERS.PriceSubscriptions, subscription);
    }
    itemsSheet.getRange(itemIndex + 2, trackingColumn).setValue(true);
    return {
      success: true,
      message: 'Price alerts are on. The first successful check records a starting price; email alerts begin when the price changes.',
      subscriberCount: countActiveSubscribers_(subscriptions.concat([subscription]), id)
    };
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

function ensureDailyPriceCheckTrigger_() {
  const existing = ScriptApp.getProjectTriggers().filter(function (trigger) {
    return trigger.getHandlerFunction() === 'runDailyPriceChecks';
  });
  if (!existing.length) {
    ScriptApp.newTrigger('runDailyPriceChecks').timeBased().everyDays(1).atHour(8).create();
    return 'Daily price checks are scheduled for the 8 AM hour.';
  }
  existing.slice(1).forEach(function (trigger) { ScriptApp.deleteTrigger(trigger); });
  return 'Daily price checks are already scheduled for the 8 AM hour.';
}

function runDailyPriceChecks(e) {
  const triggerUid = e && e.triggerUid ? String(e.triggerUid) : '';
  const isInstalledTrigger = triggerUid && ScriptApp.getProjectTriggers().some(function (trigger) {
    return String(trigger.getUniqueId()) === triggerUid && trigger.getHandlerFunction() === 'runDailyPriceChecks';
  });
  if (!isInstalledTrigger) throw new Error('Price checks can only run from the installed daily trigger.');

  const spreadsheet = getSpreadsheet_();
  const subscriptions = readRows_(spreadsheet.getSheetByName('PriceSubscriptions'));
  const subscribedItems = {};
  subscriptions.forEach(function (subscription) {
    if (isActive_(subscription.Active) && isValidEmail_(subscription.SubscriberEmail)) subscribedItems[String(subscription.ItemID)] = true;
  });
  const items = readRows_(spreadsheet.getSheetByName('WishlistItems')).filter(function (item) {
    const status = String(item.Status || '').trim().toLowerCase();
    return isActive_(item.PriceTrackingEnabled) &&
      subscribedItems[String(item.ItemID)] &&
      String(item.ItemType || '') === 'Product' &&
      isApprovedRetailerUrl_(item.ProductURL) &&
      (!status || status === 'available');
  });
  items.forEach(function (item) {
    let result;
    let change;
    try {
      result = fetchProductPrice_(String(item.ProductURL));
      if (!result) {
        recordPriceCheckFailure_(item, 'No supported product price metadata was found.');
        return;
      }
      change = recordPriceCheck_(item, result);
    } catch (error) {
      const message = error && error.message ? String(error.message).substring(0, 180) : 'The product page could not be checked.';
      recordPriceCheckFailure_(item, message);
      Logger.log('Price check failed for item ' + String(item.ItemID) + ': ' + message);
      return;
    }
    if (change.priceChanged) {
      try {
        notifyPriceSubscribers_(item, result, change.subscriptions);
      } catch (error) {
        Logger.log('Price alert delivery failed for item ' + String(item.ItemID) + ': ' +
          (error && error.message ? error.message : 'unknown email error'));
      }
    }
  });
}

function fetchProductPrice_(url) {
  let target = url;
  for (let redirects = 0; redirects <= 3; redirects++) {
    if (!isApprovedRetailerUrl_(target)) throw new Error('The product page redirected to an unapproved retailer.');
    const response = UrlFetchApp.fetch(target, {
      followRedirects: false,
      muteHttpExceptions: true,
      headers: { Accept: 'text/html,application/xhtml+xml' }
    });
    const status = response.getResponseCode();
    if (status >= 300 && status < 400) {
      const location = getHeader_(response.getHeaders(), 'location');
      if (!location || redirects === 3) throw new Error('The product page redirected too many times.');
      target = resolveRedirectUrl_(location, target);
      continue;
    }
    if (status < 200 || status >= 300) throw new Error('The product page returned HTTP ' + status + '.');
    return extractProductPrice_(response.getContentText());
  }
  throw new Error('The product page redirected too many times.');
}

function extractProductPrice_(html) {
  const scripts = String(html || '').match(/<script\b[^>]*type\s*=\s*["']application\/ld\+json["'][^>]*>[\s\S]*?<\/script\s*>/gi) || [];
  for (let i = 0; i < scripts.length; i++) {
    const body = scripts[i].replace(/^<script\b[^>]*>/i, '').replace(/<\/script\s*>$/i, '');
    try {
      const offer = findProductOffer_(JSON.parse(body));
      if (offer) return { price: offer.price, currency: offer.currency, availability: offer.availability, confidence: 'High' };
    } catch (error) {
      // Other structured data and metadata may still provide a usable price.
    }
  }

  const amount = findMetaValue_(html, ['product:price:amount', 'og:price:amount', 'price']);
  const price = parsePrice_(amount);
  if (price === null) return null;
  return {
    price: price,
    currency: findMetaValue_(html, ['product:price:currency', 'og:price:currency', 'pricecurrency']),
    availability: findMetaValue_(html, ['product:availability', 'og:availability', 'availability']),
    confidence: 'Medium'
  };
}

function findProductOffer_(node) {
  if (Array.isArray(node)) {
    for (let i = 0; i < node.length; i++) {
      const result = findProductOffer_(node[i]);
      if (result) return result;
    }
    return null;
  }
  if (!node || typeof node !== 'object') return null;

  const types = Array.isArray(node['@type']) ? node['@type'] : [node['@type']];
  const isProduct = types.some(function (type) { return String(type || '').toLowerCase() === 'product'; });
  if (node.offers) {
    const offerNodes = Array.isArray(node.offers) ? node.offers : [node.offers];
    for (let i = 0; i < offerNodes.length; i++) {
      const result = readOfferPrice_(offerNodes[i]);
      if (result) return result;
    }
  }
  if (isProduct) {
    const productPrice = readOfferPrice_(node);
    if (productPrice) return productPrice;
  }
  return Object.keys(node).filter(function (key) { return key !== 'offers'; }).reduce(function (found, key) {
    return found || findProductOffer_(node[key]);
  }, null);
}

function readOfferPrice_(offer) {
  if (!offer || typeof offer !== 'object') return null;
  const priceSpecification = offer.priceSpecification || {};
  const value = offer.price != null ? offer.price :
    (offer.lowPrice != null ? offer.lowPrice : priceSpecification.price);
  const price = parsePrice_(value);
  if (price === null) return null;
  const rawAvailability = String(offer.availability || '');
  return {
    price: price,
    currency: String(offer.priceCurrency || priceSpecification.priceCurrency || ''),
    availability: rawAvailability ? rawAvailability.split('/').pop() : '',
    confidence: 'High'
  };
}

function findMetaValue_(html, acceptedNames) {
  const tags = String(html || '').match(/<meta\b[^>]*>/gi) || [];
  for (let i = 0; i < tags.length; i++) {
    const attributes = {};
    tags[i].replace(/([a-zA-Z_:][-a-zA-Z0-9_:.]*)\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+))/g, function (match, name, doubleQuoted, singleQuoted, unquoted) {
      attributes[name.toLowerCase()] = doubleQuoted || singleQuoted || unquoted || '';
      return match;
    });
    const name = String(attributes.property || attributes.name || attributes.itemprop || '').toLowerCase();
    if (acceptedNames.indexOf(name) !== -1 && attributes.content) return attributes.content.trim();
  }
  return '';
}

function parsePrice_(value) {
  const text = String(value == null ? '' : value).trim();
  if (!/^(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d{1,2})?$/.test(text)) return null;
  const price = Number(text.replace(/,/g, ''));
  return Number.isFinite(price) && price > 0 ? price : null;
}

function priceChanged_(previous, current) {
  if (!previous) return false;
  const previousPrice = parsePrice_(previous.CurrentPrice);
  return previousPrice !== null &&
    (previousPrice !== current.price ||
      String(previous.Currency || '').toUpperCase() !== String(current.currency || '').toUpperCase());
}

function recordPriceCheck_(item, result) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const spreadsheet = getSpreadsheet_();
    const offersSheet = spreadsheet.getSheetByName('PriceOffers');
    const historySheet = spreadsheet.getSheetByName('PriceHistory');
    const subscriptionsSheet = spreadsheet.getSheetByName('PriceSubscriptions');
    if (!offersSheet || !historySheet || !subscriptionsSheet) throw new Error('Price tracking tabs are missing. Run the owner setup first.');
    const offers = readRows_(offersSheet);
    const offerIndex = offers.findIndex(function (offer) {
      return String(offer.ItemID) === String(item.ItemID) && String(offer.ProductURL) === String(item.ProductURL);
    });
    const previous = offerIndex < 0 ? null : offers[offerIndex];
    const priceChanged = priceChanged_(previous, result);
    const offerId = previous ? previous.OfferID : Utilities.getUuid();
    const now = new Date();
    const offer = {
      OfferID: offerId,
      ItemID: item.ItemID,
      Retailer: retailerName_(item.ProductURL),
      ProductURL: item.ProductURL,
      CurrentPrice: result.price,
      Currency: result.currency,
      LastChecked: now,
      Availability: result.availability,
      MatchConfidence: result.confidence,
      OfferStatus: 'OK'
    };
    if (previous) {
      offersSheet.getRange(offerIndex + 2, 1, 1, SHEET_HEADERS.PriceOffers.length)
        .setValues([SHEET_HEADERS.PriceOffers.map(function (header) { return offer[header] === undefined ? '' : offer[header]; })]);
    } else {
      appendObject_(offersSheet, SHEET_HEADERS.PriceOffers, offer);
    }
    appendObject_(historySheet, SHEET_HEADERS.PriceHistory, {
      HistoryID: Utilities.getUuid(), OfferID: offerId, CheckedDate: now,
      Price: result.price, Currency: result.currency, Availability: result.availability
    });
    return {
      priceChanged: priceChanged,
      subscriptions: priceChanged ? readRows_(subscriptionsSheet).filter(function (subscription) {
        return String(subscription.ItemID) === String(item.ItemID) && isActive_(subscription.Active) &&
          isValidEmail_(subscription.SubscriberEmail);
      }) : []
    };
  } finally {
    lock.releaseLock();
  }
}

function recordPriceCheckFailure_(item, message) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSpreadsheet_().getSheetByName('PriceOffers');
    if (!sheet) throw new Error('The PriceOffers tab is missing. Run the owner setup first.');
    const offers = readRows_(sheet);
    const index = offers.findIndex(function (offer) {
      return String(offer.ItemID) === String(item.ItemID) && String(offer.ProductURL) === String(item.ProductURL);
    });
    const previous = index < 0 ? {} : offers[index];
    const offer = {
      OfferID: previous.OfferID || Utilities.getUuid(),
      ItemID: item.ItemID,
      Retailer: retailerName_(item.ProductURL),
      ProductURL: item.ProductURL,
      CurrentPrice: previous.CurrentPrice || '',
      Currency: previous.Currency || '',
      LastChecked: new Date(),
      Availability: previous.Availability || '',
      MatchConfidence: previous.MatchConfidence || '',
      OfferStatus: 'Check failed: ' + message
    };
    if (index < 0) appendObject_(sheet, SHEET_HEADERS.PriceOffers, offer);
    else sheet.getRange(index + 2, 1, 1, SHEET_HEADERS.PriceOffers.length)
      .setValues([SHEET_HEADERS.PriceOffers.map(function (header) { return offer[header] === undefined ? '' : offer[header]; })]);
  } finally {
    lock.releaseLock();
  }
}

function notifyPriceSubscribers_(item, result, subscriptions) {
  const webAppUrl = ScriptApp.getService().getUrl();
  if (!webAppUrl) throw new Error('The web app URL is unavailable; update the deployment before sending alerts.');
  const recipients = {};
  subscriptions.forEach(function (subscription) {
    const email = String(subscription.SubscriberEmail || '').trim().toLowerCase();
    if (!email || recipients[email]) return;
    recipients[email] = true;
    const unsubscribeUrl = webAppUrl + '?unsubscribe=' + encodeURIComponent(String(subscription.UnsubscribeToken || ''));
    const price = result.currency ? result.currency + ' ' + result.price.toFixed(2) : result.price.toFixed(2);
    const body = 'The price changed for "' + String(item.ItemName || 'a wishlist item') + '".\n\n' +
      'Current price: ' + price + '\n' +
      'Product: ' + String(item.ProductURL) + '\n\n' +
      'Unsubscribe from this item: ' + unsubscribeUrl;
    try {
      MailApp.sendEmail(email, 'Wishlist price changed: ' + String(item.ItemName || 'item').replace(/[\r\n]/g, ' '), body);
      updateSubscriptionAlertDate_(subscription.SubscriptionID);
    } catch (error) {
      Logger.log('Price alert email failed for subscription ' + String(subscription.SubscriptionID) + ': ' +
        (error && error.message ? error.message : 'unknown email error'));
    }
  });
}

function updateSubscriptionAlertDate_(subscriptionId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSpreadsheet_().getSheetByName('PriceSubscriptions');
    if (!sheet) throw new Error('The PriceSubscriptions tab is missing.');
    const subscriptions = readRows_(sheet);
    const index = subscriptions.findIndex(function (subscription) { return String(subscription.SubscriptionID) === String(subscriptionId); });
    if (index >= 0) {
      const column = SHEET_HEADERS.PriceSubscriptions.indexOf('LastAlertDate') + 1;
      sheet.getRange(index + 2, column).setValue(new Date());
    }
  } finally {
    lock.releaseLock();
  }
}

function unsubscribeFromAlertLink_(token) {
  if (token.length > 100 || !/^[a-f0-9-]+$/i.test(token)) {
    return HtmlService.createHtmlOutput('<p>This unsubscribe link is invalid.</p>').setTitle('Price alerts');
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const spreadsheet = getSpreadsheet_();
    const subscriptionsSheet = spreadsheet.getSheetByName('PriceSubscriptions');
    const itemsSheet = spreadsheet.getSheetByName('WishlistItems');
    if (!subscriptionsSheet || !itemsSheet) throw new Error('Price alert data is unavailable.');
    const subscriptions = readRows_(subscriptionsSheet);
    const index = subscriptions.findIndex(function (subscription) {
      return String(subscription.UnsubscribeToken || '') === token && isActive_(subscription.Active);
    });
    if (index < 0) return HtmlService.createHtmlOutput('<p>This price alert is already unsubscribed or the link has expired.</p>').setTitle('Price alerts');
    subscriptionsSheet.getRange(index + 2, SHEET_HEADERS.PriceSubscriptions.indexOf('Active') + 1).setValue(false);
    const itemId = String(subscriptions[index].ItemID);
    const stillSubscribed = subscriptions.some(function (subscription, subscriptionIndex) {
      return subscriptionIndex !== index && String(subscription.ItemID) === itemId && isActive_(subscription.Active);
    });
    if (!stillSubscribed) {
      const items = readRows_(itemsSheet);
      const itemIndex = items.findIndex(function (item) { return String(item.ItemID) === itemId; });
      if (itemIndex >= 0) {
        itemsSheet.getRange(itemIndex + 2, SHEET_HEADERS.WishlistItems.indexOf('PriceTrackingEnabled') + 1).setValue(false);
      }
    }
    return HtmlService.createHtmlOutput('<p>You have been unsubscribed from price alerts for this item.</p>').setTitle('Price alerts');
  } finally {
    lock.releaseLock();
  }
}

function countActiveSubscribers_(subscriptions, itemId) {
  const emails = {};
  subscriptions.forEach(function (subscription) {
    const email = String(subscription.SubscriberEmail || '').trim().toLowerCase();
    if (String(subscription.ItemID) === String(itemId) && isActive_(subscription.Active) && isValidEmail_(email)) emails[email] = true;
  });
  return Object.keys(emails).length;
}

function retailerName_(url) {
  const match = String(url || '').match(/^https?:\/\/([^/:]+)/i);
  return match ? match[1].replace(/^www\./i, '') : '';
}

function dateValue_(value) {
  const date = value instanceof Date ? value.getTime() : new Date(value).getTime();
  return Number.isFinite(date) ? date : 0;
}

function formatTimestamp_(value) {
  const time = dateValue_(value);
  return time ? new Date(time).toLocaleString() : '';
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
  // Apps Script's HMAC utility expects the message and key as strings here.
  // The verifier is already a fixed-width hex string; use that consistently
  // when issuing and validating the server-side challenge.
  return base64Url_(Utilities.computeHmacSha256Signature(payload, verifierHex, Utilities.Charset.UTF_8));
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
  const structured = extractProductPrice_('<script type="application/ld+json">{"@type":"Product","offers":{"@type":"Offer","price":"24.95","priceCurrency":"USD","availability":"https://schema.org/InStock"}}</script>');
  if (!structured || structured.price !== 24.95 || structured.currency !== 'USD' || structured.availability !== 'InStock') {
    throw new Error('Structured product price extraction test failed.');
  }
  const metadata = extractProductPrice_('<meta property="og:price:amount" content="1,249.99"><meta property="og:price:currency" content="USD">');
  if (!metadata || metadata.price !== 1249.99 || metadata.currency !== 'USD') throw new Error('Product metadata price extraction test failed.');
  if (extractProductPrice_('<html><title>No price here</title></html>') !== null) throw new Error('Unsupported product page test failed.');
  if (!safePriceCheckUrl_('https://www.shop.example.com/products/item') ||
      safePriceCheckUrl_('http://www.shop.example.com/products/item') ||
      safePriceCheckUrl_('https://127.0.0.1/products/item') ||
      safePriceCheckUrl_('https://shop.local/products/item')) throw new Error('Safe price-check URL test failed.');
  if (!isAllowedRetailerHost_('www.amazon.com', DEFAULT_PRICE_CHECK_DOMAINS) ||
      isAllowedRetailerHost_('amazn.com', DEFAULT_PRICE_CHECK_DOMAINS) ||
      isAllowedRetailerHost_('amazon.com.fake-shop.com', DEFAULT_PRICE_CHECK_DOMAINS)) {
    throw new Error('Approved retailer host test failed.');
  }
  if (resolveRedirectUrl_('?variant=blue', 'https://shop.example.com/products/item?old=1') !== 'https://shop.example.com/products/item?variant=blue' ||
      resolveRedirectUrl_('/new-item', 'https://shop.example.com/products/item') !== 'https://shop.example.com/new-item') {
    throw new Error('Product redirect resolution test failed.');
  }
  const currentPrice = { price: 25, currency: 'USD' };
  if (priceChanged_(null, currentPrice) ||
      priceChanged_({ CurrentPrice: '', Currency: '' }, currentPrice) ||
      priceChanged_({ CurrentPrice: 25, Currency: 'USD' }, currentPrice) ||
      !priceChanged_({ CurrentPrice: 24.99, Currency: 'USD' }, currentPrice)) {
    throw new Error('Price change detection test failed.');
  }
  if (countActiveSubscribers_([
    { ItemID: 'item-1', SubscriberEmail: 'mom@example.com', Active: true },
    { ItemID: 'item-1', SubscriberEmail: 'MOM@example.com', Active: true },
    { ItemID: 'item-1', SubscriberEmail: 'dad@example.com', Active: false }
  ], 'item-1') !== 1) throw new Error('Active subscriber count test failed.');
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

function safePriceCheckUrl_(value) {
  const text = String(value || '').trim();
  if (!/^https:\/\//i.test(text) || !safeHttpUrl_(text)) return false;
  const match = text.match(/^https:\/\/([^/:?#]+)(?::[0-9]{1,5})?(?:[/?#]|$)/i);
  if (!match) return false;
  const host = match[1].toLowerCase().replace(/\.$/, '');
  const labels = host.split('.');
  const lastLabel = labels[labels.length - 1];
  return labels.length >= 2 &&
    !/^\d{1,3}(?:\.\d{1,3}){3}$/.test(host) &&
    ['example', 'invalid', 'test', 'localhost', 'local', 'internal', 'lan', 'home'].indexOf(lastLabel) === -1;
}

function getAllowedRetailerDomains_() {
  const configured = PropertiesService.getScriptProperties().getProperty('PRICE_CHECK_ALLOWED_DOMAINS') || '';
  const domains = DEFAULT_PRICE_CHECK_DOMAINS.slice();
  configured.split(/[,\s;]+/).filter(Boolean).forEach(function (value) {
    const domain = value.toLowerCase().replace(/^www\./, '').replace(/\.$/, '');
    if (domain.length > 253 ||
        !/^(?:[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?$/i.test(domain)) {
      throw new Error('PRICE_CHECK_ALLOWED_DOMAINS must contain domain names only, separated by commas.');
    }
    if (domains.indexOf(domain) === -1) domains.push(domain);
  });
  return domains;
}

function isApprovedRetailerUrl_(value) {
  if (!safePriceCheckUrl_(value)) return false;
  const match = String(value).trim().match(/^https:\/\/([^/:?#]+)/i);
  if (!match) return false;
  const host = match[1].toLowerCase().replace(/\.$/, '');
  return isAllowedRetailerHost_(host, getAllowedRetailerDomains_());
}

function isAllowedRetailerHost_(host, domains) {
  const normalizedHost = String(host || '').toLowerCase().replace(/\.$/, '');
  return domains.some(function (domain) {
    return normalizedHost === domain || normalizedHost.endsWith('.' + domain);
  });
}

function getHeader_(headers, name) {
  const key = Object.keys(headers || {}).filter(function (header) { return header.toLowerCase() === name.toLowerCase(); })[0];
  if (!key) return '';
  const value = headers[key];
  return Array.isArray(value) ? String(value[0] || '') : String(value || '');
}

function resolveRedirectUrl_(location, baseUrl) {
  const value = String(location || '').trim();
  if (/^https?:\/\//i.test(value)) return value;
  if (value.indexOf('//') === 0) return 'https:' + value;
  const origin = String(baseUrl).match(/^https:\/\/[^/]+/i);
  if (!origin) throw new Error('The product page returned an invalid redirect.');
  if (value.charAt(0) === '/') return origin[0] + value;
  if (value.charAt(0) === '?') return String(baseUrl).replace(/[?#].*$/, '') + value;
  if (value.charAt(0) === '#') return String(baseUrl).replace(/#.*$/, '') + value;
  const path = String(baseUrl).replace(/[?#].*$/, '');
  return path.substring(0, path.lastIndexOf('/') + 1) + value;
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
  if (entry.productUrl && !isApprovedRetailerUrl_(entry.productUrl)) throw new Error('Product URLs must use an approved HTTPS retailer domain.');
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

function enforcePriceSubscriptionRateLimit_(email) {
  const cache = CacheService.getScriptCache();
  const key = 'price-subscribe:' + sha256Text_(email);
  const count = Number(cache.get(key) || 0);
  if (count >= 8) throw new Error('Too many subscription attempts for this email. Try again later.');
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
