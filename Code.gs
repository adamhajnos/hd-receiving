/**
 * HD Backstage: back end.
 * Bind this script to the store's Google Sheet (Extensions > Apps Script).
 * Deploy as Web app: Execute as "Me", access "Anyone". See ../README.md.
 *
 * Security model: every request needs a staff PIN listed in the "Staff" tab.
 * The server looks up the name from the PIN and stamps it on every entry.
 * The tablet can only append rows and read a PII-free "today" feed.
 * Customer names and phone numbers are never sent back to the tablet.
 */

// The sheet this script works on. Leave blank when the script is bound to the sheet (Extensions > Apps Script).
// If the script is a standalone project, paste the sheet's ID here (the long part of its URL).
var SHEET_ID = '';
function ss_() { return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActive(); }

var TABS = {
  Staff:    ['Name', 'PIN'],
  Trove:    ['Artist', 'Title', 'Location', 'Price', 'Cover URL', 'Code', 'Genre', 'Status', 'Listed'],
  Log:      ['Timestamp', 'Date', 'Who', 'Task', 'Summary', 'Data (JSON)'],
  Opening:  ['Timestamp', 'Who', 'Amounts', 'Steps done', 'Could not do (step: reason)'],
  Closing:  ['Timestamp', 'Who', 'Amounts', 'Steps done', 'Could not do (step: reason)'],
  Issues:   ['Timestamp', 'Who', 'Type', 'Details', 'Photo', 'Status'],
  Cash:     ['Timestamp', 'Who', 'Drawer count', 'Deposit', 'Note'],
  Records: ['Timestamp', 'Who', 'Type', 'Artist', 'Album', 'Genre', 'Format', 'Barcode', 'MSRP', 'Price', 'Status', 'Location', 'Listed on', 'Label queued', 'Notes'],
  Expected: ['Invoice', 'Invoice date', 'PO', 'Product #', 'UPC', 'Artist', 'Title', 'MSRP', 'Qty', 'Received', 'Alerted', 'Unit cost'],
  UpcMap: ['Scanned', 'Expected UPC', 'Who', 'When'],
  Receiving: ['Timestamp', 'Who', 'Order', 'Expected', 'Scanned', 'Short', 'Over', 'Not on order', 'Packing list'],
  DropOffs: ['Timestamp', 'Who', 'Customer', 'Phone', 'Email', 'What they dropped off', 'Approx. records'],
  SquareSync: ['Timestamp', 'UPC', 'Artist', 'Title', 'Qty', 'Status', 'Note', 'Key', 'MSRP', 'Product #', 'Unit cost']
};

/** Run once from the editor to create the tabs. */
function setup() {
  var ss = ss_();
  Object.keys(TABS).forEach(function (name) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, TABS[name].length).setValues([TABS[name]]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  });
  // Keep PINs as text so a leading zero (0420) is not dropped.
  ss.getSheetByName('Staff').getRange('B:B').setNumberFormat('@');
  Logger.log('Now add one row per person to the Staff tab: Name, 4-digit PIN (unique).');
}

function doPost(e) {
  var out;
  try {
    var req = JSON.parse(e.postData.contents);
    var who = authenticate_(req.pin);
    if (req.action === 'login') out = { ok: true, name: who };
    else if (req.action === 'trove') out = { ok: true, items: troveItems_() };
    else if (req.action === 'today') out = { ok: true, entries: todayEntries_(), note: latestNote_() };
    else if (req.action === 'log') { var res = logEntry_(req, who) || {}; out = { ok: true, code: res.code || '' }; }
    else if (req.action === 'expected') out = { ok: true, lines: expectedLines_(), map: upcMap_() };
    else if (req.action === 'learnupc') out = { ok: true, saved: learnUpc_(req, who) };
    else if (req.action === 'scaninvoice') out = { ok: true, result: scanInvoice_(req) };
    else if (req.action === 'trovedel') out = { ok: true, removed: deleteTrove_(req, who) };
    else throw new Error('Unknown action');
  } catch (err) {
    out = { ok: false, error: String(err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Reads a keg invoice photo with the Google Gemini API (free tier) and returns {supplier, invoice, kegs:[{brewery,beer,size,cost}]}.
 * Script Property needed (Project Settings > Script properties): GEMINI_KEY (optional GEMINI_MODEL, default gemini-2.5-flash).
 * Only a signed-in staff PIN reaches this. Capped at 40 scans a day.
 */
function scanInvoice_(req) {
  var key = PropertiesService.getScriptProperties().getProperty('GEMINI_KEY');
  if (!key) throw new Error('Invoice reading is not set up yet');
  var img = String(req.image || '').replace(/^data:image\/\w+;base64,/, '');
  if (!img || img.length > 6000000) throw new Error('Photo missing or too large');
  var cache = CacheService.getScriptCache(), day = 'scans' + Utilities.formatDate(new Date(), 'America/New_York', 'yyyyMMdd');
  var n = Number(cache.get(day) || 0);
  if (n >= 40) throw new Error('Too many invoice scans today');
  cache.put(day, String(n + 1), 86400);
  var prompt = 'This is a photo of a beer distributor delivery invoice. Return ONLY JSON: {"supplier":string,"invoice":string,"kegs":[{"brewery":string,"beer":string,"size":string,"cost":number}]}. ' +
    'Include only kegs being delivered: lines with a product and a price that are NOT crossed out or struck through. Exclude lines starting "Empty" (returned kegs), key deposits, deposits, totals. cost is the PRICE column, not the deposit. size like "1/6 bbl". ' +
    'Product text looks like "Brewery-Beer Name 1/6 bbl" and may be cut off at the edge of the photo. If a name is cut off or abbreviated, use the matching entry from this known list (its exact beer and brewery spelling) when one clearly matches; otherwise give what is visible. Known beers (beer | brewery):\n' +
    String(req.known || '').slice(0, 4000);
  if (req.mode === 'cans' || req.mode === 'snacks') prompt = itemsPrompt_(req.mode);
  var props = PropertiesService.getScriptProperties();
  var model = props.getProperty('GEMINI_MODEL') || 'gemini-2.5-flash';
  var res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { 'x-goog-api-key': key },
    payload: JSON.stringify({
      contents: [{ parts: [{ text: prompt }, { inline_data: { mime_type: 'image/jpeg', data: img } }] }],
      generationConfig: { responseMimeType: 'application/json', temperature: 0 }
    })
  });
  if (res.getResponseCode() !== 200) throw new Error('Reader error ' + res.getResponseCode());
  var cand = (JSON.parse(res.getContentText()).candidates || [])[0] || {};
  var txt = ((cand.content || {}).parts || []).map(function (c) { return c.text || ''; }).join('');
  var m = txt.match(/\{[\s\S]*\}/);
  if (!m) throw new Error('Could not read the invoice');
  return JSON.parse(m[0]);
}

/** Prompt for drink (cans) or snack delivery invoices. Returns {supplier, invoice, items:[{product,brand,type,qty,unit}]}. */
function itemsPrompt_(mode) {
  var snacks = mode === 'snacks';
  return 'This is a photo of a delivery invoice for ' + (snacks ? 'snacks (chips, candy, jerky and similar)' : 'drinks (beer cans, THC drinks, soda, water and similar)') + '. Return ONLY JSON: ' +
    '{"supplier":string,"invoice":string,"items":[{"product":string,"brand":string,"type":' + (snacks ? '"Chips"|"Candy"|"Jerky"|"Other"' : '"Beer"|"THC"|"Soda"|"Water"|"Other"') + ',"qty":number,"unit":' + (snacks ? '"Cases"|"Boxes"' : '"Cases"|"Packs"') + '}]}. ' +
    'One entry per product line being delivered. Skip lines that are crossed out or struck through, and skip deposits, fees, taxes, discounts, empties and totals. ' +
    'qty is how many cases or packs were shipped (the quantity column, not the price). product is the product name without the pack size. brand is the maker or brewery when the line shows one, otherwise an empty string. ' +
    'unit is "Cases" for case lines' + (snacks ? ' and "Boxes" for box lines.' : ' and "Packs" for 4-pack, 6-pack or similar lines.');
}

/** Returns the staff name for a PIN, or throws. Locks guessing after 8 misses in 5 minutes. */
function authenticate_(pin) {
  var cache = CacheService.getScriptCache();
  var fails = Number(cache.get('fails') || 0);
  if (fails >= 8) throw new Error('Too many tries, wait a minute');
  var p = String(pin == null ? '' : pin).trim();
  var sh = ss_().getSheetByName('Staff');
  if (!sh) throw new Error('Run setup() first');
  var rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 2).getDisplayValues() : [];
  for (var i = 0; i < rows.length; i++) {
    if (p && String(rows[i][1]).trim() === p && String(rows[i][0]).trim()) return String(rows[i][0]).trim();
  }
  cache.put('fails', String(fails + 1), 300);
  throw new Error('Wrong PIN');
}

/** The running Trove list: one row per record in the "Trove" tab. Rows without a Code get the next TT-0001 style code. */
function troveItems_() {
  var sh = sheet_(ss_(), 'Trove');
  if (sh.getLastRow() < 2) return [];
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    sh.getRange(1, 6, 1, 4).setValues([['Code', 'Genre', 'Status', 'Listed']]).setFontWeight('bold');
    var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues();
    var max = 0;
    rows.forEach(function (r) { var n = parseInt(String(r[5]).replace(/\D/g, ''), 10); if (n > max) max = n; });
    var dirty = false;
    rows.forEach(function (r) { if ((r[0] || r[1]) && !r[5]) { max++; r[5] = trCode_(max); dirty = true; } });
    if (dirty) sh.getRange(2, 6, rows.length, 1).setValues(rows.map(function (r) { return [r[5]]; }));
    return rows.filter(function (r) { return r[0] || r[1]; })
      .map(function (r) { return { code: String(r[5]), artist: String(r[0]), title: String(r[1]), location: String(r[2]), price: r[3], cover: r[4] === 'none' ? '' : String(r[4]), genre: String(r[6] || ''), status: String(r[7] || ''), listed: String(r[8] || '') }; });
  } finally {
    lock.releaseLock();
  }
}
function trCode_(n) { return 'TT-' + ('0000' + n).slice(-4); }

/** Adds a Trove record with the next free code. Call inside the lock. Returns the code. */
function addTrove_(ss, d) {
  var sh = sheet_(ss, 'Trove');
  var max = 0;
  if (sh.getLastRow() > 1) sh.getRange(2, 6, sh.getLastRow() - 1, 1).getValues().forEach(function (r) { var n = parseInt(String(r[0]).replace(/\D/g, ''), 10); if (n > max) max = n; });
  var code = trCode_(max + 1);
  sh.appendRow([clean_(d.artist), clean_(d.album), clean_(d.location), Number(d.price) || '', '', code, clean_(d.genre), clean_(d.status) || 'Active', clean_(d.listed)]);
  return code;
}

/**
 * Fills the "Cover URL" column for any Trove row that has none, using the free
 * iTunes Search API. Rows with no match are marked "none" so they are not retried.
 * Safe to re-run: it stops before Apps Script's 6-minute limit and picks up the rest next time.
 * Clear a cell to retry it, or paste a better URL by hand.
 */
function fillCovers() {
  var sh = sheet_(ss_(), 'Trove');
  var last = sh.getLastRow();
  if (last < 2) return;
  var rows = sh.getRange(2, 1, last - 1, 5).getValues();
  var start = Date.now();
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][4] || !(rows[i][0] || rows[i][1])) continue;
    if (Date.now() - start > 4.5 * 60 * 1000) break;
    sh.getRange(i + 2, 5).setValue(lookupCover_(String(rows[i][0]), String(rows[i][1])) || 'none');
    Utilities.sleep(3000); // stay under iTunes' rate limit
  }
}

function lookupCover_(artist, title) {
  try {
    var url = 'https://itunes.apple.com/search?media=music&entity=album&limit=5&term=' + encodeURIComponent(artist + ' ' + title);
    var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return '';
    var hits = JSON.parse(res.getContentText()).results || [];
    var a = artist.toLowerCase();
    var best = hits.filter(function (h) { return String(h.artistName).toLowerCase().indexOf(a) >= 0; })[0] || hits[0];
    return best && best.artworkUrl100 ? best.artworkUrl100.replace('100x100bb', '300x300bb') : '';
  } catch (e) {
    return '';
  }
}

/** Run once: checks for new covers every hour (a no-op when none are missing). */
function installCoverTrigger() {
  ScriptApp.newTrigger('fillCovers').timeBased().everyHours(1).create();
}

function logEntry_(req, who) {
  who = clean_(who);
  var task = clean_(req.task), summary = clean_(req.summary);
  var d = req.data || {};
  if (['opening', 'closing', 'cash', 'dropoff', 'request', 'note', 'issue', 'record', 'receive',
       'reqstatus', 'dropstatus', 'feature', 'supplies', 'keg', 'kegrecv', 'stockrecv'].indexOf(task) < 0) throw new Error('Unknown task');
  if (!who) throw new Error('Missing name');

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  var result = {};
  try {
    var ss = ss_();
    var now = new Date();
    var tz = ss.getSpreadsheetTimeZone();
    var date = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    sheet_(ss, 'Log').appendRow([now, date, who, task, summary, JSON.stringify(Object.assign({}, d, { photo: d.photo ? '(photo)' : '' }))]);

    if (task === 'opening' || task === 'closing') {
      var amounts = Object.keys(d.amounts || {}).map(function (k) { return clean_(k) + ': ' + money_(d.amounts[k]); }).join('; ');
      sheet_(ss, task === 'opening' ? 'Opening' : 'Closing').appendRow([now, who, amounts, (d.done || []).length,
        (d.cant || []).map(function (c) { return clean_(c.step) + ': ' + clean_(c.reason); }).join('; ')]);
    } else if (task === 'note') {
      if (!summary) throw new Error('Empty note');
    } else if (task === 'issue') {
      var url = d.photo ? saveIssuePhoto_(String(d.photo), now, tz) : '';
      sheet_(ss, 'Issues').appendRow([now, who, clean_(d.type), clean_(d.details), url, 'Open']);
    } else if (task === 'cash') {
      sheet_(ss, 'Cash').appendRow([now, who, money_(d.drawer), money_(d.deposit), clean_(d.note)]);
    } else if (task === 'request') {
      requestEntry_(d, who, now, tz);
    } else if (task === 'receive') {
      sqQueue_(ss, d.got, now);
      applyReceived_(ss, d.got);
      var pack = d.packing ? saveIssuePhoto_(String(d.packing), now, tz) : (d.noPacking ? 'None in box' : '');
      sheet_(ss, 'Receiving').appendRow([now, who, clean_(d.order), Number(d.expected) || 0, Number(d.scanned) || 0,
        (d.short || []).map(clean_).join('; '), (d.over || []).map(clean_).join('; '), (d.extra || []).map(clean_).join('; '), pack]);
    } else if (task === 'record') {
      if (!clean_(d.artist) || !clean_(d.album) || !(Number(d.price) > 0)) throw new Error('Missing record details');
      if (clean_(d.type) === 'Trove') { result.code = addTrove_(ss, d); d.barcode = result.code; }
      sheet_(ss, 'Records').appendRow([now, who, clean_(d.type), clean_(d.artist), clean_(d.album), clean_(d.genre), clean_(d.format),
        clean_(d.barcode), Number(d.msrp) || '', Number(d.price), clean_(d.status), clean_(d.location), clean_(d.listed), d.printLabel ? 'Yes' : 'No', clean_(d.notes)]);
    } else if (task === 'reqstatus' || task === 'dropstatus' || task === 'feature' || task === 'supplies' || task === 'keg' || task === 'kegrecv' || task === 'stockrecv') {
      // Logged to the Log tab above. Nothing else to write yet.
    } else if (task === 'dropoff') {
      if (!clean_(d.name) || !(clean_(d.phone) || clean_(d.email)) || !clean_(d.items)) throw new Error('Missing drop-off details');
      sheet_(ss, 'DropOffs').appendRow([now, who, clean_(d.name), clean_(d.phone), clean_(d.email), clean_(d.items), Number(d.count) || '']);
    }
  } finally {
    lock.releaseLock();
  }
  if (task === 'receive') { try { squareSync(); } catch (e) { /* stays queued in SquareSync; the hourly retry picks it up */ } }
  if (d.autoEmail) sendAutoEmail_(d.autoEmail, task, who);
  return result;
}

/**
 * Automatic emails the app queues (customer "it's in", "can't get it", request confirmation, feature/bug notes).
 * Only a signed-in staff PIN can get here. Sent from this Google account as "Hunky Dory Records".
 * Script Properties (Project Settings): SHOP_EMAIL (reply-to, optional), FEATURE_EMAIL (where feature/bug notes go).
 * Capped at 60 emails a day as a safety net.
 */
function sendAutoEmail_(m, task, who) {
  var props = PropertiesService.getScriptProperties();
  var to = String(m.to || '').trim();
  var subject = String(m.subject || '').slice(0, 200);
  var body = String(m.body || '').slice(0, 6000);
  if (task === 'feature') to = props.getProperty('FEATURE_EMAIL') || '';
  if (!to || !subject || !body) return;
  if (!/^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/.test(to)) return;
  var cache = CacheService.getScriptCache();
  var key = 'mail-' + Utilities.formatDate(new Date(), 'UTC', 'yyyyMMdd');
  var n = Number(cache.get(key) || 0);
  if (n >= 60) throw new Error('Daily email limit reached');
  cache.put(key, String(n + 1), 86400);
  var opts = { name: 'Hunky Dory Records' };
  var reply = props.getProperty('SHOP_EMAIL') || String(m.replyTo || '').trim();
  if (reply && /^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/.test(reply)) opts.replyTo = reply;
  MailApp.sendEmail(to, subject, body, opts);
}

/** Run once from the editor to test: sends a sample email to this account. */
function testEmail() {
  MailApp.sendEmail(Session.getActiveUser().getEmail(), 'HD Backstage test', 'If you can read this, the shop can send emails.', { name: 'Hunky Dory Records' });
}

/**
 * Customer requests go into the website's "Hunky Dory Order Requests" sheet (tab "Orders"),
 * in the same columns the website form uses, so staff see one list.
 * Needs Script Properties: ORDER_SHEET_ID (the sheet's ID) and optionally SHOP (default Raleigh).
 * The store's Google account must have edit access to that sheet.
 */
function requestEntry_(d, who, now, tz) {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('ORDER_SHEET_ID');
  if (!id) return; // website order sheet not linked yet: the request is still in the Log tab
  var name = clean_(d.name), phone = clean_(d.phone), email = clean_(d.email);
  var items = (d.items && d.items.length ? d.items : [d]).slice(0, 10);
  if (!name || !(phone || email)) throw new Error('Missing request details');
  var sh = SpreadsheetApp.openById(id).getSheetByName('Orders');
  if (!sh) throw new Error('Orders tab not found');
  var date = Utilities.formatDate(now, tz, 'M/d/yyyy');
  var shop = props.getProperty('SHOP') || 'Raleigh';
  var rows = items.map(function (it) {
    var artist = clean_(it.artist), album = clean_(it.album);
    if (!artist || !album) throw new Error('Missing request details');
    var notes = (clean_(it.notes) ? clean_(it.notes) + ' ' : '') + '[In store: ' + who + ']';
    return ['Pending', date, shop, name, email, phone, artist, album, clean_(it.format) || 'Vinyl', notes, date];
  });
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

function saveIssuePhoto_(b64, now, tz) {
  if (b64.length > 4000000) throw new Error('Photo too large');
  var it = DriveApp.getFoldersByName('HD Ops Photos');
  var folder = it.hasNext() ? it.next() : DriveApp.createFolder('HD Ops Photos');
  var name = 'issue-' + Utilities.formatDate(now, tz, 'yyyy-MM-dd-HHmmss') + '.jpg';
  return folder.createFile(Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', name)).getUrl();
}

/** The newest shift note from the last 36 hours, or null. */
function latestNote_() {
  var sh = sheet_(ss_(), 'Log');
  var last = sh.getLastRow();
  if (last < 2) return null;
  var from = Math.max(2, last - 199);
  var rows = sh.getRange(from, 1, last - from + 1, 5).getValues();
  for (var i = rows.length - 1; i >= 0; i--) {
    if (rows[i][3] === 'note' && Date.now() - new Date(rows[i][0]).getTime() < 36 * 3600 * 1000) {
      return { ts: new Date(rows[i][0]).toISOString(), who: rows[i][2], text: String(rows[i][4]) };
    }
  }
  return null;
}

/** Today's log, summaries only. No amounts, names or contact info. */
function todayEntries_() {
  var ss = ss_();
  var tz = ss.getSpreadsheetTimeZone();
  var today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var sh = sheet_(ss, 'Log');
  var last = sh.getLastRow();
  if (last < 2) return [];
  var from = Math.max(2, last - 199); // newest 200 rows is plenty for one day
  var rows = sh.getRange(from, 1, last - from + 1, 5).getValues();
  return rows.filter(function (r) {
    var rd = r[1] instanceof Date ? Utilities.formatDate(r[1], tz, 'yyyy-MM-dd') : String(r[1]);
    return rd === today;
  }).map(function (r) {
    return { ts: new Date(r[0]).toISOString(), who: r[2], task: r[3], summary: safeSummary_(r[3], r[4]) };
  });
}

// Opening summaries list skipped step names, which is fine; cash/drop-off summaries are generic by design.
function safeSummary_(task, summary) { return String(summary).slice(0, 300); }

function sheet_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) throw new Error('Run setup() first');
  return sh;
}
// Strip leading = + - @ so typed text can never become a spreadsheet formula.
function clean_(v) {
  var s = String(v == null ? '' : v).trim().slice(0, 1000);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
function money_(v) {
  var n = parseFloat(v);
  return isNaN(n) ? '' : n;
}

/**
 * ONE-TIME: loads the Raleigh Trove list from the master "TREASURE TROVE DURM - CARY - IRON WORKS" sheet
 * (Raleigh Shop tab, unsold records only, with store locations; Nov 17 2026) into the Trove tab.
 * Replaces everything in the Trove tab, so run it once. Codes (TT-0001...) are assigned the next time the app loads the list.
 */
function importRaleighTrove() {
  var rows = [["Hank Mobley", "Roll Call", "Bar wall", 800, "", "Active", "Discogs"], ["Lee Konitz", "Live From Harvard Square", "Bar wall", 325, "", "Active", "Discogs"], ["Harbinger", "Second Coming", "Bar wall", 400, "", "Active", "Discogs"], ["AFI", "The Blood Album", "Bar wall", 300, "", "Active", "Discogs"], ["AFI", "DecemberUnderground (7\" Box)", "Bar wall", 165, "", "Active", ""], ["Phish", "Picture of Nectar", "Bar wall", 200, "", "Active", "Discogs"], ["Helen Merrill", "S/T", "Bar wall", 300, "", "Active", "Discogs"], ["Jay Jay Johnson", "The Eminent JJJ", "Bar wall", 90, "", "Active", ""], ["Roy Haynes", "Cymbalism", "Bar wall", 350, "", "Active", "Discogs, eBay"], ["Metallica", "Jump In The Fire", "Bar wall", 200, "", "Active", ""], ["Tash Sultana", "Flow State", "Bar wall", 90, "", "Active", ""], ["Khruangbin", "The Universe Smiles Upon You", "Bar cubbies", 175, "", "Active", ""], ["Sky Ferreira", "Night Time, My Time", "Bar cubbies", 150, "", "Active", "Discogs, eBay"], ["My Bloody Valentine", "Loveless", "Bar cubbies", 225, "", "Active", ""], ["Pink Floyd", "LIVE (Boot)", "Bar cubbies", 100, "", "Active", ""], ["2 Live Crew", "Move Somethin'", "Bar cubbies", 40, "", "Active", ""], ["Desire", "Cry At The Sky", "Bar cubbies", 80, "", "Active", ""], ["Quasimodo", "The Unseen", "Bar cubbies", 50, "", "Active", ""], ["Alien", "S/T", "Bar cubbies", 40, "", "Active", ""], ["Umphrey's McGee", "Hall of Fame", "Bar cubbies", 80, "", "Active", ""], ["The Monzas", "Hey! I Know You", "Bar cubbies", 500, "", "Active", ""], ["Bruce Springsteen", "Born to Run (Promo 45)", "Bar cubbies", 40, "", "Active", ""], ["AIR", "Music for Museum", "Bar cubbies", 350, "", "Active", "Discogs"], ["Jimmy Buffett", "Down to Earth", "Bar cubbies", 350, "", "Active", "eBay"], ["*45* Black Lips/KE$HA TP", "\"They's A Person Of The World\"", "Bar cubbies", 90, "", "Active", "Discogs"], ["Hotline Miami", "ST", "Bar cubbies", 250, "", "Active", "Discogs"], ["Fetti", "Roma", "Bar cubbies", 275, "", "Active", ""], ["Alberta Cross", "Broken Side of Time", "Bar cubbies", 120, "", "Active", ""], ["Under The Skin", "Mondo ST", "Bar cubbies", 80, "", "Active", "Discogs"], ["Bobby Hutcherson", "Enjoy The View", "Bar cubbies", 60, "", "Active", ""], ["Gene Barber and Cavaliers 45", "I Need a Love", "Bar cubbies", 100, "", "Active", ""], ["AFI (45)", "December Underground 45 Box", "Wall trove", 175, "", "Active", ""], ["Byard Lancaster", "It's Not Up to Us", "Wall trove", 200, "", "Active", "eBay"], ["Chet Baker", "Let's Get Lost ST", "Wall trove", 65, "", "Active", ""], ["Small Pools", "Love Tap!", "Wall trove", 75, "", "Active", ""], ["Tinariwen", "Aman Iman", "Wall trove", 75, "", "Active", ""], ["Major Lazer", "Peace Is The Mission (Sealed)", "Wall trove", 90, "", "Active", ""], ["Sister Hazel", "20 Stages", "Wall trove", 180, "", "Active", "Discogs"], ["Red Hot + Cool", "Stolen Moments", "Wall trove", 40, "", "Active", ""], ["De La Soul", "3 Feet High And Rising", "Wall trove", 100, "", "Active", ""], ["Status Quo", "Messages from the Status Quo", "Wall trove", 50, "", "Active", ""], ["The Wedding Present", "", "Wall trove", 60, "", "Active", ""], ["Smyrna", "Harvest Song", "Wall trove", 100, "", "Active", "Discogs"], ["Curtis Mayfield", "Curtis", "Wall trove", 100, "", "Active", ""], ["MC 900", "Hell With the Lid Off", "Wall trove", 75, "", "Active", ""], ["Kiss", "Love Gun (Promo)", "Wall trove", 75, "", "Active", ""], ["Beastie Boys", "Paul's Boutique", "Trove bin: Freshest", 65, "", "Active", ""], ["De La Soul", "Say No Go", "Trove bin: Freshest", 50, "", "Active", ""], ["De La Soul", "Say No Go", "Trove bin: Freshest", 45, "", "Active", ""], ["2 Live Crew", "Is What We Are", "Trove bin: Freshest", 60, "", "Active", ""], ["2 Live Crew", "As Nasty As They Wanna Be", "Trove bin: Freshest", 80, "", "Active", ""], ["Schoolly D", "Saturday Night The Album", "Trove bin: Freshest", 60, "", "Active", ""], ["Joni Mitchell", "Lennie and Dom Songs (Boot)", "Trove bin: Freshest", 40, "", "Active", ""], ["Antidope", "Go Pogo", "Trove bin: Freshest", 45, "", "Active", ""], ["Jimmy Smith", "The Sounds of Jimmy Smith", "Trove bin: Freshest", 100, "", "Active", ""], ["Tommy Flanagan Trio", "Moodville", "Trove bin: Freshest", 40, "", "Active", ""], ["Thomas Chapin", "Spirits Rebellious", "Trove bin: Freshest", 100, "", "Active", ""], ["Drive Like Jehu", "Headhunter", "Trove bin: Freshest", 100, "", "Active", ""], ["Jerry & Mike", "Duck 1", "Trove bin: Freshest", 40, "", "Active", ""], ["The Fall", "Extricate", "Trove bin: Freshest", 65, "", "Active", ""], ["Republic", "New World Order", "Trove bin: Freshest", 130, "", "Active", ""], ["The Smiths", "Meat Is Murder", "Trove bin: Freshest", 100, "", "Active", ""], ["Joy Division", "1977-1980", "Trove bin: Freshest", 90, "", "Active", ""], ["Dead Kennedys", "Fresh Fruit for Rotting Vegetables", "Trove bin: Freshest", 50, "", "Active", ""], ["Güngör", "Ghosts Upon The Earth", "Trove bin: Freshest", 45, "", "Active", ""], ["Ben E. King", "Don't Play That Song", "Trove bin: Freshest", 45, "", "Active", ""], ["Blue Magic", "Magic #", "Trove bin: Freshest", 60, "", "Active", ""], ["Bob Dylan", "", "Trove bin: Freshest", 50, "", "Active", ""], ["Black Ivory", "Don't Turn Around", "Trove bin: Freshest", 50, "", "Active", ""], ["Big Boi", "Boomiverse", "Trove bin: Freshest", 45, "", "Active", ""], ["Young Family", "Series 5", "Trove bin: Freshest", 40, "", "Active", ""], ["Max Roach", "4 Plays", "Trove bin: Freshest", 90, "", "Active", ""], ["Ed Sheeran", "You Need Me", "Trove bin: Freshest", 60, "", "Active", ""], ["Scribble Mural Comic Journal", "A Sunny Day in Glasgow", "Trove bin: Rock", 100, "Rock", "Active", "Discogs, eBay"], ["Robert Ashley", "Perfect Lives", "Trove bin: Rock", 40, "Rock", "Active", ""], ["Billy Bragg & Wilco", "Mermaid Avenue Vol. 2", "Trove bin: Rock", 175, "Rock", "Active", ""], ["Carl & the Passions", "So Tough", "Trove bin: Rock", 50, "Rock", "Active", ""], ["Julie Driscoll", "", "Trove bin: Rock", 40, "Rock", "Active", ""], ["Embers", "Roll", "Trove bin: Rock", 40, "Rock", "Active", ""], ["Mountain", "Burgundy", "Trove bin: Rock", 70, "Rock", "Active", ""], ["Nico", "Desertshore", "Trove bin: Rock", 75, "Rock", "Active", ""], ["Nightbird", "Cosmic Connections", "Trove bin: Rock", 50, "Rock", "Active", "eBay"], ["Sam Phillips", "Martinis & Bikinis", "Trove bin: Rock", 60, "Rock", "Active", ""], ["The Pretenders", "Live @ Santa Monica Civic", "Trove bin: Rock", 60, "Rock", "Active", ""], ["Starflyer 59", "Dial M", "Trove bin: Rock", 45, "Rock", "Active", ""], ["Tiny Habits", "All For Something", "Trove bin: Rock", 50, "Rock", "Active", ""], ["Tree Fox", "Rough Trade", "Trove bin: Rock", 45, "Rock", "Active", ""], ["Miniature Tigers", "Fortress", "Trove bin: Indie Pop", 100, "Indie Pop", "Active", "Discogs"], ["OneRepublic", "Native", "Trove bin: Indie Pop", 60, "Indie Pop", "Active", ""], ["Montaigne", "Complex (Red)", "Trove bin: Indie Pop", 60, "Indie Pop", "Active", "eBay"], ["Ed Sheeran", "Songs I Wrote With Amy", "Trove bin: Indie Pop", 50, "Indie Pop", "Active", ""], ["STRFKR", "Miracle Mile", "Trove bin: Indie Pop", 40, "Indie Pop", "Active", ""], ["Cannibal Corpse", "Red Before Black", "Trove bin: Indie Pop", 40, "Indie Pop", "Active", ""], ["White Fence", "Growing Faith", "Trove bin: Indie Pop", 40, "Indie Pop", "Active", ""], ["Big Black", "Atomizer", "Trove bin: Hard Rock/Punk/Metal", 40, "Hard Rock/Punk/Metal", "Active", ""], ["Bono Boys", "", "Trove bin: Hard Rock/Punk/Metal", 60, "Hard Rock/Punk/Metal", "Active", ""], ["Dragnet", "The Fall", "Trove bin: Hard Rock/Punk/Metal", 80, "Hard Rock/Punk/Metal", "Active", ""], ["First Strike", "Only For You (Signed)", "Trove bin: Hard Rock/Punk/Metal", 100, "Hard Rock/Punk/Metal", "Active", ""], ["Living Death", "Vengeance of Hell", "Trove bin: Hard Rock/Punk/Metal", 45, "Hard Rock/Punk/Metal", "Active", ""], ["Osker", "Idle Will Kill", "Trove bin: Hard Rock/Punk/Metal", 50, "Hard Rock/Punk/Metal", "Active", ""], ["The Saints", "Eternally Yours", "Trove bin: Hard Rock/Punk/Metal", 45, "Hard Rock/Punk/Metal", "Active", ""], ["Scratch Acid", "Berserker", "Trove bin: Hard Rock/Punk/Metal", 40, "Hard Rock/Punk/Metal", "Active", ""], ["Stormwitch", "Tales of Terror", "Trove bin: Hard Rock/Punk/Metal", 55, "Hard Rock/Punk/Metal", "Active", ""], ["Thor", "Only The Strong", "Trove bin: Hard Rock/Punk/Metal", 50, "Hard Rock/Punk/Metal", "Active", ""], ["Upright Citizens", "Open Eyes Open Ears", "Trove bin: Hard Rock/Punk/Metal", 40, "Hard Rock/Punk/Metal", "Active", ""], ["XTRMST", "S/T (Signed)", "Trove bin: Hard Rock/Punk/Metal", 95, "Hard Rock/Punk/Metal", "Active", ""], ["Aesop Rock", "Float", "Trove bin: HipHop", 95, "HipHop", "Active", ""], ["Eagle-Eye Cherry", "Save Tonight", "Trove bin: Electronic", 70, "Electronic", "Active", ""], ["AlunaGeorge", "Body Music", "Trove bin: Electronic", 50, "Electronic", "Active", ""], ["The Herbaliser", "Blow Your Headphones", "Trove bin: Electronic", 60, "Electronic", "Active", ""], ["HKE", "HK", "Trove bin: Electronic", 50, "Electronic", "Active", ""], ["Lewis", "L'Amour", "Trove bin: Electronic", 44, "Electronic", "Active", ""], ["Lizzy Mercier Descloux", "Press Color", "Trove bin: Electronic", 75, "Electronic", "Active", ""], ["Mr. Twin Sister", "", "Trove bin: Electronic", 40, "Electronic", "Active", ""], ["Report: Drug Addiction", "", "Trove bin: Electronic", 40, "Electronic", "Active", ""], ["Shpongle", "Codex VI", "Trove bin: Electronic", 40, "Electronic", "Active", ""], ["Whitehouse", "Great White Death", "Trove bin: Electronic", 50, "Electronic", "Active", ""], ["Darrell Banks", "Is Here", "Trove bin: Soul/R&B/Funk/Disco", 60, "Soul/R&B/Funk/Disco", "Active", ""], ["James Brown", "Shout and Shimmy", "Trove bin: Soul/R&B/Funk/Disco", 100, "Soul/R&B/Funk/Disco", "Active", ""], ["Dwayne Dopson", "Get Down", "Trove bin: Soul/R&B/Funk/Disco", 40, "Soul/R&B/Funk/Disco", "Active", ""], ["The Impressions", "The Never Ending", "Trove bin: Soul/R&B/Funk/Disco", 80, "Soul/R&B/Funk/Disco", "Active", ""], ["Martha and the Vandals", "Come Get These Memories", "Trove bin: Soul/R&B/Funk/Disco", 50, "Soul/R&B/Funk/Disco", "Active", ""], ["Garnet Mimms", "Cry Baby", "Trove bin: Soul/R&B/Funk/Disco", 55, "Soul/R&B/Funk/Disco", "Active", ""], ["Wilson Pickett", "The Midnight Mover", "Trove bin: Soul/R&B/Funk/Disco", 50, "Soul/R&B/Funk/Disco", "Active", ""], ["The Strikers", "The Strikers", "Trove bin: Soul/R&B/Funk/Disco", 45, "Soul/R&B/Funk/Disco", "Active", ""], ["Wilson Williams", "Up the Down Stairs", "Trove bin: Soul/R&B/Funk/Disco", 75, "Soul/R&B/Funk/Disco", "Active", ""], ["Val Anthony", "What Can I Say", "Trove bin: Jazz", 50, "Jazz", "Active", ""], ["Brian Blade", "Landmarks", "Trove bin: Jazz", 65, "Jazz", "Active", ""], ["Sonny Clark", "Cool Struttin (09 Blue Note 45 Ltd)", "Trove bin: Jazz", 150, "Jazz", "Active", ""], ["Ahmad Jamal", "73", "Trove bin: Jazz", 60, "Jazz", "Active", ""], ["Shelly Manne", "At the Blackhawk", "Trove bin: Jazz", 60, "Jazz", "Active", ""], ["Jackie McLean", "Fat Jazz", "Trove bin: Jazz", 90, "Jazz", "Active", ""], ["Harry Partch", "", "Trove bin: Jazz", 100, "Jazz", "Active", ""], ["Gabor Szabo", "Dreams", "Trove bin: Jazz", 60, "Jazz", "Active", ""], ["Larry Young", "Young Blues", "Trove bin: Jazz", 225, "Jazz", "Active", ""], ["Dan & Shay", "Obsessed", "Trove bin: Country/Folk", 135, "Country/Folk", "Active", ""], ["Scotty McCreery", "Seasons Change", "Trove bin: Country/Folk", 70, "Country/Folk", "Active", ""], ["Winding River", "", "Trove bin: Country/Folk", 60, "Country/Folk", "Active", ""], ["Women of Old Time Music", "S/T", "Trove bin: Country/Folk", 40, "Country/Folk", "Active", ""], ["Satyajit Ray", "Goopy Gyne & Bagha Byne", "Trove bin: World/Latin/Reggae", 200, "World/Latin/Reggae", "Active", ""], ["Inezita Barroso", "Vamos Falar De Brasil, Novamente…", "Trove bin: World/Latin/Reggae", 40, "World/Latin/Reggae", "Active", ""], ["Johnny Reviens", "les rocl", "Trove bin: World/Latin/Reggae", 50, "World/Latin/Reggae", "Active", ""], ["Tom MacDonald", "Exposed", "Trove bin: Misc", 60, "Misc", "Active", ""], ["Tee Vee Toons", "Television's Greatest Hits of Cable TV", "Trove bin: Misc", 50, "Misc", "Active", ""], ["Tee Vee Toons", "Television's Greatest Hits of Black and White", "Trove bin: Misc", 50, "Misc", "Active", ""], ["Tee Vee Toons", "Television's Greatest Hits of Remote Control", "Trove bin: Misc", 50, "Misc", "Active", ""], ["Sly Johnson", "Complete Mythology", "Backfill cubbie", 70, "", "Backfill", ""], ["*45* The Green Slime", "Far Beyond", "Backfill cubbie", 60, "", "Backfill", ""], ["Apples in Stereo", "Fun Trick Noisemaker", "Backfill cubbie", 200, "", "Backfill", "Discogs, eBay"], ["Cliff Jordan", "Blowing In From Chicago", "Backfill cubbie", 800, "", "Backfill", "Discogs"], ["Ronald Ingraham Concert Choir", "Saturday Morning With RICC", "Backfill cubbie", 45, "", "Backfill", ""], ["George Duke", "No Rhyme, No Reason", "Backfill cubbie", 40, "", "Backfill", ""], ["Forrest Frank", "Child of God", "Backfill cubbie", 125, "", "Backfill", ""], ["Thomas Chapin", "Spirits Rebellious", "Backfill cubbie", 100, "", "Backfill", "Discogs"], ["In Trousers", "ST", "Backfill cubbie", 100, "", "Backfill", ""], ["George Duke", "No Rhyme, No Reason", "Backfill cubbie", 40, "", "Backfill", ""], ["Taj Mahal", "Present Awakening", "Backfill cubbie", 40, "", "Backfill", ""], ["Drive Like Jehu", "S/T", "Backfill cubbie", 100, "", "Backfill", ""], ["Paul Hibbets", "Childhood Dream", "Backfill cubbie", 200, "", "Backfill", ""], ["Andrew Hill", "Judgement!", "Backfill cubbie", "", "", "Backfill", ""], ["Albert Hash", "Whitetop", "Backfill cubbie", 40, "", "Backfill", ""], ["Leaf River", "Sonshine", "Backfill cubbie", 100, "", "Backfill", ""], ["Frantic", "Conception (Sealed)", "Backfill cubbie", 150, "", "Backfill", "Discogs"], ["Eli's", "S/T", "Backfill cubbie", 45, "", "Backfill", ""], ["JuJu", "A Message From Mozambique", "Backfill cubbie", 400, "", "Backfill", "Discogs, eBay"], ["45\" The Annuals", "Hungry / Once in a Lifetime", "Backfill cubbie", 250, "", "Backfill", ""], ["*45* Richie sings", "Dream Lover", "Backfill cubbie", 100, "", "Backfill", ""], ["The Overtones", "Rustic (45)", "Backfill cubbie", 40, "", "Backfill", ""], ["Soft Rock", "Soft Rock", "Backfill cubbie", 500, "", "Backfill", "eBay"], ["The Wrens", "The Meadowlands", "Backfill cubbie", 400, "", "Backfill", "Discogs"], ["The Pyramids", "Lalibela", "Backfill cubbie", 750, "", "Backfill", "Discogs, eBay"], ["Sam Gendel and Sam Wilkes", "Music for Saxophone and Bass Guitar", "Backfill cubbie", 400, "", "Backfill", "Discogs, eBay"], ["Wild Sugar", "Bring It Here", "Backfill cubbie", 40, "", "Backfill", ""], ["65daysofstatic", "No Man's Sky ST", "Backfill cubbie", 140, "", "Backfill", "eBay"], ["Richard Pryor", "I Hope I'm Funny", "Backfill cubbie", 90, "", "Backfill", ""], ["*45* Moon Kim", "I'm Dingy Dongy Over You", "Backfill cubbie", 50, "", "Backfill", ""], ["*45* Chevells", "Pretty Little Girl", "Backfill cubbie", 300, "", "Backfill", ""], ["*45* Music Machine", "The Eagle Never Hunts the Fly", "Backfill cubbie", 75, "", "Backfill", ""]];
  var sh = sheet_(ss_(), 'Trove');
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(sh.getLastColumn(), 9)).clearContent();
  sh.getRange(1, 1, 1, 9).setValues([TABS.Trove]).setFontWeight('bold');
  var out = rows.map(function (r) { return [r[0], r[1], r[2], r[3], '', '', r[4], r[5], r[6]]; });
  sh.getRange(2, 1, out.length, 9).setValues(out);
  troveItems_();
  Logger.log('Imported ' + out.length + ' records.');
}

/**
 * ONE-TIME: syncs the Trove tab to the "trove adam version" sheet (Raleigh Shop tab, Nov 17 2026).
 * Matched records keep their code, shelf, genre, cover, status and listed flags; titles and prices take Adam's version.
 * Records not on his sheet are removed; new ones are added with no shelf and get the next TT code. Run it once.
 */
function syncAdamTrove() {
  var N = function (s) { return String(s || '').toLowerCase().replace(/[^a-z0-9]/g, ''); };
  var ALIAS = { 'mc900ftjesus': 'mc900', 'jerrygoldsmith': 'alien', 'micalevi': 'undertheskin' };
  var want = [["Hank Mobley","Roll Call",800],["Lee Konitz","Live From Harvard Square",325],["Harbinger","Second Coming",400],["AFI","The Blood Album",300],["AFI","DecemberUnderground (7\" Box)",165],["Phish","Picture of Nectar",200],["Helen Merrill","Helen Merrill",300],["Jay Jay Johnson","The Eminent Jay Jay Johnson, Vol. 2",90],["Roy Haynes","Cymbalism",350],["Metallica","Jump in the Fire",200],["Tash Sultana","Flow State",90],["Khruangbin","The Universe Smiles Upon You",175],["Sky Ferreira","Night Time, My Time",150],["My Bloody Valentine","Loveless",225],["2 Live Crew","Move Somethin'",40],["Jerry Goldsmith","Alien (Original Motion Picture Soundtrack)",40],["The Monzas","Hey! I Know You",500],["Bruce Springsteen","Born to Run (Promo 45)",40],["AIR","Music for Museum",350],["Jimmy Buffett","Down to Earth",350],["*45* Black Lips/KE$HA TP","\"They's A Person Of The World\"",90],["Hotline Miami","Soundtrack",250],["Fetti","Roma",275],["Gene Barber and Cavaliers 45","I Need a Love",100],["Quasimodo","The Unseen",50],["Bobby Hutcherson","Enjoy the View",60],["Byard Lancaster","It's Not Up to Us",200],["Chet Baker","Chet Baker Sings and Plays from the Film \"Let's Get Lost\"",65],["Small Pools","Love Tap!",75],["Tinariwen","Aman Iman: Water Is Life",75],["Major Lazer","Peace Is the Mission (Sealed)",90],["Vulfpeck","2011-2014",175],["Red Hot + Cool","Stolen Moments: Red Hot + Cool",40],["De La Soul","3 Feet High and Rising",100],["Status Quo","Messages from the Status Quo",50],["The Wedding Present","Bizarro",60],["Smyrna","Harvest Song",100],["Curtis Mayfield","Curtis",100],["MC 900 Ft. Jesus","Hell with the Lid Off",75],["KISS","Love Gun (Promo)",75],["De La Soul","Say No Go",50],["De La Soul","Say No Go",45],["2 Live Crew","2 Live Crew Is What We Are",60],["2 Live Crew","As Nasty as They Wanna Be",80],["Schoolly D","Saturday Night! The Album",60],["Antidope","Go Pogo",45],["Jimmy Smith","The Sounds of Jimmy Smith",100],["Thomas Chapin","Spirits Rebellious",100],["Drive Like Jehu","Headhunter",100],["Jerry & Mike","Duck 1",40],["The Fall","Extricate",65],["Republic","New World Order",130],["Güngör","Ghosts Upon The Earth",45],["Ben E. King","Don't Play That Song",45],["Blue Magic","Magic #",60],["Black Ivory","Don't Turn Around",50],["Big Boi","Boomiverse",45],["Young Family","Series 5",40],["Robert Ashley","Perfect Lives",40],["Billy Bragg & Wilco","Mermaid Avenue Vol. II",175],["Carl & the Passions","So Tough",50],["Julie Driscoll","1969",40],["Mountain","Burgundy",70],["Nico","Desertshore",75],["Nielsen & Pearson","If You Should Set Sail",40],["Nightbird","Cosmic Connections",50],["Sam Phillips","Martinis & Bikinis",60],["The Pretenders","Live @ Santa Monica Civic",60],["Sister Hazel","20 Stages",180],["Starflyer 59","Dial M",45],["Tiny Habits","All for Something",50],["Tree Fox","Rough Trade",45],["OneRepublic","Native",60],["Montaigne","Complex (Red)",60],["Ed Sheeran","Songs I Wrote with Amy",50],["STRFKR","Miracle Mile",40],["White Fence","Is Growing Faith",40],["Big Black","Atomizer",40],["Bono Boys","",60],["Dragnet","The Fall",80],["First Strike","Only For You (Signed)",100],["Living Death","Vengeance of Hell",45],["Osker","Idle Will Kill",50],["The Saints","Eternally Yours",45],["Savatage","Power of the Night ('85 EU)",55],["Scratch Acid","Berserker",40],["Stormwitch","Tales of Terror",55],["Thor","Only The Strong",50],["Upright Citizens","Open Eyes Open Ears",40],["XTRMST","XTRMST (Signed)",95],["Aesop Rock","Float",95],["Eagle-Eye Cherry","Save Tonight",70],["AlunaGeorge","Body Music",50],["The Herbaliser","Blow Your Headphones",60],["HKE","HK",50],["Lewis","L'Amour",44],["Lizzy Mercier Descloux","Press Color",75],["Mr. Twin Sister","Mr Twin Sister",40],["Report: Drug Addiction","A View from the Belly",40],["Shpongle","Codex VI",40],["Whitehouse","Great White Death",50],["Darrell Banks","Darrell Banks Is Here!",60],["James Brown","Shout and Shimmy",100],["Dwayne Dopson","Get Down",40],["The Impressions","The Never Ending Impressions",80],["Garnet Mimms","Cry Baby",55],["Wilson Pickett","The Midnight Mover",50],["The Strikers","The Strikers",45],["Wilson Williams","Up the Down Stairs",75],["Val Anthony","What Can I Say",50],["Brian Blade","Landmarks",65],["Sonny Clark","Cool Struttin (09 Blue Note 45 Ltd)",150],["Ahmad Jamal","73",60],["Shelly Manne","At the Blackhawk",60],["Jackie McLean","Fat Jazz",90],["Harry Partch","30 Years of Lyrical and Dramatic Music",100],["Gabor Szabo","Dreams",60],["Larry Young","Young Blues",225],["Dan & Shay","Obsessed",135],["Scotty McCreery","Seasons Change",70],["Winding River","Winding River",60],["Women of Old Time Music","Women of Old Time Music",40],["Inezita Barroso","Vamos Falar De Brasil, Novamente…",40],["Johnny Reviens","les rocl",50],["Tom MacDonald","Exposed",60],["Tee Vee Toons","Television's Greatest Hits of Cable TV",50],["Tee Vee Toons","Television's Greatest Hits of Black and White",50],["Tee Vee Toons","Television's Greatest Hits of Remote Control",50],["Mica Levi","Under the Skin (Original Motion Picture Soundtrack)",80],["Apples in Stereo","Fun Trick Noisemaker",200],["Cliff Jordan","Blowing in from Chicago",800],["Ronald Ingraham Concert Choir","Saturday Morning with RICC",45],["George Duke","No Rhyme, No Reason",40],["Forrest Frank","Child of God",125],["Thomas Chapin","Spirits Rebellious",100],["In Trousers","ST",100],["George Duke","No Rhyme, No Reason",40],["Taj Mahal","Present Awakening",40],["Drive Like Jehu","Drive Like Jehu",100],["Paul Hibbets","Childhood Dream",200],["Andrew Hill","Judgement!",""],["Albert Hash","Whitetop",40],["Leaf River","Sonshine",100],["Frantic","Conception (Sealed)",150],["Eli's","S/T",45],["JuJu","A Message from Mozambique",400],["45\" The Annuals","Hungry / Once in a Lifetime",250],["*45* Richie sings","Dream Lover",100],["The Overtones","Rustic (45)",40],["Soft Rock","Soft Rock",500],["The Wrens","The Meadowlands",400],["The Pyramids","Lalibela",750],["Sam Gendel and Sam Wilkes","Music for Saxophone and Bass Guitar",400],["Wild Sugar","Bring It Here",40],["65daysofstatic","No Man's Sky: Music for an Infinite Universe",140],["Richard Pryor","I Hope I'm Funny",90]];
  var sh = sheet_(ss_(), 'Trove');
  var old = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues().filter(function (r) { return r[0] || r[1]; }) : [];
  var used = {}, keep = {}, added = [];
  want.forEach(function (w, wi) {
    var k = ALIAS[N(w[0])] || N(w[0]), hit = -1, i;
    for (i = 0; i < old.length && hit < 0; i++) if (!used[i] && N(old[i][0]) === k && Number(old[i][3]) === Number(w[2])) hit = i;
    for (i = 0; i < old.length && hit < 0; i++) if (!used[i] && N(old[i][0]) === k) hit = i;
    if (hit >= 0) { used[hit] = 1; keep[hit] = w; } else added.push(w);
  });
  var out = [], removed = [];
  old.forEach(function (r, i) {
    if (keep[i]) { r[0] = keep[i][0]; r[1] = keep[i][1]; r[3] = keep[i][2]; out.push(r); } else removed.push(r[0] + ' - ' + r[1]);
  });
  added.forEach(function (w) { out.push([w[0], w[1], 'No location', w[2], '', '', '', 'Active', '']); });
  sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 9).clearContent();
  sh.getRange(2, 1, out.length, 9).setValues(out);
  troveItems_();
  Logger.log('Kept ' + (out.length - added.length) + ', added ' + added.length + ', removed ' + removed.length + '. Removed: ' + removed.join('; '));
}


/** Deletes one Trove row by code (falls back to artist + title). A copy goes to the 'Trove Deleted' tab first. */
function deleteTrove_(req, who) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = ss_(), sh = sheet_(ss, 'Trove');
    if (sh.getLastRow() < 2) return 0;
    var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues();
    var code = String(req.code || '').trim(), at = String(req.artist || '').trim().toLowerCase(), ti = String(req.title || '').trim().toLowerCase();
    var idx = -1;
    for (var i = 0; i < rows.length; i++) {
      if (code && String(rows[i][5]).trim() === code) { idx = i; break; }
      if (!code && at && String(rows[i][0]).trim().toLowerCase() === at && String(rows[i][1]).trim().toLowerCase() === ti) { idx = i; break; }
    }
    if (idx < 0) return 0;
    var bin = ss.getSheetByName('Trove Deleted') || ss.insertSheet('Trove Deleted');
    if (bin.getLastRow() === 0) bin.getRange(1, 1, 1, 11).setValues([['Deleted at', 'Deleted by', 'Artist', 'Title', 'Location', 'Price', 'Cover URL', 'Code', 'Genre', 'Status', 'Listed']]).setFontWeight('bold');
    bin.appendRow([new Date(), who].concat(rows[idx]));
    sh.deleteRow(idx + 2);
    return 1;
  } finally {
    lock.releaseLock();
  }
}


/* ---------- Expected records (Raleigh = Alliance account 066847) ---------- */

/** Open lines: everything invoiced and not yet scanned in. Oldest first. */
function expectedLines_() {
  var sh = ss_().getSheetByName('Expected');
  if (!sh || sh.getLastRow() < 2) return [];
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 11).getValues(), out = [];
  rows.forEach(function (r) {
    var left = Number(r[8]) - Number(r[9] || 0);
    if (r[4] && left > 0) out.push({ inv: String(r[0]), date: dstr_(r[1]), po: String(r[2]), prod: String(r[3]), upc: String(r[4]),
      artist: String(r[5]), title: String(r[6]), list: Number(r[7]), qty: left, media: 'LP' });
  });
  return out;
}

/** Add scanned counts ({upc: n}) to the oldest open lines for that UPC. Runs inside the log lock. */
function applyReceived_(ss, got) {
  if (!got || typeof got !== 'object') return;
  var sh = ss.getSheetByName('Expected');
  if (!sh || sh.getLastRow() < 2) return;
  var rg = sh.getRange(2, 1, sh.getLastRow() - 1, 11), rows = rg.getValues(), changed = false;
  Object.keys(got).forEach(function (u) {
    var n = Math.floor(Number(got[u]) || 0);
    for (var i = 0; i < rows.length && n > 0; i++) {
      if (String(rows[i][4]) !== String(u)) continue;
      var room = Number(rows[i][8]) - Number(rows[i][9] || 0);
      if (room <= 0) continue;
      var take = Math.min(room, n);
      rows[i][9] = Number(rows[i][9] || 0) + take; n -= take; changed = true;
    }
  });
  if (changed) rg.setValues(rows);
}

/** Add invoice lines. Each: {inv, date:'yyyy-mm-dd', po, prod, upc, artist, title, list, qty, cost (invoice unit cost)}. Skips lines already there. */
function addExpected_(lines) {
  var ss = ss_(), sh = sheet_(ss, 'Expected');
  var have = {};
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues().forEach(function (r) { have[r[0] + '|' + r[3]] = 1; });
  var rows = [];
  sh.getRange('A:E').setNumberFormat('@'); // text, so leading zeros and dates survive
  lines.forEach(function (l) {
    if (have[l.inv + '|' + l.prod]) return;
    rows.push([l.inv, l.date, l.po, l.prod, String(l.upc), l.artist, l.title, l.list, l.qty, Math.min(Number(l.got) || 0, l.qty), '', Number(l.cost) || '']);
  });
  if (!rows.length) return 0;
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, 12).setValues(rows);
  return rows.length;
}

/**
 * Runs every morning. Any invoiced record still not scanned in after 14 days gets listed in one email
 * to Forrest and Bell (Script Property OVERDUE_TO, comma separated). Each line is only reported once.
 */
function checkOverdue() {
  var props = PropertiesService.getScriptProperties();
  var to = props.getProperty('OVERDUE_TO') || props.getProperty('SHOP_EMAIL') || '';
  var sh = ss_().getSheetByName('Expected');
  if (!to || !sh || sh.getLastRow() < 2) return;
  var rg = sh.getRange(2, 1, sh.getLastRow() - 1, 11), rows = rg.getValues();
  var tz = ss_().getSpreadsheetTimeZone(), today = new Date(), todayStr = Utilities.formatDate(today, tz, 'yyyy-MM-dd');
  var byInv = {}, hit = 0;
  rows.forEach(function (r) {
    var left = Number(r[8]) - Number(r[9] || 0);
    if (left <= 0 || r[10]) return;
    var d = new Date(dstr_(r[1]) + 'T12:00:00');
    if (isNaN(d) || (today - d) / 864e5 <= 14) return;
    (byInv[r[0]] = byInv[r[0]] || { date: dstr_(r[1]), lines: [] }).lines.push(r[5] + ' - ' + r[6] + (left > 1 ? ' (' + left + ' missing)' : ''));
    r[10] = todayStr; hit++;
  });
  if (!hit) return;
  var body = 'These records were on a Raleigh Alliance invoice (account 066847) but still have not been scanned in after 14 days.\n\n';
  Object.keys(byInv).forEach(function (k) {
    body += 'Invoice ' + k + ' (' + byInv[k].date + '):\n  ' + byInv[k].lines.join('\n  ') + '\n\n';
  });
  body += 'If a box is sitting somewhere, scan it in. If it never arrived, it is a claim with Alliance (claims must be made within 5 days of the invoice date, so call Katie Curtis now).\n';
  MailApp.sendEmail(to, 'Hunky Dory Raleigh: ' + hit + ' record' + (hit === 1 ? '' : 's') + ' invoiced but not received', body, { name: 'Hunky Dory Records' });
  rg.setValues(rows);
}

/** Run once from the editor: runs checkOverdue every morning around 8. */
function installOverdueTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'checkOverdue') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('checkOverdue').timeBased().everyDays(1).atHour(8).create();
}

function dstr_(v) { return v instanceof Date ? Utilities.formatDate(v, 'UTC', 'yyyy-MM-dd') : String(v); }

/** Barcodes staff matched by hand (or Discogs matched): scanned code -> the UPC on the invoice. Shared by every iPad. */
function upcMap_() {
  var sh = ss_().getSheetByName('UpcMap'), out = {};
  if (sh && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) { if (r[0]) out[String(r[0])] = String(r[1]); });
  return out;
}
function learnUpc_(req, who) {
  var a = String(req.scanned || '').replace(/\D/g, '').replace(/^0+/, ''), b = String(req.to || '').replace(/\D/g, '');
  if (!a || !b || a.length > 14 || b.length > 14) return false;
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var sh = sheet_(ss_(), 'UpcMap');
    sh.getRange('A:B').setNumberFormat('@');
    if (upcMap_()[a]) return true;
    sh.appendRow([a, b, clean_(who), new Date()]);
  } finally { lock.releaseLock(); }
  return true;
}


/* ---------- Square sync: received records go straight into the Square catalog + stock ----------
 * Script Property needed: SQUARE_TOKEN (Square developer dashboard > your app > Production access token;
 * needs ITEMS_READ, ITEMS_WRITE, INVENTORY_READ, INVENTORY_WRITE). Until it is set, receipts just wait in the SquareSync tab.
 * Each scanned UPC: exists in Square -> add stock (un-archives it first if needed); new -> create the item, then add stock.
 * Safe to re-run: every stock change carries a fixed idempotency key, so a retry can never double-count.
 */
var SQ = { loc: 'L7WFX5VKQEQ29', cat: 'MC5G6VGYDMHIDS3QHHTIDZ76', tax: 'UIFR5UKZEUSS566Z644HZN6D', vendor: 'ZJSCSOCI6C2JMLMX', ver: '2025-10-16' };

/** Queue one row per scanned UPC (called inside the log lock, before applyReceived_ changes the counts). */
function sqQueue_(ss, got, now) {
  if (!got || typeof got !== 'object') return;
  var ex = ss.getSheetByName('Expected'), info = {};
  if (ex && ex.getLastRow() > 1) ex.getRange(2, 1, ex.getLastRow() - 1, 12).getValues().forEach(function (r) { if (r[4] && !info[String(r[4])]) info[String(r[4])] = r; });
  var rows = [];
  Object.keys(got).forEach(function (u) {
    var n = Math.floor(Number(got[u]) || 0), r = info[String(u)];
    if (n < 1 || !r) return;
    rows.push([now, String(u), r[5], r[6], n, 'Pending', '', 'hd-recv-' + now.getTime() + '-' + u, Number(r[7]) || 0, String(r[3]), Number(r[11]) || '']);
  });
  if (!rows.length) return;
  var sh = sheet_(ss, 'SquareSync');
  sh.getRange('B:B').setNumberFormat('@'); sh.getRange('J:J').setNumberFormat('@');
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

function sq_(method, path, body) {
  var tok = PropertiesService.getScriptProperties().getProperty('SQUARE_TOKEN');
  if (!tok) throw new Error('SQUARE_TOKEN not set');
  var r = UrlFetchApp.fetch('https://connect.squareup.com' + path, { method: method, contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + tok, 'Square-Version': SQ.ver }, payload: body ? JSON.stringify(body) : undefined });
  var j = JSON.parse(r.getContentText() || '{}');
  if (r.getResponseCode() >= 300) throw new Error((j.errors || []).map(function (e) { return e.detail || e.code; }).join('; ') || ('HTTP ' + r.getResponseCode()));
  return j;
}

/** Push every Pending/Error row to Square. Runs after each receipt and hourly (installSquareTrigger). */
function squareSync() {
  if (!PropertiesService.getScriptProperties().getProperty('SQUARE_TOKEN')) return;
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var sh = ss_().getSheetByName('SquareSync');
    if (!sh || sh.getLastRow() < 2) return;
    var rg = sh.getRange(2, 1, sh.getLastRow() - 1, 11), rows = rg.getValues();
    var todo = rows.filter(function (r) { return r[5] !== 'Done'; });
    if (!todo.length) return;
    var upcs = todo.map(function (r) { return String(r[1]); }).filter(function (u, i, a) { return a.indexOf(u) === i; });
    var found = {}, items = {};
    for (var i = 0; i < upcs.length; i += 50) {
      var res = sq_('post', '/v2/catalog/search', { object_types: ['ITEM_VARIATION'], include_related_objects: true,
        query: { set_query: { attribute_name: 'upc', attribute_values: upcs.slice(i, i + 50) } } });
      (res.objects || []).forEach(function (o) { found[o.item_variation_data.upc] = o; });
      (res.related_objects || []).forEach(function (o) { if (o.type === 'ITEM') items[o.id] = o; });
    }
    rows.forEach(function (r) {
      if (r[5] === 'Done') return;
      try {
        var u = String(r[1]), v = found[u];
        if (!v) { v = sqCreate_(r); found[u] = v; r[6] = v.noCost ? 'Created new item, NO COST (none on invoice line): add it in Square' : 'Created new item with cost'; }
        else {
          var it = items[v.item_variation_data.item_id];
          if (it && it.item_data.is_archived) { it.item_data.is_archived = false; sq_('post', '/v2/catalog/object', { idempotency_key: r[7] + '-un', object: it }); r[6] = 'Restored archived item'; }
          else r[6] = 'Added stock';
        }
        sq_('post', '/v2/inventory/changes/batch-create', { idempotency_key: r[7], changes: [{ type: 'ADJUSTMENT', adjustment: {
          catalog_object_id: v.id, from_state: 'NONE', to_state: 'IN_STOCK', location_id: SQ.loc, quantity: String(r[4]), occurred_at: new Date().toISOString() } }] });
        r[5] = 'Done';
      } catch (err) { r[5] = 'Error'; r[6] = String(err.message || err).slice(0, 200); }
    });
    rg.setValues(rows);
  } finally { lock.releaseLock(); }
}

/** New catalog item in the house style: "Artist: Title", SKU = product #, price = MSRP rounded up + $1, New Records, taxed, Raleigh only. */
function sqCreate_(r) {
  var price = Math.round((Math.ceil(Number(r[8])) + 1) * 100);
  if (!(price > 100)) throw new Error('No MSRP on the invoice line, so no price. Add the item in Square by hand.');
  var res = sq_('post', '/v2/catalog/object', { idempotency_key: r[7] + '-new', object: { type: 'ITEM', id: '#i', present_at_all_locations: false, present_at_location_ids: [SQ.loc],
    item_data: { name: String(r[2]) + ': ' + String(r[3]), tax_ids: [SQ.tax], categories: [{ id: SQ.cat }], variations: [{ type: 'ITEM_VARIATION', id: '#v', present_at_all_locations: false,
      present_at_location_ids: [SQ.loc], item_variation_data: { item_id: '#i', name: '', sku: String(r[9]), upc: String(r[1]), pricing_type: 'FIXED_PRICING',
        price_money: { amount: price, currency: 'USD' }, track_inventory: true } }] } } });
  var v = res.catalog_object.item_data.variations[0];
  if (Number(r[10]) > 0) sq_('post', '/v2/catalog/object', { idempotency_key: r[7] + '-cost', object: { type: 'ITEM_VARIATION_VENDOR_INFO', id: '#c',
    present_at_all_locations: false, present_at_location_ids: [SQ.loc], item_variation_vendor_info_data: { ordinal: 1,
      price_money: { amount: Math.round(Number(r[10]) * 100), currency: 'USD' }, item_variation_id: v.id, vendor_id: SQ.vendor } } });
  else v.noCost = true;
  return v;
}

function installSquareTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'squareSync') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('squareSync').timeBased().everyHours(1).create();
}

/** One-time: fill Unit cost on the 4 lines already in Expected. Run once from the editor. */
function backfillCosts() {
  var cost = { 'DEDO383C1.1': 19.83, 'PKFL872027.1': 24.47, 'ISCB002333701.1': 26.41, 'MERNB002349201.1': 24.55 };
  var sh = sheet_(ss_(), 'Expected');
  sh.getRange(1, 12).setValue('Unit cost').setFontWeight('bold');
  if (sh.getLastRow() < 2) return;
  var rg = sh.getRange(2, 1, sh.getLastRow() - 1, 12), rows = rg.getValues();
  rows.forEach(function (r) { if (!r[11] && cost[String(r[3])]) r[11] = cost[String(r[3])]; });
  rg.setValues(rows);
}
