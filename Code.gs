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
  Staff:    ['Name', 'PIN', 'Role'],
  Trove:    ['Artist', 'Title', 'Location', 'Price', 'Cover URL', 'Code', 'Genre', 'Status', 'Listed'],
  Log:      ['Timestamp', 'Date', 'Who', 'Task', 'Summary', 'Data (JSON)'],
  Opening:  ['Timestamp', 'Who', 'Amounts', 'Steps done', 'Could not do (step: reason)', 'Register note'],
  Closing:  ['Timestamp', 'Who', 'Amounts', 'Steps done', 'Could not do (step: reason)', 'Register note'],
  Issues:   ['Timestamp', 'Who', 'Type', 'Details', 'Photo', 'Status'],
  Cash:     ['Timestamp', 'Who', 'Drawer count', 'Deposit', 'Note', 'Type', 'Source'],
  Records: ['Timestamp', 'Who', 'Type', 'Artist', 'Album', 'Genre', 'Format', 'Barcode', 'MSRP', 'Price', 'Status', 'Location', 'Listed on', 'Label queued', 'Notes'],
  Expected: ['Invoice', 'Invoice date', 'PO', 'Product #', 'UPC', 'Artist', 'Title', 'MSRP', 'Qty', 'Received', 'Alerted', 'Unit cost'],
  UpcMap: ['Scanned', 'Expected UPC', 'Who', 'When'],
  Receiving: ['Timestamp', 'Who', 'Order', 'Expected', 'Scanned', 'Short', 'Over', 'Not on order', 'Packing list'],
  DropOffs: ['Timestamp', 'Who', 'Customer', 'Phone', 'Email', 'What they dropped off', 'Approx. records', 'ID', 'Status', 'Pay', 'Unwanted', 'Offer', 'Offer sent', 'Status changed', 'Last action by', 'Leftovers done', 'History', 'Card', 'Card issued', 'Photo'],
  Requests: ['ID', 'Created', 'Source', 'Shop', 'Name', 'Phone', 'Email', 'Artist', 'Album', 'Format', 'Notes', 'Status', 'Status changed', 'Quote', 'Quote note', 'Last action by', 'Hold until', 'Closed reason', 'Reminded'],
  Covers: ['Key', 'URL', 'Updated'],
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
    if (req.action === 'webrequest') return json_(webRequest_(req)); // public website form: no PIN, strict checks inside
    var who = authenticate_(req.pin);
    if (req.action === 'login') out = { ok: true, name: who, role: AUTH_ROLE_ };
    else if (req.action === 'reqlist') out = { ok: true, requests: reqList_(req) };
    else if (req.action === 'requpdate') out = { ok: true, updated: reqUpdate_(req, who) };
    else if (req.action === 'trove') { try { troveSoldSync_(false); } catch (e) { /* list still loads; sold check retries next time */ } out = { ok: true, items: troveItems_(), covers: coversMap_() }; }
    else if (req.action === 'droplist') out = { ok: true, drops: dropList_() };
    else if (req.action === 'dropphoto') out = { ok: true, photo: dropPhoto_(req) };
    else if (req.action === 'dropupdate') out = { ok: true, drop: dropUpdate_(req, who) };
    else if (req.action === 'coverset') out = { ok: true, saved: coverSet_(req.entries) };
    else if (req.action === 'registerlog') { if (AUTH_ROLE_ !== 'manager') throw new Error('Managers only'); out = { ok: true, counts: registerLog_() }; }
    else if (req.action === 'supplystate') out = { ok: true, state: supplyState_() };
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
  return json_(out);
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

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
var AUTH_ROLE_ = 'staff';
function authenticate_(pin) {
  var cache = CacheService.getScriptCache();
  var fails = Number(cache.get('fails') || 0);
  if (fails >= 8) throw new Error('Too many tries, wait a minute');
  var p = String(pin == null ? '' : pin).trim();
  var sh = ss_().getSheetByName('Staff');
  if (!sh) throw new Error('Run setup() first');
  var rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 3).getDisplayValues() : [];
  for (var i = 0; i < rows.length; i++) {
    if (p && String(rows[i][1]).trim() === p && String(rows[i][0]).trim()) {
      AUTH_ROLE_ = String(rows[i][2] || '').trim().toLowerCase() === 'manager' ? 'manager' : 'staff';
      return String(rows[i][0]).trim();
    }
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
        (d.cant || []).map(function (c) { return clean_(c.step) + ': ' + clean_(c.reason); }).join('; '), clean_(d.regnote)]);
      if (task === 'opening') { try { registerMismatch_(who, now, d.amounts); } catch (e) { /* an alert problem must never block opening */ } }
    } else if (task === 'note') {
      if (!summary) throw new Error('Empty note');
    } else if (task === 'issue') {
      var url = d.photo ? saveIssuePhoto_(String(d.photo), now, tz) : '';
      sheet_(ss, 'Issues').appendRow([now, who, clean_(d.type), clean_(d.details), url, 'Open']);
    } else if (task === 'cash') {
      sheet_(ss, 'Cash').appendRow([now, who, money_(d.drawer), money_(d.deposit), clean_(d.note), 'Count', '']);
    } else if (task === 'request') {
      requestEntry_(d, who, now, tz);
    } else if (task === 'receive') {
      try { sqQueue_(ss, d.got, now); } catch (e) { /* Square problems must never block a receipt */ }
      try { sqQueueExtras_(ss, d.extras, now); } catch (e) { /* same: never block a receipt */ }
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
      var dem = clean_(d.email), dpay = clean_(d.payment), dleft = clean_(d.unwanted);
      if (!clean_(d.name) || !clean_(d.items) || ['Cash', 'Store credit'].indexOf(dpay) < 0 || ['Donate', 'Return'].indexOf(dleft) < 0) throw new Error('Missing drop-off details');
      if (!EMAIL_RE_.test(dem)) throw new Error('A valid email is required'); // offers and store credit go by email
      var dsh = dropSheet_(), drow = new Array(DR_W).fill('');
      drow[DR.TS] = now; drow[DR.WHO] = who; drow[DR.NAME] = clean_(d.name); drow[DR.PHONE] = phoneFmt_(d.phone); drow[DR.EMAIL] = dem; drow[DR.ITEMS] = clean_(d.items);
      drow[DR.COUNT] = Number(d.count) || ''; drow[DR.ID] = 'D' + Utilities.getUuid().replace(/-/g, '').slice(0, 7).toUpperCase(); drow[DR.STATUS] = 'New';
      drow[DR.PAY] = dpay; drow[DR.LEFT] = dleft; drow[DR.CHANGED] = now; drow[DR.BY] = who; drow[DR.HIST] = '[]';
      drow[DR.PHOTO] = d.photo ? saveIssuePhoto_(String(d.photo), now, tz) : '';
      dsh.getRange(dsh.getLastRow() + 1, 1, 1, DR_W).setValues([drow]);
      result.drop = drow;
    }
  } finally {
    lock.releaseLock();
  }
  if (clean_(d.type) === 'Trove') { try { troveSquareSync_(false); } catch (e) { /* the TroveSync tab shows what is pending; re-run troveSquareRun */ } }
  if (result.drop) { try { dropAlert_(result.drop); } catch (e) { /* the drop-off is saved; a mail problem must never block it */ } }
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

/** In-store requests (Customer request tile) go to the shared Requests tab, same as the website form. */
function requestEntry_(d, who, now, tz) {
  var shop = PropertiesService.getScriptProperties().getProperty('SHOP') || 'Raleigh';
  addRequests_(ss_(), 'Store', shop, who, d, now, false);
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
/**
 * Emails Bell right away when this morning's opening count is not what the last close left behind
 * (last close count minus the cash moved to the deposit). Recipient: Script Property REGISTER_ALERT_TO, else the script owner.
 */
function registerMismatch_(who, now, amounts) {
  var m = /^\s*\$?(-?[\d,.]+)/.exec(String((amounts || {})['Count the register'] || ''));
  if (!m) return;
  var opened = parseFloat(m[1].replace(/,/g, ''));
  var prev = null, log = registerLog_();
  for (var i = 0; i < log.length; i++) {
    if (log[i].kind === 'Close' && log[i].amount !== '' && new Date(log[i].ts) < now) { prev = log[i]; break; }
  }
  if (!prev) return;
  var dep = prev.deposit === '' ? 0 : prev.deposit, expected = Math.round((prev.amount - dep) * 100) / 100, diff = Math.round((opened - expected) * 100) / 100;
  if (Math.abs(diff) < 0.005) return;
  var tz = ss_().getSpreadsheetTimeZone(), fmt = function (n) { return (n < 0 ? '-$' : '$') + Math.abs(n).toFixed(2); };
  var to = PropertiesService.getScriptProperties().getProperty('REGISTER_ALERT_TO') || Session.getEffectiveUser().getEmail();
  var body = 'The register does not match between the last close and this morning\'s open.\n\n' +
    'Last close: ' + fmt(prev.amount) + ' (' + prev.who + ', ' + Utilities.formatDate(new Date(prev.ts), tz, 'EEE MMM d') + ')\n' +
    'Moved to deposit: ' + fmt(dep) + '\n' +
    'Expected opening: ' + fmt(expected) + '\n' +
    'Counted at open: ' + fmt(opened) + ' (' + who + ', ' + Utilities.formatDate(now, tz, 'EEE MMM d, h:mm a') + ')\n' +
    'Difference: ' + (diff > 0 ? '+' : '') + fmt(diff) + (prev.note ? '\n\nNote on last close: ' + prev.note : '') +
    '\n\nThe full history is under Register count in HD Backstage.';
  MailApp.sendEmail(to, 'Hunky Dory Raleigh: register is off by ' + fmt(diff), body, { name: 'Hunky Dory Records' });
}

/** Running register-count log for managers: opening and closing "Count the register" amounts plus counts logged from the Register count tile. Newest first, max 600. */
function registerLog_() {
  var ss = ss_(), out = [];
  function tail_(name, cols, n) {
    var sh = ss.getSheetByName(name), last = sh ? sh.getLastRow() : 0;
    if (last < 2) return [];
    var from = Math.max(2, last - ((n || 200) - 1));
    return sh.getRange(from, 1, last - from + 1, cols).getValues();
  }
  ['Opening', 'Closing'].forEach(function (n) {
    tail_(n, 6).forEach(function (r) {
      var m = /Count the register:\s*\$?(-?[\d,.]+)/i.exec(String(r[2] || '')), dp = /Cash moved to the deposit[^:]*:\s*\$?(-?[\d,.]+)/i.exec(String(r[2] || ''));
      if (m && r[0]) out.push({ ts: new Date(r[0]).toISOString(), who: String(r[1]), kind: n === 'Opening' ? 'Open' : 'Close', amount: parseFloat(m[1].replace(/,/g, '')), deposit: dp ? parseFloat(dp[1].replace(/,/g, '')) : '', note: String(r[5] || '') });
    });
  });
  tail_('Cash', 7, 700).forEach(function (r) {
    if (!r[0]) return;
    var amt = r[2] === '' ? NaN : parseFloat(r[2]);
    if (isNaN(amt) && r[3] === '') return;
    out.push({ ts: new Date(r[0]).toISOString(), who: String(r[1]), kind: String(r[5] || 'Count'), amount: isNaN(amt) ? '' : amt, deposit: r[3] === '' ? '' : parseFloat(r[3]), note: String(r[4] || ''), src: String(r[6] || '') });
  });
  // The old register sheet is the record for the days it covers: drop live Opening/Closing rows on a day that has an imported row of the same kind.
  var tz = ss.getSpreadsheetTimeZone(), day = function (iso) { return Utilities.formatDate(new Date(iso), tz, 'yyyy-MM-dd'); }, imported = {};
  out.forEach(function (c) { if (c.src === 'Import') imported[c.kind + day(c.ts)] = true; });
  out = out.filter(function (c) { return c.src === 'Import' || c.src === 'Cash' || !imported[c.kind + day(c.ts)] || c.kind === 'Count'; });
  // The history is one Open and one Close per day: the first open and the last close. Mid-day counts are not part of it.
  out.sort(function (a, b) { return a.ts < b.ts ? -1 : 1; });
  var keep = {}, kept = [];
  out.forEach(function (c) {
    if (c.kind !== 'Open' && c.kind !== 'Close') return;
    var k = c.kind + day(c.ts);
    if (c.kind === 'Open') { if (keep[k] === undefined) { keep[k] = kept.length; kept.push(c); } }
    else { if (keep[k] === undefined) keep[k] = kept.length; kept[keep[k]] = c; }
  });
  out = kept;
  out.sort(function (a, b) { return a.ts < b.ts ? 1 : -1; });
  return out.slice(0, 600);
}

/** Open supply restock requests, from the Log: newest event per item wins (a "Restocked" line clears it). Looks back through the last 1500 log rows. */
function supplyState_() {
  var ss = ss_(), sh = sheet_(ss, 'Log'), last = sh.getLastRow(), state = {};
  if (last < 2) return state;
  var from = Math.max(2, last - 1499);
  sh.getRange(from, 1, last - from + 1, 6).getValues().forEach(function (r) {
    if (r[3] !== 'supplies') return;
    var d = {};
    try { d = JSON.parse(r[5] || '{}'); } catch (e) { return; }
    var name = clean_(d.item);
    if (!name) return;
    if (d.cancelled) delete state[name];
    else if (d.request) state[name] = { need: d.need || [], ts: new Date(r[0]).toISOString() };
  });
  return state;
}

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

/**
 * End-of-day supplies email. Reads today's "supplies" lines from the Log (restock requests and low-count alerts),
 * keeps the latest state per item (a cancelled request drops off), and sends ONE email only if something is open.
 * Recipients: Script Property SUPPLY_TO (comma separated), else OVERDUE_TO, else SHOP_EMAIL.
 */
function sendSupplyDigest() {
  var props = PropertiesService.getScriptProperties();
  var to = props.getProperty('SUPPLY_TO') || props.getProperty('OVERDUE_TO') || props.getProperty('SHOP_EMAIL') || '';
  if (!to) return;
  var ss = ss_(), tz = ss.getSpreadsheetTimeZone();
  var today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var sh = sheet_(ss, 'Log'), last = sh.getLastRow();
  if (last < 2) return;
  var from = Math.max(2, last - 499);
  var rows = sh.getRange(from, 1, last - from + 1, 6).getValues();
  var items = {}, order = [];
  rows.forEach(function (r) {
    var rd = r[1] instanceof Date ? Utilities.formatDate(r[1], tz, 'yyyy-MM-dd') : String(r[1]);
    if (rd !== today || r[3] !== 'supplies') return;
    var d = {};
    try { d = JSON.parse(r[5] || '{}'); } catch (e) { return; }
    var name = clean_(d.item);
    if (!name) return;
    if (order.indexOf(name) < 0) order.push(name);
    if (d.cancelled) { items[name] = null; return; }
    items[name] = d.request
      ? 'restock requested' + (d.need && d.need.length ? ' (' + d.need.join(', ') + ')' : '')
      : 'running low (' + d.count + ' left, reorder at ' + d.reorderAt + ')';
  });
  var lines = order.filter(function (n) { return items[n]; }).map(function (n) { return '- ' + n + ': ' + items[n]; });
  if (!lines.length) return;
  var body = 'Supplies flagged at Hunky Dory Raleigh today:\n\n' + lines.join('\n') +
    '\n\nThe full list is under Supplies in HD Backstage (To reorder tab).';
  MailApp.sendEmail(to, 'Hunky Dory Raleigh: supplies to reorder (' + lines.length + ')', body, { name: 'Hunky Dory Records' });
}

/** Run once from the editor: sends the supplies email every evening around 9. */
function installSupplyDigestTrigger() {
  setSupplyRecipients();
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'sendSupplyDigest') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('sendSupplyDigest').timeBased().everyDays(1).atHour(21).create();
}

/** Who gets the evening supplies email (also run by installSupplyDigestTrigger). */
function setSupplyRecipients() {
  PropertiesService.getScriptProperties().setProperty('SUPPLY_TO', 'forrest@hunkydoryvinyl.com,hunkydoryraleigh@gmail.com');
  Logger.log(PropertiesService.getScriptProperties().getProperty('SUPPLY_TO'));
}

/** Read-only check: Trove sheet vs Square. Logs counts; writes the TroveReview tab (Square Trove items in stock that are not on the sheet). */
function troveAudit() {
  var sh = ss_().getSheetByName('Trove');
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues();
  var statusOf = {}, nm = {}, by = {};
  rows.forEach(function (r) {
    if (!(r[0] || r[1])) return;
    var st = String(r[7] || '').trim() || 'Active'; by[st] = (by[st] || 0) + 1;
    var code = String(r[5] || '').trim(); if (!code) return;
    statusOf[code] = st; nm[code] = r[0] + ' - ' + r[1];
  });
  Logger.log('SHEET status counts: ' + JSON.stringify(by));
  var items = [], cursor = null;
  do {
    var b = { category_ids: [TS.cat], limit: 100 }; if (cursor) b.cursor = cursor;
    var s = sq_('post', '/v2/catalog/search-catalog-items', b);
    (s.items || []).forEach(function (it) { (it.item_data.variations || []).forEach(function (v) { items.push({ id: v.id, sku: v.item_variation_data.sku, name: it.item_data.name, price: v.item_variation_data.price_money ? v.item_variation_data.price_money.amount / 100 : '' }); }); });
    cursor = s.cursor;
  } while (cursor);
  var qty = {};
  for (var j = 0; j < items.length; j += 100) {
    var body = { catalog_object_ids: items.slice(j, j + 100).map(function (x) { return x.id; }), location_ids: [SQ.loc], states: ['IN_STOCK'] }, cur = null;
    do { if (cur) body.cursor = cur; var c = sq_('post', '/v2/inventory/counts/batch-retrieve', body);
      (c.counts || []).forEach(function (k) { qty[k.catalog_object_id] = Number(k.quantity); }); cur = c.cursor; } while (cur);
  }
  var groups = {}, samples = {};
  items.forEach(function (x) {
    var st = statusOf[x.sku] || 'NOT ON SHEET', stock = qty[x.id] > 0 ? 'in stock' : 'no stock';
    var k = st + ' / ' + stock; groups[k] = (groups[k] || 0) + 1;
    if (!(/^active/i.test(st) && stock === 'in stock')) { (samples[k] = samples[k] || []).length < 6 && samples[k].push(x.sku + ' ' + x.name); }
  });
  Logger.log('SQUARE Trove category: ' + items.length + ' items. By sheet status / Square stock: ' + JSON.stringify(groups));
  Object.keys(samples).forEach(function (k) { Logger.log(k + ' e.g. ' + samples[k].join(' ; ')); });
  var rev = items.filter(function (x) { return !statusOf[x.sku] && qty[x.id] > 0; }).map(function (x) { return [x.sku, x.name, x.price, qty[x.id]]; });
  var ts = ss_().getSheetByName('TroveReview') || ss_().insertSheet('TroveReview');
  ts.clear(); ts.getRange(1, 1, 1, 4).setValues([['Square SKU', 'Square item name', 'Price', 'Square stock']]).setFontWeight('bold');
  if (rev.length) ts.getRange(2, 1, rev.length, 4).setValues(rev);
  Logger.log('TroveReview tab written: ' + rev.length + ' rows');
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
  var sh = ss.getSheetByName('SquareSync');
  if (!sh) { sh = ss.insertSheet('SquareSync'); sh.getRange(1, 1, 1, TABS.SquareSync.length).setValues([TABS.SquareSync]).setFontWeight('bold'); sh.setFrozenRows(1); }
  sh.getRange('B:B').setNumberFormat('@'); sh.getRange('J:J').setNumberFormat('@');
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

/** Records scanned that were on no invoice (staff added artist/title/price): queue them for Square like any other receipt. SKU = the barcode. */
function sqQueueExtras_(ss, extras, now) {
  if (!extras || !extras.length) return;
  var rows = [];
  extras.forEach(function (e) {
    var u = String(e.upc || '').replace(/\D/g, ''), n = Math.floor(Number(e.qty) || 1), price = Math.round(Number(e.price) || 0);
    if (!u || price < 1) return;
    // The sync prices as ceil(MSRP)+1, so store price-1 to land on exactly the price staff typed.
    rows.push([now, u, clean_(e.artist), clean_(e.title), n, 'Pending', '', 'hd-recv-' + now.getTime() + '-x-' + u, price - 1, u, '']);
  });
  if (!rows.length) return;
  var sh = ss.getSheetByName('SquareSync');
  if (!sh) { sh = ss.insertSheet('SquareSync'); sh.getRange(1, 1, 1, TABS.SquareSync.length).setValues([TABS.SquareSync]).setFontWeight('bold'); sh.setFrozenRows(1); }
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

/* ---------- Trove -> Square ----------
 * Every Trove record (Trove tab) becomes its own Square item: name "Artist: Title", SKU = its TT code, price = the Trove price, stock 1, category Treasure Trove.
 * Step 1: run troveSquarePreview() - lists what it WOULD create in the "TroveSync" tab (status Planned). Nothing is sent to Square.
 * Step 2: after approval, run troveSquareRun() - creates the Planned rows (safe to re-run; skips codes already in Square; picks up where it stopped).
 * New Trove records added in the app sync automatically. Sold records are skipped.
 */
var TS = { cat: 'RFY4M6FBLOC2W3ZJKAIN5AK7', head: ['Code', 'Artist', 'Title', 'Price', 'Trove status', 'Sync status', 'Square item', 'Note'] };

function troveSquarePreview() { troveSquareSync_(true); }
function troveSquareRun() { troveSquareSync_(false); }

function troveSquareSync_(preview) {
  if (!preview && !PropertiesService.getScriptProperties().getProperty('SQUARE_TOKEN')) throw new Error('SQUARE_TOKEN not set');
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var ss = ss_(), tr = ss.getSheetByName('Trove');
    if (!tr || tr.getLastRow() < 2) return;
    var tab = ss.getSheetByName('TroveSync');
    if (!tab) { tab = ss.insertSheet('TroveSync'); tab.getRange(1, 1, 1, TS.head.length).setValues([TS.head]).setFontWeight('bold'); tab.setFrozenRows(1); }
    tab.getRange('A:A').setNumberFormat('@');
    var have = {};
    if (tab.getLastRow() > 1) tab.getRange(2, 1, tab.getLastRow() - 1, 1).getValues().forEach(function (r) { have[String(r[0])] = true; });
    var add = [];
    tr.getRange(2, 1, tr.getLastRow() - 1, 9).getValues().forEach(function (r) {
      var code = String(r[5] || '').trim();
      if (!code || !(r[0] || r[1]) || have[code]) return;
      var st = String(r[7] || 'Active'), price = Number(r[3]) || 0, sold = /sold/i.test(st);
      add.push([code, String(r[0]), String(r[1]), price || '', st, sold ? 'Skipped' : (price > 0 ? 'Planned' : 'Skipped'), '', sold ? 'Sold' : (price > 0 ? '' : 'No price')]);
    });
    if (add.length) tab.getRange(tab.getLastRow() + 1, 1, add.length, TS.head.length).setValues(add);
    if (preview) return;
    var rg = tab.getRange(2, 1, tab.getLastRow() - 1, TS.head.length), rows = rg.getValues(), t0 = Date.now();
    var todo = rows.filter(function (r) { return r[5] === 'Planned' || r[5] === 'Error'; });
    var found = {}, codes = todo.map(function (r) { return String(r[0]); });
    for (var i = 0; i < codes.length; i += 50) {
      var res = sq_('post', '/v2/catalog/search', { object_types: ['ITEM_VARIATION'], query: { set_query: { attribute_name: 'sku', attribute_values: codes.slice(i, i + 50) } } });
      (res.objects || []).forEach(function (o) { found[o.item_variation_data.sku] = o; });
    }
    todo.forEach(function (r) {
      if (Date.now() - t0 > 5 * 60 * 1000) return; // out of time: re-run to continue
      var code = String(r[0]);
      try {
        if (found[code]) { r[5] = 'Done'; r[6] = found[code].item_variation_data.item_id; r[7] = 'Already in Square'; return; }
        var res = sq_('post', '/v2/catalog/object', { idempotency_key: 'hd-trove-' + code, object: { type: 'ITEM', id: '#i', present_at_all_locations: false, present_at_location_ids: [SQ.loc],
          item_data: { name: String(r[1]) + ': ' + String(r[2]), tax_ids: [SQ.tax], categories: [{ id: TS.cat }], variations: [{ type: 'ITEM_VARIATION', id: '#v', present_at_all_locations: false,
            present_at_location_ids: [SQ.loc], item_variation_data: { item_id: '#i', name: '', sku: code, pricing_type: 'FIXED_PRICING',
              price_money: { amount: Math.round(Number(r[3]) * 100), currency: 'USD' }, track_inventory: true } }] } } });
        var v = res.catalog_object.item_data.variations[0];
        sq_('post', '/v2/inventory/changes/batch-create', { idempotency_key: 'hd-trove-' + code + '-stock', changes: [{ type: 'ADJUSTMENT', adjustment: {
          catalog_object_id: v.id, from_state: 'NONE', to_state: 'IN_STOCK', location_id: SQ.loc, quantity: '1', occurred_at: new Date().toISOString() } }] });
        r[5] = 'Done'; r[6] = res.catalog_object.id; r[7] = '';
      } catch (err) { r[5] = 'Error'; r[7] = String(err.message || err).slice(0, 200); }
    });
    rg.setValues(rows);
  } finally { lock.releaseLock(); }
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

/* ---------- Trove: mark sold records ----------
 * Looks up every Active Trove record in Square by its TT code (SKU). If the Square item exists but has no stock left,
 * the record sold at the register: its Status in the Trove tab becomes "Sold" and it drops off the scan list.
 * Runs when the app loads the Trove list (at most once every 5 minutes). Never writes to Square. */
function troveSoldSync_(force) {
  var cache = CacheService.getScriptCache();
  if (!force && cache.get('trovesold')) return 0;
  if (!PropertiesService.getScriptProperties().getProperty('SQUARE_TOKEN')) return 0;
  var sh = ss_().getSheetByName('Trove');
  if (!sh || sh.getLastRow() < 2) return 0;
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues(), idx = {}, codes = [];
  rows.forEach(function (r, n) {
    var code = String(r[5] || '').trim(), st = String(r[7] || 'Active');
    if (code && /^(active|)$/i.test(st.trim())) { idx[code] = n; codes.push(code); }
  });
  var varId = {};
  for (var i = 0; i < codes.length; i += 50) {
    var res = sq_('post', '/v2/catalog/search', { object_types: ['ITEM_VARIATION'], query: { set_query: { attribute_name: 'sku', attribute_values: codes.slice(i, i + 50) } } });
    (res.objects || []).forEach(function (o) { varId[o.id] = o.item_variation_data.sku; });
  }
  var ids = Object.keys(varId), inStock = {};
  for (var j = 0; j < ids.length; j += 100) {
    var body = { catalog_object_ids: ids.slice(j, j + 100), location_ids: [SQ.loc], states: ['IN_STOCK'] }, cur = null;
    do {
      if (cur) body.cursor = cur;
      var c = sq_('post', '/v2/inventory/counts/batch-retrieve', body);
      (c.counts || []).forEach(function (k) { if (Number(k.quantity) > 0) inStock[k.catalog_object_id] = true; });
      cur = c.cursor;
    } while (cur);
  }
  var sold = 0;
  ids.forEach(function (id) {
    if (inStock[id]) return;
    var n = idx[varId[id]];
    if (n === undefined) return;
    sh.getRange(n + 2, 8).setValue('Sold'); sold++;
  });
  cache.put('trovesold', '1', 300);
  return sold;
}
/** Run once by hand to check right now. */
function troveSoldNow() { Logger.log(troveSoldSync_(true) + ' records marked Sold'); }

/* =====================================================================================
 * CUSTOMER REQUESTS: one shared list (Requests tab) for the website's Request Line and the in-store tile.
 * Flow: New > Quoted (price entered; Forrest contacts the customer himself) > Ordered (customer said yes)
 *       > Waiting for pickup (arrived; held 10 days) > Picked up.   Closed = declined, can't get it, or no reply in 14 days.
 * Every shop's requests are stored (Shop column); a manager PIN only sees and changes the shop set in Script Property SHOP (default Raleigh).
 * Managers are rows in the Staff tab with Role = manager (third column).
 * ===================================================================================== */
var SHOPS = ['Raleigh', 'Cary', 'Durham', 'Apex'];
var QUOTE_EXPIRE_DAYS = 14, HOLD_DAYS = 10, HOLD_REMIND_DAY = 8;
var RQ = { ID: 0, CREATED: 1, SOURCE: 2, SHOP: 3, NAME: 4, PHONE: 5, EMAIL: 6, ARTIST: 7, ALBUM: 8, FORMAT: 9, NOTES: 10, STATUS: 11,
           CHANGED: 12, QUOTE: 13, QNOTE: 14, BY: 15, HOLD: 16, REASON: 17, REMINDED: 18 };
var EMAIL_RE_ = /^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/;

function myShop_() { return PropertiesService.getScriptProperties().getProperty('SHOP') || 'Raleigh'; }
function phoneFmt_(v) {
  var dg = String(v || '').replace(/\D/g, '');
  if (dg.length === 11 && dg.charAt(0) === '1') dg = dg.slice(1);
  return dg.length === 10 ? '(' + dg.slice(0, 3) + ') ' + dg.slice(3, 6) + '-' + dg.slice(6) : clean_(v);
}
function firstName_(n) { return String(n || '').trim().split(/\s+/)[0] || 'there'; }

/** Appends one row per record. web=true applies the website's stricter rules (phone and email both required, terms ticked). */
function addRequests_(ss, source, shop, who, d, now, web) {
  var name = clean_(d.name), phone = clean_(d.phone), email = clean_(d.email);
  if (SHOPS.indexOf(shop) < 0) throw new Error('Pick a shop');
  if (!name) throw new Error('Name is required');
  if (web) {
    if (String(phone).replace(/\D/g, '').length < 10) throw new Error('A 10-digit phone number is required');
    if (!EMAIL_RE_.test(email)) throw new Error('A valid email is required');
    if (d.terms !== true) throw new Error('Please agree to the request terms');
  } else {
    if (!(phone || email)) throw new Error('Phone or email is required');
    if (email && !EMAIL_RE_.test(email)) throw new Error('That email does not look right');
  }
  var items = (d.items && d.items.length ? d.items : [d]).slice(0, 10);
  var sh = sheet_(ss, 'Requests');
  var rows = items.map(function (it) {
    var artist = clean_(it.artist), album = clean_(it.album);
    if (!artist || !album) throw new Error('Artist and album are required');
    var row = new Array(19).fill('');
    row[RQ.ID] = 'R' + Utilities.getUuid().replace(/-/g, '').slice(0, 7).toUpperCase();
    row[RQ.CREATED] = now; row[RQ.SOURCE] = source; row[RQ.SHOP] = shop; row[RQ.NAME] = name;
    row[RQ.PHONE] = phoneFmt_(phone); row[RQ.EMAIL] = email; row[RQ.ARTIST] = artist; row[RQ.ALBUM] = album;
    row[RQ.FORMAT] = ['Vinyl', 'CD', 'Cassette', 'Tape', 'Other'].indexOf(clean_(it.format)) >= 0 ? clean_(it.format) : 'Vinyl';
    row[RQ.NOTES] = clean_(it.notes); row[RQ.STATUS] = 'New'; row[RQ.CHANGED] = now; row[RQ.BY] = who;
    return row;
  });
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  return rows;
}

/** Website Request Line posts here with no PIN: {action:'webrequest', shop, name, phone, email, terms, items:[{artist,album,format,notes}], hp}. */
function webRequest_(req) {
  if (clean_(req.hp)) return { ok: true }; // honeypot ("leave this empty") filled in: a bot. Pretend it worked, save nothing.
  var cache = CacheService.getScriptCache();
  var day = 'webreq-' + Utilities.formatDate(new Date(), 'UTC', 'yyyyMMdd');
  var n = Number(cache.get(day) || 0);
  if (n >= 100) return { ok: false, error: 'Too many requests today. Please stop by the shop.' };
  var em = clean_(req.email).toLowerCase(), ek = 'webmail-' + em;
  if (Number(cache.get(ek) || 0) >= 5) return { ok: false, error: 'Too many requests from this email today.' };
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  var rows;
  try {
    rows = addRequests_(ss_(), 'Web', clean_(req.shop), 'Website', req, new Date(), true);
    cache.put(day, String(n + 1), 86400); cache.put(ek, String(Number(cache.get(ek) || 0) + 1), 86400);
  } catch (err) {
    return { ok: false, error: String(err.message || err) };
  } finally {
    lock.releaseLock();
  }
  try { sendRequestConfirmation_(rows); } catch (err) { /* the request is saved even if the email fails */ }
  return { ok: true, count: rows.length };
}

function sendRequestConfirmation_(rows) {
  var r = rows[0], to = String(r[RQ.EMAIL] || '');
  if (!EMAIL_RE_.test(to)) return;
  var list = rows.map(function (x) { return '  • ' + x[RQ.ARTIST] + ', ' + x[RQ.ALBUM] + ' (' + x[RQ.FORMAT] + ')'; }).join('\n');
  sendAutoEmail_({
    to: to,
    subject: 'We got your request' + (rows.length > 1 ? 's' : '') + ': Hunky Dory Records ' + r[RQ.SHOP],
    body: 'Hi ' + firstName_(r[RQ.NAME]) + ',\n\nThanks for your request at Hunky Dory Records ' + r[RQ.SHOP] + ':\n\n' + list +
      '\n\nWe will contact you with a price quote. We do not order anything until you say yes. This is a request, not a guarantee.\n\nReply to this email if you have questions.\n\nHunky Dory Records'
  }, 'request', 'Website');
}

function requireManager_() { if (AUTH_ROLE_ !== 'manager') throw new Error('Managers only'); }

function reqObj_(row, rowNum) {
  var iso = function (v) { return v instanceof Date ? v.toISOString() : String(v || ''); };
  return { id: row[RQ.ID], row: rowNum, created: iso(row[RQ.CREATED]), source: row[RQ.SOURCE], shop: row[RQ.SHOP], name: row[RQ.NAME], phone: row[RQ.PHONE],
    email: row[RQ.EMAIL], artist: row[RQ.ARTIST], album: row[RQ.ALBUM], format: row[RQ.FORMAT], notes: row[RQ.NOTES], status: row[RQ.STATUS],
    changed: iso(row[RQ.CHANGED]), quote: row[RQ.QUOTE] === '' ? null : Number(row[RQ.QUOTE]), quoteNote: row[RQ.QNOTE], by: row[RQ.BY],
    holdUntil: iso(row[RQ.HOLD]), reason: row[RQ.REASON] };
}
function reqKey_(r) { return (String(r.artist) + '|' + String(r.album)).toLowerCase().replace(/[^a-z0-9|]/g, ''); }

/** Managers get EVERY request: all shops, every status, no date cutoff. Staff get only Ordered and Waiting for pickup for this shop, trimmed (name, record,
 *  price, hold date; phone only when there is no email on file, so they can call) so they can check people off at the counter. "asked" = open requests for the same record. */
function reqList_(req) {
  var isMgr = AUTH_ROLE_ === 'manager';
  var shop = myShop_(), sh = sheet_(ss_(), 'Requests');
  var vals = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 19).getValues() : [];
  var out = [], open = {};
  vals.forEach(function (row, i) {
    var o = reqObj_(row, i + 2), closed = ['Picked up', 'Closed'].indexOf(o.status) >= 0;
    if (!isMgr) {
      if (row[RQ.SHOP] !== shop || ['Ordered', 'Waiting for pickup'].indexOf(o.status) < 0) return;
      var hasEmail = EMAIL_RE_.test(String(o.email));
      o = { id: o.id, shop: o.shop, name: o.name, artist: o.artist, album: o.album, format: o.format, status: o.status, quote: o.quote,
            holdUntil: o.holdUntil, changed: o.changed, hasEmail: hasEmail, phone: hasEmail ? '' : o.phone };
    }
    if (!closed) open[o.shop + '|' + reqKey_(o)] = (open[o.shop + '|' + reqKey_(o)] || 0) + 1;
    out.push(o);
  });
  out.forEach(function (o) { o.asked = open[o.shop + '|' + reqKey_(o)] || 0; });
  return out;
}

var REQ_MOVES_ = { // action: [from statuses, new status]
  quote:    [['New', 'Quoted'], 'Quoted'],
  accepted: [['Quoted'], 'Ordered'],
  declined: [['Quoted'], 'Closed'],
  cantget:  [['New', 'Quoted', 'Ordered'], 'Closed'],
  arrived:  [['Ordered'], 'Waiting for pickup'],
  pickedup: [['Waiting for pickup'], 'Picked up']
};

/** Staff may only mark Picked up and Arrived (receiving); managers can do every step. {ids:[...], op:'quote'|'accepted'|'declined'|'cantget'|'arrived'|'pickedup', price, note}. Several ids = the same record asked for by several people. */
function reqUpdate_(req, who) {
  var op = clean_(req.op), mv = REQ_MOVES_[op], setSt = op === 'setstatus' ? clean_(req.status) : '';
  if (AUTH_ROLE_ !== 'manager' && op !== 'pickedup' && op !== 'arrived') throw new Error('Managers only'); // staff: check people off, and mark arrivals from receiving
  if (setSt && ['New', 'Quoted', 'Ordered', 'Waiting for pickup', 'Picked up', 'Closed'].indexOf(setSt) < 0) throw new Error('Unknown status');
  if (!mv && !setSt) throw new Error('Unknown request action');
  var ids = (req.ids || []).map(String).slice(0, 25);
  if (!ids.length) throw new Error('No request chosen');
  var price = Number(req.price);
  if (op === 'quote' && !(price > 0)) throw new Error('Enter the quoted price');
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  var changed = [], shop = myShop_(), now = new Date(), mails = [];
  try {
    var ss = ss_(), sh = sheet_(ss, 'Requests');
    var vals = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 19).getValues() : [];
    vals.forEach(function (row, i) {
      if (ids.indexOf(String(row[RQ.ID])) < 0) return;
      if (AUTH_ROLE_ !== 'manager' && row[RQ.SHOP] !== shop) throw new Error('That request belongs to another shop'); // managers work every shop's list
      if (!setSt && mv[0].indexOf(row[RQ.STATUS]) < 0) throw new Error(row[RQ.ARTIST] + ' is ' + row[RQ.STATUS] + ', so that step is not allowed');
      var ns = setSt || mv[1], was = row[RQ.STATUS];
      row[RQ.STATUS] = ns; row[RQ.CHANGED] = now; row[RQ.BY] = who;
      if (setSt) { // manager sets the status directly: no emails
        if (ns === 'Waiting for pickup') { row[RQ.HOLD] = new Date(now.getTime() + HOLD_DAYS * 86400000); row[RQ.REMINDED] = 'manual'; } else if (was === 'Waiting for pickup') row[RQ.HOLD] = '';
        row[RQ.REASON] = ns === 'Closed' ? (row[RQ.REASON] || 'Closed by manager') : '';
      }
      if (op === 'quote') { row[RQ.QUOTE] = price; row[RQ.QNOTE] = clean_(req.note); }
      if (op === 'declined') row[RQ.REASON] = 'Customer declined the quote';
      if (op === 'cantget') row[RQ.REASON] = "Can't get it";
      if (op === 'arrived') { var h = new Date(now.getTime() + HOLD_DAYS * 86400000); row[RQ.HOLD] = h; row[RQ.REMINDED] = ''; }
      sh.getRange(i + 2, 1, 1, 19).setValues([row]);
      changed.push(row);
      sheet_(ss, 'Log').appendRow([now, Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd'), who, 'reqstatus',
        'Request ' + row[RQ.NAME] + ' (' + row[RQ.ARTIST] + ' - ' + row[RQ.ALBUM] + '): ' + ns, JSON.stringify({ id: row[RQ.ID], op: op })]);
    });
    if (!changed.length) throw new Error('Request not found');
  } finally { lock.releaseLock(); }
  if (op === 'arrived' || op === 'cantget') {
    var byCust = {};
    changed.forEach(function (r) { if (EMAIL_RE_.test(String(r[RQ.EMAIL]))) (byCust[r[RQ.EMAIL]] = byCust[r[RQ.EMAIL]] || []).push(r); });
    Object.keys(byCust).forEach(function (em) { try { (op === 'arrived' ? sendArrival_ : sendCantGet_)(byCust[em]); } catch (e) { /* status already saved */ } });
  }
  return changed.map(function (r) { return { id: r[RQ.ID], status: r[RQ.STATUS] }; });
}

function dateLong_(d) { return Utilities.formatDate(d, 'America/New_York', 'EEEE, MMMM d'); }
function sendArrival_(rows) {
  var r = rows[0], one = rows.length === 1, hold = rows[0][RQ.HOLD] instanceof Date ? rows[0][RQ.HOLD] : new Date(rows[0][RQ.HOLD]);
  var list = rows.map(function (x) { return '  • ' + x[RQ.ARTIST] + ', ' + x[RQ.ALBUM]; }).join('\n');
  sendAutoEmail_({ to: r[RQ.EMAIL],
    subject: one ? "It's in: " + r[RQ.ARTIST] + ', ' + r[RQ.ALBUM] : 'Your ' + rows.length + ' records are in',
    body: 'Hi ' + firstName_(r[RQ.NAME]) + ',\n\n' + (one ? 'Your record is' : 'Your records are') + ' in at Hunky Dory Records ' + r[RQ.SHOP] + ':\n\n' + list +
      '\n\nWe will hold ' + (one ? 'it' : 'them') + ' for you through ' + dateLong_(hold) + '. Reply to this email if you would like to cancel.\n\nHunky Dory Records' }, 'request', 'Backstage');
}
function sendCantGet_(rows) {
  var r = rows[0];
  rows.forEach(function (x) {
    sendAutoEmail_({ to: x[RQ.EMAIL], subject: 'Update on your record request: ' + x[RQ.ARTIST] + ', ' + x[RQ.ALBUM],
      body: 'Hi ' + firstName_(x[RQ.NAME]) + ',\n\nThanks for asking about ' + x[RQ.ARTIST] + ', ' + x[RQ.ALBUM] + '. It is currently out of stock or backordered and we are not able to get it.\n\nHunky Dory Records' }, 'request', 'Backstage');
  });
}

/** Daily: quotes with no answer after 14 days close themselves; held records get a reminder on day 8. Run installRequestsTrigger() once. */
function requestsDaily() {
  var ss = ss_(), sh = sheet_(ss, 'Requests');
  if (sh.getLastRow() < 2) return;
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, 19).getValues(), now = new Date(), closed = 0, reminded = 0;
  vals.forEach(function (row, i) {
    var changed = new Date(row[RQ.CHANGED]).getTime(), dirty = false;
    if (row[RQ.STATUS] === 'Quoted' && now.getTime() - changed >= QUOTE_EXPIRE_DAYS * 86400000) {
      row[RQ.STATUS] = 'Closed'; row[RQ.REASON] = 'No response to quote (auto-closed)'; row[RQ.CHANGED] = now; row[RQ.BY] = 'Auto'; dirty = true; closed++;
    } else if (row[RQ.STATUS] === 'Waiting for pickup' && !row[RQ.REMINDED] && now.getTime() - changed >= HOLD_REMIND_DAY * 86400000 && EMAIL_RE_.test(String(row[RQ.EMAIL]))) {
      try {
        var hold = row[RQ.HOLD] instanceof Date ? row[RQ.HOLD] : new Date(row[RQ.HOLD]);
        sendAutoEmail_({ to: row[RQ.EMAIL], subject: 'Reminder: ' + row[RQ.ARTIST] + ', ' + row[RQ.ALBUM] + ' is waiting for you',
          body: 'Hi ' + firstName_(row[RQ.NAME]) + ',\n\nA reminder that ' + row[RQ.ARTIST] + ', ' + row[RQ.ALBUM] + ' is being held for you at Hunky Dory Records ' + row[RQ.SHOP] +
            ' through ' + dateLong_(hold) + '. Reply to this email if you would like to cancel.\n\nHunky Dory Records' }, 'request', 'Backstage');
        row[RQ.REMINDED] = now; dirty = true; reminded++;
      } catch (e) { /* tries again tomorrow */ }
    }
    if (dirty) sh.getRange(i + 2, 1, 1, 19).setValues([row]);
  });
  Logger.log(closed + ' quotes auto-closed, ' + reminded + ' pickup reminders sent');
}
function installRequestsTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'requestsDaily') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('requestsDaily').timeBased().everyDays(1).atHour(9).create();
}


/* =====================================================================================
 * CUSTOMER DROP-OFFS: one shared queue Bell works from.
 * Status: New > Offer made > Accepted | Declined > Paid > Done. Rows with no Status are old history (shown as Done, never changed).
 * Everyone signed in sees a trimmed list (no phone, email or amounts). Managers (Bell) see everything and make the offer, record the answer,
 * pay cash, or issue store credit as a Square gift card (once per drop-off, even if tapped twice).
 * Script Properties: DROPOFF_ALERT_TO (who is emailed about each new drop-off; falls back to REGISTER_ALERT_TO, then the script owner),
 * SQUARE_TOKEN (needs gift card write permission) for store credit.
 * ===================================================================================== */
var DR = { TS: 0, WHO: 1, NAME: 2, PHONE: 3, EMAIL: 4, ITEMS: 5, COUNT: 6, ID: 7, STATUS: 8, PAY: 9, LEFT: 10, OFFER: 11, SENT: 12, CHANGED: 13, BY: 14,
           LEFTDONE: 15, HIST: 16, CARD: 17, CARDOK: 18, PHOTO: 19 };
var DR_W = 20;

function dropSheet_() {
  var sh = sheet_(ss_(), 'DropOffs');
  if (sh.getMaxColumns() < DR_W) sh.insertColumnsAfter(sh.getMaxColumns(), DR_W - sh.getMaxColumns());
  if (String(sh.getRange(1, DR.ID + 1).getValue()) !== TABS.DropOffs[DR.ID]) sh.getRange(1, 1, 1, DR_W).setValues([TABS.DropOffs]).setFontWeight('bold');
  return sh;
}

function dropHist_(v) { try { var a = JSON.parse(String(v || '[]')); return Array.isArray(a) ? a : []; } catch (e) { return []; } }

function dropObj_(row, n, mgr) {
  var iso = function (v) { return v instanceof Date ? v.toISOString() : String(v || ''); };
  var legacy = !row[DR.STATUS];
  var o = { id: String(row[DR.ID] || 'L' + n), ts: iso(row[DR.TS]), who: String(row[DR.WHO]), name: String(row[DR.NAME]), items: String(row[DR.ITEMS]),
    pay: String(row[DR.PAY] || ''), left: String(row[DR.LEFT] || ''), status: legacy ? 'Done' : String(row[DR.STATUS]),
    sts: iso(row[DR.CHANGED] || row[DR.TS]), leftDone: row[DR.LEFTDONE] === 'Yes', legacy: legacy, hasCard: !!row[DR.CARD], hasPhoto: !!row[DR.PHOTO] };
  if (mgr) {
    o.phone = String(row[DR.PHONE] || ''); o.email = String(row[DR.EMAIL] || ''); o.offer = row[DR.OFFER] === '' ? '' : String(Number(row[DR.OFFER]));
    o.sent = iso(row[DR.SENT]); o.hist = dropHist_(row[DR.HIST]); o.photo = String(row[DR.PHOTO] || ''); o.cardOk = !!row[DR.CARDOK];
  }
  return o;
}

function dropList_() {
  var mgr = AUTH_ROLE_ === 'manager', sh = dropSheet_(), last = sh.getLastRow();
  if (last < 2) return [];
  var from = Math.max(2, last - 399), vals = sh.getRange(from, 1, last - from + 1, DR_W).getValues();
  return vals.map(function (r, i) { return dropObj_(r, from + i, mgr); }).filter(function (o) { return o.name; });
}

function dropAlert_(r) {
  var to = PropertiesService.getScriptProperties().getProperty('DROPOFF_ALERT_TO') || PropertiesService.getScriptProperties().getProperty('REGISTER_ALERT_TO') || Session.getEffectiveUser().getEmail();
  var body = 'A customer dropped off records at Hunky Dory Records ' + myShop_() + '.\n\n' +
    'Customer: ' + r[DR.NAME] + '\nEmail: ' + r[DR.EMAIL] + '\nPhone: ' + (r[DR.PHONE] || 'none') + '\nWhat they dropped off: ' + r[DR.ITEMS] + '\n' +
    'Wants to be paid in: ' + r[DR.PAY] + '\nRecords we do not want: ' + String(r[DR.LEFT]).toLowerCase() + '\nLogged by: ' + r[DR.WHO] + '\n' + (r[DR.PHOTO] ? 'Photo: ' + r[DR.PHOTO] + '\n' : '') +
    '\nMake the offer under Drop-off list in HD Backstage.';
  MailApp.sendEmail(to, 'New drop-off: ' + r[DR.NAME], body, { name: 'Hunky Dory Records' });
}

/** The photo staff took at drop-off, as a base64 JPEG (any signed-in user). */
function dropPhoto_(req) {
  var id = clean_(req.id), sh = dropSheet_(), last = sh.getLastRow();
  var vals = last > 1 ? sh.getRange(2, 1, last - 1, DR_W).getValues() : [];
  for (var i = 0; i < vals.length; i++) if (String(vals[i][DR.ID]) === id) {
    var m = /[-\w]{25,}/.exec(String(vals[i][DR.PHOTO] || ''));
    if (!m) return '';
    return Utilities.base64Encode(DriveApp.getFileById(m[0]).getBlob().getBytes());
  }
  throw new Error('Drop-off not found');
}

/** {id, op, amount}. Managers: offer, accepted, declined, paid (cash), credit (store credit gift card), undo. Any staff: leftdone, done. Returns the changed drop-off. */
function dropUpdate_(req, who) {
  var op = clean_(req.op), id = clean_(req.id);
  if (['offer', 'accepted', 'declined', 'paid', 'credit', 'undo'].indexOf(op) >= 0) requireManager_();
  else if (['leftdone', 'done'].indexOf(op) < 0) throw new Error('Unknown drop-off action');
  var amount = Math.round(Number(req.amount) * 100) / 100;
  if (op === 'offer' && !(amount > 0 && amount < 100000)) throw new Error('Enter the offer amount');
  var FROM = { offer: ['New', 'Offer made'], accepted: ['Offer made'], declined: ['Offer made'], paid: ['Accepted'], credit: ['Accepted'], leftdone: ['Paid', 'Declined'], done: ['Paid', 'Declined'] };
  var lock = LockService.getScriptLock(); lock.waitLock(15000);
  var row, rowNum, now = new Date(), mail = null, card = null;
  try {
    var ss = ss_(), sh = dropSheet_(), last = sh.getLastRow();
    var vals = last > 1 ? sh.getRange(2, 1, last - 1, DR_W).getValues() : [];
    for (var i = 0; i < vals.length; i++) if (String(vals[i][DR.ID]) === id) { row = vals[i]; rowNum = i + 2; break; }
    if (!row) throw new Error('Drop-off not found');
    var st = String(row[DR.STATUS]);
    if (!st) throw new Error('That is an old record and cannot be changed');
    var hist = dropHist_(row[DR.HIST]);
    if (op === 'undo') {
      if (!hist.length || st === 'New') throw new Error('Nothing to undo');
      if (row[DR.CARD]) throw new Error('A gift card was already issued, so this cannot be undone');
      row[DR.STATUS] = hist.pop(); row[DR.LEFTDONE] = '';
    } else if (op === 'leftdone') {
      if (FROM.leftdone.indexOf(st) < 0) throw new Error('Not ready for that step');
      row[DR.LEFTDONE] = row[DR.LEFTDONE] === 'Yes' ? '' : 'Yes';
    } else {
      if (FROM[op].indexOf(st) < 0) throw new Error('This drop-off is ' + st + ', so that step is not allowed');
      var ns = { offer: 'Offer made', accepted: 'Accepted', declined: 'Declined', paid: 'Paid', credit: 'Paid', done: 'Done' }[op];
      if (op === 'paid' && row[DR.PAY] !== 'Cash') throw new Error('This customer chose store credit');
      if (op === 'credit' && row[DR.PAY] !== 'Store credit') throw new Error('This customer chose cash');
      if (op === 'done' && st === 'Paid' && row[DR.LEFTDONE] !== 'Yes') throw new Error('Handle the leftover records first');
      if (op === 'offer') { row[DR.OFFER] = amount; row[DR.SENT] = now; mail = 'offer'; }
      if (op === 'credit') {
        try { card = dropIssueCard_(row, id); } catch (e) { sh.getRange(rowNum, 1, 1, DR_W).setValues([row]); throw e; } // keep the card id if only the activation failed
        mail = 'card';
      }
      if (st !== ns) hist.push(st);
      row[DR.STATUS] = ns;
    }
    row[DR.HIST] = JSON.stringify(hist.slice(-12)); row[DR.CHANGED] = now; row[DR.BY] = who;
    sh.getRange(rowNum, 1, 1, DR_W).setValues([row]);
    sheet_(ss, 'Log').appendRow([now, Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd'), who, 'dropstatus', 'Drop-off ' + row[DR.ID] + ': ' + row[DR.STATUS], JSON.stringify({ id: id, op: op })]);
  } finally { lock.releaseLock(); }
  try { if (mail === 'offer') sendDropOffer_(row); else if (mail === 'card') sendDropCard_(row, card); } catch (e) { /* the step is saved; the customer email can be resent by hand */ }
  return dropObj_(row, rowNum, true);
}

/** Creates and loads a Square digital gift card for the accepted offer. Safe to call twice: the card id is saved before loading, and both Square calls carry fixed idempotency keys. */
function dropIssueCard_(row, id) {
  var cents = Math.round(Number(row[DR.OFFER]) * 100);
  if (!(cents > 0)) throw new Error('No offer amount on this drop-off');
  var gid = '', gan = '';
  if (row[DR.CARD]) { var p = String(row[DR.CARD]).split('|'); gid = p[0]; gan = p[1] || ''; }
  else {
    var c = sq_('POST', '/v2/gift-cards', { idempotency_key: 'hd-drop-' + id, location_id: SQ.loc, gift_card: { type: 'DIGITAL' } }).gift_card || {};
    if (!c.id) throw new Error('Square did not create the gift card');
    gid = c.id; gan = c.gan || ''; row[DR.CARD] = gid + '|' + gan;
  }
  if (!row[DR.CARDOK]) {
    sq_('POST', '/v2/gift-cards/activities', { idempotency_key: 'hd-drop-act-' + id, gift_card_activity: { type: 'ACTIVATE', location_id: SQ.loc, gift_card_id: gid,
      activate_activity_details: { amount_money: { amount: cents, currency: 'USD' }, buyer_payment_instrument_ids: ['hd-dropoff-' + id], reference_id: id } } });
    row[DR.CARDOK] = new Date();
  }
  return { gid: gid, gan: gan, amount: cents / 100 };
}

function dropMoney_(n) { return '$' + Number(n).toFixed(2); }

function sendDropOffer_(row) {
  var cash = row[DR.PAY] === 'Cash', shop = myShop_();
  sendAutoEmail_({ to: row[DR.EMAIL], subject: 'Your offer from Hunky Dory Records ' + shop,
    body: 'Hi ' + firstName_(row[DR.NAME]) + ',\n\nThanks for bringing your records in to Hunky Dory Records ' + shop + '. We looked through them and can offer you ' + dropMoney_(row[DR.OFFER]) +
      (cash ? ' in cash' : ' in store credit') + ' for the records we are keeping.\n\n' +
      (row[DR.LEFT] === 'Donate' ? 'The records we are not buying will be donated.' : 'The records we are not buying are ours to hold for you. You can pick them up at the shop.') +
      '\n\nReply to this email to accept or decline. If we do not hear back within 10 days, we may donate or dispose of your records, as in the drop-off terms you agreed to.' +
      (cash ? '' : '\n\nStore credit is sent to this email address as a gift card.') + '\n\nHunky Dory Records' }, 'dropoff', 'Backstage');
}

function sendDropCard_(row, card) {
  var gan = String(card.gan || '').replace(/(\d{4})(?=\d)/g, '$1 ');
  sendAutoEmail_({ to: row[DR.EMAIL], subject: 'Your Hunky Dory Records store credit: ' + dropMoney_(card.amount),
    body: 'Hi ' + firstName_(row[DR.NAME]) + ',\n\nYour store credit is ready: ' + dropMoney_(card.amount) + ' at Hunky Dory Records.\n\nGift card number: ' + gan +
      '\n\nGive this number to us at the register and we will take it off your total. Keep this email; it is your gift card.\n\nHunky Dory Records' }, 'dropoff', 'Backstage');
}

/** Shared album-cover links so every iPad looks a cover up once, ever. */
function coversSheet_() {
  var ss = ss_(), sh = ss.getSheetByName('Covers');
  if (!sh) { sh = ss.insertSheet('Covers'); sh.getRange(1, 1, 1, 3).setValues([TABS.Covers]).setFontWeight('bold'); sh.setFrozenRows(1); }
  return sh;
}
function coversMap_() {
  var sh = coversSheet_(), n = sh.getLastRow(), out = {};
  if (n < 2) return out;
  sh.getRange(2, 1, n - 1, 2).getValues().forEach(function (r) { if (r[0] && r[1]) out[String(r[0])] = String(r[1]); });
  return out;
}
function coverSet_(entries) {
  if (!Array.isArray(entries)) return 0;
  var have = coversMap_(), rows = [], now = new Date();
  entries.slice(0, 200).forEach(function (e) {
    var k = String(e && e.k || '').slice(0, 200), u = String(e && e.u || '').trim();
    if (!k || have[k] || !/^https:\/\/[^\s]{4,400}$/.test(u)) return;
    have[k] = u; rows.push([clean_(k), u, now]);
  });
  if (rows.length) { var sh = coversSheet_(); sh.getRange(sh.getLastRow() + 1, 1, rows.length, 3).setValues(rows); }
  return rows.length;
}
