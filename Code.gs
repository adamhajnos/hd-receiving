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

var TABS = {
  Staff:    ['Name', 'PIN'],
  Trove:    ['Artist', 'Title', 'Location', 'Price', 'Cover URL'],
  Log:      ['Timestamp', 'Date', 'Who', 'Task', 'Summary', 'Data (JSON)'],
  Opening:  ['Timestamp', 'Who', 'Amounts', 'Steps done', 'Could not do (step: reason)'],
  Closing:  ['Timestamp', 'Who', 'Amounts', 'Steps done', 'Could not do (step: reason)'],
  Issues:   ['Timestamp', 'Who', 'Type', 'Details', 'Photo', 'Status'],
  Cash:     ['Timestamp', 'Who', 'Drawer count', 'Deposit', 'Note'],
  Records: ['Timestamp', 'Who', 'Type', 'Artist', 'Album', 'Genre', 'Format', 'Barcode', 'MSRP', 'Price', 'Status', 'Location', 'Listed on', 'Label queued', 'Notes'],
  Receiving: ['Timestamp', 'Who', 'Order', 'Expected', 'Scanned', 'Short', 'Over', 'Not on order', 'Packing list'],
  DropOffs: ['Timestamp', 'Who', 'Customer', 'Phone', 'Email', 'What they dropped off', 'Approx. records']
};

/** Run once from the editor to create the tabs. */
function setup() {
  var ss = SpreadsheetApp.getActive();
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
    else if (req.action === 'log') { logEntry_(req, who); out = { ok: true }; }
    else throw new Error('Unknown action');
  } catch (err) {
    out = { ok: false, error: String(err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/** Returns the staff name for a PIN, or throws. Locks guessing after 8 misses in 5 minutes. */
function authenticate_(pin) {
  var cache = CacheService.getScriptCache();
  var fails = Number(cache.get('fails') || 0);
  if (fails >= 8) throw new Error('Too many tries, wait a minute');
  var p = String(pin == null ? '' : pin).trim();
  var sh = SpreadsheetApp.getActive().getSheetByName('Staff');
  if (!sh) throw new Error('Run setup() first');
  var rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 2).getDisplayValues() : [];
  for (var i = 0; i < rows.length; i++) {
    if (p && String(rows[i][1]).trim() === p && String(rows[i][0]).trim()) return String(rows[i][0]).trim();
  }
  cache.put('fails', String(fails + 1), 300);
  throw new Error('Wrong PIN');
}

/** The running Trove list: one row per record in the "Trove" tab. */
function troveItems_() {
  var sh = sheet_(SpreadsheetApp.getActive(), 'Trove');
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 5).getDisplayValues()
    .filter(function (r) { return r[0] || r[1]; })
    .map(function (r) { return { artist: r[0], title: r[1], location: r[2], price: r[3], cover: r[4] === 'none' ? '' : r[4] }; });
}

/**
 * Fills the "Cover URL" column for any Trove row that has none, using the free
 * iTunes Search API. Rows with no match are marked "none" so they are not retried.
 * Safe to re-run: it stops before Apps Script's 6-minute limit and picks up the rest next time.
 * Clear a cell to retry it, or paste a better URL by hand.
 */
function fillCovers() {
  var sh = sheet_(SpreadsheetApp.getActive(), 'Trove');
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
  try {
    var ss = SpreadsheetApp.getActive();
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
      var pack = d.packing ? saveIssuePhoto_(String(d.packing), now, tz) : (d.noPacking ? 'None in box' : '');
      sheet_(ss, 'Receiving').appendRow([now, who, clean_(d.order), Number(d.expected) || 0, Number(d.scanned) || 0,
        (d.short || []).map(clean_).join('; '), (d.over || []).map(clean_).join('; '), (d.extra || []).map(clean_).join('; '), pack]);
    } else if (task === 'record') {
      if (!clean_(d.artist) || !clean_(d.album) || !(Number(d.price) > 0)) throw new Error('Missing record details');
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
  if (d.autoEmail) sendAutoEmail_(d.autoEmail, task, who);
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
  var sh = sheet_(SpreadsheetApp.getActive(), 'Log');
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
  var ss = SpreadsheetApp.getActive();
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
