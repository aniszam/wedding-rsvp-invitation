/**
 * RSVP backend for the 12-Dec-2026 page — deploy this as its OWN Apps
 * Script project bound to a NEW Google Sheet, separate from the main
 * wedding's RSVP sheet and from the gift registry sheet.
 *
 * Sheet: one tab named "RSVP" with header row (row 1):
 *   name | attending | paxAdult | paxKids | pax | message | submittedAt
 *
 * paxAdult/paxKids are the separate adult/kids counts from the RSVP form;
 * pax is their sum, kept for convenience.
 *
 * The sheet/tab is created automatically on first submission if missing,
 * so you don't need to set up the header row by hand — just create a
 * blank Google Sheet and deploy this script against it.
 *
 * Setup:
 * 1. Create a new Google Sheet (blank is fine).
 * 2. Extensions -> Apps Script, delete the placeholder code, paste this
 *    file's contents in, save.
 * 3. Deploy -> New deployment -> type "Web app".
 *      - Execute as: Me
 *      - Who has access: Anyone
 * 4. Copy the /exec URL it gives you and send it back so it can be
 *    wired into 12-Dec-2026/index.html (window.__GOOGLE_SHEET_URL).
 * 5. Re-deploy (new version, not just save) any time you edit this
 *    script — Apps Script Web Apps don't pick up edits otherwise.
 */

var RSVP_CACHE_KEY = 'rsvp_12dec_entries';
var RSVP_CACHE_SECONDS = 15;

function doGet(e) {
  var cache = CacheService.getScriptCache();
  var cached = cache.get(RSVP_CACHE_KEY);
  if (cached) {
    return jsonOutput(JSON.parse(cached));
  }

  var sheet = getSheet();
  var rows = sheet.getDataRange().getValues();
  var headers = rows.shift();
  var entries = rows
    .filter(function (r) { return r[0]; })
    .map(function (r) {
      var obj = {};
      headers.forEach(function (h, i) { obj[h] = r[i]; });
      return obj;
    });

  cache.put(RSVP_CACHE_KEY, JSON.stringify(entries), RSVP_CACHE_SECONDS);
  return jsonOutput(entries);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var name = e.parameter.name || '';
    if (!name) {
      return jsonOutput({ ok: false, error: 'missing_name' });
    }

    var sheet = getSheet();
    sheet.appendRow([
      name,
      e.parameter.attending || '',
      e.parameter.paxAdult || '0',
      e.parameter.paxKids || '0',
      e.parameter.pax || '0',
      e.parameter.message || '',
      e.parameter.submittedAt || new Date().toISOString()
    ]);

    CacheService.getScriptCache().remove(RSVP_CACHE_KEY);
    return jsonOutput({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('RSVP');
  if (!sheet) {
    sheet = ss.insertSheet('RSVP');
    sheet.appendRow(['name', 'attending', 'paxAdult', 'paxKids', 'pax', 'message', 'submittedAt']);
  }
  return sheet;
}
