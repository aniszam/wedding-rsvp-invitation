/**
 * Gift Registry backend — deploy this as its OWN Apps Script project bound
 * to a NEW Google Sheet (keep it separate from the RSVP sheet/script).
 *
 * The item catalog AND claim status both live in one sheet tab named
 * "Items". Edit the sheet directly to add, remove, or reorder gifts — no
 * code changes needed.
 *
 * Sheet columns (row 1 = exact header names below):
 *   itemId | name | price | link | icon | names | claimedAt
 *
 * - itemId:    unique slug (e.g. "stand-mixer"). Never change it once a
 *              guest may have claimed that row, or the claim link breaks.
 * - name:      display name, e.g. "KitchenAid Stand Mixer"
 * - price:     display price, e.g. "RM 2,500"
 * - link:      URL to the item (optional — leave blank to hide the
 *              "View Item" link on that card)
 * - icon:      Font Awesome class, e.g. "fa-solid fa-mug-hot" (optional —
 *              leave blank to use the default gift icon)
 * - names:     leave BLANK. Filled in automatically when a guest claims it.
 * - claimedAt: leave BLANK. Filled in automatically when claimed.
 *
 * "Cash Fund" is hardcoded on the front end (gift/index.html) as a
 * mandatory first option — it is NOT a row in the Items sheet. Every
 * contribution is appended as its own row to an auto-created "CashFundLog"
 * tab (columns: names | claimedAt) instead of the single-claim-per-row
 * logic used for the other items, since any number of guests can
 * contribute to it.
 *
 * Setup:
 * 1. Create a new Google Sheet. Rename the first tab to "Items".
 * 2. Add the header row above, then one row per gift item (leave the
 *    names/claimedAt cells blank for every row).
 * 3. Extensions -> Apps Script, paste this file's contents in, save.
 * 4. Deploy -> New deployment -> type "Web app".
 *      - Execute as: Me
 *      - Who has access: Anyone
 * 5. Copy the /exec URL it gives you into gift/index.html:
 *      window.__GOOGLE_SHEET_URL = 'https://script.google.com/macros/s/XXXX/exec';
 * 6. Re-deploy (new version) any time you edit this script.
 */

function doGet(e) {
  var sheet = getSheet();
  var rows = sheet.getDataRange().getValues();
  var headers = rows.shift();
  var items = rows
    .filter(function (r) { return r[0]; })
    .map(function (r) {
      var obj = {};
      headers.forEach(function (h, i) { obj[h] = r[i]; });
      return obj;
    });
  return jsonOutput(items);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var itemId = e.parameter.itemId;
    if (!itemId) {
      return jsonOutput({ ok: false, error: 'missing_item_id' });
    }

    // "Cash Fund" is a hardcoded special case on the front end (not a row
    // in the Items sheet) — every contribution is logged, never blocked,
    // since any number of guests can contribute.
    if (itemId === 'cash-fund') {
      var logSheet = getCashFundLogSheet();
      logSheet.appendRow([
        e.parameter.names || '',
        e.parameter.claimedAt || new Date().toISOString()
      ]);
      return jsonOutput({ ok: true });
    }

    var sheet = getSheet();
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var itemIdCol = headers.indexOf('itemId');
    var namesCol = headers.indexOf('names');
    var claimedAtCol = headers.indexOf('claimedAt');

    for (var i = 1; i < data.length; i++) {
      if (data[i][itemIdCol] === itemId) {
        if (data[i][namesCol]) {
          return jsonOutput({ ok: false, error: 'already_claimed' });
        }
        sheet.getRange(i + 1, namesCol + 1).setValue(e.parameter.names || '');
        sheet.getRange(i + 1, claimedAtCol + 1).setValue(e.parameter.claimedAt || new Date().toISOString());
        return jsonOutput({ ok: true });
      }
    }

    return jsonOutput({ ok: false, error: 'item_not_found' });
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
  var sheet = ss.getSheetByName('Items');
  if (!sheet) {
    sheet = ss.insertSheet('Items');
    sheet.appendRow(['itemId', 'name', 'price', 'link', 'icon', 'names', 'claimedAt']);
  }
  return sheet;
}

function getCashFundLogSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('CashFundLog');
  if (!sheet) {
    sheet = ss.insertSheet('CashFundLog');
    sheet.appendRow(['names', 'claimedAt']);
  }
  return sheet;
}
