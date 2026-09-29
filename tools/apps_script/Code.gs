/**
 * CMC Report Automation — patient lookup (Google Apps Script web app).
 *
 * The desktop app sends a PIN (Anderson ID); this script, running as your
 * Google account, reads the PRIVATE patient sheet and returns only that one
 * patient's demography fields. The sheet itself never has to be shared.
 *
 * SET UP (once)
 *   1. Open the patient sheet → Extensions → Apps Script.
 *   2. Replace everything in Code.gs with this file and click Save.
 *   3. Deploy → New deployment → gear icon → Web app.
 *        Description:     CMC patient lookup
 *        Execute as:      Me
 *        Who has access:  Anyone
 *      Click Deploy, then Authorize access (choose your account →
 *      Advanced → Go to project → Allow).
 *   4. Copy the Web app URL (ends in /exec) and send it to be put in the app.
 *   5. Test in a browser:   <Web app URL>?pin=ADK0000001234
 *   6. Then restrict the sheet: Share → General access → Restricted.
 *
 * AFTER EDITING THIS SCRIPT
 *   Deploy → Manage deployments → pencil → Version: New version → Deploy.
 *   (Keeps the same /exec URL, so the app needs no change.)
 *
 * RESPONSES (JSON)
 *   {"ok": true, "found": true,  "row": {"Anderson ID": "...", ...}}
 *   {"ok": true, "found": false}
 *   {"ok": false, "error": "..."}
 */

// Tab that holds the patient list (the "gid=" number in the sheet's URL).
var SHEET_GID = 534459671;

// The only columns ever returned. Matched by header text (case/space
// insensitive), so moving columns around doesn't break anything.
var RETURN_COLUMNS = [
  'Anderson ID',
  'Sample Number',
  'Name',
  'Received Date',
  'Sample Type',
  'Client name',
  'Client Doctor Name'
];

var PIN_COLUMN = 'Anderson ID';

// Protects against someone using the URL to page through every PIN.
var MAX_LOOKUPS_PER_MINUTE = 60;


function doGet(e) {
  try {
    var pin = String((e && e.parameter && e.parameter.pin) || '').trim().toUpperCase();
    if (!/^[A-Z]{2,5}\d{6,14}$/.test(pin)) {
      return json_({ ok: false, error: 'Invalid or missing PIN.' });
    }
    if (!withinRateLimit_()) {
      return json_({ ok: false, error: 'Too many lookups - please wait a minute and try again.' });
    }

    var sheet = findSheet_();
    if (!sheet) {
      return json_({ ok: false, error: 'Patient tab not found (check SHEET_GID).' });
    }

    // Display values keep dates exactly as shown in the sheet (e.g. 29-08-2026).
    var values = sheet.getDataRange().getDisplayValues();
    if (values.length < 2) {
      return json_({ ok: true, found: false });
    }

    var header = values[0].map(norm_);
    var pinCol = header.indexOf(norm_(PIN_COLUMN));
    if (pinCol < 0) {
      return json_({ ok: false, error: 'Column "' + PIN_COLUMN + '" not found in the header row.' });
    }

    for (var r = 1; r < values.length; r++) {
      if (String(values[r][pinCol]).trim().toUpperCase() !== pin) continue;
      var row = {};
      RETURN_COLUMNS.forEach(function (name) {
        var c = header.indexOf(norm_(name));
        row[name] = c >= 0 ? String(values[r][c]).trim() : '';
      });
      return json_({ ok: true, found: true, row: row });
    }
    return json_({ ok: true, found: false });
  } catch (err) {
    return json_({ ok: false, error: 'Lookup failed: ' + err });
  }
}


function findSheet_() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === SHEET_GID) return sheets[i];
  }
  return null;
}


function withinRateLimit_() {
  var cache = CacheService.getScriptCache();
  var key = 'lookups_' + Math.floor(Date.now() / 60000);
  var lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    var count = Number(cache.get(key) || 0) + 1;
    cache.put(key, String(count), 120);
    return count <= MAX_LOOKUPS_PER_MINUTE;
  } finally {
    lock.releaseLock();
  }
}


function norm_(text) {
  return String(text || '').replace(/\s+/g, ' ').trim().toLowerCase();
}


function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
