/**
 * CMC Report Automation — patient lookup (Google Apps Script web app).
 *
 * The desktop app sends a PIN (Anderson ID); this script, running as your
 * Google account, reads the PRIVATE patient sheet and sequencing QC sheet
 * and returns only that one patient's demography and QC (After Data / Q30)
 * fields. Neither sheet has to be shared.
 *
 * SET UP (once)
 *   1. Open the script project (a standalone project from script.google.com
 *      works: the sheets are opened by ID below).
 *   2. Replace everything in Code.gs with this file, fill in
 *      PATIENT_SPREADSHEET_ID and click Save.
 *   3. Deploy → New deployment → gear icon → Web app.
 *        Description:     CMC patient lookup
 *        Execute as:      Me
 *        Who has access:  Anyone
 *      Click Deploy, then Authorize access (choose your account →
 *      Advanced → Go to project → Allow).
 *   4. Copy the Web app URL (ends in /exec) and send it to be put in the app.
 *   5. Test in a browser:   <Web app URL>?pin=ADK0000001234
 *      and for the Whole Exome Sequencing sheet:  ...?pin=ADK0000001234&sheet=wes
 *   6. Then restrict both sheets: Share → General access → Restricted.
 *
 * AFTER EDITING THIS SCRIPT
 *   Deploy → Manage deployments → pencil → Version: New version → Deploy.
 *   (Keeps the same /exec URL, so the app needs no change.)
 *
 * RESPONSES (JSON)
 *   {"ok": true, "found": true,  "row": {"Anderson ID": "...", ...}, "qc": {...}}
 *   {"ok": true, "found": false, "qc": {"After Data": "9.96 GB", "Q30": "-95.85%"}}
 *   {"ok": true, "found": false, "qc": null}
 *   {"ok": false, "error": "..."}
 *   With sheet=wes the reply also has "sheet": "wes" and "found"/"row" come
 *   from the WES sheet (its "Client" column is returned as "Client name").
 *   "found" is about the patient sheet; "qc" is null when the PIN isn't in
 *   the QC sheet. A QC-only problem comes back as "qc_error" next to the row.
 */

// Patient list: the spreadsheet ID (between /d/ and /edit in its URL) and
// the tab's "gid=" number.
var PATIENT_SPREADSHEET_ID = 'PASTE_PATIENT_SPREADSHEET_ID_HERE';
var PATIENT_GID = 534459671;

// Sequencing QC sheet (Anderson_ID / After Data / Q30). Every tab whose
// header row has QC_PIN_COLUMN is searched (one tab per batch), left to right.
var QC_SPREADSHEET_ID = 'PASTE_QC_SPREADSHEET_ID_HERE';
var QC_PIN_COLUMN = 'Anderson_ID';
var QC_COLUMNS = ['After Data', 'Q30'];

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

// Whole Exome Sequencing sheet (?sheet=wes). Its header is on row 2, under a
// group row that heads the otherwise blank client column "Client". A PIN
// listed more than once gives its latest (lowest) row.
var WES_SPREADSHEET_ID = 'PASTE_WES_SPREADSHEET_ID_HERE';
var WES_GID = 0;
var WES_HEADER_ROW = 2;
var WES_RETURN_COLUMNS = [
  'Anderson ID',
  'Sample Number',
  'Name',
  'Received Date',
  'Sample Type',
  'Client'
];

// Protects against someone using the URL to page through every PIN.
var MAX_LOOKUPS_PER_MINUTE = 60;


function doGet(e) {
  try {
    var pin = String((e && e.parameter && e.parameter.pin) || '').trim().toUpperCase();
    if (!/^[A-Z0-9]{3,6}\d{6,14}$/.test(pin)) {
      return json_({ ok: false, error: 'Invalid or missing PIN.' });
    }
    if (!withinRateLimit_()) {
      return json_({ ok: false, error: 'Too many lookups - please wait a minute and try again.' });
    }

    var qc = null, qcError = '';
    try {
      qc = lookupAllTabs_(QC_SPREADSHEET_ID, QC_PIN_COLUMN, QC_COLUMNS, pin);
    } catch (err) {
      qcError = 'QC sheet: ' + (err.message || err);
    }
    var wes = String((e.parameter && e.parameter.sheet) || '').toLowerCase() === 'wes';
    var row = wes
      ? lookup_(WES_SPREADSHEET_ID, WES_GID, PIN_COLUMN, WES_RETURN_COLUMNS, pin,
                WES_HEADER_ROW, true)
      : lookup_(PATIENT_SPREADSHEET_ID, PATIENT_GID, PIN_COLUMN, RETURN_COLUMNS, pin);
    if (row && wes) {
      row['Client name'] = row['Client'];
      delete row['Client'];
    }
    var reply = { ok: true, found: !!row, qc: qc };
    if (wes) reply.sheet = 'wes';
    if (row) reply.row = row;
    if (qcError) reply.qc_error = qcError;
    return json_(reply);
  } catch (err) {
    return json_({ ok: false, error: 'Lookup failed: ' + (err.message || err) });
  }
}


// Finder for a cell holding exactly `pin` (surrounding spaces allowed).
// Searching with Find is much faster than reading the whole sheet.
function pinFinder_(range, pin) {
  return range.createTextFinder('^\\s*' + pin + '\\s*$')
    .useRegularExpression(true).matchCase(false);
}


// `columns` (by header name) from one row's values.
function pick_(header, values, columns) {
  var row = {};
  columns.forEach(function (name) {
    var c = header.indexOf(norm_(name));
    row[name] = c >= 0 ? String(values[c]).trim() : '';
  });
  return row;
}


// The `columns` of the row whose `pinColumn` equals `pin`, or null.
// headerRow: 1-based row of the column names (default 1); a blank name takes
// the label in the row above it. last: the PIN's lowest row, not its first.
function lookup_(spreadsheetId, gid, pinColumn, columns, pin, headerRow, last) {
  var sheet = findSheet_(spreadsheetId, gid);
  var top = headerRow || 1;
  var lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
  if (lastRow <= top || lastCol < 1) return null;
  // Display values keep dates exactly as shown in the sheet (e.g. 29-08-2026).
  var head = sheet.getRange(top > 1 ? top - 1 : 1, 1, top > 1 ? 2 : 1, lastCol)
    .getDisplayValues();
  var above = top > 1 ? head[0] : [];
  var header = head[head.length - 1].map(function (h, c) {
    return norm_(h) || norm_(above[c]);
  });
  var pinCol = header.indexOf(norm_(pinColumn));
  if (pinCol < 0) {
    throw new Error('Column "' + pinColumn + '" not found in the header row of ' +
                    sheet.getName() + '.');
  }
  var hits = pinFinder_(sheet.getRange(top + 1, pinCol + 1, lastRow - top, 1), pin).findAll();
  if (!hits.length) return null;
  var r = hits[last ? hits.length - 1 : 0].getRow();
  return pick_(header, sheet.getRange(r, 1, 1, lastCol).getDisplayValues()[0], columns);
}


// The same over every tab of a spreadsheet: the first cell holding `pin`
// (tabs left to right) that sits in a column headed `pinColumn` in row 1.
function lookupAllTabs_(spreadsheetId, pinColumn, columns, pin) {
  var hits = pinFinder_(SpreadsheetApp.openById(spreadsheetId), pin).findAll();
  for (var i = 0; i < hits.length; i++) {
    var sheet = hits[i].getSheet();
    var lastCol = sheet.getLastColumn();
    var header = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(norm_);
    if (header[hits[i].getColumn() - 1] !== norm_(pinColumn)) continue;
    return pick_(header, sheet.getRange(hits[i].getRow(), 1, 1, lastCol)
                           .getDisplayValues()[0], columns);
  }
  return null;
}


function findSheet_(spreadsheetId, gid) {
  var sheets = SpreadsheetApp.openById(spreadsheetId).getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === gid) return sheets[i];
  }
  throw new Error('Tab gid=' + gid + ' not found in spreadsheet ' + spreadsheetId + '.');
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
