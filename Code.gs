// ============================================================
// Code.gs  —  Jewellery Purchase Register
// Mahalaxmi Jewellers | v1.0.0
//
// Architecture (inherited & extended from reference project):
//   doGet()  → JSON status page
//   doPost() → REST API entry point
//   routeAction_() → action dispatcher
//
// Sheets expected in the Google Spreadsheet:
//   Config  (Key | Value)
//   Users   (Username | Password | Role | Active)
//   Records (15 columns — see getRecords() for schema)
//
// Config sheet only needs "Login Required" to start.
// All other config keys use defaults from CFG_DEFAULTS.
// "Next Record Number" is appended automatically on first save.
// ============================================================


// ── Config Defaults ──────────────────────────────────────────
// Values used when a key is absent from the Config sheet.
var CFG_DEFAULTS = {
  'Business Name':       'Mahalaxmi Jewellers',
  'Business Address':    '',
  'Business Phone':      '',
  'Record Prefix':       'MJ',
  'Next Record Number':  1,
  'Currency Symbol':     '₹',
  'Theme Color':         '#033C3C',
  'Session Timeout':     480,
  'Version':             '1.0.0',
  'Login Required':      'TRUE'
};


// ── ID Validation ─────────────────────────────────────────────
// Texts that indicate a column header rather than a real record ID.
var RECORD_SKIP_TEXTS = ['record id', 'id', 'recordid'];

/**
 * Returns true if val looks like a real Record ID:
 *   - non-empty string
 *   - contains at least one digit
 *   - is not a header label
 */
function isValidId_(val) {
  if (val === null || val === undefined) return false;
  var s = String(val).trim();
  return s !== '' && RECORD_SKIP_TEXTS.indexOf(s.toLowerCase()) === -1 && /\d/.test(s);
}


// ── Web App Entry Points ──────────────────────────────────────


// ── CORS Helper ───────────────────────────────────────────────
/**
 * Builds a JSON response with the CORS headers required so that a
 * Vercel-hosted frontend (or any cross-origin caller) can read the
 * response without being blocked by the browser.
 *
 * Google Apps Script's ContentService does not support setHeader(),
 * so we use HtmlService which does — but we still return plain JSON
 * text so the frontend's res.json() call works normally.
 */
function buildCorsResponse_(payload) {
  var json = JSON.stringify(payload);
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * GET handler.
 * Handles two cases:
 *   1. ?action=... query param → routes to the action (useful for testing in browser)
 *   2. No params → returns a simple status JSON (health check)
 *
 * ⚠️  Always deploy with "Execute as: Me" + "Who has access: Anyone"
 *     and use the /exec URL, NOT the /dev URL.
 *     /dev requires Google login and causes CORS errors from Vercel.
 */
function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    var result = routeAction_(
      e.parameter.action,
      e.parameter.data ? JSON.parse(e.parameter.data) : null
    );
    return buildCorsResponse_(result);
  }
  return buildCorsResponse_({
    ok:      true,
    message: 'Jewellery Purchase Register API v1.0 is running.',
    hint:    'Use POST with { action, data } body to call API functions.'
  });
}

/**
 * POST handler — primary API entry point called by the Vercel frontend.
 * Expected request body (JSON string with Content-Type: text/plain):
 *   { action: 'actionName', data: { ...payload } }
 *
 * Using Content-Type: text/plain avoids CORS preflight (OPTIONS) requests
 * while still letting us JSON.parse the body on the server side.
 */
function doPost(e) {
  var result;
  try {
    var body   = JSON.parse(e.postData.contents);
    var action = body.action;
    var data   = body.data || null;
    result = routeAction_(action, data);
  } catch (err) {
    result = { ok: false, error: 'doPost error: ' + err.message };
  }
  return buildCorsResponse_(result);
}

/**
 * Routes an action string to the appropriate handler function.
 * Add new cases here as the app grows.
 */
function routeAction_(action, data) {
  switch (action) {
    case 'getConfig':     return getConfig();
    case 'login':         return login(data);
    case 'getRecords':    return getRecords();
    case 'saveRecord':    return saveRecord(data);
    case 'updateRecord':  return updateRecord(data);
    case 'deleteRecord':  return deleteRecord(data && data.id);
    case 'deleteRecords': return deleteRecords(data && data.ids);
    default:              return { ok: false, error: 'Unknown action: ' + action };
  }
}




// ── getConfig ─────────────────────────────────────────────────
/**
 * Returns all config values as a flat key→value map.
 * Merges Config sheet values on top of CFG_DEFAULTS,
 * so the sheet only needs to contain overrides.
 *
 * Config sheet format:
 *   Row 1 (optional header): Key | Value
 *   Row 2+: actual key-value pairs
 */
function getConfig() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getConfigSheet_(ss);

    // Start with a full copy of defaults
    var cfg = {};
    var defaultKeys = Object.keys(CFG_DEFAULTS);
    for (var i = 0; i < defaultKeys.length; i++) {
      cfg[defaultKeys[i]] = CFG_DEFAULTS[defaultKeys[i]];
    }

    // Override with whatever is in the Config sheet
    if (sheet) {
      var vals = sheet.getDataRange().getValues();
      for (var j = 0; j < vals.length; j++) {
        var k = String(vals[j][0]).trim();
        var v = String(vals[j][1]).trim();
        // Skip blank rows and any optional header row
        if (k && k.toLowerCase() !== 'key') {
          cfg[k] = v;
        }
      }
    }

    // Coerce numeric fields to their proper types
    cfg['Next Record Number'] = parseInt(cfg['Next Record Number']) || 1;
    cfg['Session Timeout']    = parseInt(cfg['Session Timeout'])    || 480;

    return { ok: true, data: cfg };
  } catch (e) {
    return { ok: false, error: 'getConfig: ' + e.message };
  }
}


// ── login ─────────────────────────────────────────────────────
/**
 * Validates credentials against the Users sheet.
 *
 * Users sheet columns (row 1 = header):
 *   A: Username  B: Password  C: Role  D: Active (TRUE/FALSE)
 *
 * Returns: { ok: true, user: { username, role } }
 *       or { ok: false, error: '...' }
 */
function login(data) {
  try {
    if (!data || !data.username || !data.password) {
      return { ok: false, error: 'Username and password are required.' };
    }

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getUsersSheet_(ss);

    if (!sheet) {
      return {
        ok: false,
        error: 'Users sheet not found. Create a sheet named "Users" with columns: Username | Password | Role | Active'
      };
    }

    var rows  = sheet.getDataRange().getValues();
    var uname = String(data.username).trim().toLowerCase();
    var upass = String(data.password).trim();

    // Row 0 is the header — start from row 1
    for (var i = 1; i < rows.length; i++) {
      var rowUser   = String(rows[i][0]).trim().toLowerCase();
      var rowPass   = String(rows[i][1]).trim();
      var rowRole   = String(rows[i][2]).trim();
      var rowActive = String(rows[i][3]).trim().toUpperCase();

      if (rowUser === uname && rowPass === upass && rowActive === 'TRUE') {
        return {
          ok:   true,
          user: { username: String(rows[i][0]).trim(), role: rowRole }
        };
      }
    }

    return { ok: false, error: 'Invalid credentials or account is inactive.' };
  } catch (e) {
    return { ok: false, error: 'login: ' + e.message };
  }
}


// ── getRecords ────────────────────────────────────────────────
/**
 * Returns all purchase records from the Records sheet.
 *
 * Records sheet columns (row 1 = header, data from row 2):
 *   A: Record ID       B: Date            C: Customer Name
 *   D: Address         E: Mobile          F: Metal Type
 *   G: Gross WT        H: Deduction WT    I: Before Melt WT
 *   J: After Melt WT   K: Final Price     L: Purity
 *   M: Remarks         N: Created By      O: Created Time
 */
function getRecords() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getRecordsSheet_(ss);
    if (!sheet) return { ok: false, error: 'Records sheet not found. Create a sheet named "Records".' };

    var lastRow = getLastDataRow_(sheet);
    if (lastRow < 2) return { ok: true, data: [] }; // No data rows yet

    var vals = sheet.getRange(2, 1, lastRow - 1, 15).getValues();
    var tz   = Session.getScriptTimeZone();
    var rows = [];

    for (var i = 0; i < vals.length; i++) {
      var r = vals[i];
      if (!isValidId_(r[0])) continue; // skip header-like or blank rows

      // Parse and format the purchase date (column B)
      var dateVal = '';
      if (r[1]) {
        try {
          var d = new Date(r[1]);
          if (!isNaN(d.getTime())) {
            dateVal = Utilities.formatDate(d, tz, 'dd/MM/yyyy');
          }
        } catch (_) {}
      }

      // Parse and format the created timestamp (column O)
      var createdTime = '';
      if (r[14]) {
        try {
          var ct = new Date(r[14]);
          if (!isNaN(ct.getTime())) {
            createdTime = Utilities.formatDate(ct, tz, 'dd/MM/yyyy HH:mm');
          }
        } catch (_) {}
      }

      rows.push({
        recordId:     String(r[0]).trim(),
        date:         dateVal,
        customerName: String(r[2]  || ''),
        address:      String(r[3]  || ''),
        mobile:       String(r[4]  || ''),
        metalType:    String(r[5]  || 'Gold'),
        grossWT:      Number(r[6])  || 0,
        deductionWT:  Number(r[7])  || 0,
        beforeMeltWT: Number(r[8])  || 0,
        afterMeltWT:  Number(r[9])  || 0,
        finalPrice:   Number(r[10]) || 0,
        purity:       String(r[11] || ''),
        remarks:      String(r[12] || ''),
        createdBy:    String(r[13] || ''),
        createdTime:  createdTime
      });
    }

    return { ok: true, data: rows };
  } catch (e) {
    return { ok: false, error: 'getRecords: ' + e.message };
  }
}


// ── saveRecord ────────────────────────────────────────────────
/**
 * Saves a new purchase record.
 * Auto-generates the Record ID (e.g. MJ000001).
 * Stamps Created By and Created Time automatically.
 * Returns: { ok: true, recordId: 'MJ000001' }
 */
function saveRecord(data) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getRecordsSheet_(ss);
    if (!sheet) return { ok: false, error: 'Records sheet not found.' };

    var id      = generateRecordId_(ss);
    var lastRow = getLastDataRow_(sheet);
    var newRow  = lastRow + 1;
    var now     = new Date();

    sheet.getRange(newRow, 1, 1, 15).setValues([[
      id,                                         // A: Record ID (auto)
      data.date ? new Date(data.date) : now,      // B: Date
      String(data.customerName || ''),            // C: Customer Name
      String(data.address      || ''),            // D: Address
      String(data.mobile       || ''),            // E: Mobile
      String(data.metalType    || 'Gold'),        // F: Metal Type
      Number(data.grossWT)      || 0,             // G: Gross WT
      Number(data.deductionWT)  || 0,             // H: Deduction WT
      Number(data.beforeMeltWT) || 0,             // I: Before Melt WT
      Number(data.afterMeltWT)  || 0,             // J: After Melt WT
      Number(data.finalPrice)   || 0,             // K: Final Price
      String(data.purity       || ''),            // L: Purity
      String(data.remarks      || ''),            // M: Remarks
      String(data.createdBy    || ''),            // N: Created By
      now                                         // O: Created Time (auto)
    ]]);

    SpreadsheetApp.flush();
    return { ok: true, recordId: id };
  } catch (e) {
    return { ok: false, error: 'saveRecord: ' + e.message };
  }
}


// ── updateRecord ──────────────────────────────────────────────
/**
 * Updates an existing record by its Record ID.
 * Preserves the original Created By and Created Time.
 * Never locks records — always editable per user permissions.
 */
function updateRecord(data) {
  try {
    if (!isValidId_(data.recordId)) return { ok: false, error: 'Invalid ID: ' + data.recordId };

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getRecordsSheet_(ss);
    if (!sheet) return { ok: false, error: 'Records sheet not found.' };

    var row = findRowById_(sheet, data.recordId);
    if (!row) return { ok: false, error: 'Record not found: ' + data.recordId };

    // Read existing row to preserve audit columns (N and O)
    var existing = sheet.getRange(row, 1, 1, 15).getValues()[0];

    sheet.getRange(row, 1, 1, 15).setValues([[
      data.recordId,                              // A: Record ID (unchanged)
      data.date ? new Date(data.date) : existing[1], // B: Date
      String(data.customerName || ''),            // C: Customer Name
      String(data.address      || ''),            // D: Address
      String(data.mobile       || ''),            // E: Mobile
      String(data.metalType    || 'Gold'),        // F: Metal Type
      Number(data.grossWT)      || 0,             // G: Gross WT
      Number(data.deductionWT)  || 0,             // H: Deduction WT
      Number(data.beforeMeltWT) || 0,             // I: Before Melt WT
      Number(data.afterMeltWT)  || 0,             // J: After Melt WT
      Number(data.finalPrice)   || 0,             // K: Final Price
      String(data.purity       || ''),            // L: Purity
      String(data.remarks      || ''),            // M: Remarks
      existing[13],                               // N: Created By (preserved)
      existing[14]                                // O: Created Time (preserved)
    ]]);

    SpreadsheetApp.flush();
    return { ok: true };
  } catch (e) {
    return { ok: false, error: 'updateRecord: ' + e.message };
  }
}


// ── deleteRecord ──────────────────────────────────────────────
/**
 * Deletes a single record row by its Record ID.
 * Access control (Admin only) is enforced on the frontend;
 * this function does not re-validate role.
 */
function deleteRecord(id) {
  try {
    if (!isValidId_(id)) return { ok: false, error: 'Not a valid Record ID: ' + id };

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getRecordsSheet_(ss);
    if (!sheet) return { ok: false, error: 'Records sheet not found.' };

    var row = findRowById_(sheet, id);
    if (!row) return { ok: false, error: 'Record not found: ' + id };

    sheet.deleteRow(row);
    return { ok: true };
  } catch (e) {
    return { ok: false, error: 'deleteRecord: ' + e.message };
  }
}


// ── deleteRecords (Bulk) ──────────────────────────────────────
/**
 * Deletes multiple records by an array of Record IDs.
 * Rows are sorted descending before deletion to prevent
 * index-shifting from invalidating subsequent row numbers.
 */
function deleteRecords(ids) {
  try {
    if (!Array.isArray(ids) || ids.length === 0) {
      return { ok: false, error: 'No IDs provided.' };
    }

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getRecordsSheet_(ss);
    if (!sheet) return { ok: false, error: 'Records sheet not found.' };

    // Collect valid row numbers
    var rowsToDelete = [];
    for (var i = 0; i < ids.length; i++) {
      if (isValidId_(ids[i])) {
        var row = findRowById_(sheet, ids[i]);
        if (row) rowsToDelete.push(row);
      }
    }

    // Sort rows descending — delete from bottom to top
    rowsToDelete.sort(function (a, b) { return b - a; });
    for (var j = 0; j < rowsToDelete.length; j++) {
      sheet.deleteRow(rowsToDelete[j]);
    }

    SpreadsheetApp.flush();
    return { ok: true, deleted: rowsToDelete.length };
  } catch (e) {
    return { ok: false, error: 'deleteRecords: ' + e.message };
  }
}


// ── Record ID Generator ───────────────────────────────────────
/**
 * Generates the next sequential Record ID, e.g. MJ000001.
 *
 * Algorithm:
 *   1. Read 'Record Prefix' and 'Next Record Number' from Config sheet.
 *   2. If 'Next Record Number' row is found → increment and update it.
 *   3. If 'Next Record Number' row is missing → append it (handles the
 *      case where Config only has "Login Required" so far).
 *   4. Return prefix + zero-padded number.
 */
function generateRecordId_(ss) {
  var sheet = getConfigSheet_(ss);

  var prefix     = CFG_DEFAULTS['Record Prefix']; // 'MJ'
  var lastNum    = 0;
  var lastNumRow = -1; // 0-based index within getDataRange values

  if (sheet) {
    var vals = sheet.getDataRange().getValues();

    for (var i = 0; i < vals.length; i++) {
      var k = String(vals[i][0]).trim();
      if (k === 'Record Prefix')      prefix     = String(vals[i][1]).trim() || prefix;
      if (k === 'Next Record Number') { lastNum = parseInt(vals[i][1]) || 0; lastNumRow = i; }
    }

    var newNum = lastNum + 1;

    if (lastNumRow >= 0) {
      // "Next Record Number" row already exists — update column B
      // Index i is 0-based; sheet row = i + 1
      sheet.getRange(lastNumRow + 1, 2).setValue(newNum);
    } else {
      // Row does not exist — append it to the Config sheet automatically
      var lastDataRow = getLastDataRow_(sheet);
      sheet.getRange(lastDataRow + 1, 1).setValue('Next Record Number');
      sheet.getRange(lastDataRow + 1, 2).setValue(newNum);
    }

    return prefix + String(newNum).padStart(6, '0');
  }

  // Fallback: Config sheet is missing entirely
  return prefix + String(Date.now()).slice(-6);
}


// ── Sheet Accessors ───────────────────────────────────────────

/**
 * Gets the Config sheet.
 * Tries named range 'RANGECONFIG' first, then falls back to sheet name 'Config'.
 */
function getConfigSheet_(ss) {
  var rng = ss.getRangeByName('RANGECONFIG');
  if (rng) { try { return rng.getSheet(); } catch (e) {} }
  return ss.getSheetByName('Config') || null;
}

/**
 * Gets the Users sheet.
 * Tries named range 'RANGEUSERS' first, then falls back to sheet name 'Users'.
 */
function getUsersSheet_(ss) {
  var rng = ss.getRangeByName('RANGEUSERS');
  if (rng) { try { return rng.getSheet(); } catch (e) {} }
  return ss.getSheetByName('Users') || null;
}

/**
 * Gets the Records sheet.
 * Tries named range 'RANGERECORDS' first, then falls back to sheet name 'Records'.
 */
function getRecordsSheet_(ss) {
  var rng = ss.getRangeByName('RANGERECORDS');
  if (rng) { try { return rng.getSheet(); } catch (e) {} }
  return ss.getSheetByName('Records') || null;
}


// ── Row Helpers ───────────────────────────────────────────────

/**
 * Returns the last row index (1-based) that has a non-empty value in column A.
 * Walks backwards from sheet.getLastRow() to skip trailing empty rows.
 * Returns 1 (header row) if the sheet has no data.
 */
function getLastDataRow_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return 1;
  var vals = sheet.getRange(1, 1, lastRow, 1).getValues();
  for (var i = vals.length - 1; i >= 0; i--) {
    if (vals[i][0] !== '') return i + 1;
  }
  return 1;
}

/**
 * Finds the sheet row number (1-based) of a record by its ID in column A.
 * Uses TextFinder for efficiency on large sheets.
 * Returns null if not found.
 */
function findRowById_(sheet, id) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  var hit = sheet.getRange(2, 1, lastRow - 1, 1)
    .createTextFinder(String(id).trim())
    .matchEntireCell(true)
    .findNext();
  return hit ? hit.getRow() : null;
}
