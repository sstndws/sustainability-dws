/**
 * MERGED: sustain-dashboard repo — includes Mill QUARTER/YEAR ↔ Quarter/Year mapping in GENERIC CRUD (getData / addRow / updateRow).
 * Source: user Apps Script v3 + patch.
 */

/**
 * Google Apps Script — backend SDD / Mill / TTP / Grievance
 *
 * PATCH v3 (coordinate round-trip fix — ROOT CAUSE FIXED):
 *
 *   ROOT CAUSE (v1 & v2 missed this):
 *   ws.appendRow(["0.318318"]) in Google Sheets ALWAYS coerces string "0.318318"
 *   back to a Number, regardless of any setNumberFormat('@') pre-applied to the column.
 *   This is a documented GAS appendRow() behaviour — it bypasses cell format.
 *   Under locale id_ID (dot = thousands separator), Sheets then DISPLAYS that number
 *   as "318.318" (thousands-formatted), which the JS recoverCoord() previously
 *   returned as-is (also fixed in main.js v3).
 *
 *   FIX: Replace EVERY ws.appendRow(row) that may contain coordinates with
 *   safeAppendRow_(ws, headers, row) — a helper that:
 *     1. Calls ws.appendRow(row) to create the new physical row.
 *     2. Immediately overwrites ONLY the coordinate columns on that new row with
 *        ws.getRange(newRowNum, col, 1, 1).setNumberFormat('@').setValue(stringVal).
 *        Because the cell now has Text format AND a string value, Sheets stores it
 *        as plain text permanently — no locale conversion possible.
 *
 *   Affected callers fixed in v3:
 *     appendRelRow_   → safeAppendRow_
 *     insertSddRow    → safeAppendRow_
 *     upsertSDD       → safeAppendRow_
 *     bulkUpsertSDD   → safeAppendRange_ (setNumberFormat then setValues for bulk)
 *
 *   Also fixed in v3:
 *     fixExistingCoordColumns — now ACTUALLY REWRITES integer values using the
 *     same integer-recovery heuristic as JS recoverCoord(), instead of just logging.
 *
 *   Run fixExistingCoordColumns() once from the GAS editor after deploy.
 */

// ═══════════════════════════════════════════════════════════
//  SHEET NAME CONFIG
// ═══════════════════════════════════════════════════════════

const SHEETS = {
  mill      : 'Mill Onboarding Profile',
  ttp       : 'Monitoring TTP/TTM',
  grievance : 'Grievance Monitoring',
  sdd       : 'SDD Data',
  sddMain   : 'SDD_MAIN',
  sddMill   : 'SDD_MILL_LIST',
  sddFfb    : 'SDD_FFB_LIST',
  contactSupplier : 'Contact List Supplier',
  nbl             : 'NBL',
  unileverNbl     : 'Unilever NBL',
  supplyDraft     : 'Supply Import Draft',
};

// ── Supply Import Draft headers ──────────────────────────────
const SUPPLY_DRAFT_HEADERS = [
  'draft_id', 'batch_id', 'status',        // 'draft' | 'submitted'
  'quarter', 'year',
  'created_at', 'updated_at', 'created_by',
  'match_status',                           // 'matched' | 'new'
  // Mirror of MILL_FIELDS subset that may be pre-filled or user-edited:
  'QUARTER', 'YEAR', 'COMPANY CODE', 'TRADER NAME', 'GROUP NAME',
  'COMPANY NAME', 'MILL NAME', 'UML ID', 'ADDRESS', 'PROVINCE',
  'COORDINATES', 'MILL CATEGORY', 'MILL CAPACITY (TON/HOUR)',
  'HGU/HGB', 'IZIN LOKASI', 'IUP', 'IZIN LINGKUNGAN', 'SCORE',
  'MILL LOC', 'COMPLIMENT/NOT COMPLIMENT',
  'DEFORESTATION SPATIAL', 'BURN AREA SPATIAL', 'PEAT', 'LEGALITY',
  'DEFORESTATION GRIEVANCES', 'BURN AREA GRIEVANCES',
  'HUMAN RIGHT', 'SAFETY', 'SOCIAL', 'ENVIRONMENT', 'TOTAL GRIEVANCES',
  'NDPE', 'HRDD', 'TOTAL POLICY',
  'CERTIFICATION', 'TOTAL CERTIFICATION', 'TOTAL SCORE',
  'SUPPLIER LEVEL', 'BUYER NO BUY LIST', 'VOLUME SUPPLY STATUS',
  'RECOMMENDATION LEVEL', 'SIGN', 'SUPPLIER STATUS', 'RISK LEVEL',
  'RESULT RISK LEVEL', 'FACILITY NAME CPO', 'FACILITY NAME PK',
  'PRODUCT SUPPLY', 'SUPPLY CPO',
];

const NBL_HEADERS = [
  'Riser',
  'Group Name NBL',
  'Company Name NBL',
  'SOURCE',
];

const UNILEVER_NBL_HEADERS = [
  'NO.',
  'UML ID',
  'COMPANY NAME',
  'MILL NAME',
  'COUNTRY',
  'PROVINCE',
  'DISTRICT / REGENCY',
  'LAT.',
  'LONG.',
];

const CONTACT_SUPPLIER_HEADERS = [
  'submission_id',
  'Group Name',
  'Company Name',
  'Supplier Type',
  'Sustainability PIC',
  'Phone Number',
  'statusSDD',
  'approved_at',
  'updated_at',
  'updated_by',
];

// ═══════════════════════════════════════════════════════════
//  COORDINATE COLUMN NAMES
// ═══════════════════════════════════════════════════════════

const COORD_COLUMN_NAMES = [
  'Latitude', 'Longitude',
  'TML - Latitude', 'TML - Longitude',
  'FFB - Latitude', 'FFB - Longitude',
];

const FORCE_STRING_FIELDS = new Set(COORD_COLUMN_NAMES);

// ═══════════════════════════════════════════════════════════
//  RELATIONAL HEADER DEFINITIONS
// ═══════════════════════════════════════════════════════════

const RELATIONAL_HEADERS = {
  sddMain: [
    'submission_id', 'supplier_type', 'Supplier Type',
    'created_at', 'updated_at', 'created_by', 'updated_by', 'is_deleted',
    'Mill ID', 'Date Imported', 'Imported By',
    'Group Name', 'Company Name', 'Current Owner', 'Previous Owner',
    'Take Over From', 'Office Address', 'Office Phone', 'Office Fax',
    'Office Email', 'Company Website',
    'Contact Person', 'Contact Position', 'Contact Mobile', 'Contact Email',
    'Sustainability PIC', 'Sustainability PIC Mobile',
    'Mill Name', 'Mill Address', 'Distance to Refinery (km)', 'Distance',
    'Latitude', 'Longitude', 'Mill Category',
    'Mill Capacity (Ton/Hour)', 'Sterilizer Type',
    'Storage Tank Capacity', 'Silo Capacity', 'Commissioning Date',
    'Main Product', 'Main Product Avg Production/Month (Ton)', 'Main Product Yield',
    'Main Product 2', 'Main Product 2 Avg Production/Month (Ton)', 'Main Product 2 Yield',
    'Main Product 3', 'Main Product 3 Avg Production/Month (Ton)', 'Main Product 3 Yield',
    'Other Product 1', 'Other Product 1 Avg/Month (Ton)',
    'Other Product 2', 'Other Product 2 Avg/Month (Ton)',
    'Other Product 3', 'Other Product 3 Avg/Month (Ton)',
    'Other Product 4', 'Other Product 4 Avg/Month (Ton)',
    'Other Product 5', 'Other Product 5 Avg/Month (Ton)',
    'Other Product 6', 'Other Product 6 Avg/Month (Ton)',
    'Other Product 7', 'Other Product 7 Avg/Month (Ton)',
    'Other Product 8', 'Other Product 8 Avg/Month (Ton)',
    'Other Product 9', 'Other Product 9 Avg/Month (Ton)',
    'Other Product 10', 'Other Product 10 Avg/Month (Ton)',
    'Other Product 11', 'Other Product 11 Avg/Month (Ton)',
    'Other Product 12', 'Other Product 12 Avg/Month (Ton)',
    'SDD - Product Lines JSON',
    'ISPO (Y/N)', 'ISPO Certificate No', 'ISPO Certificate Date',
    'RSPO (Y/N)', 'RSPO Certificate No', 'RSPO Certificate Date',
    'ISCC (Y/N)', 'ISCC Certificate No', 'ISCC Certificate Date',
    'HGU/HGB Type', 'HGU/HGB No', 'HGU/HGB Issue Date', 'HGU/HGB Expiry Date', 'HGU/HGB Area (Ha)',
    'IUP Type', 'IUP No', 'IUP Issue Date', 'IUP Expiry Date',
    'Izin Lokasi No', 'Izin Lokasi Issue Date', 'Izin Lokasi Expiry Date',
    'Izin Lingkungan No', 'Izin Lingkungan Issue Date',
    'NIB No',
    'CPO Quality - FFA', 'CPO Quality - Moisture', 'CPO Quality - Dirt',
    'GRV1 - Source', 'GRV1 - Description', 'GRV1 - Publisher', 'GRV1 - Date Publish', 'GRV1 - Status', 'GRV1 - Attachment',
    'GRV2 - Source', 'GRV2 - Description', 'GRV2 - Publisher', 'GRV2 - Date Publish', 'GRV2 - Status', 'GRV2 - Attachment',
    'GRV3 - Source', 'GRV3 - Description', 'GRV3 - Publisher', 'GRV3 - Date Publish', 'GRV3 - Status', 'GRV3 - Attachment',
    'GRV4 - Source', 'GRV4 - Description', 'GRV4 - Publisher', 'GRV4 - Date Publish', 'GRV4 - Status', 'GRV4 - Attachment',
    'GRV5 - Source', 'GRV5 - Description', 'GRV5 - Publisher', 'GRV5 - Date Publish', 'GRV5 - Status', 'GRV5 - Attachment',
    'GRV6 - Source', 'GRV6 - Description', 'GRV6 - Publisher', 'GRV6 - Date Publish', 'GRV6 - Status', 'GRV6 - Attachment',
    'GRV7 - Source', 'GRV7 - Description', 'GRV7 - Publisher', 'GRV7 - Date Publish', 'GRV7 - Status', 'GRV7 - Attachment',
    'GRV8 - Source', 'GRV8 - Description', 'GRV8 - Publisher', 'GRV8 - Date Publish', 'GRV8 - Status', 'GRV8 - Attachment',
    'GRV9 - Source', 'GRV9 - Description', 'GRV9 - Publisher', 'GRV9 - Date Publish', 'GRV9 - Status', 'GRV9 - Attachment',
    'GRV10 - Source', 'GRV10 - Description', 'GRV10 - Publisher', 'GRV10 - Date Publish', 'GRV10 - Status', 'GRV10 - Attachment',
    'No of Open Grievance', 'Grievance Detail',
    'PRI1 - Company', 'PRI1 - Description', 'PRI1 - Publisher', 'PRI1 - Date Publish', 'PRI1 - Attachment', 'PRI1 - Action Request',
    'PRI2 - Company', 'PRI2 - Description', 'PRI2 - Publisher', 'PRI2 - Date Publish', 'PRI2 - Attachment', 'PRI2 - Action Request',
    'PRI3 - Company', 'PRI3 - Description', 'PRI3 - Publisher', 'PRI3 - Date Publish', 'PRI3 - Attachment', 'PRI3 - Action Request',
    'PRI4 - Company', 'PRI4 - Description', 'PRI4 - Publisher', 'PRI4 - Date Publish', 'PRI4 - Attachment', 'PRI4 - Action Request',
    'PRI5 - Company', 'PRI5 - Description', 'PRI5 - Publisher', 'PRI5 - Date Publish', 'PRI5 - Attachment', 'PRI5 - Action Request',
    'PRI6 - Company', 'PRI6 - Description', 'PRI6 - Publisher', 'PRI6 - Date Publish', 'PRI6 - Attachment', 'PRI6 - Action Request',
    'PRI7 - Company', 'PRI7 - Description', 'PRI7 - Publisher', 'PRI7 - Date Publish', 'PRI7 - Attachment', 'PRI7 - Action Request',
    'PRI8 - Company', 'PRI8 - Description', 'PRI8 - Publisher', 'PRI8 - Date Publish', 'PRI8 - Attachment', 'PRI8 - Action Request',
    'PRI9 - Company', 'PRI9 - Description', 'PRI9 - Publisher', 'PRI9 - Date Publish', 'PRI9 - Attachment', 'PRI9 - Action Request',
    'PRI10 - Company', 'PRI10 - Description', 'PRI10 - Publisher', 'PRI10 - Date Publish', 'PRI10 - Attachment', 'PRI10 - Action Request',
    'PRI Status', 'PRI Date', 'PRI Notes',
    'statusSDD',
    'noteBossDecision',
    'statusBossDecision',
    'SCR - Screening Status', 'SCR - Screening Date', 'SCR - Screened By',
    'SCR - Last Updated', 'SCR - Notes', 'SCR - Requested Data', 'SCR - Recommendation',
    'SCR - Overall Risk Level',
    'SCR - Forest Area (Ha)', 'SCR - Peatland (Ha)',
    'SCR - Moratorium', 'SCR - Moratorium (Ha)',
    'SCR - Deforestation Buffer 50KM (Ha)',
  ],
  sddMill: [
    'submission_id', 'line_id', 'supplier_type',
    'created_at', 'updated_at', 'created_by', 'updated_by', 'is_deleted',
    'TML - Company Name', 'TML - Mill Name', 'TML - UML ID',
    'TML - Village', 'TML - Sub District', 'TML - District',
    'TML - Capacity (Ton/Hour)',
    'TML - Latitude', 'TML - Longitude', 'TML - Legality',
    'TML - ISPO (Y/N)', 'TML - RSPO (Y/N)', 'TML - ISCC (Y/N)',
    'TML - Total Supply CPO/PK (Ton)',
    'SCR - TML Valid Coordinate',
    'SCR - TML Forest Area', 'SCR - TML Peatland',
    'SCR - TML Moratorium', 'SCR - TML Moratorium (Ha)',
    'SCR - TML Deforestation Buffer 50KM (Ha)',
    'SCR - TML Screening Status', 'SCR - TML Screening Date',
  ],
  sddFfb: [
    'submission_id', 'line_id', 'supplier_type',
    'created_at', 'updated_at', 'created_by', 'updated_by', 'is_deleted',
    'FFB - Mill Name', 'FFB - Supplier Group Name', 'FFB - Supplier Name',
    'FFB - Village', 'FFB - Sub District', 'FFB - District',
    'FFB - Supplier Category',
    'FFB - Concession Area (Ha)', 'FFB - Planted Area (Ha)',
    'FFB - Number of Smallholders', 'FFB - Planted Year',
    'FFB - Legality',
    'FFB - Latitude', 'FFB - Longitude',
    'FFB - Distance to Mill (Km)',
    'FFB - Valid Coordinate', 'FFB - Forest Area', 'FFB - Peatland',
    'FFB - Moratorium', 'FFB - Moratorium (Ha)',
    'FFB - Deforestation (Ha)', 'FFB - Burn Area (Ha)', 'FFB - Village Risk',
    'FFB - Screening Status', 'FFB - Screening Date',
    'FFB - ISPO (Y/N)', 'FFB - RSPO (Y/N)', 'FFB - ISCC (Y/N)',
    'FFB - Total Supply FFB (Ton)',
  ],
};

// ═══════════════════════════════════════════════════════════
//  CANONICAL ALIAS MAP
// ═══════════════════════════════════════════════════════════

const CANONICAL_ALIASES = {
  'Mil ID':                       'Mill ID',
  'MILL ID':                      'Mill ID',
  'Grup Name':                    'Group Name',
  'HGU / HGB Type':               'HGU/HGB Type',
  'HGU HGB Type':                 'HGU/HGB Type',
  'HGU/HGB':                      'HGU/HGB Type',
  'HGU / HGB No':                 'HGU/HGB No',
  'HGU HGB No':                   'HGU/HGB No',
  'HGU/HGB Number':               'HGU/HGB No',
  'HGU / HGB Number':             'HGU/HGB No',
  'HGU / HGB Issue Date':         'HGU/HGB Issue Date',
  'HGU HGB Issue Date':           'HGU/HGB Issue Date',
  'IUP / IUP - P Type':           'IUP Type',
  'IUP / IUP - P No':             'IUP No',
  'IUP Number':                   'IUP No',
  'IUP / IUP - P Issue Date':     'IUP Issue Date',
  'Izin Lokasi Number':           'Izin Lokasi No',
  'IZIN LOKASI No':               'Izin Lokasi No',
  'IZIN LOKASI':                  'Izin Lokasi No',
  'IZIN LOKASI Issue Date':       'Izin Lokasi Issue Date',
  'Izin Lingkungan Number':       'Izin Lingkungan No',
  'IZIN LINGKUNGAN No':           'Izin Lingkungan No',
  'IZIN LINGKUNGAN':              'Izin Lingkungan No',
  'IZIN LINGKUNGAN Issue Date':   'Izin Lingkungan Issue Date',
  'NIB Number':                   'NIB No',
};

const TTP_HEADERS = [
  'NO', 'COMPANY CODE', 'GROUP NAME', 'COMPANY NAME', 'MILL NAME', 'UML ID',
  'FFB SUPPLIER GROUP NAME', 'FFB SUPPLIER NAME', 'CATEGORY', 'LAT', 'LONG',
  'VILLAGE ID', 'VILLAGE', 'SUBDISTRICT', 'DISTRICT', 'PROVINCE',
  'CONCESION AREA', 'PLANTED AREA', 'NUMBER OD SMALLHOLDERS', 'TAHUN TANAM',
  'LEGALITAS', 'ISPO (Y/N)', 'RSPO (Y/N)', 'ISCC (Y/N)',
  'FFB SUPPLY to MILL (TON)', 'CONVERSION FFB to PK (5%)', 'PK SUPPLY to KCP',
  'CONVERSION FFB to CPO (20%)', 'CPO SUPPLY to REFINERY',
  '% PK TRACEABLE', '% CPO TRACEABLE',
  'Total PK % Traceable', 'Total CPO % Traceable',
  'MSD', 'PK Traceable Volume', 'CPO Traceable Volume',
  'submission_id', 'ffb_line_id', 'supplier_type', 'synced_at', 'synced_by',
];

/** TTP columns filled by monitoring team — not overwritten on SDD re-sync when already set. */
const TTP_MONITORING_PRESERVE_KEYS = [
  'NO', 'COMPANY CODE', 'PROVINCE', 'VILLAGE ID',
  '% PK TRACEABLE', '% CPO TRACEABLE',
  'Total PK % Traceable', 'Total CPO % Traceable',
  'MSD', 'PK Traceable Volume', 'CPO Traceable Volume',
  'CONVERSION FFB to PK (5%)', 'PK SUPPLY to KCP',
  'CONVERSION FFB to CPO (20%)', 'CPO SUPPLY to REFINERY',
];

// ═══════════════════════════════════════════════════════════
//  ENTRY POINTS
// ═══════════════════════════════════════════════════════════

function doGet(e) {
  try {
    const action   = (e && e.parameter && e.parameter.action)  || '';
    const sheetKey = (e && e.parameter && e.parameter.sheet)   || '';

    if (action === 'getAll') {
      if (sheetKey === 'ttp') ensureTtpHeaders_();
      if (sheetKey === 'contactSupplier') ensureContactSupplierHeaders_();
      if (sheetKey === 'nbl') ensureNblHeaders_();
      if (sheetKey === 'unileverNbl') ensureUnileverNblHeaders_();
      if (sheetKey === 'supplyDraft') ensureSupplyDraftHeaders_();
      return respond(getData(sheetKey));
    }
    if (action === 'getByMillId')       return respond(getByMillId(e.parameter.millId));
    if (action === 'ping')              return respond({ success: true, message: 'Apps Script is alive' });
    if (action === 'getSubmissionById') return respond(getSubmissionById(e.parameter.submission_id));
    if (action === 'listSubmissions')   return respond(listSubmissions(e.parameter));

    return respond({ success: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return respond({ success: false, error: err.message });
  }
}

function doPost(e) {
  try {
    const body     = parsePostBody_(e);
    const action   = body.action   || '';
    const sheetKey = body.sheet    || '';

    if (action === 'add')    return respond(addRow(sheetKey, body.data || {}));
    if (action === 'update') return respond(updateRow(sheetKey, body.row, body.data || {}));
    if (action === 'delete') {
      var multi = body.rows;
      if (Array.isArray(multi) && multi.length) return respond(bulkDelete(sheetKey, multi));
      return respond(deleteRow(sheetKey, body.row));
    }
    if (action === 'bulkDelete')    return respond(bulkDelete(sheetKey, body.rows || []));
    if (action === 'bulkUpsertSDD') return respond(bulkUpsertSDD(body.rows || []));
    if (action === 'insertSDD') return respond(legacyAdapterInsert_(body.data || {}));
    if (action === 'updateSDD') return respond(legacyAdapterUpdate_(body.row, body.data || {}));
    if (action === 'upsertSDD') return respond(legacyAdapterUpsert_(body.data || {}));
    if (action === 'createSubmission')    return respond(createSubmission(body.payload    || {}));
    if (action === 'updateSubmission')    return respond(updateSubmission(body.payload    || {}));
    if (action === 'setSubmissionStatus') return respond(setSubmissionStatus(body.payload || {}));
    if (action === 'deleteSubmission')    return respond(deleteSubmission(body.payload    || {}));
    if (action === 'saveSupplyDraft')   return respond(saveSupplyDraft_(body.rows || [], body.batch_id || '', body.meta || {}));
    if (action === 'submitSupplyDraft') return respond(submitSupplyDraft_(body.batch_id || '', body.rows || []));
    if (action === 'deleteSupplyDraft') return respond(deleteSupplyDraft_(body.draft_id || '', body.batch_id || ''));

    return respond({ success: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return respond({ success: false, error: err.message });
  }
}

function doOptions() {
  return respond({ success: true });
}

// ═══════════════════════════════════════════════════════════
//  COORDINATE COLUMN FORMAT GUARD
// ═══════════════════════════════════════════════════════════

/**
 * Set coordinate columns to Plain Text format ('@') so Google Sheets
 * never auto-converts values under any locale.
 * Idempotent — safe to call on every header init.
 */
function setCoordColumnsAsText_(ws, headers) {
  if (!ws || !Array.isArray(headers)) return;
  const lastRow = Math.max(ws.getLastRow(), 2);
  COORD_COLUMN_NAMES.forEach(function(name) {
    const colIdx = headers.indexOf(name);
    if (colIdx < 0) return;
    const colNum = colIdx + 1;
    ws.getRange(2, colNum, lastRow - 1, 1).setNumberFormat('@');
  });
}

// ═══════════════════════════════════════════════════════════
//  COORDINATE STRING ENFORCER
// ═══════════════════════════════════════════════════════════

/**
 * forceCoordStrings_ — convert coordinate values in a row array to plain
 * dot-decimal strings immediately before any write call.
 */
function forceCoordStrings_(headers, rowArr) {
  COORD_COLUMN_NAMES.forEach(function(name) {
    var colIdx = headers.indexOf(name);
    if (colIdx < 0) return;
    var v = rowArr[colIdx];
    if (v === null || v === undefined) { rowArr[colIdx] = ''; return; }
    var s = String(v).trim();
    if (s.indexOf(',') !== -1 && s.indexOf('.') === -1) s = s.replace(',', '.');
    rowArr[colIdx] = s;
  });
}

// ═══════════════════════════════════════════════════════════
//  PATCH v3 — SAFE APPEND HELPERS
//
//  ROOT CAUSE: ws.appendRow(["0.318318"]) ignores any pre-set '@' (Text)
//  number format and coerces the string to a Number. Under locale id_ID
//  (dot = thousands separator) Sheets then stores INTEGER 318318.
//
//  FIX: After appendRow(), immediately overwrite coordinate columns with
//  setNumberFormat('@') + setValue(stringVal) on each affected cell.
//  setNumberFormat('@') on an existing cell + setValue(string) stores text.
// ═══════════════════════════════════════════════════════════

/**
 * safeAppendRow_ — append a single row, then patch coordinate columns as Text.
 *
 * @param {Sheet}    ws       Target worksheet.
 * @param {string[]} headers  Header row array.
 * @param {Array}    rowArr   Data row (already processed by forceCoordStrings_).
 */
function safeAppendRow_(ws, headers, rowArr) {
  ws.appendRow(rowArr);
  // appendRow() ignores column number format — re-write coord cells as explicit Text.
  const newRowNum = ws.getLastRow();
  COORD_COLUMN_NAMES.forEach(function(name) {
    const colIdx = headers.indexOf(name);
    if (colIdx < 0) return;
    const v = rowArr[colIdx];
    if (v === null || v === undefined) return;
    const s = String(v).trim();
    if (!s) return;
    const cell = ws.getRange(newRowNum, colIdx + 1);
    cell.setNumberFormat('@');
    cell.setValue(s);
  });
}

/**
 * safeAppendRange_ — write multiple new rows via setValues, then patch
 * coordinate columns as Text on all new rows.
 *
 * @param {Sheet}    ws         Target worksheet.
 * @param {string[]} headers    Header row array.
 * @param {Array[]}  rowArrays  Array of data rows.
 * @param {number}   startRow   1-based sheet row where the first new row goes.
 */
function safeAppendRange_(ws, headers, rowArrays, startRow) {
  if (!rowArrays.length) return;
  ws.getRange(startRow, 1, rowArrays.length, headers.length).setValues(rowArrays);
  // Re-apply Text format + re-write coord cells on every newly inserted row.
  COORD_COLUMN_NAMES.forEach(function(name) {
    const colIdx = headers.indexOf(name);
    if (colIdx < 0) return;
    const colNum = colIdx + 1;
    // Build a column of string values for this coord column across all new rows.
    const coordValues = rowArrays.map(function(row) {
      const v = row[colIdx];
      if (v === null || v === undefined) return [''];
      return [String(v).trim()];
    });
    const coordRange = ws.getRange(startRow, colNum, rowArrays.length, 1);
    coordRange.setNumberFormat('@');
    coordRange.setValues(coordValues);
  });
}

// ═══════════════════════════════════════════════════════════
//  ONE-TIME MIGRATION — run once from GAS editor after deploy
// ═══════════════════════════════════════════════════════════

/**
 * recoverCoordValue_ — GAS-side equivalent of JS recoverCoord().
 * Converts corrupted integer / thousands-formatted string back to dot-decimal.
 *
 * @param {*}      v         Raw cell value (number, string, etc.)
 * @param {string} colName   Column header name (used to detect lat vs lng).
 * @returns {string}  Dot-decimal coordinate string, or original string if unrecoverable.
 */
function recoverCoordValue_(v, colName) {
  if (v === null || v === undefined) return '';
  var s = String(v).trim();
  if (!s) return '';

  // Comma-decimal → dot
  if (s.indexOf(',') !== -1 && s.indexOf('.') === -1) return s.replace(',', '.');

  // Has dot — check WGS-84 range
  if (s.indexOf('.') !== -1) {
    var dotParsed = parseFloat(s);
    if (!isNaN(dotParsed) && Math.abs(dotParsed) <= 180) return s; // valid decimal
    // Out-of-range: strip dots (id_ID thousands seps) and recover as integer
    var stripped = s.replace(/\./g, '');
    if (/^-?\d+$/.test(stripped)) return recoverCoordValue_(stripped, colName);
    return s;
  }

  // Pure integer
  if (!/^-?\d+$/.test(s)) return s;
  var absN = Math.abs(parseInt(s, 10));
  if (absN <= 180) return s; // valid small integer coordinate

  var isLng  = /long|lng/i.test(colName);
  var maxInt = isLng ? 180 : 90;
  var neg    = s.charAt(0) === '-';
  var pfx    = neg ? '-' : '';
  var digits = String(absN);
  var candidates = [];

  for (var k = 1; k <= digits.length; k++) {
    var pow      = Math.pow(10, k);
    var leftInt  = Math.floor(absN / pow);
    if (leftInt > maxInt) continue;
    var rightStr = String(absN % pow);
    while (rightStr.length < k) rightStr = '0' + rightStr;
    candidates.push({ leftInt: leftInt, k: k, valStr: pfx + leftInt + '.' + rightStr });
  }

  if (!candidates.length) return s;

  var best;
  if (isLng) {
    best = candidates.reduce(function(a, b) { return b.leftInt > a.leftInt ? b : a; });
  } else {
    best = candidates.reduce(function(a, b) { return b.k > a.k ? b : a; });
  }
  return best.valStr;
}

/**
 * fixExistingCoordColumns — run manually ONCE from GAS editor after deploy.
 *
 * For every coordinate column in SDD_MAIN, SDD_MILL_LIST, SDD_FFB_LIST:
 *   1. Sets column number format to '@' (Text).
 *   2. Reads all existing values.
 *   3. Recovers any corrupted integer / thousands-formatted values using
 *      recoverCoordValue_() (same heuristic as JS recoverCoord()).
 *   4. Re-writes the column as plain text strings.
 *
 * This permanently heals rows corrupted by previous appendRow() calls.
 */
function fixExistingCoordColumns() {
  var totalFixed = 0;
  ['sddMain', 'sddMill', 'sddFfb'].forEach(function(key) {
    var ws      = getSheet(key);
    var lastCol = ws.getLastColumn();
    if (lastCol < 1) return;
    var headers = ws.getRange(1, 1, 1, lastCol).getValues()[0]
                    .map(function(h) { return String(h || '').trim(); });
    var lastRow = ws.getLastRow();
    if (lastRow < 2) {
      setCoordColumnsAsText_(ws, headers);
      return;
    }

    COORD_COLUMN_NAMES.forEach(function(name) {
      var colIdx = headers.indexOf(name);
      if (colIdx < 0) return;
      var colNum = colIdx + 1;

      // Set Text format on entire column (rows 2 onward)
      ws.getRange(2, colNum, lastRow - 1, 1).setNumberFormat('@');

      // Read, recover, and rewrite
      var range  = ws.getRange(2, colNum, lastRow - 1, 1);
      var vals   = range.getValues();
      var fixed  = vals.map(function(row, i) {
        var v   = row[0];
        var raw = String(v === null || v === undefined ? '' : v).trim();
        var recovered = recoverCoordValue_(v, name);
        if (recovered !== raw) {
          Logger.log('FIXED sheet=' + key + ' col=' + name +
                     ' row=' + (i + 2) + ' from=' + raw + ' to=' + recovered);
          totalFixed++;
        }
        return [recovered];
      });
      range.setValues(fixed);
    });
  });
  Logger.log('fixExistingCoordColumns complete. Total cells recovered: ' + totalFixed);
}

// ═══════════════════════════════════════════════════════════
//  RELATIONAL HEADER AUTO-INIT
// ═══════════════════════════════════════════════════════════

function ensureRelationalHeaders_(sheetKey) {
  const ws       = getSheet(sheetKey);
  const wantList = RELATIONAL_HEADERS[sheetKey];
  if (!wantList || !wantList.length) return;

  const lastCol = ws.getLastColumn();

  if (lastCol === 0 || ws.getLastRow() === 0) {
    ws.getRange(1, 1, 1, wantList.length).setValues([wantList]);
    setCoordColumnsAsText_(ws, wantList);
    return;
  }

  const existing    = ws.getRange(1, 1, 1, lastCol).getValues()[0]
                        .map(function(h) { return String(h || '').trim(); });
  const existingSet = new Set(existing.filter(Boolean));

  const missing = wantList.filter(function(h) { return !existingSet.has(h); });
  if (missing.length) {
    const startCol = existing.filter(Boolean).length + 1;
    ws.insertColumnsAfter(startCol - 1, missing.length);
    ws.getRange(1, startCol, 1, missing.length).setValues([missing]);
  }

  setCoordColumnsAsText_(ws, existing);
}

function ensureAllRelationalHeaders_() {
  ensureRelationalHeaders_('sddMain');
  ensureRelationalHeaders_('sddMill');
  ensureRelationalHeaders_('sddFfb');
}

function ensureSheetHeadersGeneric_(sheetKey, wantHeaders) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const name = SHEETS[sheetKey];
  if (!name) throw new Error('Sheet key not found: ' + sheetKey);
  let ws     = ss.getSheetByName(name);
  if (!ws) ws = ss.insertSheet(name);

  const lastCol = ws.getLastColumn();
  if (lastCol === 0 || ws.getLastRow() === 0) {
    ws.getRange(1, 1, 1, wantHeaders.length).setValues([wantHeaders]);
    return ws;
  }

  const existing = ws.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function(h) { return String(h || '').trim(); });
  const existingSet = new Set(existing.filter(Boolean));
  const missing = wantHeaders.filter(function(h) { return !existingSet.has(h); });
  if (missing.length) {
    const startCol = existing.filter(Boolean).length + 1;
    ws.insertColumnsAfter(Math.max(startCol - 1, 0), missing.length);
    ws.getRange(1, startCol, 1, missing.length).setValues([missing]);
  }
  return ws;
}

function ensureContactSupplierHeaders_() {
  return ensureSheetHeadersGeneric_('contactSupplier', CONTACT_SUPPLIER_HEADERS);
}

function ensureNblHeaders_() {
  return ensureSheetHeadersGeneric_('nbl', NBL_HEADERS);
}

function ensureUnileverNblHeaders_() {
  return ensureSheetHeadersGeneric_('unileverNbl', UNILEVER_NBL_HEADERS);
}

function normalizeSddDecisionLabel_(raw) {
  const s = String(raw || '').trim().toLowerCase();
  if (s === 'approve' || s === 'approved') return 'APPROVED';
  if (s === 'hold' || s === 'on hold') return 'ON HOLD';
  if (s === 'reject' || s === 'rejected') return 'REJECTED';
  return String(raw || '').trim().toUpperCase();
}

function readContactSupplierRows_() {
  const ws = ensureContactSupplierHeaders_();
  const range = ws.getDataRange();
  const values = range.getValues();
  if (!values.length) return { ws: ws, headers: CONTACT_SUPPLIER_HEADERS.slice(), rows: [] };

  const headers = values[0].map(function(h) { return String(h || '').trim(); });
  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const empty = !row.some(function(c) { return String(c || '').trim() !== ''; });
    if (empty) continue;
    const obj = { _row: r + 1 };
    headers.forEach(function(h, i) {
      if (h) obj[h] = row[i];
    });
    rows.push(obj);
  }
  return { ws: ws, headers: headers, rows: rows };
}

function findContactRowBySubmissionId_(sid) {
  const result = readContactSupplierRows_();
  for (let i = 0; i < result.rows.length; i++) {
    if (String(result.rows[i]['submission_id'] || '').trim() === sid) {
      return { ws: result.ws, headers: result.headers, row: result.rows[i], _sheetRow: result.rows[i]._row };
    }
  }
  return null;
}

function appendContactSupplierRow_(ws, headers, obj) {
  const row = headers.map(function(h) {
    return obj[h] !== undefined && obj[h] !== null ? obj[h] : '';
  });
  ws.appendRow(row);
}

function patchContactSupplierRow_(ws, headers, sheetRow, patch) {
  const headerIndex = {};
  headers.forEach(function(h, i) { if (h) headerIndex[h] = i; });
  Object.keys(patch).forEach(function(key) {
    const col = headerIndex[key];
    if (col === undefined) return;
    ws.getRange(sheetRow, col + 1).setValue(patch[key]);
  });
}

/**
 * When SDD decision is APPROVED, copy Sustainability PIC + Phone Number into Contact List Supplier.
 * Upserts by submission_id (re-approve updates the same row).
 */
function syncContactFromApprovedSubmission_(sid, mainObj, user, now) {
  sid = String(sid || '').trim();
  if (!sid) return { synced: false, skipped: true, reason: 'missing_submission_id' };

  const merged = mainObj || {};
  const decision = normalizeSddDecisionLabel_(
    merged['statusSDD'] || merged['statusBossDecision'] || ''
  );
  if (decision !== 'APPROVED') {
    return { synced: false, skipped: true, reason: 'not_approved', decision: decision };
  }

  const pic = String(merged['Sustainability PIC'] || '').trim();
  const phone = String(merged['Sustainability PIC Mobile'] || merged['Phone Number'] || '').trim();
  if (!pic && !phone) {
    return { synced: false, skipped: true, reason: 'empty_contact_fields' };
  }

  const wsResult = readContactSupplierRows_();
  const ws = wsResult.ws;
  const headers = wsResult.headers;
  const patch = {
    'submission_id': sid,
    'Group Name': String(merged['Group Name'] || '').trim(),
    'Company Name': String(merged['Company Name'] || '').trim(),
    'Supplier Type': String(merged['supplier_type'] || merged['Supplier Type'] || '').trim(),
    'Sustainability PIC': pic,
    'Phone Number': phone,
    'statusSDD': 'APPROVED',
    'approved_at': now,
    'updated_at': now,
    'updated_by': user,
  };

  const hit = findContactRowBySubmissionId_(sid);
  if (hit) {
    patchContactSupplierRow_(ws, headers, hit._sheetRow, patch);
    return { synced: true, updated: true, submission_id: sid };
  }

  appendContactSupplierRow_(ws, headers, patch);
  return { synced: true, inserted: true, submission_id: sid };
}

// ═══════════════════════════════════════════════════════════
//  TTP / TTM  ─ SYNC FROM APPROVED SDD (FFB rows)
// ═══════════════════════════════════════════════════════════

function isBlankTtpCell_(v) {
  const s = String(v === undefined || v === null ? '' : v).trim();
  if (!s || s === '—' || s === '-') return true;
  if (/^please select$/i.test(s)) return true;
  if (/^silahkan pilih$/i.test(s)) return true;
  return false;
}

function isMeaningfulSddFfbRow_(ffb) {
  if (!ffb || typeof ffb !== 'object') return false;
  if (String(ffb['is_deleted'] || '') === '1') return false;
  const name = String(ffb['FFB - Supplier Name'] || '').trim();
  const grp  = String(ffb['FFB - Supplier Group Name'] || '').trim();
  return !!(name || grp);
}

function readTtpRows_() {
  ensureTtpHeaders_();
  const ws = getSheet('ttp');
  const range = ws.getDataRange();
  const values = range.getValues();
  if (!values.length) {
    return { ws: ws, headers: TTP_HEADERS.slice(), rows: [] };
  }

  const headers = values[0].map(function(h) { return String(h || '').trim(); });
  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const empty = !row.some(function(c) { return String(c || '').trim() !== ''; });
    if (empty) continue;
    const obj = { _row: r + 1 };
    headers.forEach(function(h, i) {
      if (h) obj[h] = row[i];
    });
    rows.push(obj);
  }
  return { ws: ws, headers: headers, rows: rows };
}

function findTtpRowBySyncKeys_(rows, sid, lineId, patch) {
  sid = String(sid || '').trim();
  lineId = String(lineId || '').trim();
  const supName = String((patch && patch['FFB SUPPLIER NAME']) || '').trim().toLowerCase();
  const supGrp  = String((patch && patch['FFB SUPPLIER GROUP NAME']) || '').trim().toLowerCase();

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (String(row['submission_id'] || '').trim() !== sid) continue;
    if (lineId && String(row['ffb_line_id'] || '').trim() === lineId) {
      return { row: row, _sheetRow: row._row };
    }
  }
  if (supName || supGrp) {
    for (let j = 0; j < rows.length; j++) {
      const row2 = rows[j];
      if (String(row2['submission_id'] || '').trim() !== sid) continue;
      const n = String(row2['FFB SUPPLIER NAME'] || '').trim().toLowerCase();
      const g = String(row2['FFB SUPPLIER GROUP NAME'] || '').trim().toLowerCase();
      if (supName && n === supName && (!supGrp || g === supGrp)) {
        return { row: row2, _sheetRow: row2._row };
      }
    }
  }
  return null;
}

function appendTtpRow_(ws, headers, obj) {
  const row = headers.map(function(h) {
    return obj[h] !== undefined && obj[h] !== null ? obj[h] : '';
  });
  ws.appendRow(row);
}

function patchTtpRow_(ws, headers, sheetRow, patch) {
  const headerIndex = {};
  headers.forEach(function(h, i) { if (h) headerIndex[h] = i; });
  Object.keys(patch).forEach(function(key) {
    const col = headerIndex[key];
    if (col === undefined) return;
    ws.getRange(sheetRow, col + 1).setValue(patch[key]);
  });
}

function mergeTtpPreserveMonitoring_(existingRow, patch) {
  const out = Object.assign({}, patch);
  TTP_MONITORING_PRESERVE_KEYS.forEach(function(key) {
    if (existingRow && !isBlankTtpCell_(existingRow[key])) {
      out[key] = existingRow[key];
    }
  });
  return out;
}

function buildTtpPatchFromSddFfb_(main, ffb, sid, user, now) {
  const supplierType = String(main['supplier_type'] || main['Supplier Type'] || '').trim();
  let millName = String(main['Mill Name'] || '').trim();
  if (!millName && /trader/i.test(supplierType)) {
    millName = String(main['Company Name'] || '').trim();
  }

  let lat = String(ffb['FFB - Latitude'] || '').trim();
  let lng = String(ffb['FFB - Longitude'] || '').trim();
  if (isBlankTtpCell_(lat)) lat = String(main['Latitude'] || '').trim();
  if (isBlankTtpCell_(lng)) lng = String(main['Longitude'] || '').trim();

  return {
    'GROUP NAME'              : String(main['Group Name'] || '').trim(),
    'COMPANY NAME'            : String(main['Company Name'] || '').trim(),
    'MILL NAME'               : millName,
    'FFB SUPPLIER GROUP NAME' : String(ffb['FFB - Supplier Group Name'] || '').trim(),
    'FFB SUPPLIER NAME'       : String(ffb['FFB - Supplier Name'] || '').trim(),
    'CATEGORY'                : String(ffb['FFB - Supplier Category'] || '').trim(),
    'LAT'                     : lat,
    'LONG'                    : lng,
    'VILLAGE'                 : String(ffb['FFB - Village'] || '').trim(),
    'SUBDISTRICT'             : String(ffb['FFB - Sub District'] || '').trim(),
    'DISTRICT'                : String(ffb['FFB - District'] || '').trim(),
    'CONCESION AREA'          : String(ffb['FFB - Concession Area (Ha)'] || '').trim(),
    'PLANTED AREA'            : String(ffb['FFB - Planted Area (Ha)'] || '').trim(),
    'NUMBER OD SMALLHOLDERS'  : String(ffb['FFB - Number of Smallholders'] || '').trim(),
    'TAHUN TANAM'             : String(ffb['FFB - Planted Year'] || '').trim(),
    'LEGALITAS'               : String(ffb['FFB - Legality'] || '').trim(),
    'ISPO (Y/N)'              : String(ffb['FFB - ISPO (Y/N)'] || '').trim(),
    'RSPO (Y/N)'              : String(ffb['FFB - RSPO (Y/N)'] || '').trim(),
    'ISCC (Y/N)'              : String(ffb['FFB - ISCC (Y/N)'] || '').trim(),
    'FFB SUPPLY to MILL (TON)' : String(ffb['FFB - Total Supply FFB (Ton)'] || '').trim(),
    'submission_id'           : sid,
    'ffb_line_id'             : String(ffb['line_id'] || '').trim(),
    'supplier_type'           : supplierType,
    'synced_at'               : now,
    'synced_by'               : user,
  };
}

/**
 * When SDD decision is APPROVED, upsert Monitoring TTP/TTM rows (one per FFB child row).
 * Does not overwrite monitoring-only columns when already filled.
 */
function syncTtpFromApprovedSubmission_(sid, mainObj, user, now) {
  sid = String(sid || '').trim();
  if (!sid) return { synced: false, skipped: true, reason: 'missing_submission_id' };

  const merged = mainObj || {};
  const decision = normalizeSddDecisionLabel_(
    merged['statusSDD'] || merged['statusBossDecision'] || ''
  );
  if (decision !== 'APPROVED') {
    return { synced: false, skipped: true, reason: 'not_approved', decision: decision };
  }

  const ffbRows = findChildRows_('sddFfb', sid)
    .map(function(r) { return r.obj; })
    .filter(isMeaningfulSddFfbRow_);

  if (!ffbRows.length) {
    return { synced: false, skipped: true, reason: 'no_ffb_rows', submission_id: sid };
  }

  const ttpResult = readTtpRows_();
  let inserted = 0;
  let updated = 0;
  const errors = [];

  ffbRows.forEach(function(ffb) {
    try {
      const lineId = String(ffb['line_id'] || '').trim();
      const patch = buildTtpPatchFromSddFfb_(merged, ffb, sid, user, now);
      const hit = findTtpRowBySyncKeys_(ttpResult.rows, sid, lineId, patch);
      if (hit) {
        const mergedPatch = mergeTtpPreserveMonitoring_(hit.row, patch);
        patchTtpRow_(ttpResult.ws, ttpResult.headers, hit._sheetRow, mergedPatch);
        Object.assign(hit.row, mergedPatch);
        updated++;
      } else {
        appendTtpRow_(ttpResult.ws, ttpResult.headers, patch);
        ttpResult.rows.push(Object.assign({ _row: ttpResult.ws.getLastRow() }, patch));
        inserted++;
      }
    } catch (err) {
      errors.push({
        line_id: String(ffb['line_id'] || ''),
        error: String(err && err.message ? err.message : err),
      });
    }
  });

  return {
    synced: true,
    submission_id: sid,
    inserted: inserted,
    updated: updated,
    total_ffb: ffbRows.length,
    errors: errors.length ? errors : undefined,
  };
}

// ═══════════════════════════════════════════════════════════
//  TTP / TTM  ─ AUTO-INIT HEADERS
// ═══════════════════════════════════════════════════════════

function ensureTtpHeaders_() {
  const sheet   = getSheet('ttp');
  const lastCol = sheet.getLastColumn();

  if (lastCol === 0) {
    sheet.getRange(1, 1, 1, TTP_HEADERS.length).setValues([TTP_HEADERS]);
    return sheet;
  }

  const existing    = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
                           .map(function(h) { return String(h || '').trim(); });
  const existingSet = new Set(existing.filter(Boolean));
  const missing     = TTP_HEADERS.filter(function(h) { return !existingSet.has(h); });

  if (missing.length) {
    const start = existing.length + 1;
    sheet.insertColumnsAfter(existing.length, missing.length);
    sheet.getRange(1, start, 1, missing.length).setValues([missing]);
  }
  return sheet;
}

// ═══════════════════════════════════════════════════════════
//  SDD UPSERT  (single row)
// ═══════════════════════════════════════════════════════════

function upsertSDD(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    sanitizeSddPayload_(data);
    const sheet = getSheet('sdd');
    ensureOtherProductExtensionHeadersOrdered_(sheet);
    ensureMainProductExtensionHeadersOrdered_(sheet);
    ensureSddHeadersForPayloads_(sheet, [data || {}]);

    const rows        = sheet.getDataRange().getValues();
    if (!rows.length) throw new Error('SDD sheet is empty (no headers)');

    const headers     = rows[0];
    const headerIndex = indexByHeader_(headers);

    normalizeSubmittedFlags_(data);
    const matchIdx = findMatchingSddRowIndex_(rows, headers, headerIndex, data);

    if (matchIdx > 0) {
      const existingObj = rowToObject_(headers, rows[matchIdx]);
      assertSubmittedNotDowngraded_(existingObj, data);
      const updatedRow = headers.map(function(h, j) {
        return data[h] !== undefined ? data[h] : rows[matchIdx][j];
      });
      forceCoordStrings_(headers, updatedRow);
      sheet.getRange(matchIdx + 1, 1, 1, updatedRow.length).setValues([updatedRow]);
      const millId = updatedRow[headerIndex['Mill ID']] || '';
      return { success: true, action: 'updated', row: matchIdx + 1, millId: millId };
    }

    const millIdCol = headerIndex['Mill ID'];
    if (millIdCol !== undefined && !normalize_(data['Mill ID'])) {
      data['Mill ID'] = generateMillId(rows, millIdCol);
    }

    if (!data['Date Imported']) {
      data['Date Imported'] = Utilities.formatDate(
        new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'
      );
    }

    const newRow = headers.map(function(h) { return data[h] !== undefined ? data[h] : ''; });
    forceCoordStrings_(headers, newRow);
    // PATCH v3: use safeAppendRow_ to prevent appendRow() number coercion
    safeAppendRow_(sheet, headers, newRow);
    return { success: true, action: 'inserted', millId: data['Mill ID'] || '' };
  } finally {
    try { lock.releaseLock(); } catch (e) { /* ignore */ }
  }
}

// ═══════════════════════════════════════════════════════════
//  insertSddRow
// ═══════════════════════════════════════════════════════════

function insertSddRow(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    sanitizeSddPayload_(data);
    const sheet = getSheet('sdd');
    ensureOtherProductExtensionHeadersOrdered_(sheet);
    ensureMainProductExtensionHeadersOrdered_(sheet);
    ensureSddHeadersForPayloads_(sheet, [data || {}]);

    const rows    = sheet.getDataRange().getValues();
    if (!rows.length) throw new Error('SDD sheet is empty (no headers)');

    const headers     = rows[0];
    const headerIndex = indexByHeader_(headers);
    const millIdCol   = headerIndex['Mill ID'];

    if (millIdCol !== undefined && !normalize_(data['Mill ID'])) {
      data['Mill ID'] = generateMillId(rows, millIdCol);
    }

    if (!data['Date Imported']) {
      data['Date Imported'] = Utilities.formatDate(
        new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'
      );
    }

    const row = headers.map(function(h) { return data[h] !== undefined ? data[h] : ''; });
    forceCoordStrings_(headers, row);
    // PATCH v3: use safeAppendRow_
    safeAppendRow_(sheet, headers, row);

    return {
      success : true,
      action  : 'inserted',
      row     : sheet.getLastRow(),
      millId  : data['Mill ID'] || '',
    };
  } finally {
    try { lock.releaseLock(); } catch (e) { /* ignore */ }
  }
}

// ═══════════════════════════════════════════════════════════
//  updateSddRow
// ═══════════════════════════════════════════════════════════

function updateSddRow(rowNum, data) {
  sanitizeSddPayload_(data);
  const r = Number(rowNum);
  if (!r || r < 2) throw new Error('Invalid row number for updateSDD: ' + rowNum);
  const sheet = getSheet('sdd');
  ensureOtherProductExtensionHeadersOrdered_(sheet);
  ensureMainProductExtensionHeadersOrdered_(sheet);
  ensureSddHeadersForPayloads_(sheet, [data || {}]);

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const current = sheet.getRange(r, 1, 1, headers.length).getValues()[0];
  const updated = headers.map(function(h, i) {
    return Object.prototype.hasOwnProperty.call(data, h) ? data[h] : current[i];
  });
  forceCoordStrings_(headers, updated);
  // setValues() on existing Text-formatted cells is safe (no re-coercion)
  sheet.getRange(r, 1, 1, headers.length).setValues([updated]);

  const headerIndex = indexByHeader_(headers);
  const millIdCol   = headerIndex['Mill ID'];
  const millId      = millIdCol !== undefined ? (updated[millIdCol] || '') : '';
  return { success: true, action: 'updated', row: r, millId: millId };
}

// ═══════════════════════════════════════════════════════════
//  SDD BULK UPSERT
// ═══════════════════════════════════════════════════════════

function bulkUpsertSDD(rowsIn) {
  if (!Array.isArray(rowsIn)) throw new Error('bulkUpsertSDD expects an array of rows');
  if (!rowsIn.length) return { success: true, count: 0, results: [] };

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = getSheet('sdd');
    ensureOtherProductExtensionHeadersOrdered_(sheet);
    ensureMainProductExtensionHeadersOrdered_(sheet);
    rowsIn.forEach(function(r) { sanitizeSddPayload_(r); });
    ensureSddHeadersForPayloads_(sheet, rowsIn);

    const raw = sheet.getDataRange().getValues();
    if (!raw.length) throw new Error('SDD sheet is empty (no headers)');

    const headers     = raw[0];
    const headerIndex = indexByHeader_(headers);
    const millIdCol   = headerIndex['Mill ID'];
    const working     = raw.map(function(r) { return r.slice(); });

    let maxExistingNum = 0;
    if (millIdCol !== undefined) {
      for (let i = 1; i < working.length; i++) {
        const m = String(working[i][millIdCol] || '').match(/^SDD-(\d+)$/);
        if (m) maxExistingNum = Math.max(maxExistingNum, parseInt(m[1], 10));
      }
    }

    const todayStamp = Utilities.formatDate(
      new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'
    );

    const updates = [];
    const inserts = [];
    const results = [];

    rowsIn.forEach(function(data) {
      if (!data || typeof data !== 'object') {
        results.push({ success: false, error: 'Empty row payload' });
        return;
      }
      normalizeSubmittedFlags_(data);
      if (!data['Date Imported']) data['Date Imported'] = todayStamp;

      const matchIdx = findMatchingSddRowIndex_(working, headers, headerIndex, data);

      if (matchIdx > 0) {
        const existingObj = rowToObject_(headers, working[matchIdx]);
        assertSubmittedNotDowngraded_(existingObj, data);
        const updatedRow = headers.map(function(h, j) {
          return data[h] !== undefined ? data[h] : working[matchIdx][j];
        });
        forceCoordStrings_(headers, updatedRow);
        working[matchIdx] = updatedRow;
        updates.push({ sheetRow: matchIdx + 1, values: updatedRow });
        results.push({
          success : true, action: 'updated', row: matchIdx + 1,
          millId  : millIdCol !== undefined ? (updatedRow[millIdCol] || '') : '',
        });
      } else {
        if (millIdCol !== undefined && !normalize_(data['Mill ID'])) {
          maxExistingNum++;
          data['Mill ID'] = 'SDD-' + String(maxExistingNum).padStart(4, '0');
        }
        const newRow = headers.map(function(h) { return data[h] !== undefined ? data[h] : ''; });
        forceCoordStrings_(headers, newRow);
        working.push(newRow);
        inserts.push(newRow);
        results.push({
          success : true, action: 'inserted',
          millId  : millIdCol !== undefined ? (newRow[millIdCol] || '') : '',
        });
      }
    });

    // setValues() on existing rows is safe (cells already have '@' format)
    updates.forEach(function(u) {
      sheet.getRange(u.sheetRow, 1, 1, headers.length).setValues([u.values]);
    });

    if (inserts.length) {
      const startRow = sheet.getLastRow() + 1;
      // PATCH v3: use safeAppendRange_ to prevent number coercion on new rows
      safeAppendRange_(sheet, headers, inserts, startRow);
    }

    return { success: true, count: results.length, results: results };
  } finally {
    try { lock.releaseLock(); } catch (e) { /* ignore */ }
  }
}

// ═══════════════════════════════════════════════════════════
//  SDD MATCHING & HELPERS
// ═══════════════════════════════════════════════════════════

function findMatchingSddRowIndex_(rows, headers, idx, data) {
  const millId      = normalize_(data['Mill ID']);
  const millIdCol   = idx['Mill ID'];
  const payloadKind = classifySddPayload_(data);

  function millIdOk(i) {
    if (!millId || millIdCol === undefined) return true;
    return normalize_(rows[i][millIdCol]) === millId;
  }
  function rowKind(i) { return classifySddPayload_(rowToObject_(headers, rows[i])); }
  function sameKind(i) { return rowKind(i) === payloadKind; }

  const ffbId = normalize_(data['FFB - ID']);
  if (payloadKind === 'FFB' && ffbId && idx['FFB - ID'] !== undefined) {
    for (let i = 1; i < rows.length; i++) {
      if (!sameKind(i)) continue;
      if (normalize_(rows[i][idx['FFB - ID']]) === ffbId && millIdOk(i)) return i;
    }
  }

  const ffbSupplier = normalize_(data['FFB - Supplier Name']);
  const ffbMill     = normalize_(data['FFB - Mill Name']);
  if (payloadKind === 'FFB' && ffbSupplier && ffbMill
      && idx['FFB - Supplier Name'] !== undefined && idx['FFB - Mill Name'] !== undefined) {
    for (let i = 1; i < rows.length; i++) {
      if (!sameKind(i)) continue;
      if (normalize_(rows[i][idx['FFB - Supplier Name']]) === ffbSupplier
          && normalize_(rows[i][idx['FFB - Mill Name']]) === ffbMill
          && millIdOk(i)) return i;
    }
  }

  const uml = normalize_(data['UML ID']) || normalize_(data['TML - UML ID']);
  if (payloadKind === 'TML' && uml) {
    const candidates = ['UML ID', 'TML - UML ID'];
    for (let c = 0; c < candidates.length; c++) {
      const h = candidates[c];
      if (idx[h] === undefined) continue;
      for (let i = 1; i < rows.length; i++) {
        if (!sameKind(i)) continue;
        if (normalize_(rows[i][idx[h]]) === uml && millIdOk(i)) return i;
      }
    }
  }

  const tmlMill = normalize_(data['TML - Mill Name']);
  if (payloadKind === 'TML' && tmlMill && idx['TML - Mill Name'] !== undefined) {
    for (let i = 1; i < rows.length; i++) {
      if (!sameKind(i)) continue;
      if (normalize_(rows[i][idx['TML - Mill Name']]) === tmlMill && millIdOk(i)) return i;
    }
  }

  const company = normalize_(data['Company Name']);
  const mill    = normalize_(data['Mill Name']);
  if (payloadKind === 'MAIN' && company && mill
      && idx['Company Name'] !== undefined && idx['Mill Name'] !== undefined) {
    for (let i = 1; i < rows.length; i++) {
      if (!sameKind(i)) continue;
      if (normalize_(rows[i][idx['Company Name']]) === company
          && normalize_(rows[i][idx['Mill Name']]) === mill
          && millIdOk(i)) return i;
    }
  }

  const hasNaturalKey = Boolean(
    (ffbSupplier && ffbMill) || ffbId || uml || tmlMill || (company && mill)
  );
  if (millId && millIdCol !== undefined && !hasNaturalKey) {
    let hit = -1, count = 0;
    for (let i = 1; i < rows.length; i++) {
      if (normalize_(rows[i][millIdCol]) === millId && sameKind(i)) { hit = i; count++; }
    }
    if (count === 1) return hit;
  }

  return -1;
}

function normalize_(v) {
  return String(v === undefined || v === null ? '' : v).trim();
}

function isSubmitted_(rowObj) {
  const st   = normalize_(rowObj['SCR - Screening Status']).toLowerCase();
  const flag = normalize_(rowObj['isSubmitted']).toLowerCase();
  return st === 'submitted' || flag === 'true';
}

function normalizeSubmittedFlags_(payload) {
  if (!payload || typeof payload !== 'object') return;
  const submitted = isSubmitted_(payload);
  if (submitted) {
    payload['SCR - Screening Status'] = 'Submitted';
    payload['isSubmitted'] = 'true';
  } else {
    if (normalize_(payload['SCR - Screening Status']).toLowerCase() === 'draft') {
      payload['SCR - Screening Status'] = 'Draft';
    }
    if (payload['isSubmitted'] === true || payload['isSubmitted'] === false) {
      payload['isSubmitted'] = payload['isSubmitted'] ? 'true' : 'false';
    }
  }
}

function assertSubmittedNotDowngraded_(existingRow, incomingRow) {
  if (isSubmitted_(existingRow) && !isSubmitted_(incomingRow)) {
    throw new Error('Row already submitted and locked; draft overwrite is blocked.');
  }
}

function rowToObject_(headers, rowValues) {
  const out = {};
  for (let i = 0; i < headers.length; i++) out[headers[i]] = rowValues[i];
  return out;
}

function classifySddPayload_(rowObj) {
  const hasFfb = normalize_(rowObj['FFB - Supplier Name']) || normalize_(rowObj['FFB - ID']);
  const hasTml = normalize_(rowObj['TML - Mill Name'])
              || normalize_(rowObj['TML - UML ID'])
              || normalize_(rowObj['UML ID']);
  if (hasFfb) return 'FFB';
  if (hasTml) return 'TML';
  return 'MAIN';
}

function indexByHeader_(headers) {
  const map = {};
  headers.forEach(function(h, i) { map[String(h)] = i; });
  return map;
}

// ═══════════════════════════════════════════════════════════
//  SANITIZE PAYLOAD
// ═══════════════════════════════════════════════════════════

function sanitizeSddPayload_(data) {
  if (!data || typeof data !== 'object') return;
  delete data._row;
  delete data.__row;

  Object.keys(CANONICAL_ALIASES).forEach(function(alias) {
    if (!Object.prototype.hasOwnProperty.call(data, alias)) return;
    const canonical = CANONICAL_ALIASES[alias];
    const aliasVal  = data[alias];
    if (aliasVal !== undefined && aliasVal !== null && String(aliasVal).trim() !== '') {
      if (!data[canonical] || String(data[canonical]).trim() === '') {
        data[canonical] = aliasVal;
      }
    }
    delete data[alias];
  });
}

// ═══════════════════════════════════════════════════════════
//  SDD HEADER AUTO-EXPAND
// ═══════════════════════════════════════════════════════════

function getSddOtherProductExtensionHeaders_() {
  var list = [];
  for (var k = 4; k <= 12; k++) {
    list.push('Other Product ' + k);
    list.push('Other Product ' + k + ' Avg/Month (Ton)');
  }
  return list;
}

function getSddMainProductExtensionHeaders_() {
  var list = [];
  for (var k = 2; k <= 3; k++) {
    list.push('Main Product ' + k);
    list.push('Main Product ' + k + ' Avg Production/Month (Ton)');
    list.push('Main Product ' + k + ' Yield');
  }
  return list;
}

function ensureOtherProductExtensionHeadersOrdered_(sheet) {
  if (!sheet) return;
  var lastCol = Math.max(1, sheet.getLastColumn());
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var want    = getSddOtherProductExtensionHeaders_();
  var present = {};
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i] || '').trim();
    if (h) present[h] = true;
  }
  var missing = want.filter(function(w) { return !present[w]; });
  if (!missing.length) return;

  var cpoCol = -1;
  for (var j = 0; j < headers.length; j++) {
    if (String(headers[j]).trim() === 'CPO Quality - FFA') { cpoCol = j + 1; break; }
  }
  if (cpoCol > 0) {
    sheet.insertColumnsBefore(cpoCol, missing.length);
    sheet.getRange(1, cpoCol, 1, missing.length).setValues([missing]);
  } else {
    sheet.insertColumnsAfter(headers.length, missing.length);
    sheet.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
  }
}

function ensureMainProductExtensionHeadersOrdered_(sheet) {
  if (!sheet) return;
  var lastCol = Math.max(1, sheet.getLastColumn());
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var want    = getSddMainProductExtensionHeaders_();
  var present = {};
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i] || '').trim();
    if (h) present[h] = true;
  }
  var missing = want.filter(function(w) { return !present[w]; });
  if (!missing.length) return;

  var cpoCol = -1;
  for (var j = 0; j < headers.length; j++) {
    if (String(headers[j]).trim() === 'CPO Quality - FFA') { cpoCol = j + 1; break; }
  }
  if (cpoCol > 0) {
    sheet.insertColumnsBefore(cpoCol, missing.length);
    sheet.getRange(1, cpoCol, 1, missing.length).setValues([missing]);
  } else {
    sheet.insertColumnsAfter(headers.length, missing.length);
    sheet.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
  }
}

function ensureSddHeadersForPayloads_(sheet, payloadRows) {
  if (!sheet) return;
  if (!Array.isArray(payloadRows) || !payloadRows.length) return;

  const lastCol = Math.max(1, sheet.getLastColumn());
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0] || [];
  const existing = new Set(
    headers.map(function(h) { return String(h || '').trim(); }).filter(Boolean)
  );

  const missing = [];
  payloadRows.forEach(function(row) {
    if (!row || typeof row !== 'object') return;
    Object.keys(row).forEach(function(k) {
      const key = String(k || '').trim();
      if (!key) return;
      if (CANONICAL_ALIASES[key]) return;
      if (!existing.has(key)) { existing.add(key); missing.push(key); }
    });
  });
  if (!missing.length) return;

  const start = headers.length + 1;
  sheet.insertColumnsAfter(headers.length, missing.length);
  sheet.getRange(1, start, 1, missing.length).setValues([missing]);
}

function generateMillId(rows, millIdCol) {
  let maxNum = 0;
  for (let i = 1; i < rows.length; i++) {
    const id    = String(rows[i][millIdCol] || '');
    const match = id.match(/^SDD-(\d+)$/);
    if (match) maxNum = Math.max(maxNum, parseInt(match[1], 10));
  }
  return 'SDD-' + String(maxNum + 1).padStart(4, '0');
}

// ═══════════════════════════════════════════════════════════
//  GET BY MILL ID
// ═══════════════════════════════════════════════════════════

function getByMillId(millId) {
  const want = String(millId || '').trim();
  if (!want) return { success: false, error: 'millId is required' };

  try {
    ensureRelationalHeaders_('sddMain');
    const mainResult = readRelSheet_('sddMain', false);
    for (let i = 0; i < mainResult.rows.length; i++) {
      const obj = mainResult.rows[i].obj;
      if (String(obj['Mill ID'] || '').trim() === want) {
        const sid   = String(obj['submission_id'] || '').trim();
        const mills = sid ? findChildRows_('sddMill', sid).map(function(r) { return r.obj; }) : [];
        const ffbs  = sid ? findChildRows_('sddFfb',  sid).map(function(r) { return r.obj; }) : [];
        return {
          success  : true,
          source   : 'relational',
          _row     : mainResult.rows[i]._sheetRow,
          main     : obj,
          mills    : mills,
          ffb_rows : ffbs,
        };
      }
    }
  } catch (relErr) {
    // fall through to legacy
  }

  const sheet = getSheet('sdd');
  const rows  = sheet.getDataRange().getValues();
  if (!rows.length) return { success: false, error: 'Empty sheet' };

  const headers   = rows[0];
  const millIdCol = headers.indexOf('Mill ID');
  if (millIdCol < 0) {
    return { success: false, error: '"Mill ID" header not found — check sheet column order' };
  }

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][millIdCol]) === want) {
      const obj = { success: true, source: 'legacy', _row: i + 1 };
      headers.forEach(function(h, j) { obj[h] = rows[i][j]; });
      return obj;
    }
  }

  return { success: false, error: 'Not found' };
}

// ═══════════════════════════════════════════════════════════
//  GENERIC CRUD
//  (+ Mill Quarter/Year: dashboard sends QUARTER/YEAR; sheet may use Quarter/Year)
// ═══════════════════════════════════════════════════════════

/**
 * Map payload keys QUARTER / YEAR (from sustain-dashboard) onto the exact
 * header names present in row 1 of the Mill sheet (e.g. Quarter, Year).
 */
function resolveMillQuarterYearKeys_(data, headers) {
  if (!data || typeof data !== 'object' || !Array.isArray(headers)) return;
  var list = headers.map(function(x) { return String(x || '').trim(); });
  function findQuarterCol() {
    for (var i = 0; i < list.length; i++) {
      if (/^quarter$/i.test(list[i])) return list[i];
    }
    return null;
  }
  function findYearCol() {
    for (var j = 0; j < list.length; j++) {
      var h = list[j];
      if (h === 'Year' || h === 'YEAR') return h;
    }
    for (var k = 0; k < list.length; k++) {
      var y = list[k];
      if (!/^year$/i.test(y)) continue;
      if (/planted|capacity|mill|tanam|ffb|tml|issue|expiry|average/i.test(y)) continue;
      return y;
    }
    return null;
  }
  var qCol = findQuarterCol();
  var yCol = findYearCol();
  if (qCol && data['QUARTER'] !== undefined && String(data['QUARTER']).trim() !== '' &&
      (data[qCol] === undefined || String(data[qCol]).trim() === '')) {
    data[qCol] = data['QUARTER'];
  }
  if (yCol && data['YEAR'] !== undefined && String(data['YEAR']).trim() !== '' &&
      (data[yCol] === undefined || String(data[yCol]).trim() === '')) {
    data[yCol] = data['YEAR'];
  }
}

/** Expose QUARTER / YEAR on getAll('mill') for clients that expect uppercase keys. */
function mirrorMillQuarterYearOnRead_(obj) {
  if (!obj || typeof obj !== 'object') return;
  Object.keys(obj).forEach(function(k) {
    if (k === '_row') return;
    if (/^quarter$/i.test(String(k).trim()) && obj['QUARTER'] === undefined) obj['QUARTER'] = obj[k];
  });
  Object.keys(obj).forEach(function(k) {
    if (k === '_row') return;
    var t = String(k).trim();
    if ((t === 'Year' || t === 'YEAR') && obj['YEAR'] === undefined) obj['YEAR'] = obj[k];
  });
}

function getData(sheetKey) {
  const sheet = getSheet(sheetKey);
  const range = sheet.getDataRange();
  const rows  = range.getValues();
  if (!rows.length) return [];
  const headers = rows[0];

  // For Mill data we need display-preserved numerics (e.g. "66.000")
  // so UI matches the exact visual formatting in Google Sheets.
  const dispRows = sheetKey === 'mill' ? range.getDisplayValues() : null;

  return rows.slice(1).map(function(row, i) {
    const sourceRow = dispRows ? dispRows[i + 1] : row;
    const obj = { _row: i + 2 };
    headers.forEach(function(h, j) { obj[h] = sourceRow[j]; });
    if (sheetKey === 'mill') mirrorMillQuarterYearOnRead_(obj);
    return obj;
  });
}

function addRow(sheetKey, data) {
  const sheet   = getSheet(sheetKey);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (sheetKey === 'mill') resolveMillQuarterYearKeys_(data, headers);
  const newRow  = headers.map(function(h) { return data[h] !== undefined ? data[h] : ''; });
  sheet.appendRow(newRow);
  return { success: true };
}

function updateRow(sheetKey, rowNum, data) {
  const sheet = getSheet(sheetKey);
  const r     = Number(rowNum);
  if (!r || r < 2) throw new Error('Invalid row number for update: ' + rowNum);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (sheetKey === 'mill') resolveMillQuarterYearKeys_(data, headers);
  const current = sheet.getRange(r, 1, 1, headers.length).getValues()[0];
  const updated = headers.map(function(h, j) {
    return data[h] !== undefined ? data[h] : current[j];
  });
  sheet.getRange(r, 1, 1, updated.length).setValues([updated]);
  return { success: true };
}

function deleteRow(sheetKey, rowNum) {
  const sheet = getSheet(sheetKey);
  const r     = Number(rowNum);
  if (!r || r < 2) throw new Error('Invalid row number for delete: ' + rowNum);
  sheet.deleteRow(r);
  return { success: true };
}

function bulkDelete(sheetKey, rowNums) {
  if (!Array.isArray(rowNums) || !rowNums.length) return { success: true, deleted: 0 };
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = getSheet(sheetKey);
    let validRows = rowNums
      .map(function(r) { return Number(r); })
      .filter(function(r) { return r >= 2 && !isNaN(r); });
    const seen = {};
    validRows = validRows.filter(function(r) {
      if (seen[r]) return false;
      seen[r] = true;
      return true;
    });
    if (!validRows.length) return { success: true, deleted: 0 };
    validRows.sort(function(a, b) { return b - a; });
    validRows.forEach(function(r) { sheet.deleteRow(r); });
    return { success: true, deleted: validRows.length };
  } finally {
    try { lock.releaseLock(); } catch (e) { /* ignore */ }
  }
}

// ═══════════════════════════════════════════════════════════
//  UTILITIES
// ═══════════════════════════════════════════════════════════

function parsePostBody_(e) {
  if (!e || !e.postData || !e.postData.contents) throw new Error('Missing request body');
  try {
    return JSON.parse(e.postData.contents);
  } catch (err) {
    throw new Error('Invalid JSON body: ' + String(e.postData.contents).slice(0, 200));
  }
}

function getSheet(sheetKey) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const name = SHEETS[sheetKey];
  if (!name) throw new Error('Sheet key not found: ' + sheetKey);
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('Tab not found: ' + name);
  return sheet;
}

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ═══════════════════════════════════════════════════════════════════════════
//  RELATIONAL SDD LAYER
// ═══════════════════════════════════════════════════════════════════════════

function generateSubmissionId_() {
  const d   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
  const rnd = Math.random().toString(16).slice(2, 8).toUpperCase();
  return 'SUB-' + d + '-' + rnd;
}

function nowIso_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
}

function callerEmail_() {
  try { return Session.getActiveUser().getEmail() || 'system'; } catch (e) { return 'system'; }
}

function assertRequiredKeys_(obj, keys, ctx) {
  keys.forEach(function(k) {
    if (!obj || obj[k] === undefined || obj[k] === null || String(obj[k]).trim() === '') {
      throw new Error('[' + ctx + '] Missing required field: ' + k);
    }
  });
}

function assertSupplierType_(val) {
  const VALID = ['MILL', 'KCP', 'TRADER'];
  const v     = String(val || '').trim().toUpperCase();
  if (VALID.indexOf(v) === -1) {
    throw new Error('Invalid supplier_type "' + val + '". Must be one of: ' + VALID.join(', '));
  }
  return v;
}

/**
 * readRelSheet_ — normalises coordinate fields to dot-decimal strings on read.
 */
function readRelSheet_(sheetKey, includeDeleted) {
  ensureRelationalHeaders_(sheetKey);
  const ws   = getSheet(sheetKey);
  const data = ws.getDataRange().getValues();
  if (data.length < 1) return { ws: ws, headers: [], rows: [] };

  const headers = data[0].map(function(h) { return String(h || '').trim(); });
  const delIdx  = headers.indexOf('is_deleted');

  const coordIdx = new Set(
    COORD_COLUMN_NAMES.map(function(n) { return headers.indexOf(n); })
                      .filter(function(i) { return i >= 0; })
  );

  const rows = [];
  for (let r = 1; r < data.length; r++) {
    if (!includeDeleted && delIdx >= 0 && String(data[r][delIdx]) === '1') continue;
    const rowData = data[r].map(function(v, i) {
      if (!coordIdx.has(i)) return v;
      // Normalise to dot-decimal on read — also recovers any residual corruption
      var s = String(v === null || v === undefined ? '' : v).trim();
      return recoverCoordValue_(s, headers[i]);
    });
    rows.push({ _sheetRow: r + 1, obj: rowToObject_(headers, rowData) });
  }
  return { ws: ws, headers: headers, rows: rows };
}

/**
 * appendRelRow_ — PATCH v3: uses safeAppendRow_ to prevent appendRow() coercion.
 */
function appendRelRow_(ws, headers, obj) {
  const row = headers.map(function(h) {
    const v = obj[h];
    if (v === undefined || v === null) return '';
    if (FORCE_STRING_FIELDS.has(h)) return String(v);
    return v;
  });
  // PATCH v3: safeAppendRow_ re-writes coord cells as Text after appendRow()
  safeAppendRow_(ws, headers, row);
}

/**
 * patchRelRow_ — setValues() on existing Text-formatted rows is safe.
 */
function patchRelRow_(ws, headers, sheetRow, patch) {
  const current = ws.getRange(sheetRow, 1, 1, headers.length).getValues()[0];
  const updated = headers.map(function(h, i) {
    if (Object.prototype.hasOwnProperty.call(patch, h)) {
      const v = patch[h];
      if (FORCE_STRING_FIELDS.has(h)) return String(v === null || v === undefined ? '' : v);
      return v;
    }
    return current[i];
  });
  ws.getRange(sheetRow, 1, 1, headers.length).setValues([updated]);
}

function softDelRelRow_(ws, headers, sheetRow, user, now) {
  patchRelRow_(ws, headers, sheetRow, {
    is_deleted : '1',
    updated_at : now,
    updated_by : user,
  });
}

function injectTechKeys_(obj, meta) {
  obj['submission_id'] = meta.submission_id;
  if (meta.line_id !== undefined) obj['line_id'] = meta.line_id;
  obj['updated_at'] = meta.now;
  obj['updated_by'] = meta.user;
  if (meta.isNew) {
    obj['created_at'] = meta.now;
    obj['created_by'] = meta.user;
    obj['is_deleted'] = '0';
  }
}

function findMainRow_(sheetKey, sid) {
  if (!sid) return null;
  const result = readRelSheet_(sheetKey, true);
  for (let i = 0; i < result.rows.length; i++) {
    if (String(result.rows[i].obj['submission_id'] || '').trim() === sid) {
      return {
        ws        : result.ws,
        headers   : result.headers,
        _sheetRow : result.rows[i]._sheetRow,
        obj       : result.rows[i].obj,
      };
    }
  }
  return null;
}

function findChildRows_(sheetKey, sid) {
  const result = readRelSheet_(sheetKey, false);
  return result.rows.filter(function(r) {
    return String(r.obj['submission_id'] || '').trim() === sid;
  });
}

function nextLineId_(sheetKey, sid) {
  const result = readRelSheet_(sheetKey, true);
  let max = 0;
  result.rows.forEach(function(r) {
    if (String(r.obj['submission_id'] || '').trim() !== sid) return;
    const lid = parseInt(r.obj['line_id'], 10);
    if (!isNaN(lid) && lid > max) max = lid;
  });
  return max + 1;
}

function bulkInsertChildren_(sheetKey, sid, childRows, supplierType, user, now, nextLid) {
  if (!Array.isArray(childRows) || !childRows.length) return 0;
  const result  = readRelSheet_(sheetKey, true);
  const ws      = result.ws;
  const headers = result.headers;

  childRows.forEach(function(row, i) {
    const obj = Object.assign({}, row);
    injectTechKeys_(obj, {
      submission_id : sid,
      line_id       : nextLid + i,
      now           : now,
      user          : user,
      isNew         : true,
    });
    obj['supplier_type'] = supplierType;
    appendRelRow_(ws, headers, obj);
  });
  return childRows.length;
}

// ───────────────────────────────────────────────────────────────────────────
//  1. createSubmission
// ───────────────────────────────────────────────────────────────────────────
function createSubmission(payload) {
  assertRequiredKeys_(payload,      ['main'],          'createSubmission');
  assertRequiredKeys_(payload.main, ['supplier_type'], 'createSubmission.main');

  const supplierType = assertSupplierType_(payload.main.supplier_type);
  const sid          = generateSubmissionId_();
  const now          = nowIso_();
  const user         = callerEmail_();

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    ensureAllRelationalHeaders_();

    const mainResult = readRelSheet_('sddMain', true);
    const mainObj    = Object.assign({}, payload.main);
    injectTechKeys_(mainObj, { submission_id: sid, now: now, user: user, isNew: true });
    mainObj['supplier_type'] = supplierType;
    if (!mainObj['Date Imported']) {
      mainObj['Date Imported'] = Utilities.formatDate(
        new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'
      );
    }
    if (!mainObj['Imported By']) mainObj['Imported By'] = user;
    appendRelRow_(mainResult.ws, mainResult.headers, mainObj);

    const millsInserted = bulkInsertChildren_(
      'sddMill', sid, payload.mills    || [], supplierType, user, now,
      nextLineId_('sddMill', sid)
    );
    const ffbInserted = bulkInsertChildren_(
      'sddFfb',  sid, payload.ffb_rows || [], supplierType, user, now,
      nextLineId_('sddFfb', sid)
    );

    return {
      success        : true,
      submission_id  : sid,
      mills_inserted : millsInserted,
      ffb_inserted   : ffbInserted,
    };
  } finally {
    try { lock.releaseLock(); } catch (e) { /* ignore */ }
  }
}

// ───────────────────────────────────────────────────────────────────────────
//  2. updateSubmission
// ───────────────────────────────────────────────────────────────────────────
function updateSubmission(payload) {
  assertRequiredKeys_(payload, ['submission_id'], 'updateSubmission');
  const sid  = String(payload.submission_id).trim();
  const now  = nowIso_();
  const user = callerEmail_();

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    ensureAllRelationalHeaders_();

    const mainHit = findMainRow_('sddMain', sid);
    if (!mainHit) throw new Error('submission_id not found: ' + sid);
    if (String(mainHit.obj['is_deleted']) === '1') {
      throw new Error('Cannot update a deleted submission: ' + sid);
    }

    if (payload.main && Object.keys(payload.main).length > 0) {
      const patch = Object.assign({}, payload.main);
      delete patch['submission_id'];
      delete patch['created_at'];
      delete patch['created_by'];
      delete patch['is_deleted'];
      patch['updated_at'] = now;
      patch['updated_by'] = user;
      patchRelRow_(mainHit.ws, mainHit.headers, mainHit._sheetRow, patch);

      if (payload.main.statusSDD !== undefined || payload.main.statusBossDecision !== undefined) {
        const mergedMain = Object.assign({}, mainHit.obj, patch);
        syncContactFromApprovedSubmission_(sid, mergedMain, user, now);
        syncTtpFromApprovedSubmission_(sid, mergedMain, user, now);
      }
    }

    const supplierType = String(mainHit.obj['supplier_type'] || 'MILL');

    const millStats = payload.mills !== undefined
      ? upsertChildSheet_('sddMill', sid, payload.mills,    supplierType, user, now)
      : { upserted: 0, inserted: 0, deleted: 0 };

    const ffbStats = payload.ffb_rows !== undefined
      ? upsertChildSheet_('sddFfb',  sid, payload.ffb_rows, supplierType, user, now)
      : { upserted: 0, inserted: 0, deleted: 0 };

    let ttpSync = null;
    if (payload.ffb_rows !== undefined) {
      const mainHitAfter = findMainRow_('sddMain', sid);
      if (mainHitAfter) {
        const decisionAfter = normalizeSddDecisionLabel_(
          mainHitAfter.obj['statusSDD'] || mainHitAfter.obj['statusBossDecision'] || ''
        );
        if (decisionAfter === 'APPROVED') {
          ttpSync = syncTtpFromApprovedSubmission_(sid, mainHitAfter.obj, user, now);
        }
      }
    }

    return {
      success       : true,
      submission_id : sid,
      mills         : millStats,
      ffb           : ffbStats,
      ttp_sync      : ttpSync,
    };
  } finally {
    try { lock.releaseLock(); } catch (e) { /* ignore */ }
  }
}

function upsertChildSheet_(sheetKey, sid, incomingRows, supplierType, user, now) {
  const result  = readRelSheet_(sheetKey, true);
  const ws      = result.ws;
  const headers = result.headers;

  const existingMap = {};
  let maxLid = 0;
  result.rows.forEach(function(r) {
    if (String(r.obj['submission_id'] || '').trim() !== sid) return;
    const lid    = String(r.obj['line_id'] || '').trim();
    const lidNum = parseInt(lid, 10);
    if (!isNaN(lidNum) && lidNum > maxLid) maxLid = lidNum;
    if (String(r.obj['is_deleted']) !== '1' && lid) {
      existingMap[lid] = r;
    }
  });

  const touchedLineIds = {};
  let upserted = 0;
  let inserted = 0;

  (incomingRows || []).forEach(function(row) {
    const inLid = row['line_id'] !== undefined ? String(row['line_id']).trim() : '';

    if (inLid && existingMap[inLid]) {
      touchedLineIds[inLid] = true;
      const patch = Object.assign({}, row);
      delete patch['submission_id'];
      delete patch['created_at'];
      delete patch['created_by'];
      delete patch['is_deleted'];
      patch['updated_at']    = now;
      patch['updated_by']    = user;
      patch['supplier_type'] = supplierType;
      patchRelRow_(ws, headers, existingMap[inLid]._sheetRow, patch);
      upserted++;
    } else {
      maxLid++;
      const obj = Object.assign({}, row);
      injectTechKeys_(obj, {
        submission_id : sid,
        line_id       : inLid ? parseInt(inLid, 10) : maxLid,
        now           : now,
        user          : user,
        isNew         : true,
      });
      obj['supplier_type'] = supplierType;
      if (inLid) touchedLineIds[inLid] = true;
      appendRelRow_(ws, headers, obj);
      inserted++;
    }
  });

  let deleted = 0;
  Object.keys(existingMap).forEach(function(lid) {
    if (!touchedLineIds[lid]) {
      softDelRelRow_(ws, headers, existingMap[lid]._sheetRow, user, now);
      deleted++;
    }
  });

  return { upserted: upserted, inserted: inserted, deleted: deleted };
}

// ───────────────────────────────────────────────────────────────────────────
//  3. getSubmissionById
// ───────────────────────────────────────────────────────────────────────────
function getSubmissionById(sid) {
  if (!sid || String(sid).trim() === '') {
    throw new Error('getSubmissionById: submission_id is required');
  }
  sid = String(sid).trim();

  const mainHit = findMainRow_('sddMain', sid);
  if (!mainHit)                                  throw new Error('submission_id not found: ' + sid);
  if (String(mainHit.obj['is_deleted']) === '1') throw new Error('Submission is deleted: ' + sid);

  const mills   = findChildRows_('sddMill', sid).map(function(r) { return r.obj; });
  const ffbRows = findChildRows_('sddFfb',  sid).map(function(r) { return r.obj; });

  function byLineId(a, b) {
    return parseInt(a['line_id'] || 0, 10) - parseInt(b['line_id'] || 0, 10);
  }
  mills.sort(byLineId);
  ffbRows.sort(byLineId);

  return {
    success  : true,
    main     : mainHit.obj,
    mills    : mills,
    ffb_rows : ffbRows,
  };
}

// ───────────────────────────────────────────────────────────────────────────
//  4. listSubmissions
// ───────────────────────────────────────────────────────────────────────────
function listSubmissions(params) {
  params = params || {};

  const page         = Math.max(1,   parseInt(params.page      || 1,   10));
  const pageSize     = Math.min(200, Math.max(1, parseInt(params.page_size || 50, 10)));
  const filterType   = String(params.supplier_type || '').toUpperCase();
  const filterStatus = String(params.status        || '');
  const filterScr    = String(params.scr_status    || '');
  const search       = String(params.search        || '').toLowerCase();

  let rows = readRelSheet_('sddMain', false).rows;

  if (filterType) {
    rows = rows.filter(function(r) {
      return String(r.obj['supplier_type'] || '').toUpperCase() === filterType;
    });
  }
  if (filterStatus) {
    rows = rows.filter(function(r) {
      return String(r.obj['statusSDD'] || '') === filterStatus;
    });
  }
  if (filterScr) {
    rows = rows.filter(function(r) {
      return String(r.obj['SCR - Screening Status'] || '') === filterScr;
    });
  }
  if (search) {
    rows = rows.filter(function(r) {
      const c = String(r.obj['Company Name'] || '').toLowerCase();
      const m = String(r.obj['Mill Name']    || '').toLowerCase();
      return c.indexOf(search) >= 0 || m.indexOf(search) >= 0;
    });
  }

  const total = rows.length;
  const start = (page - 1) * pageSize;
  const data  = rows.slice(start, start + pageSize).map(function(r) { return r.obj; });

  return {
    success   : true,
    total     : total,
    page      : page,
    page_size : pageSize,
    data      : data,
  };
}

// ───────────────────────────────────────────────────────────────────────────
//  5. setSubmissionStatus
// ───────────────────────────────────────────────────────────────────────────
function setSubmissionStatus(payload) {
  assertRequiredKeys_(payload, ['submission_id'], 'setSubmissionStatus');
  const sid  = String(payload.submission_id).trim();
  const now  = nowIso_();
  const user = callerEmail_();

  const mainHit = findMainRow_('sddMain', sid);
  if (!mainHit) throw new Error('submission_id not found: ' + sid);
  if (String(mainHit.obj['is_deleted']) === '1') {
    throw new Error('Cannot update status of a deleted submission: ' + sid);
  }

  const patch = { updated_at: now, updated_by: user };

  if (payload.statusSDD          !== undefined) patch['statusSDD']                = payload.statusSDD;
  if (payload.scr_status         !== undefined) patch['SCR - Screening Status']   = payload.scr_status;
  if (payload.scr_notes          !== undefined) patch['SCR - Notes']              = payload.scr_notes;
  if (payload.scr_recommendation !== undefined) patch['SCR - Recommendation']     = payload.scr_recommendation;
  if (payload.scr_risk_level     !== undefined) patch['SCR - Overall Risk Level'] = payload.scr_risk_level;

  if (payload.scr_status !== undefined || payload.scr_recommendation !== undefined) {
    patch['SCR - Last Updated']   = now;
    patch['SCR - Screened By']    = user;
    patch['SCR - Screening Date'] = now;
  }

  patchRelRow_(mainHit.ws, mainHit.headers, mainHit._sheetRow, patch);

  let contactSync = null;
  let ttpSync = null;
  if (payload.statusSDD !== undefined) {
    const mergedMain = Object.assign({}, mainHit.obj, patch);
    contactSync = syncContactFromApprovedSubmission_(sid, mergedMain, user, now);
    ttpSync = syncTtpFromApprovedSubmission_(sid, mergedMain, user, now);
  }

  return {
    success        : true,
    submission_id  : sid,
    patched_fields : Object.keys(patch),
    contact_sync   : contactSync,
    ttp_sync       : ttpSync,
  };
}

// ───────────────────────────────────────────────────────────────────────────
//  6. deleteSubmission
// ───────────────────────────────────────────────────────────────────────────
function deleteSubmission(payload) {
  assertRequiredKeys_(payload, ['submission_id'], 'deleteSubmission');
  const sid  = String(payload.submission_id).trim();
  const now  = nowIso_();
  const user = callerEmail_();

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const mainHit = findMainRow_('sddMain', sid);
    if (!mainHit) throw new Error('submission_id not found: ' + sid);
    if (String(mainHit.obj['is_deleted']) === '1') {
      throw new Error('Submission already deleted: ' + sid);
    }

    softDelRelRow_(mainHit.ws, mainHit.headers, mainHit._sheetRow, user, now);

    let millsDeleted = 0;
    const millResult = readRelSheet_('sddMill', false);
    millResult.rows.forEach(function(r) {
      if (String(r.obj['submission_id'] || '').trim() === sid) {
        softDelRelRow_(millResult.ws, millResult.headers, r._sheetRow, user, now);
        millsDeleted++;
      }
    });

    let ffbDeleted = 0;
    const ffbResult = readRelSheet_('sddFfb', false);
    ffbResult.rows.forEach(function(r) {
      if (String(r.obj['submission_id'] || '').trim() === sid) {
        softDelRelRow_(ffbResult.ws, ffbResult.headers, r._sheetRow, user, now);
        ffbDeleted++;
      }
    });

    return {
      success        : true,
      submission_id  : sid,
      mills_deleted  : millsDeleted,
      ffb_deleted    : ffbDeleted,
    };
  } finally {
    try { lock.releaseLock(); } catch (e) { /* ignore */ }
  }
}

// ═══════════════════════════════════════════════════════════════════════════
//  LEGACY ADAPTERS
// ═══════════════════════════════════════════════════════════════════════════

function splitLegacyPayload_(flat) {
  const main = {};
  const mill = {};
  const ffb  = {};

  Object.keys(flat).forEach(function(k) {
    const v = flat[k];
    if (k.indexOf('TML - ') === 0 || k.indexOf('SCR - TML') === 0) {
      mill[k] = v;
    } else if (k.indexOf('FFB - ') === 0) {
      ffb[k] = v;
    } else {
      main[k] = v;
    }
  });

  if (!main['supplier_type'] && main['Supplier Type']) {
    main['supplier_type'] = String(main['Supplier Type']).trim().toUpperCase();
    delete main['Supplier Type'];
  }
  if (!main['supplier_type']) main['supplier_type'] = 'MILL';

  const result = { main: main };
  if (Object.keys(mill).length > 0) result.mills    = [mill];
  if (Object.keys(ffb).length  > 0) result.ffb_rows = [ffb];
  return result;
}

function resolveSidByMillId_(millId, companyName, millName) {
  const targetMillId = String(millId   || '').trim();
  const targetCo     = String(companyName || '').trim().toLowerCase();
  const targetMill   = String(millName    || '').trim().toLowerCase();

  const result = readRelSheet_('sddMain', false);

  if (targetMillId) {
    for (let i = 0; i < result.rows.length; i++) {
      if (String(result.rows[i].obj['Mill ID'] || '').trim() === targetMillId) {
        return String(result.rows[i].obj['submission_id'] || '').trim() || null;
      }
    }
  }

  if (targetCo && targetMill) {
    for (let i = 0; i < result.rows.length; i++) {
      const co   = String(result.rows[i].obj['Company Name'] || '').trim().toLowerCase();
      const mill = String(result.rows[i].obj['Mill Name']    || '').trim().toLowerCase();
      if (co === targetCo && mill === targetMill) {
        return String(result.rows[i].obj['submission_id'] || '').trim() || null;
      }
    }
  }

  return null;
}

function legacyAdapterInsert_(data) {
  const legacyResult = insertSddRow(data);
  try {
    ensureAllRelationalHeaders_();
    const split     = splitLegacyPayload_(data);
    const relResult = createSubmission(split);
    legacyResult['submission_id'] = relResult.submission_id;
    legacyResult['relational']    = { mills_inserted: relResult.mills_inserted, ffb_inserted: relResult.ffb_inserted };
  } catch (err) {
    legacyResult['relational_error'] = err.message;
  }
  return legacyResult;
}

function legacyAdapterUpdate_(rowNum, data) {
  const legacyResult = updateSddRow(rowNum, data);
  try {
    ensureAllRelationalHeaders_();
    const sid   = resolveSidByMillId_(data['Mill ID'], data['Company Name'], data['Mill Name']);
    const split = splitLegacyPayload_(data);
    let relResult;
    if (sid) {
      split['submission_id'] = sid;
      relResult = updateSubmission(split);
    } else {
      relResult = createSubmission(split);
    }
    legacyResult['submission_id'] = relResult.submission_id || sid;
    legacyResult['relational']    = relResult;
  } catch (err) {
    legacyResult['relational_error'] = err.message;
  }
  return legacyResult;
}

function legacyAdapterUpsert_(data) {
  const legacyResult = upsertSDD(data);
  try {
    ensureAllRelationalHeaders_();
    const sid   = resolveSidByMillId_(data['Mill ID'], data['Company Name'], data['Mill Name']);
    const split = splitLegacyPayload_(data);
    let relResult;
    if (sid) {
      split['submission_id'] = sid;
      relResult = updateSubmission(split);
    } else {
      relResult = createSubmission(split);
    }
    legacyResult['submission_id'] = relResult.submission_id || sid;
    legacyResult['relational']    = relResult;
  } catch (err) {
    legacyResult['relational_error'] = err.message;
  }
  return legacyResult;
}
// ═══════════════════════════════════════════════════════════
//  SUPPLY IMPORT DRAFT
// ═══════════════════════════════════════════════════════════

function ensureSupplyDraftHeaders_() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var name = SHEETS.supplyDraft;
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, SUPPLY_DRAFT_HEADERS.length).setValues([SUPPLY_DRAFT_HEADERS]);
    sheet.setFrozenRows(1);
    return sheet;
  }
  var existing    = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0]
                         .map(function(h) { return String(h || '').trim(); });
  var existingSet = new Set(existing.filter(Boolean));
  var missing     = SUPPLY_DRAFT_HEADERS.filter(function(h) { return !existingSet.has(h); });
  if (missing.length) {
    var start = existing.filter(Boolean).length + 1;
    sheet.getRange(1, start, 1, missing.length).setValues([missing]);
  }
  return sheet;
}

function saveSupplyDraft_(rows, batchId, meta) {
  ensureSupplyDraftHeaders_();
  var sheet   = getSheet('supplyDraft');
  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h || '').trim(); });
  var draftIdCol = headers.indexOf('draft_id');
  var now  = nowIso_();
  var user = callerEmail_();
  var saved = 0;

  rows.forEach(function(row) {
    var draftId = String(row.draft_id || '').trim();
    if (!draftId) return;
    var existingSheetRow = -1;
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][draftIdCol] || '').trim() === draftId) {
        existingSheetRow = r + 1;
        break;
      }
    }
    var rowArr = headers.map(function(h) {
      if (h === 'updated_at') return now;
      if (h === 'batch_id' && !row[h]) return batchId;
      if (h === 'created_at' && existingSheetRow < 0) return now;
      if (h === 'created_by' && existingSheetRow < 0) return user;
      if (h === 'status' && !row[h]) return 'draft';
      var v = row[h];
      return (v !== undefined && v !== null) ? v : '';
    });
    if (existingSheetRow > 0) {
      sheet.getRange(existingSheetRow, 1, 1, rowArr.length).setValues([rowArr]);
    } else {
      sheet.appendRow(rowArr);
    }
    saved++;
  });
  return { success: true, saved: saved, batch_id: batchId };
}

function submitSupplyDraft_(batchId, rows) {
  if (!batchId) throw new Error('batch_id required');
  ensureSupplyDraftHeaders_();

  // Ensure all Mill Onboarding headers exist in target sheet
  ensureRelationalHeaders_('mill');

  var millSheet   = getSheet('mill');
  // Re-read headers AFTER ensureRelationalHeaders_ so we have all columns
  var millHeaders = millSheet.getDataRange().getValues()[0].map(function(h) { return String(h || '').trim(); });

  var draftSheet   = getSheet('supplyDraft');
  var draftData    = draftSheet.getDataRange().getValues();
  var draftHeaders = draftData[0].map(function(h) { return String(h || '').trim(); });
  var draftIdColD  = draftHeaders.indexOf('draft_id');
  var statusColD   = draftHeaders.indexOf('status');
  var now          = nowIso_();
  var submitted    = 0;

  rows.forEach(function(row) {
    var millRow = millHeaders.map(function(h) {
      // Handle Quarter/Year case variants
      if (h === 'Quarter' || h === 'QUARTER') return row['QUARTER'] || row['quarter'] || '';
      if (h === 'Year'    || h === 'YEAR')    return row['YEAR']    || row['year']    || '';
      var v = row[h];
      return (v !== undefined && v !== null) ? v : '';
    });
    millSheet.appendRow(millRow);

    // Mark draft row as submitted
    var draftId = String(row.draft_id || '').trim();
    if (draftId) {
      for (var r = 1; r < draftData.length; r++) {
        if (String(draftData[r][draftIdColD] || '').trim() === draftId) {
          draftSheet.getRange(r + 1, statusColD + 1).setValue('submitted');
          var updIdx = draftHeaders.indexOf('updated_at');
          if (updIdx >= 0) draftSheet.getRange(r + 1, updIdx + 1).setValue(now);
          break;
        }
      }
    }
    submitted++;
  });
  return { success: true, submitted: submitted, batch_id: batchId };
}

function deleteSupplyDraft_(draftId, batchId) {
  ensureSupplyDraftHeaders_();
  var sheet   = getSheet('supplyDraft');
  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h || '').trim(); });
  var draftIdCol = headers.indexOf('draft_id');
  var batchIdCol = headers.indexOf('batch_id');
  var toDelete = [];
  for (var r = data.length - 1; r >= 1; r--) {
    var rowDId = String(data[r][draftIdCol] || '').trim();
    var rowBId = String(data[r][batchIdCol] || '').trim();
    if ((draftId && rowDId === draftId) || (batchId && !draftId && rowBId === batchId)) {
      toDelete.push(r + 1);
    }
  }
  toDelete.forEach(function(sheetRow) { sheet.deleteRow(sheetRow); });
  return { success: true, deleted: toDelete.length };
}
