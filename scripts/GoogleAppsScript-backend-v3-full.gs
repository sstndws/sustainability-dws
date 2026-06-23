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
  blMonitoring    : 'BL Monitoring',
  blReference     : 'BL Reference',
  questionnaireMonitoring : 'Questionnaire Monitoring',
  eudrPotential           : 'EUDR Potential',
  eudrStatusFormula       : 'EUDR Status Formula',
  facilityProfile         : 'Facility Profile',
};

/** Tab name for BL Monitoring (must match spreadsheet tab exactly). */
const BL_MONITORING_TAB = 'BL Monitoring';

/** API sheet keys → spreadsheet tab names (includes aliases). */
const SHEET_TAB_ALIASES = {
  blMonitoring: BL_MONITORING_TAB,
  bl: BL_MONITORING_TAB,
  'BL Monitoring': BL_MONITORING_TAB,
  blReference: 'BL Reference',
  'BL Reference': 'BL Reference',
};

/**
 * Resolve API sheet key to spreadsheet tab name.
 * Supports SHEETS map, SHEET_TAB_ALIASES, or direct tab name if it exists.
 */
function resolveSheetTabName_(sheetKey) {
  const key = String(sheetKey || '').trim();
  if (!key) return '';
  if (SHEETS[key]) return SHEETS[key];
  if (SHEET_TAB_ALIASES[key]) return SHEET_TAB_ALIASES[key];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss && ss.getSheetByName(key)) return key;
  return '';
}

/** Physical sheet columns (matches Excel — two columns both named NATION). */
const BL_MONITORING_SHEET_HEADERS = [
  'RECORD TYPE',
  'NO',
  'TOTAL BL',
  'LOADING PORT',
  'NATION',
  'BUYER',
  'NATION',
  'BL NO.',
  'RECEIVED DATE',
  'SENT TO REVIEW (DATE)',
  'SEND TO EXIM (DATE)',
  'BL DATE',
  'VESSEL',
  'COMODITY',
  'VOLUME (TON)',
  'REQUEST TYPE',
  'STATUS',
  'SENT TO REQUESTER',
  'PERIOD',
  'COMMODITY SUPPLY',
  'TTM',
  'TTP',
];

const BL_MONITORING_JSON_HEADERS = ['TTM LINKS JSON', 'TTP LINKS JSON'];

const BL_MONITORING_HEADERS = BL_MONITORING_SHEET_HEADERS.concat(BL_MONITORING_JSON_HEADERS);

/** BL dropdown lists (commodity + buyer) — TYPE + NAME rows. */
const BL_REFERENCE_HEADERS = ['TYPE', 'NAME'];

const BL_REFERENCE_DEFAULTS = [
  ['Commodity', 'CPO'],
  ['Commodity', 'PK'],
  ['Commodity', 'RBDPO'],
  ['Commodity', 'RBDPOLEIN'],
  ['Commodity', 'RBDPS'],
  ['Commodity', 'PFAD'],
  ['Commodity', 'CPKO'],
  ['Commodity', 'PKE'],
  ['Buyer', 'COFCO'],
  ['Buyer', 'FIRST RESOURCES'],
  ['Buyer', 'ADM'],
  ['Buyer', '3F INDUSTRIES LIMITED'],
  ['Buyer', 'EMAMI'],
  ['Buyer', 'CARGILL'],
  ['Buyer', 'IOI'],
  ['Buyer', 'APICAL'],
  ['Buyer', 'CITY EDIBLE OIL.LTD'],
  ['Buyer', 'RAJ INDUSTRIES'],
  ['Buyer', 'SANTHOSHIMATHAA'],
  ['Buyer', 'HDPC BANK Ltd'],
  ['Buyer', 'STA GRUP'],
];

/** Columns appended to existing BL Monitoring tabs when missing. */
const BL_MONITORING_ENSURE_HEADERS = [
  'RECORD TYPE',
  'TOTAL BL',
  'SENT TO REQUESTER',
  'PERIOD',
  'COMMODITY SUPPLY',
].concat(BL_MONITORING_JSON_HEADERS);

// ── Supply Import Draft headers ──────────────────────────────
const SUPPLY_DRAFT_HEADERS = [
  'draft_id', 'batch_id', 'status',        // 'draft' | 'submitted'
  'month', 'year',
  'created_at', 'updated_at', 'created_by',
  'match_status',                           // 'matched' | 'new' | 'group_mismatch'
  'supply_type',                            // 'CPO' | 'PK'
  'target_mill_row',
  'SUPPLY_QTY',
  'SUPPLY CPO',
  'SUPPLY PK',
  'PERCENTAGE SUPPLY CPO',
  'PERCENTAGE SUPPLY PK',
  // Mirror of MILL_FIELDS subset that may be pre-filled or user-edited:
  'MONTH', 'YEAR', 'COMPANY CODE', 'TRADER NAME', 'GROUP NAME',
  'COMPANY NAME', 'MILL NAME', 'UML ID', 'ADDRESS', 'PROVINCE',
  'COORDINATES', 'MILL CATEGORY', 'MILL CAPACITY (TON/HOUR)',
  'HGU/HGB', 'IZIN LOKASI', 'IUP', 'IZIN LINGKUNGAN', 'SCORE',
  'MILL LOC', 'COMPLIMENT/NOT COMPLIMENT',
  'DEFORESTATION WIDTH', 'BURN AREA WIDTH', 'PEAT WIDTH', 'LEGALITY',
  'DEFORESTATION GRIEVANCES', 'BURN AREA GRIEVANCES',
  'HUMAN RIGHT', 'SAFETY', 'SOCIAL', 'ENVIRONMENT', 'TOTAL GRIEVANCES',
  'NDPE', 'HRDD', 'TOTAL POLICY',
  'CERTIFICATION', 'TOTAL CERTIFICATION', 'TOTAL SCORE',
  'SUPPLIER LEVEL', 'BUYER NO BUY LIST', 'VOLUME SUPPLY STATUS',
  'RECOMMENDATION LEVEL', 'PRIORITY ENGAGEMENT', 'SIGN', 'SUPPLIER STATUS', 'RISK LEVEL',
  'RESULT RISK LEVEL', 'FACILITY NAME CPO', 'FACILITY NAME PK',
  'PRODUCT SUPPLY', 'SUPPLY CPO', 'SUPPLY PK',
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
  'Email',
  'statusSDD',
  'approved_at',
  'updated_at',
  'updated_by',
];

/** Grievance Monitoring — risk score columns (auto-added to sheet row 1 when missing). */
const GRIEVANCE_RISK_HEADERS = [
  'Publish Grievance',
  'Subject Grievance',
  'Repeat Grievance',
  'Consequence',
  'Group Scale',
  'No Response',
  'Total Score',
  'Risk Classification',
];

/** Facility Profile / Company Profile List (tab: Facility Profile). */
const FACILITY_PROFILE_HEADERS = [
  'PLANT',
  'COMPANY NAME',
  'SITE NAME',
  'ADDRESS',
  'CAPACITY',
  'COORDINATE',
  'FACILITY',
  'CERTIFICATION',
];

const QUESTIONNAIRE_MONITORING_HEADERS = [
  'GROUP NAME',
  'COMPANY NAME',
  'MILL NAME',
  'UML ID',
  'STATUS',
  'PROGRESS',
  'LAST UPDATE',
  'DATE OPEN',
  'DATE SEND EMAIL',
  'DATE RECEIVED QUESTIONNAIRE',
  'DATE SEND BACK WITH FEEDBACK',
  'DATE COMPLETED',
  'UPDATED BY',
];

const EUDR_POTENTIAL_HEADERS = [
  'GROUP NAME',
  'COMPANY NAME',
  'MILL NAME',
  'UML ID',
  'PROVINCE',
  'SUPPLY TO',
  'MILL CAPACITY',
  'STATUS',
  'DEFORESTATION (AFTER 2020)',
  'COMPLETE QUESTIONNAIRE',
  'NDPE IRF DEFORESTATION',
  'NDPE IRF PEAT',
  'SATELLITE MONITORING',
  '3RD PARTY VERIFICATION NDPE',
  '3RD PARTY VERIFICATION TRACEABILITY',
  '3RD PARTY VERIFICATION DEFORESTATION FREE',
  'LAST UPDATE',
  'UPDATED BY',
];

const EUDR_STATUS_FORMULA_HEADERS = ['CRITERION KEY', 'ENABLED', 'LABEL', 'THRESHOLD', 'CONFIG'];

const EUDR_FFB_FORMULA_DEFAULT_CONFIG = {
  mode: 'combined',
  combined: {
    categories: ['own_estate', 'plasma'],
    operator: 'gte',
    threshold: 70,
  },
  categories: {
    own_estate: { enabled: true, operator: 'gte', threshold: 70 },
    external_estate: { enabled: false, operator: 'gte', threshold: 70 },
    plasma: { enabled: true, operator: 'gte', threshold: 70 },
    dealer: { enabled: false, operator: 'gte', threshold: 70 },
    cooperative: { enabled: false, operator: 'gte', threshold: 70 },
  },
};

const EUDR_STATUS_FORMULA_DEFAULTS = [
  ['legality', 'Yes', 'Legality = Complete (1)', '', ''],
  ['millCategory', 'Yes', 'Mill Category = Integrated', '', ''],
  ['ownPlasmaFfb', 'Yes', '% FFB by Category', '70', JSON.stringify(EUDR_FFB_FORMULA_DEFAULT_CONFIG)],
  ['resultRiskLevel', 'Yes', 'Result Risk Level = Low', '', ''],
  ['millLocation', 'Yes', 'Mill Location = APL', '', ''],
  ['certification', 'Yes', 'Certification (at least 1 certificate)', '', ''],
  ['grievance', 'Yes', 'Grievance = No', '', ''],
  ['ndpePolicy', 'Yes', 'NDPE Policy = Yes', '', ''],
  ['noBuyList', 'Yes', 'No Buy List = No', '', ''],
  ['deforestation', 'Yes', 'Deforestation (max Ha)', '10', ''],
];

/** Identity fields merged from existing row when omitted on upsert. */
const EUDR_IDENTITY_FIELDS = [
  'GROUP NAME',
  'COMPANY NAME',
  'MILL NAME',
  'UML ID',
  'PROVINCE',
  'SUPPLY TO',
  'MILL CAPACITY',
];

/** Sheet-only fields preserved when sync identity from Mill Onboarding. */
const EUDR_SHEET_PRESERVE_FIELDS = [
  'STATUS',
  'DEFORESTATION (AFTER 2020)',
  'COMPLETE QUESTIONNAIRE',
  'NDPE IRF DEFORESTATION',
  'NDPE IRF PEAT',
  'SATELLITE MONITORING',
  '3RD PARTY VERIFICATION NDPE',
  '3RD PARTY VERIFICATION TRACEABILITY',
  '3RD PARTY VERIFICATION DEFORESTATION FREE',
  'LAST UPDATE',
  'UPDATED BY',
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
    'mill_added', 'mill_added_lines',
  ],
  sddMill: [
    'submission_id', 'line_id', 'supplier_type',
    'created_at', 'updated_at', 'created_by', 'updated_by', 'is_deleted',
    'TML - Company Group Name', 'TML - Company Name', 'TML - Mill Name', 'TML - UML ID',
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

/** Identity-only TTP row when SDD has no FFB suppliers (header stub). */
const TTP_MILL_HEADER_LINE_ID = '__mill_header__';
/** Per-mill TTP row id prefix for TRADER (mirrors Mill Onboarding, one row per TML line). */
const TTP_TRADER_TML_LINE_PREFIX = 'trader_tml_';
/** Per-mill TTP header stub prefix for MILL/KCP (one TML line, FFB rows use real ffb line_id). */
const TTP_MILL_TML_LINE_PREFIX = 'mill_tml_';

function ttpTraderMirrorLineId_(tmlLineId) {
  return TTP_TRADER_TML_LINE_PREFIX + String(tmlLineId || '').trim();
}

function ttpMillTmlHeaderLineId_(tmlLineId) {
  return TTP_MILL_TML_LINE_PREFIX + String(tmlLineId || '').trim();
}

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
//  API SECURITY (set Script property API_SECRET — min 32 random chars)
// ═══════════════════════════════════════════════════════════

function getApiSecret_() {
  return String(PropertiesService.getScriptProperties().getProperty('API_SECRET') || '').trim();
}

function extractApiToken_(e, body) {
  if (e && e.parameter && e.parameter.token) return String(e.parameter.token).trim();
  if (body && body.token) return String(body.token).trim();
  return '';
}

function extractActor_(e, body) {
  if (body && body._actor) return String(body._actor).trim();
  if (e && e.parameter && e.parameter._actor) return String(e.parameter._actor).trim();
  try { return Session.getActiveUser().getEmail() || 'system'; } catch (err) { return 'system'; }
}

function assertApiAuth_(e, body) {
  const secret = getApiSecret_();
  // Belum set API_SECRET di Script Properties → mode kompatibel (dashboard tetap jalan).
  // Set API_SECRET + GAS_API_SECRET di Vercel untuk mengaktifkan proteksi penuh.
  if (!secret) return;
  const token = extractApiToken_(e, body);
  if (!token || token !== secret) {
    throw new Error('Unauthorized');
  }
}

function assertRateLimit_(actor) {
  const key = 'rl:' + String(actor || 'anon').slice(0, 120);
  const cache = CacheService.getScriptCache();
  const raw = cache.get(key);
  const count = raw ? parseInt(raw, 10) + 1 : 1;
  if (count > 150) throw new Error('Rate limit exceeded — try again in a minute');
  cache.put(key, String(count), 60);
}

var MRD_WRITE_ACTIONS_ = {
  add: 1, update: 1, delete: 1, bulkDelete: 1, bulkUpsertSDD: 1,
  insertSDD: 1, updateSDD: 1, upsertSDD: 1,
  createSubmission: 1, updateSubmission: 1, setSubmissionStatus: 1, deleteSubmission: 1,
  saveSupplyDraft: 1, submitSupplyDraft: 1, deleteSupplyDraft: 1,
  addTtpBatch: 1, upsertQuestionnaire: 1, syncEudrPotential: 1, upsertEudr: 1,
  saveEudrStatusFormula: 1,
};

function auditLog_(verb, action, sheetKey, actor, detail) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;
    const tab = 'AUDIT_LOG';
    let sh = ss.getSheetByName(tab);
    if (!sh) {
      sh = ss.insertSheet(tab);
      sh.appendRow(['timestamp', 'verb', 'action', 'sheet', 'actor', 'detail']);
      sh.setFrozenRows(1);
    }
    const detailStr = String(detail || '');
    sh.appendRow([
      nowIso_(),
      String(verb || ''),
      String(action || ''),
      String(sheetKey || ''),
      String(actor || 'system'),
      detailStr.length > 800 ? detailStr.slice(0, 800) : detailStr,
    ]);
  } catch (err) {
    console.warn('[auditLog_]', err.message);
  }
}

// ═══════════════════════════════════════════════════════════
//  ENTRY POINTS
// ═══════════════════════════════════════════════════════════

function doGet(e) {
  try {
    assertApiAuth_(e, null);
    const actor = extractActor_(e, null);
    assertRateLimit_(actor);

    const action   = (e && e.parameter && e.parameter.action)  || '';
    const sheetKey = (e && e.parameter && e.parameter.sheet)   || '';

    if (action === 'getAll') {
      if (sheetKey === 'ttp') ensureTtpHeaders_();
      if (sheetKey === 'contactSupplier') ensureContactSupplierHeaders_();
      if (sheetKey === 'nbl') ensureNblHeaders_();
      if (sheetKey === 'unileverNbl') ensureUnileverNblHeaders_();
      if (sheetKey === 'supplyDraft') ensureSupplyDraftHeaders_();
      if (sheetKey === 'blMonitoring') ensureBlMonitoringHeaders_();
      if (sheetKey === 'blReference') ensureBlReferenceHeaders_();
      if (sheetKey === 'questionnaireMonitoring') ensureQuestionnaireMonitoringHeaders_();
      if (sheetKey === 'eudrPotential') ensureEudrPotentialHeaders_();
      if (sheetKey === 'eudrStatusFormula') ensureEudrStatusFormulaHeaders_();
      if (sheetKey === 'facilityProfile') ensureFacilityProfileHeaders_();
      if (sheetKey === 'grievance') ensureGrievanceRiskHeaders_();
      return respond(getData(sheetKey));
    }
    if (action === 'getByMillId')          return respond(getByMillId(e.parameter.millId));
    if (action === 'ping') {
      var grvHdr = false;
      try {
        grvHdr = !!ensureGrievanceRiskHeaders_();
      } catch (grvErr) {
        grvHdr = false;
      }
      return respond({
        success: true,
        message: 'Apps Script is alive',
        version: 'v3-ttp-header-row3',
        blMonitoring: !!resolveSheetTabName_('blMonitoring'),
        questionnaireMonitoring: !!resolveSheetTabName_('questionnaireMonitoring'),
        eudrPotential: !!resolveSheetTabName_('eudrPotential'),
        eudrStatusFormula: !!resolveSheetTabName_('eudrStatusFormula'),
        grievanceRiskHeaders: grvHdr,
      });
    }
    if (action === 'getSubmissionById')    return respond(getSubmissionById(e.parameter.submission_id));
    if (action === 'listSubmissions')      return respond(listSubmissions(e.parameter));
    if (action === 'listSuppliedCpoSheets') return respond(listSuppliedCpoSheets_());
    if (action === 'getSuppliedCpo')       return respond(getSuppliedCpoData_(e.parameter.sheet || ''));
    if (action === 'listSuppliedPkSheets') return respond(listSuppliedPkSheets_());
    if (action === 'getSuppliedPk')        return respond(getSuppliedPkData_(e.parameter.sheet || ''));
    if (action === 'getFacilityMapImage')  return respond(getFacilityMapImage_(e.parameter || {}));
    if (action === 'seedBlReference')      return respond(seedBlReferenceDefaults_());

    return respond({ success: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return respond({ success: false, error: err.message });
  }
}

function doPost(e) {
  try {
    const body     = parsePostBody_(e);
    assertApiAuth_(e, body);
    const actor = extractActor_(e, body);
    assertRateLimit_(actor);

    const action   = body.action   || '';
    const sheetKey = body.sheet    || '';

    if (MRD_WRITE_ACTIONS_[action]) {
      auditLog_('POST', action, sheetKey, actor, JSON.stringify({
        row: body.row,
        rows_count: Array.isArray(body.rows) ? body.rows.length : undefined,
        submission_id: (body.payload && body.payload.submission_id) || body.submission_id,
      }));
    }

    if (sheetKey === 'blMonitoring') ensureBlMonitoringHeaders_();
    if (sheetKey === 'blReference') ensureBlReferenceHeaders_();
    if (sheetKey === 'questionnaireMonitoring') ensureQuestionnaireMonitoringHeaders_();
    if (sheetKey === 'eudrPotential') ensureEudrPotentialHeaders_();
    if (sheetKey === 'facilityProfile') ensureFacilityProfileHeaders_();
    if (sheetKey === 'grievance') ensureGrievanceRiskHeaders_();
    if (action === 'upsertQuestionnaire') return respond(upsertQuestionnaireRow_(body.data || {}));
    if (action === 'syncEudrPotential') return respond(syncEudrPotentialRows_(body.mills || []));
    if (action === 'upsertEudr') return respond(upsertEudrPotentialRow_(body.data || {}));
    if (action === 'saveEudrStatusFormula') return respond(saveEudrStatusFormula_(body.criteria || []));
    if (action === 'addTtpBatch') return respond(addTtpBatch_(body.rows || []));
    if (action === 'upsertBlReference') return respond(upsertBlReferenceItem_(body.type, body.name));
    if (action === 'seedBlReference') return respond(seedBlReferenceDefaults_());
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

/**
 * safeInsertRowAt_ — write one row at a fixed sheet row, then patch coord columns as Text.
 */
function safeInsertRowAt_(ws, headers, rowArr, targetRow) {
  ws.getRange(targetRow, 1, 1, rowArr.length).setValues([rowArr]);
  COORD_COLUMN_NAMES.forEach(function(name) {
    const colIdx = headers.indexOf(name);
    if (colIdx < 0) return;
    const v = rowArr[colIdx];
    if (v === null || v === undefined) return;
    const s = String(v).trim();
    if (!s) return;
    const cell = ws.getRange(targetRow, colIdx + 1);
    cell.setNumberFormat('@');
    cell.setValue(s);
  });
}

var MILL_IDENTITY_HEADERS_ = [
  'MILL NAME', 'COMPANY NAME', 'GROUP NAME', 'UML ID', 'COMPANY CODE'
];

var MILL_IDENTITY_PLACEHOLDERS_ = { 'NO DATA': true, '-': true, '—': true, 'N/A': true, 'NA': true };

/** Kolom formula/default (SUPPLY %, dll.) tidak menentukan zona aktif vs kosong. */
var MILL_NON_IDENTITY_HEADERS_ = {
  'MONTH': true, 'QUARTER': true, 'YEAR': true, 'SOURCE TYPE': true, 'TRADER NAME': true,
  'SUPPLY CPO': true, 'SUPPLY PK': true, 'PRODUCT SUPPLY': true,
  'PERCENTAGE SUPPLY CPO': true, 'PERCENTAGE SUPPLY PK': true,
  'PERCENTAGE': true, 'SUPPLY': true, 'SCORE': true, 'TOTAL SCORE': true,
  'PRIORITY ENGAGEMENT': true, 'SIGN': true, 'DATE IMPORTED': true, 'IMPORTED BY': true,
};

function millIdentityColumnIndices_(headers) {
  var out = [];
  var seen = {};
  (headers || []).forEach(function(h, i) {
    var key = String(h || '').trim().toUpperCase();
    if (!key || seen[key] || MILL_NON_IDENTITY_HEADERS_[key]) return;
    for (var k = 0; k < MILL_IDENTITY_HEADERS_.length; k++) {
      if (MILL_IDENTITY_HEADERS_[k] === key) {
        seen[key] = true;
        out.push(i);
        break;
      }
    }
  });
  return out.length ? out : null;
}

function millIsPlaceholderIdentity_(v) {
  var t = String(v === null || v === undefined ? '' : v).trim().toUpperCase();
  return !t || MILL_IDENTITY_PLACEHOLDERS_[t] === true;
}

/** Baris dianggap berisi mill hanya jika kolom identitas terisi (bukan SUPPLY % / formula). */
function millRowHasIdentityContent_(rowValues, identityColIndices) {
  if (!rowValues || !rowValues.length) return false;
  if (identityColIndices && identityColIndices.length) {
    for (var i = 0; i < identityColIndices.length; i++) {
      if (!millIsPlaceholderIdentity_(rowValues[identityColIndices[i]])) return true;
    }
    return false;
  }
  return false;
}

function millKeyColumnIndices_(headers) {
  return millIdentityColumnIndices_(headers);
}

function millRowValuesHaveContent_(rowValues, keyColIndices) {
  return millRowHasIdentityContent_(rowValues, keyColIndices);
}

/**
 * Active zone vs archive zone on Mill sheet:
 *   - Active (atas): baris dengan MILL/COMPANY/GROUP/UML terisi + slot kosong berikutnya
 *   - Archive (bawah): blok copy/manual jauh di bawah — tidak disentuh insert web
 *
 * Baris yang hanya punya formula SUPPLY % / PERCENTAGE (0%, 0.0) dianggap KOSONG
 * untuk penentuan row insert — supaya web masuk ke row 258, bukan row ujung sheet.
 */
var MILL_ARCHIVE_GAP_ROWS = 10;

function millFindActiveZoneLastRow_(sheet, headers) {
  var headerRow = 1;
  var lastRow = sheet.getLastRow();
  if (lastRow <= headerRow) return headerRow;

  var numCols = Math.max(headers.length, sheet.getLastColumn());
  var rowCount = lastRow - headerRow;
  if (rowCount <= 0) return headerRow;

  var allData = sheet.getRange(headerRow + 1, 1, rowCount, numCols).getValues();
  var idCols = millIdentityColumnIndices_(headers);
  var lastDataRow = headerRow;
  var emptyStreak = 0;

  for (var i = 0; i < allData.length; i++) {
    if (millRowHasIdentityContent_(allData[i], idCols)) {
      lastDataRow = headerRow + 1 + i;
      emptyStreak = 0;
    } else {
      emptyStreak++;
      if (emptyStreak >= MILL_ARCHIVE_GAP_ROWS && lastDataRow > headerRow) break;
    }
  }
  return lastDataRow;
}

function millFindNextAppendRow_(sheet, headers) {
  return millFindActiveZoneLastRow_(sheet, headers) + 1;
}

var TTP_IDENTITY_HEADERS_ = [
  'VILLAGE', 'FFB SUPPLIER NAME', 'FFB SUPPLIER GROUP NAME', 'CATEGORY',
  'UML ID', 'MILL NAME', 'COMPANY NAME', 'GROUP NAME', 'COMPANY CODE'
];

var TTP_NON_IDENTITY_HEADERS_ = {
  'QUARTER': true, 'YEAR': true, 'LAT': true, 'LONG': true, 'LATITUDE': true, 'LONGITUDE': true,
  'PERCENTAGE TRACEABILITY': true, '% CPO TRACEABLE': true, '% PK TRACEABLE': true,
  'PERCENTAGE': true, 'SUPPLY': true, 'FFB SUPPLY TO MILL (TON)': true,
  'TOTAL SUPPLY FFB (TON)': true, 'ISPO (Y/N)': true, 'RSPO (Y/N)': true, 'ISCC (Y/N)': true,
};

function ttpIdentityColumnIndices_(headers) {
  var out = [];
  var seen = {};
  (headers || []).forEach(function(h, i) {
    var key = String(h || '').trim().toUpperCase();
    if (!key || seen[key] || TTP_NON_IDENTITY_HEADERS_[key]) return;
    for (var k = 0; k < TTP_IDENTITY_HEADERS_.length; k++) {
      if (TTP_IDENTITY_HEADERS_[k] === key) {
        seen[key] = true;
        out.push(i);
        break;
      }
    }
  });
  return out.length ? out : null;
}

function ttpFindActiveZoneLastRow_(sheet, headers, headerRow) {
  headerRow = headerRow || 1;
  var lastRow = sheet.getLastRow();
  if (lastRow <= headerRow) return headerRow;

  var numCols = Math.max(headers.length, sheet.getLastColumn());
  var rowCount = lastRow - headerRow;
  if (rowCount <= 0) return headerRow;

  var allData = sheet.getRange(headerRow + 1, 1, rowCount, numCols).getValues();
  var idCols = ttpIdentityColumnIndices_(headers);
  var lastDataRow = headerRow;
  var emptyStreak = 0;

  for (var i = 0; i < allData.length; i++) {
    if (millRowHasIdentityContent_(allData[i], idCols)) {
      lastDataRow = headerRow + 1 + i;
      emptyStreak = 0;
    } else {
      emptyStreak++;
      if (emptyStreak >= MILL_ARCHIVE_GAP_ROWS && lastDataRow > headerRow) break;
    }
  }
  return lastDataRow;
}

function ttpFindNextAppendRow_(sheet, headers, headerRow) {
  return ttpFindActiveZoneLastRow_(sheet, headers, headerRow) + 1;
}

/** Ringkasan zona aktif vs archive copy (jalankan dari GAS editor). */
function millSheetTailReport_() {
  var sheet = getSheet('mill');
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var lastRow = sheet.getLastRow();
  var nextWebInsert = millFindNextAppendRow_(sheet, headers);
  var activeLastRow = Math.max(1, nextWebInsert - 1);
  var keyCols = millKeyColumnIndices_(headers);
  var numCols = Math.max(headers.length, sheet.getLastColumn());
  var archiveRows = 0;
  var emptyGapRows = 0;

  for (var r = nextWebInsert; r <= lastRow; r++) {
    var vals = sheet.getRange(r, 1, 1, numCols).getValues()[0];
    if (millRowValuesHaveContent_(vals, keyCols)) archiveRows++;
    else emptyGapRows++;
  }

  return {
    sheetLastRow: lastRow,
    activeZoneLastRow: activeLastRow,
    nextWebInsertRow: nextWebInsert,
    emptyGapRows: emptyGapRows,
    archiveCopyRows: archiveRows,
    hint: archiveRows > 0
      ? 'Archive copy di bawah aman. Input web masuk ke row ' + nextWebInsert + '.'
      : 'Tidak ada archive copy terpisah; input web masuk ke row ' + nextWebInsert + '.',
  };
}

/**
 * Hapus HANYA baris kosong di celah antara zona aktif dan archive copy.
 * Baris berisi data mill (archive copy) tidak dihapus.
 * trimMillSheetEmptyGapRows_(true)  — dry-run
 * trimMillSheetEmptyGapRows_(false) — eksekusi
 */
function trimMillSheetEmptyGapRows_(dryRun) {
  var sheet = getSheet('mill');
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var lastRow = sheet.getLastRow();
  var fromRow = millFindNextAppendRow_(sheet, headers);
  if (fromRow > lastRow) {
    return { ok: true, message: 'Nothing to trim', nextWebInsertRow: fromRow };
  }

  var numCols = Math.max(headers.length, sheet.getLastColumn());
  var keyCols = millKeyColumnIndices_(headers);
  var emptyRows = [];
  var preservedRows = 0;

  for (var r = fromRow; r <= lastRow; r++) {
    var vals = sheet.getRange(r, 1, 1, numCols).getValues()[0];
    if (millRowValuesHaveContent_(vals, keyCols)) preservedRows++;
    else emptyRows.push(r);
  }

  if (dryRun !== false) {
    return {
      ok: true, dryRun: true,
      wouldDeleteEmptyRows: emptyRows.length,
      preservedArchiveRows: preservedRows,
      fromRow: fromRow, toRow: lastRow,
    };
  }

  for (var i = emptyRows.length - 1; i >= 0; i--) {
    sheet.deleteRow(emptyRows[i]);
  }
  return {
    ok: true, dryRun: false,
    deletedEmptyRows: emptyRows.length,
    preservedArchiveRows: preservedRows,
    newLastRow: sheet.getLastRow(),
  };
}

/** @deprecated Pakai trimMillSheetEmptyGapRows_ — jangan hapus archive copy. */
function trimMillSheetOrphanTail_(dryRun) {
  return trimMillSheetEmptyGapRows_(dryRun);
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

function ensureFacilityProfileHeaders_() {
  return ensureSheetHeadersGeneric_('facilityProfile', FACILITY_PROFILE_HEADERS);
}

function facilityProfileRowHasContent_(obj) {
  if (!obj) return false;
  return !!(
    String(obj.PLANT || '').trim() ||
    String(obj['COMPANY NAME'] || '').trim() ||
    String(obj['SITE NAME'] || '').trim()
  );
}

function ttpRowHasContent_(obj) {
  if (!obj) return false;
  if (String(obj['COMPANY NAME'] || obj['Company Name'] || '').trim()) return true;
  if (String(obj['MILL NAME'] || '').trim()) return true;
  if (String(obj['FFB SUPPLIER NAME'] || '').trim()) return true;
  if (String(obj['GROUP NAME'] || '').trim()) return true;
  return false;
}

function millRowHasContent_(obj) {
  if (!obj) return false;
  var co = String(obj['COMPANY NAME'] || obj['Company Name'] || '').trim();
  if (!co) {
    Object.keys(obj).forEach(function(k) {
      if (k === '_row' || co) return;
      var nk = String(k).replace(/\s+/g, ' ').trim();
      if (!/^company(\s*name)?$/i.test(nk)) return;
      if (/group|trader|code|facility|supply|nbl|profile|owner|tml|ffb/i.test(nk)) return;
      var v = String(obj[k] || '').trim();
      if (v && v !== '—' && v !== '-') co = v;
    });
  }
  if (co && co !== '—' && co !== '-') return true;
  return !!(
    String(obj['MILL NAME'] || '').trim() ||
    String(obj['GROUP NAME'] || '').trim() ||
    String(obj['UML ID'] || '').trim() ||
    String(obj['COMPANY CODE'] || '').trim()
  );
}

function ensureQuestionnaireMonitoringHeaders_() {
  return ensureSheetHeadersGeneric_('questionnaireMonitoring', QUESTIONNAIRE_MONITORING_HEADERS);
}

function ensureEudrPotentialHeaders_() {
  return ensureSheetHeadersGeneric_('eudrPotential', EUDR_POTENTIAL_HEADERS);
}

function ensureEudrStatusFormulaHeaders_() {
  const sheet = ensureSheetHeadersGeneric_('eudrStatusFormula', EUDR_STATUS_FORMULA_HEADERS);
  const lastCol = sheet.getLastColumn();
  const existing = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function(h) { return String(h || '').trim(); });
  const existingSet = {};
  existing.forEach(function(h) { if (h) existingSet[h] = true; });
  const missing = EUDR_STATUS_FORMULA_HEADERS.filter(function(h) { return !existingSet[h]; });
  if (missing.length) {
    const startCol = lastCol + 1;
    sheet.getRange(1, startCol, 1, missing.length).setValues([missing]);
  }
  if (sheet.getLastRow() <= 1) {
    const rows = EUDR_STATUS_FORMULA_DEFAULTS.map(function(r) {
      return EUDR_STATUS_FORMULA_HEADERS.map(function(_, idx) {
        return r[idx] != null ? String(r[idx]) : '';
      });
    });
    if (rows.length) {
      sheet.getRange(2, 1, rows.length, EUDR_STATUS_FORMULA_HEADERS.length).setValues(rows);
    }
  }
  return sheet;
}

function saveEudrStatusFormula_(criteria) {
  ensureEudrStatusFormulaHeaders_();
  const sheet = getSheet('eudrStatusFormula');
  const headers = EUDR_STATUS_FORMULA_HEADERS;
  const labelMap = {};
  const defaultThresholdMap = {};
  EUDR_STATUS_FORMULA_DEFAULTS.forEach(function(r) {
    labelMap[r[0]] = r[2];
    defaultThresholdMap[r[0]] = r[3] != null ? String(r[3]) : '';
  });
  const enabledMap = {};
  const thresholdMap = {};
  const configMap = {};
  (criteria || []).forEach(function(c) {
    if (!c || !c.key) return;
    var key = String(c.key).trim();
    if (key === 'cpoTraceable') key = 'ownPlasmaFfb';
    enabledMap[key] = c.enabled !== false && String(c.enabled).toLowerCase() !== 'no';
    if (c.threshold !== undefined && c.threshold !== null && String(c.threshold).trim() !== '') {
      thresholdMap[key] = String(c.threshold).trim();
    }
    if (c.config !== undefined && c.config !== null) {
      configMap[key] = typeof c.config === 'string' ? c.config : JSON.stringify(c.config);
    }
  });
  const rows = EUDR_STATUS_FORMULA_DEFAULTS.map(function(def) {
    const key = def[0];
    const enabled = enabledMap.hasOwnProperty(key) ? enabledMap[key] : true;
    const thr = thresholdMap.hasOwnProperty(key)
      ? thresholdMap[key]
      : (defaultThresholdMap[key] || '');
    const cfg = configMap.hasOwnProperty(key)
      ? configMap[key]
      : (def[4] != null ? String(def[4]) : '');
    return [key, enabled ? 'Yes' : 'No', labelMap[key] || def[2], thr, cfg];
  });
  const last = Math.max(sheet.getLastRow(), 1);
  if (last > 1) sheet.deleteRows(2, last - 1);
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  return { success: true, saved: rows.length };
}

function eudrEntityKey_(group, company, mill) {
  return [group, company, mill].map(function(s) {
    return String(s || '').trim().toLowerCase();
  }).join('|');
}

function findEudrRowByEntity_(rows, data) {
  const key = eudrEntityKey_(data['GROUP NAME'], data['COMPANY NAME'], data['MILL NAME']);
  for (var i = 0; i < (rows || []).length; i++) {
    var r = rows[i];
    if (eudrEntityKey_(r['GROUP NAME'], r['COMPANY NAME'], r['MILL NAME']) === key) return r;
  }
  return null;
}

function eudrIdentityPatch_(m) {
  m = m || {};
  return {
    'GROUP NAME': String(m['GROUP NAME'] || '').trim(),
    'COMPANY NAME': String(m['COMPANY NAME'] || '').trim(),
    'MILL NAME': String(m['MILL NAME'] || '').trim(),
    'UML ID': String(m['UML ID'] || '').trim(),
    'PROVINCE': String(m['PROVINCE'] || '').trim(),
    'SUPPLY TO': String(m['SUPPLY TO'] || '').trim(),
    'MILL CAPACITY': String(m['MILL CAPACITY'] || '').trim(),
  };
}

function syncEudrPotentialRows_(mills) {
  ensureEudrPotentialHeaders_();
  const sheet = getSheet('eudrPotential');
  const headers = EUDR_POTENTIAL_HEADERS;
  const rows = getData('eudrPotential');
  const existing = {};
  (rows || []).forEach(function(r) {
    existing[eudrEntityKey_(r['GROUP NAME'], r['COMPANY NAME'], r['MILL NAME'])] = r;
  });
  var inserted = 0;
  var updated = 0;
  var appendRows = [];
  var updateOps = [];

  (mills || []).forEach(function(m) {
    const patch = eudrIdentityPatch_(m);
    if (!patch['MILL NAME'] && !patch['COMPANY NAME']) return;
    const key = eudrEntityKey_(patch['GROUP NAME'], patch['COMPANY NAME'], patch['MILL NAME']);
    const hit = existing[key];
    if (hit && hit._row) {
      const merged = {};
      headers.forEach(function(h) {
        if (EUDR_SHEET_PRESERVE_FIELDS.indexOf(h) !== -1) {
          merged[h] = hit[h] != null ? hit[h] : '';
        } else if (patch[h] !== undefined && String(patch[h]).trim() !== '') {
          merged[h] = patch[h];
        } else {
          merged[h] = hit[h] != null ? hit[h] : '';
        }
      });
      updateOps.push({
        row: hit._row,
        values: headers.map(function(h) { return merged[h] !== undefined ? merged[h] : ''; }),
      });
      updated++;
    } else {
      patch['STATUS'] = '';
      appendRows.push(headers.map(function(h) { return patch[h] !== undefined ? patch[h] : ''; }));
      inserted++;
    }
  });

  if (appendRows.length) {
    const startRow = Math.max(sheet.getLastRow(), 1) + 1;
    sheet.getRange(startRow, 1, appendRows.length, headers.length).setValues(appendRows);
  }

  updateOps.forEach(function(op) {
    sheet.getRange(op.row, 1, 1, headers.length).setValues([op.values]);
  });

  return { success: true, inserted: inserted, updated: updated };
}

function stampEudrRowMeta_(data, user, now) {
  data = data || {};
  now = now || nowIso_();
  user = user || callerEmail_();
  data['LAST UPDATE'] = now.slice(0, 10);
  data['UPDATED BY'] = user;
  return data;
}

function upsertEudrPotentialRow_(data) {
  ensureEudrPotentialHeaders_();
  const user = callerEmail_();
  const now = nowIso_();
  const rows = getData('eudrPotential');
  const hit = findEudrRowByEntity_(rows, data || {});
  var patch = Object.assign({}, data || {});
  if (hit) {
    EUDR_IDENTITY_FIELDS.forEach(function(h) {
      if (patch[h] === undefined || patch[h] === null || String(patch[h]).trim() === '') {
        if (hit[h] != null && String(hit[h]).trim() !== '') patch[h] = hit[h];
      }
    });
  } else {
    EUDR_SHEET_PRESERVE_FIELDS.forEach(function(h) {
      if (patch[h] === undefined || patch[h] === null) patch[h] = '';
    });
  }
  patch = stampEudrRowMeta_(patch, user, now);
  if (hit && hit._row) {
    updateRow('eudrPotential', hit._row, patch);
    return { success: true, mode: 'update', row: hit._row };
  }
  addRow('eudrPotential', patch);
  return { success: true, mode: 'insert' };
}

function qmEntityKey_(group, company, mill) {
  return [group, company, mill].map(function(s) {
    return String(s || '').trim().toLowerCase();
  }).join('|');
}

function findQmRowByEntity_(rows, data) {
  const key = qmEntityKey_(data['GROUP NAME'], data['COMPANY NAME'], data['MILL NAME']);
  for (var i = 0; i < (rows || []).length; i++) {
    var r = rows[i];
    if (qmEntityKey_(r['GROUP NAME'], r['COMPANY NAME'], r['MILL NAME']) === key) return r;
  }
  return null;
}

function stampQmRowMeta_(data, user, now) {
  data = data || {};
  now = now || nowIso_();
  user = user || callerEmail_();
  data['LAST UPDATE'] = now.slice(0, 10);
  data['UPDATED BY'] = user;
  if (!String(data['STATUS'] || '').trim()) data['STATUS'] = 'On Progress';
  if (!String(data['PROGRESS'] || '').trim()) data['PROGRESS'] = 'Open';
  return data;
}

function qmProgressDateMap_() {
  return {
    'Open': 'DATE OPEN',
    'Send Email': 'DATE SEND EMAIL',
    'Received Questionnaire': 'DATE RECEIVED QUESTIONNAIRE',
    'Send Back With Feedback': 'DATE SEND BACK WITH FEEDBACK',
    'Completed': 'DATE COMPLETED',
  };
}

function qmApplyProgressDates_(data, now, prevRow) {
  data = data || {};
  now = now || nowIso_();
  const map = qmProgressDateMap_();
  const order = ['Open', 'Send Email', 'Received Questionnaire', 'Send Back With Feedback', 'Completed'];
  const idx = order.indexOf(String(data['PROGRESS'] || '').trim());
  if (idx < 0) return data;
  const today = now.slice(0, 10);
  for (var i = 0; i <= idx; i++) {
    const col = map[order[i]];
    const incoming = String(data[col] != null ? data[col] : '').trim();
    if (incoming) continue;
    const prev = prevRow ? String(prevRow[col] != null ? prevRow[col] : '').trim() : '';
    data[col] = prev || today;
  }
  return data;
}

function upsertQuestionnaireRow_(data) {
  ensureQuestionnaireMonitoringHeaders_();
  const user = callerEmail_();
  const now = nowIso_();
  const rows = getData('questionnaireMonitoring');
  const hit = findQmRowByEntity_(rows, data || {});
  var patch = Object.assign({}, data || {});
  if (hit) {
    QUESTIONNAIRE_MONITORING_HEADERS.forEach(function(h) {
      if (patch[h] === undefined || patch[h] === null || String(patch[h]).trim() === '') {
        if (hit[h] != null && String(hit[h]).trim() !== '') patch[h] = hit[h];
      }
    });
  }
  patch = stampQmRowMeta_(patch, user, now);
  patch = qmApplyProgressDates_(patch, now, hit);
  if (hit && hit._row) {
    updateRow('questionnaireMonitoring', hit._row, patch);
    return { success: true, mode: 'update', row: hit._row };
  }
  addRow('questionnaireMonitoring', patch);
  return { success: true, mode: 'insert' };
}

function normalizeNblLegacyHeaders_(ws) {
  if (!ws) return ws;
  const lastCol = Math.max(ws.getLastColumn(), 1);
  const headers = ws.getRange(1, 1, 1, lastCol).getValues()[0];
  const renameMap = {
    'raiser': 'Riser',
    'riser': 'Riser',
    'group name': 'Group Name NBL',
    'group name nbl': 'Group Name NBL',
    'company name': 'Company Name NBL',
    'company name nbl': 'Company Name NBL',
    'source': 'SOURCE',
  };
  let changed = false;
  const out = headers.map(function(h) {
    const trimmed = String(h || '').replace(/\s+/g, ' ').trim();
    const canon = renameMap[trimmed.toLowerCase()];
    if (canon && trimmed !== canon) {
      changed = true;
      return canon;
    }
    return trimmed;
  });
  if (changed) {
    ws.getRange(1, 1, 1, out.length).setValues([out]);
  }
  return ws;
}

function ensureNblHeaders_() {
  const ws = ensureSheetHeadersGeneric_('nbl', NBL_HEADERS);
  return normalizeNblLegacyHeaders_(ws);
}

function ensureUnileverNblHeaders_() {
  return ensureSheetHeadersGeneric_('unileverNbl', UNILEVER_NBL_HEADERS);
}

function ensureBlReferenceHeaders_() {
  const sheet = ensureSheetHeadersGeneric_('blReference', BL_REFERENCE_HEADERS);
  if (sheet.getLastRow() <= 1 && BL_REFERENCE_DEFAULTS.length) {
    const startRow = 2;
    const numRows = BL_REFERENCE_DEFAULTS.length;
    const numCols = BL_REFERENCE_HEADERS.length;
    sheet.getRange(startRow, 1, numRows, numCols).setValues(BL_REFERENCE_DEFAULTS);
  }
  return sheet;
}

function seedBlReferenceDefaults_() {
  ensureBlReferenceHeaders_();
  return { success: true, count: getData('blReference').length };
}

function upsertBlReferenceItem_(type, name) {
  const t = String(type || '').trim();
  const n = String(name || '').trim();
  if (!t || !n) throw new Error('TYPE and NAME are required for BL Reference.');
  const sheet = ensureBlReferenceHeaders_();
  const rows = sheet.getDataRange().getValues();
  if (!rows.length) throw new Error('BL Reference sheet has no headers.');
  const headers = rows[0].map(function(h) { return String(h || '').trim(); });
  const typeCol = headers.indexOf('TYPE');
  const nameCol = headers.indexOf('NAME');
  if (typeCol < 0 || nameCol < 0) throw new Error('BL Reference headers must include TYPE and NAME.');
  const tLower = t.toLowerCase();
  const nLower = n.toLowerCase();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][typeCol] || '').trim().toLowerCase() === tLower
        && String(rows[i][nameCol] || '').trim().toLowerCase() === nLower) {
      return { success: true, existed: true };
    }
  }
  sheet.appendRow([t, n]);
  return { success: true, added: true };
}

function detectBlHeaderRow_(sheet) {
  return detectHeaderRowByKeys_(
    sheet,
    ['LOADING PORT', 'BL NO.', 'VESSEL', 'REQUEST TYPE'],
    25
  );
}

function getBlHeaderInfo_(sheet) {
  const hdr = detectBlHeaderRow_(sheet);
  if (hdr && hdr.headers && hdr.headers.length) return hdr;
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function(h) { return String(h || '').trim(); });
  return { headerRow: 1, headers: headers };
}

/** Map sheet row (duplicate NATION headers) → canonical API keys. */
function blCanonicalizeRow_(headers, sourceRow) {
  const obj = {};
  let nationIdx = 0;
  headers.forEach(function(h, j) {
    const key = String(h || '').trim();
    if (!key) return;
    if (/^nation$/i.test(key)) {
      nationIdx++;
      obj[nationIdx === 1 ? 'NATION' : 'NATION (BUYER)'] = sourceRow[j];
      return;
    }
    obj[key] = sourceRow[j];
  });
  if (obj['NO'] != null && String(obj['NO']).trim() !== '' && (!obj['TOTAL BL'] || String(obj['TOTAL BL']).trim() === '')) {
    obj['TOTAL BL'] = obj['NO'];
  }
  if (!obj['RECORD TYPE'] || String(obj['RECORD TYPE']).trim() === '') {
    obj['RECORD TYPE'] = 'Shipping';
  }
  return obj;
}

/** Map canonical API payload → physical sheet row (handles duplicate NATION). */
function blExpandCanonicalToRow_(headers, data, current) {
  let nationIdx = 0;
  return headers.map(function(h, j) {
    const key = String(h || '').trim();
    if (/^nation$/i.test(key)) {
      nationIdx++;
      const canonKey = nationIdx === 1 ? 'NATION' : 'NATION (BUYER)';
      if (data && data[canonKey] !== undefined) return data[canonKey];
      if (data && nationIdx === 1 && data['LOADING NATION'] !== undefined) return data['LOADING NATION'];
      if (data && nationIdx === 2 && data['BUYER NATION'] !== undefined) return data['BUYER NATION'];
      return current ? current[j] : '';
    }
    if (/^no$/i.test(key) && data && data['TOTAL BL'] !== undefined) return data['TOTAL BL'];
    if (data && data[key] !== undefined) return data[key];
    return current ? current[j] : '';
  });
}

function blRowHasContent_(obj) {
  if (!obj || typeof obj !== 'object') return false;
  return Object.keys(obj).some(function(k) {
    if (k === '_row') return false;
    return String(obj[k] != null ? obj[k] : '').trim() !== '';
  });
}

function ensureBlMonitoringHeaders_() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const name = resolveSheetTabName_('blMonitoring') || BL_MONITORING_TAB;
  let ws     = ss.getSheetByName(name);
  if (!ws) ws = ss.insertSheet(name);

  const hdr = detectBlHeaderRow_(ws);
  if (!hdr) {
    ws.getRange(1, 1, 1, BL_MONITORING_HEADERS.length).setValues([BL_MONITORING_HEADERS]);
    return ws;
  }

  const headers = hdr.headers.slice();
  const headerRow = hdr.headerRow;
  BL_MONITORING_ENSURE_HEADERS.forEach(function(col) {
    if (headers.indexOf(col) !== -1) return;
    const newCol = headers.length + 1;
    ws.getRange(headerRow, newCol).setValue(col);
    headers.push(col);
  });
  return ws;
}

function normalizeSddDecisionLabel_(raw) {
  const s = String(raw || '').trim().toLowerCase();
  if (s === 'approve' || s === 'approved') return 'APPROVED';
  if (s === 'hold' || s === 'on hold') return 'ON HOLD';
  if (s === 'reject' || s === 'rejected') return 'REJECTED';
  return String(raw || '').trim().toUpperCase();
}

/** Read SDD MAIN decision from any legacy/alternate column name. */
function readSddMainDecisionLabel_(main) {
  const merged = main || {};
  return normalizeSddDecisionLabel_(
    merged['statusSDD'] || merged['statusSdd'] || merged['Status SDD'] ||
    merged['statusBossDecision'] || merged['Status Boss Decision'] || ''
  );
}

function readSddMainScrStatus_(main) {
  return String((main || {})['SCR - Screening Status'] || '').trim().toLowerCase();
}

function readSddMainSupplierType_(main) {
  return String(
    (main || {})['supplier_type'] || (main || {})['Supplier Type'] || ''
  ).trim().toUpperCase();
}

function millTtpPeriodPatch_(identity) {
  const out = {};
  const m = String((identity || {})['MONTH'] || (identity || {})['Month'] || '').trim();
  const q = String((identity || {})['QUARTER'] || (identity || {})['Quarter'] || '').trim();
  const y = String((identity || {})['YEAR'] || (identity || {})['Year'] || '').trim();
  const qDerived = m ? 'Q' + Math.ceil(parseInt(m, 10) / 3) : q;
  if (qDerived) {
    out['Quarter'] = qDerived;
    out['QUARTER'] = qDerived;
  }
  if (m) {
    out['Month'] = m;
    out['MONTH'] = m;
  }
  if (y) {
    out['Year'] = y;
    out['YEAR'] = y;
  }
  return out;
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

function isMeaningfulTmlRow_(tml) {
  if (!tml || typeof tml !== 'object') return false;
  if (String(tml['is_deleted'] || '') === '1') return false;
  return !!(
    String(tml['TML - Mill Name'] || '').trim()
    || String(tml['TML - UML ID'] || '').trim()
    || String(tml['TML - Company Name'] || '').trim()
  );
}

function parseMillAddedLines_(raw) {
  return String(raw || '').split(',').map(function(s) {
    return String(s || '').trim();
  }).filter(Boolean);
}

/**
 * Track per-mill completion for Task List (one submission, many TML rows).
 * When all meaningful TML rows are marked, sets mill_added=true on MAIN.
 */
function applyMillAddedLine_(sid, lineId, mainHit, user, now) {
  sid = String(sid || '').trim();
  lineId = String(lineId || '').trim();
  if (!sid || !lineId) throw new Error('mill_added_line requires submission_id and line_id');

  const supplierType = String(mainHit.obj['supplier_type'] || mainHit.obj['Supplier Type'] || '').trim().toUpperCase();
  if (supplierType !== 'TRADER' && supplierType !== 'MILL' && supplierType !== 'KCP') {
    throw new Error('mill_added_line is only valid for TRADER, MILL, or KCP submissions');
  }

  const tmlRows = findChildRows_('sddMill', sid)
    .map(function(r) { return r.obj; })
    .filter(isMeaningfulTmlRow_);
  const knownIds = {};
  tmlRows.forEach(function(r) {
    const lid = String(r['line_id'] || '').trim();
    if (lid) knownIds[lid] = true;
  });
  if (!knownIds[lineId]) {
    throw new Error('mill_added_line not found in submission mill list: ' + lineId);
  }

  const lines = parseMillAddedLines_(mainHit.obj['mill_added_lines']);
  if (lines.indexOf(lineId) === -1) lines.push(lineId);

  const patch = {
    mill_added_lines: lines.join(','),
    updated_at: now,
    updated_by: user,
  };

  const allDone = tmlRows.length > 0 && tmlRows.every(function(r) {
    return lines.indexOf(String(r['line_id'] || '').trim()) !== -1;
  });
  if (allDone) patch.mill_added = 'true';

  patchRelRow_(mainHit.ws, mainHit.headers, mainHit._sheetRow, patch);
  Object.assign(mainHit.obj, patch);

  return {
    success: true,
    submission_id: sid,
    line_id: lineId,
    completed_lines: lines.length,
    total_mills: tmlRows.length,
    all_done: allDone,
    mill_added: allDone,
  };
}

/** Ensure TRADER mill_added_line references a real TML row before TTP mirror / completion. */
function assertTraderTmlLineKnown_(sid, lineId) {
  sid = String(sid || '').trim();
  lineId = String(lineId || '').trim();
  if (!sid || !lineId) throw new Error('mill_added_line requires submission_id and line_id');

  const tmlRows = findChildRows_('sddMill', sid)
    .map(function(r) { return r.obj; })
    .filter(isMeaningfulTmlRow_);
  const known = tmlRows.some(function(r) {
    return String(r['line_id'] || '').trim() === lineId;
  });
  if (!known) {
    throw new Error('mill_added_line not found in submission mill list: ' + lineId);
  }
}

function readTtpRows_() {
  ensureTtpHeaders_();
  const ws = getSheet('ttp');
  const range = ws.getDataRange();
  const values = range.getValues();
  if (!values.length) {
    return { ws: ws, headers: TTP_HEADERS.slice(), rows: [] };
  }

  const hdr = detectTtpHeaderRow_(ws);
  const headerRowNum = hdr ? hdr.headerRow : 1;
  const headers = (hdr && hdr.headers && hdr.headers.length)
    ? hdr.headers
    : values[0].map(function(h) { return String(h || '').trim(); });
  const rows = [];
  for (let r = headerRowNum; r < values.length; r++) {
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
  const hdr = detectTtpHeaderRow_(ws);
  const headerRow = hdr && hdr.headerRow ? hdr.headerRow : 1;
  const hdrs = headers && headers.length ? headers : (hdr && hdr.headers ? hdr.headers : TTP_HEADERS.slice());
  const newRow = hdrs.map(function(h) {
    return obj[h] !== undefined && obj[h] !== null ? obj[h] : '';
  });
  forceCoordStrings_(hdrs, newRow);
  const targetRow = ttpFindNextAppendRow_(ws, hdrs, headerRow);
  safeInsertRowAt_(ws, hdrs, newRow, targetRow);
  return targetRow;
}

/** Add multiple TTP rows in one request (Dealer → one row per village, consecutive active-zone rows). */
function addTtpBatch_(rows) {
  if (!rows || !rows.length) throw new Error('No TTP rows to add');
  ensureTtpHeaders_();
  const sheet = getSheet('ttp');
  const hdr = detectTtpHeaderRow_(sheet);
  const headers = hdr && hdr.headers && hdr.headers.length ? hdr.headers : TTP_HEADERS.slice();
  const headerRow = hdr && hdr.headerRow ? hdr.headerRow : 1;
  var nextRow = ttpFindNextAppendRow_(sheet, headers, headerRow);
  var rowArrays = rows.map(function(data) {
    const newRow = headers.map(function(h) {
      return data[h] !== undefined && data[h] !== null ? data[h] : '';
    });
    forceCoordStrings_(headers, newRow);
    return newRow;
  });
  rowArrays.forEach(function(rowArr, i) {
    safeInsertRowAt_(sheet, headers, rowArr, nextRow + i);
  });
  return {
    success: true,
    count: rows.length,
    startRow: nextRow,
    endRow: nextRow + rows.length - 1,
  };
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

function buildTtpPatchFromSddFfb_(main, ffb, sid, user, now, millIdentity) {
  const identity = millIdentity || {};
  const supplierType = readSddMainSupplierType_(main);
  let millName = String(identity['MILL NAME'] || main['Mill Name'] || '').trim();
  if (!millName && /trader/i.test(supplierType)) {
    millName = String(identity['COMPANY NAME'] || main['Company Name'] || '').trim();
  }

  let lat = String(ffb['FFB - Latitude'] || '').trim();
  let lng = String(ffb['FFB - Longitude'] || '').trim();
  if (isBlankTtpCell_(lat)) lat = String(main['Latitude'] || '').trim();
  if (isBlankTtpCell_(lng)) lng = String(main['Longitude'] || '').trim();

  return Object.assign({
    'GROUP NAME'              : String(identity['GROUP NAME'] || main['Group Name'] || '').trim(),
    'COMPANY NAME'            : String(identity['COMPANY NAME'] || main['Company Name'] || '').trim(),
    'MILL NAME'               : millName,
    'UML ID'                  : String(identity['UML ID'] || '').trim(),
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
  }, millTtpPeriodPatch_(identity));
}

function buildTtpMillHeaderPatch_(main, millIdentity, sid, user, now) {
  const identity = millIdentity || {};
  const supplierType = readSddMainSupplierType_(main);
  let millName = String(identity['MILL NAME'] || main['Mill Name'] || '').trim();
  if (!millName && /trader/i.test(supplierType)) {
    millName = String(identity['COMPANY NAME'] || main['Company Name'] || '').trim();
  }
  return Object.assign({
    'GROUP NAME'    : String(identity['GROUP NAME'] || main['Group Name'] || '').trim(),
    'COMPANY NAME'  : String(identity['COMPANY NAME'] || main['Company Name'] || '').trim(),
    'MILL NAME'     : millName,
    'UML ID'        : String(identity['UML ID'] || '').trim(),
    'submission_id' : sid,
    'ffb_line_id'   : TTP_MILL_HEADER_LINE_ID,
    'supplier_type' : supplierType,
    'synced_at'     : now,
    'synced_by'     : user,
  }, millTtpPeriodPatch_(identity));
}

function removeTtpMillHeaderStub_(ttpResult, sid) {
  sid = String(sid || '').trim();
  if (!sid || !ttpResult || !ttpResult.rows) return;
  for (let i = ttpResult.rows.length - 1; i >= 0; i--) {
    const row = ttpResult.rows[i];
    if (String(row['submission_id'] || '').trim() !== sid) continue;
    if (String(row['ffb_line_id'] || '').trim() !== TTP_MILL_HEADER_LINE_ID) continue;
    ttpResult.ws.deleteRow(row._row);
    ttpResult.rows.splice(i, 1);
  }
}

/** Whitelist + trim mill identity for TTP sync (Task List → Traceability Data). */
function sanitizeMillTtpIdentity_(raw) {
  if (!raw || typeof raw !== 'object') {
    throw new Error('mill_ttp_sync identity is required');
  }
  const MAX = 500;
  function clip(v) {
    return String(v === undefined || v === null ? '' : v).trim().slice(0, MAX);
  }
  const out = {
    'GROUP NAME'   : clip(raw['GROUP NAME']),
    'COMPANY NAME' : clip(raw['COMPANY NAME']),
    'MILL NAME'    : clip(raw['MILL NAME']),
    'UML ID'       : clip(raw['UML ID']),
  };
  if (!out['GROUP NAME'])   throw new Error('GROUP NAME is required for Traceability sync');
  if (!out['COMPANY NAME']) throw new Error('COMPANY NAME is required for Traceability sync');
  if (!out['MILL NAME'])    throw new Error('MILL NAME is required for Traceability sync');
  if (!out['UML ID'])       throw new Error('UML ID is required for Traceability sync');
  const m = clip(raw['MONTH'] || raw['Month']);
  const q = clip(raw['QUARTER'] || raw['Quarter']);
  const y = clip(raw['YEAR'] || raw['Year']);
  if (m) out['MONTH'] = m;
  if (q) out['QUARTER'] = q; // backward compat
  if (y) out['YEAR'] = y;
  return out;
}

/** Mirror payload from Mill Onboarding save (TRADER) — identity + optional location. */
function sanitizeMillTtpMirrorFromOnboarding_(raw) {
  const base = sanitizeMillTtpIdentity_(raw);
  const MAX = 500;
  function clip(v) {
    return String(v === undefined || v === null ? '' : v).trim().slice(0, MAX);
  }
  const out = Object.assign({}, base);
  const lat = clip(raw['LAT'] || raw['Latitude']);
  const lng = clip(raw['LONG'] || raw['Longitude']);
  if (lat) out['LAT'] = lat;
  if (lng) out['LONG'] = lng;
  return out;
}

function buildTtpTraderMillMirrorPatch_(millIdentity, sid, tmlLineId, user, now) {
  const identity = millIdentity || {};
  return Object.assign({
    'GROUP NAME'    : String(identity['GROUP NAME'] || '').trim(),
    'COMPANY NAME'  : String(identity['COMPANY NAME'] || '').trim(),
    'MILL NAME'     : String(identity['MILL NAME'] || '').trim(),
    'UML ID'        : String(identity['UML ID'] || '').trim(),
    'LAT'           : String(identity['LAT'] || '').trim(),
    'LONG'          : String(identity['LONG'] || '').trim(),
    'submission_id' : sid,
    'ffb_line_id'   : ttpTraderMirrorLineId_(tmlLineId),
    'supplier_type' : 'TRADER',
    'synced_at'     : now,
    'synced_by'     : user,
  }, millTtpPeriodPatch_(identity));
}

/**
 * TRADER: mirror one Mill Onboarding row → one TTP row (identity only, no FFB list).
 */
function syncTtpMirrorTraderMillFromOnboarding_(sid, tmlLineId, millIdentity, mainObj, user, now) {
  sid = String(sid || '').trim();
  tmlLineId = String(tmlLineId || '').trim();
  if (!sid) return { synced: false, skipped: true, reason: 'missing_submission_id' };
  if (!tmlLineId) return { synced: false, skipped: true, reason: 'missing_tml_line_id' };

  const merged = mainObj || {};
  const decision = readSddMainDecisionLabel_(merged);
  if (decision !== 'APPROVED') {
    return { synced: false, skipped: true, reason: 'not_approved', decision: decision };
  }

  const scrSt = readSddMainScrStatus_(merged);
  if (scrSt !== 'submitted') {
    return { synced: false, skipped: true, reason: 'not_submitted', scr_status: scrSt };
  }

  const supplierType = readSddMainSupplierType_(merged);
  if (supplierType !== 'TRADER') {
    return { synced: false, skipped: true, reason: 'not_trader', supplier_type: supplierType };
  }

  const patch = buildTtpTraderMillMirrorPatch_(millIdentity, sid, tmlLineId, user, now);
  const ttpResult = readTtpRows_();
  const mirrorLineId = ttpTraderMirrorLineId_(tmlLineId);
  let inserted = 0;
  let updated = 0;

  try {
    const hit = findTtpRowBySyncKeys_(ttpResult.rows, sid, mirrorLineId, patch);
    if (hit) {
      const mergedPatch = mergeTtpPreserveMonitoring_(hit.row, patch);
      patchTtpRow_(ttpResult.ws, ttpResult.headers, hit._sheetRow, mergedPatch);
      updated++;
    } else {
      const newRow = appendTtpRow_(ttpResult.ws, ttpResult.headers, patch);
      ttpResult.rows.push(Object.assign({ _row: newRow }, patch));
      inserted++;
    }
  } catch (err) {
    return {
      synced: false,
      submission_id: sid,
      inserted: 0,
      updated: 0,
      total_ffb: 0,
      errors: [{ line_id: tmlLineId, error: String(err && err.message ? err.message : err) }],
      reason: 'mirror_failed',
    };
  }

  return {
    synced: true,
    submission_id: sid,
    inserted: inserted,
    updated: updated,
    total_ffb: 0,
    trader_mirror: true,
    tml_line_id: tmlLineId,
  };
}

function assertMillTtpSyncSucceeded_(ttpSync) {
  if (!ttpSync) {
    throw new Error('Traceability sync did not run');
  }
  if (!ttpSync.synced) {
    const reason = String(ttpSync.reason || ttpSync.decision || 'unknown').trim();
    throw new Error('Traceability sync skipped: ' + reason);
  }
  if (ttpSync.errors && ttpSync.errors.length) {
    const detail = ttpSync.errors.map(function(e) {
      return String(e.line_id || '') + ': ' + String(e.error || '');
    }).join('; ');
    throw new Error('Traceability sync failed: ' + detail);
  }
  if ((ttpSync.inserted || 0) + (ttpSync.updated || 0) < 1) {
    throw new Error('Traceability sync produced no rows');
  }
}

/**
 * After Mill Onboarding save (Task List), upsert Monitoring TTP/TTM rows.
 * Header identity (GROUP/COMPANY/MILL/UML) comes from Mill Onboarding; FFB rows from SDD.
 * Does not overwrite monitoring-only columns when already filled.
 */
function syncTtpFromMillOnboarding_(sid, millIdentity, mainObj, user, now) {
  sid = String(sid || '').trim();
  if (!sid) return { synced: false, skipped: true, reason: 'missing_submission_id' };

  const merged = mainObj || {};
  const decision = readSddMainDecisionLabel_(merged);
  if (decision !== 'APPROVED') {
    return { synced: false, skipped: true, reason: 'not_approved', decision: decision };
  }

  const scrSt = readSddMainScrStatus_(merged);
  if (scrSt !== 'submitted') {
    return { synced: false, skipped: true, reason: 'not_submitted', scr_status: scrSt };
  }

  const supplierType = readSddMainSupplierType_(merged);
  if (supplierType !== 'MILL' && supplierType !== 'KCP') {
    return { synced: false, skipped: true, reason: 'supplier_type_not_mill_or_kcp', supplier_type: supplierType };
  }

  const identity = millIdentity || {};
  const ffbRows = findChildRows_('sddFfb', sid)
    .map(function(r) { return r.obj; })
    .filter(isMeaningfulSddFfbRow_);

  const ttpResult = readTtpRows_();
  let inserted = 0;
  let updated = 0;
  const errors = [];

  if (ffbRows.length) {
    removeTtpMillHeaderStub_(ttpResult, sid);
    ffbRows.forEach(function(ffb) {
      try {
        const lineId = String(ffb['line_id'] || '').trim();
        const patch = buildTtpPatchFromSddFfb_(merged, ffb, sid, user, now, identity);
        const hit = findTtpRowBySyncKeys_(ttpResult.rows, sid, lineId, patch);
        if (hit) {
          const mergedPatch = mergeTtpPreserveMonitoring_(hit.row, patch);
          patchTtpRow_(ttpResult.ws, ttpResult.headers, hit._sheetRow, mergedPatch);
          Object.assign(hit.row, mergedPatch);
          updated++;
        } else {
          const newRow = appendTtpRow_(ttpResult.ws, ttpResult.headers, patch);
          ttpResult.rows.push(Object.assign({ _row: newRow }, patch));
          inserted++;
        }
      } catch (err) {
        errors.push({
          line_id: String(ffb['line_id'] || ''),
          error: String(err && err.message ? err.message : err),
        });
      }
    });
  } else {
    try {
      const patch = buildTtpMillHeaderPatch_(merged, identity, sid, user, now);
      const hit = findTtpRowBySyncKeys_(ttpResult.rows, sid, TTP_MILL_HEADER_LINE_ID, patch);
      if (hit) {
        const mergedPatch = mergeTtpPreserveMonitoring_(hit.row, patch);
        patchTtpRow_(ttpResult.ws, ttpResult.headers, hit._sheetRow, mergedPatch);
        Object.assign(hit.row, mergedPatch);
        updated++;
      } else {
        const newRow = appendTtpRow_(ttpResult.ws, ttpResult.headers, patch);
        ttpResult.rows.push(Object.assign({ _row: newRow }, patch));
        inserted++;
      }
    } catch (err) {
      errors.push({
        line_id: TTP_MILL_HEADER_LINE_ID,
        error: String(err && err.message ? err.message : err),
      });
    }
  }

  return {
    synced: errors.length ? false : true,
    submission_id: sid,
    inserted: inserted,
    updated: updated,
    total_ffb: ffbRows.length,
    errors: errors.length ? errors : undefined,
    reason: errors.length ? 'partial_or_total_failure' : undefined,
  };
}

function normalizeTtpMatchKey_(raw) {
  return String(raw || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function tmlMatchKeys_(tml) {
  const keys = [];
  ['TML - Mill Name', 'TML - Company Name', 'TML - Company Group Name'].forEach(function(k) {
    const v = normalizeTtpMatchKey_(tml[k]);
    if (v && keys.indexOf(v) === -1) keys.push(v);
  });
  return keys;
}

/** Carry-forward FFB - Mill Name (blank / — / - = same mill as row above). */
function resolveFfbEffectiveMillNames_(ffbRows) {
  let carry = '';
  return (ffbRows || []).filter(isMeaningfulSddFfbRow_).map(function(ffb) {
    const raw = String(ffb['FFB - Mill Name'] || '').trim();
    if (!isBlankTtpCell_(raw)) carry = raw;
    return { ffb: ffb, effectiveMill: carry };
  });
}

/** FFB rows from Traceability B that belong to one TML line (Mill List A). */
function ffbRowsForTmlLine_(tmlRows, ffbRows, lineId) {
  lineId = String(lineId || '').trim();
  const orderedTml = (tmlRows || []).filter(isMeaningfulTmlRow_);
  const targetIdx = orderedTml.findIndex(function(t) {
    return String(t['line_id'] || '').trim() === lineId;
  });
  if (targetIdx < 0) return [];

  const targetKeys = tmlMatchKeys_(orderedTml[targetIdx]);
  if (!targetKeys.length) return [];

  const resolved = resolveFfbEffectiveMillNames_(ffbRows);
  return resolved.filter(function(entry) {
    const eff = normalizeTtpMatchKey_(entry.effectiveMill);
    if (!eff) return false;
    return targetKeys.indexOf(eff) !== -1;
  }).map(function(entry) { return entry.ffb; });
}

function countMeaningfulTmlRowsForSid_(sid) {
  return findChildRows_('sddMill', sid)
    .map(function(r) { return r.obj; })
    .filter(isMeaningfulTmlRow_).length;
}

/**
 * MILL/KCP: sync one Mill List (TML) row → Mill Onboarding identity + matching FFB rows.
 */
function syncTtpFromMillOnboardingForTmlLine_(sid, tmlLineId, millIdentity, mainObj, user, now) {
  sid = String(sid || '').trim();
  tmlLineId = String(tmlLineId || '').trim();
  if (!sid) return { synced: false, skipped: true, reason: 'missing_submission_id' };
  if (!tmlLineId) return { synced: false, skipped: true, reason: 'missing_tml_line_id' };

  const merged = mainObj || {};
  const decision = readSddMainDecisionLabel_(merged);
  if (decision !== 'APPROVED') {
    return { synced: false, skipped: true, reason: 'not_approved', decision: decision };
  }

  const scrSt = readSddMainScrStatus_(merged);
  if (scrSt !== 'submitted') {
    return { synced: false, skipped: true, reason: 'not_submitted', scr_status: scrSt };
  }

  const supplierType = readSddMainSupplierType_(merged);
  if (supplierType !== 'MILL' && supplierType !== 'KCP') {
    return { synced: false, skipped: true, reason: 'not_mill_or_kcp', supplier_type: supplierType };
  }

  const tmlRows = findChildRows_('sddMill', sid).map(function(r) { return r.obj; });
  const tmlHit = tmlRows.filter(isMeaningfulTmlRow_).find(function(t) {
    return String(t['line_id'] || '').trim() === tmlLineId;
  });
  if (!tmlHit) {
    return { synced: false, skipped: true, reason: 'tml_line_not_found', tml_line_id: tmlLineId };
  }

  const identity = millIdentity || {};
  const allFfb = findChildRows_('sddFfb', sid).map(function(r) { return r.obj; });
  const ffbRows = ffbRowsForTmlLine_(tmlRows, allFfb, tmlLineId);
  const headerLineId = ttpMillTmlHeaderLineId_(tmlLineId);

  const ttpResult = readTtpRows_();
  let inserted = 0;
  let updated = 0;
  const errors = [];

  if (ffbRows.length) {
    ffbRows.forEach(function(ffb) {
      try {
        const lineId = String(ffb['line_id'] || '').trim();
        const patch = buildTtpPatchFromSddFfb_(merged, ffb, sid, user, now, identity);
        const hit = findTtpRowBySyncKeys_(ttpResult.rows, sid, lineId, patch);
        if (hit) {
          const mergedPatch = mergeTtpPreserveMonitoring_(hit.row, patch);
          patchTtpRow_(ttpResult.ws, ttpResult.headers, hit._sheetRow, mergedPatch);
          Object.assign(hit.row, mergedPatch);
          updated++;
        } else {
          const newRow = appendTtpRow_(ttpResult.ws, ttpResult.headers, patch);
          ttpResult.rows.push(Object.assign({ _row: newRow }, patch));
          inserted++;
        }
      } catch (err) {
        errors.push({
          line_id: String(ffb['line_id'] || ''),
          error: String(err && err.message ? err.message : err),
        });
      }
    });
  } else {
    try {
      const patch = buildTtpMillHeaderPatch_(merged, identity, sid, user, now);
      patch['ffb_line_id'] = headerLineId;
      const hit = findTtpRowBySyncKeys_(ttpResult.rows, sid, headerLineId, patch);
      if (hit) {
        const mergedPatch = mergeTtpPreserveMonitoring_(hit.row, patch);
        patchTtpRow_(ttpResult.ws, ttpResult.headers, hit._sheetRow, mergedPatch);
        Object.assign(hit.row, mergedPatch);
        updated++;
      } else {
        const newRow = appendTtpRow_(ttpResult.ws, ttpResult.headers, patch);
        ttpResult.rows.push(Object.assign({ _row: newRow }, patch));
        inserted++;
      }
    } catch (err) {
      errors.push({
        line_id: headerLineId,
        error: String(err && err.message ? err.message : err),
      });
    }
  }

  return {
    synced: errors.length ? false : true,
    submission_id: sid,
    inserted: inserted,
    updated: updated,
    total_ffb: ffbRows.length,
    tml_line_id: tmlLineId,
    mill_tml: true,
    errors: errors.length ? errors : undefined,
    reason: errors.length ? 'partial_or_total_failure' : undefined,
  };
}

// ═══════════════════════════════════════════════════════════
//  TTP / TTM  ─ AUTO-INIT HEADERS
// ═══════════════════════════════════════════════════════════

function normalizeTtpLegacyHeaders_(ws) {
  if (!ws) return ws;
  const hdrRowInfo = detectTtpHeaderRow_(ws);
  const headerRow = hdrRowInfo ? hdrRowInfo.headerRow : 1;
  const lastCol = Math.max(ws.getLastColumn(), 1);
  const headers = ws.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  let changed = false;
  const out = headers.map(function(h, idx) {
    const trimmed = String(h || '').replace(/\s+/g, ' ').trim();
    const upper = trimmed.toUpperCase();
    if (upper === 'GROUP NAME' || upper === 'GROUP NAME NBL') return trimmed;
    if (trimmed === 'NAME' || trimmed === ' NAME' || upper === 'GROUP NAME') {
      changed = true;
      return 'GROUP NAME';
    }
    if (/^group\s*name$/i.test(trimmed)) {
      changed = true;
      return 'GROUP NAME';
    }
    return trimmed;
  });
  if (changed) {
    ws.getRange(headerRow, 1, 1, out.length).setValues([out]);
  }
  return ws;
}

function ensureTtpHeaders_() {
  const sheet   = getSheet('ttp');
  const lastCol = sheet.getLastColumn();

  if (lastCol === 0) {
    sheet.getRange(1, 1, 1, TTP_HEADERS.length).setValues([TTP_HEADERS]);
    return sheet;
  }

  normalizeTtpLegacyHeaders_(sheet);

  const hdrRowInfo = detectTtpHeaderRow_(sheet);
  const headerRow  = hdrRowInfo ? hdrRowInfo.headerRow : 1;
  const existing    = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0]
                           .map(function(h) { return String(h || '').trim(); });
  const existingSet = new Set(existing.filter(Boolean));
  const missing     = TTP_HEADERS.filter(function(h) { return !existingSet.has(h); });

  if (missing.length) {
    const start = existing.length + 1;
    sheet.insertColumnsAfter(existing.length, missing.length);
    sheet.getRange(headerRow, start, 1, missing.length).setValues([missing]);
  }
  return sheet;
}

function detectHeaderRowByKeys_(sheet, requiredKeys, scanMaxRows) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (!lastRow || !lastCol) return null;
  const scanRows = Math.min(Math.max(scanMaxRows || 10, 1), lastRow);
  const data = sheet.getRange(1, 1, scanRows, lastCol).getValues();
  const wanted = requiredKeys.map(function(k) { return String(k || '').trim().toLowerCase(); });
  for (var r = 0; r < data.length; r++) {
    const row = data[r].map(function(c) { return String(c || '').trim().toLowerCase(); });
    var hit = 0;
    for (var i = 0; i < wanted.length; i++) {
      if (row.indexOf(wanted[i]) !== -1) hit++;
    }
    // Consider this row as header if enough key columns are present.
    if (hit >= Math.min(4, wanted.length)) {
      const headers = data[r].map(function(h) { return String(h || '').trim(); });
      if (isTtpInvalidHeaderRow_(headers)) continue;
      return { headerRow: r + 1, headers: headers };
    }
  }
  return null;
}

function isTtpInvalidHeaderRow_(headers) {
  var list = (headers || []).map(function(h) { return String(h || '').trim(); }).filter(Boolean);
  if (!list.length) return true;
  for (var i = 0; i < list.length; i++) {
    var h = list[i];
    var hu = h.toUpperCase();
    if (/^TOTAL TRACEABLE/i.test(h)) return true;
    if (/^TOTAL CPO|^TOTAL PK/i.test(hu)) return true;
    if (/^[\d.,]+$/.test(h)) return true;
  }
  return false;
}

function ttpHeaderRowLooksCanonical_(rowLower) {
  if (!rowLower || !rowLower.length) return false;
  return rowLower.indexOf('quarter') !== -1
    && rowLower.indexOf('year') !== -1
    && rowLower.indexOf('company name') !== -1
    && (rowLower.indexOf('mill name') !== -1 || rowLower.indexOf('uml id') !== -1);
}

function detectTtpHeaderRow_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (!lastRow || !lastCol) return null;
  const scanRows = Math.min(25, lastRow);
  const data = sheet.getRange(1, 1, scanRows, lastCol).getValues();
  const primaryKeys = ['GROUP NAME', 'COMPANY NAME', 'MILL NAME', 'FFB SUPPLIER NAME', 'CATEGORY'];
  const legacyKeys = ['Quarter', 'Year', 'UML ID', 'FFB SUPPLIER NAME', 'CATEGORY'];
  var best = null;
  var bestScore = 0;

  // Standard layout: KPI rows 1–2, header row 3 (NO | Quarter | Year | … | COMPANY NAME | MILL NAME).
  for (var r = 0; r < data.length; r++) {
    const headers = data[r].map(function(h) { return String(h || '').trim(); });
    if (!headers.some(Boolean)) continue;
    if (isTtpInvalidHeaderRow_(headers)) continue;
    const rowLower = headers.map(function(h) { return h.toLowerCase(); });
    if (ttpHeaderRowLooksCanonical_(rowLower)) {
      return { headerRow: r + 1, headers: headers };
    }
  }

  for (var r2 = 0; r2 < data.length; r2++) {
    const headers = data[r2].map(function(h) { return String(h || '').trim(); });
    if (!headers.some(Boolean)) continue;
    if (isTtpInvalidHeaderRow_(headers)) continue;

    const rowLower = headers.map(function(h) { return h.toLowerCase(); });
    var hitPrimary = 0;
    for (var pi = 0; pi < primaryKeys.length; pi++) {
      if (rowLower.indexOf(String(primaryKeys[pi]).trim().toLowerCase()) !== -1) hitPrimary++;
    }
    var hitLegacy = 0;
    for (var li = 0; li < legacyKeys.length; li++) {
      if (rowLower.indexOf(String(legacyKeys[li]).trim().toLowerCase()) !== -1) hitLegacy++;
    }

    const hasNo = rowLower.indexOf('no') !== -1 || rowLower.indexOf('no.') !== -1;
    const hasCompany = rowLower.indexOf('company name') !== -1;
    const hasMill = rowLower.indexOf('mill name') !== -1;
    const hasFfb = rowLower.indexOf('ffb supplier name') !== -1;
    const hasGroup = rowLower.indexOf('group name') !== -1;
    const hasQuarterYear = rowLower.indexOf('quarter') !== -1 && rowLower.indexOf('year') !== -1;

    var score = hitPrimary;
    if (hasQuarterYear) score += 25;
    if (hasCompany && hasMill) score += 6;
    if (hasFfb) score += 3;
    if (hasGroup) score += 2;
    if (hasNo) score += 2;
    if (hitLegacy >= 4 && rowLower.indexOf('uml id') !== -1) score = Math.max(score, hitLegacy + 4);

    if (score > bestScore) {
      bestScore = score;
      best = { headerRow: r2 + 1, headers: headers };
    }
  }

  if (best && bestScore >= 4) return best;
  return detectHeaderRowByKeys_(sheet, primaryKeys, 15) || detectHeaderRowByKeys_(sheet, legacyKeys, 15);
}

function looksLikeTtpYearOrQuarterValue_(s) {
  var t = String(s || '').trim();
  if (!t) return true;
  if (/^(19|20)\d{2}$/.test(t)) return true;
  if (/^Q[1-4]$/i.test(t)) return true;
  return false;
}

function looksLikeTtpGroupNameValue_(s) {
  return !looksLikeTtpYearOrQuarterValue_(s);
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
//  (+ Mill Month/Year: dashboard sends MONTH/YEAR; sheet may use Month/Year or Quarter/Year for legacy data)
// ═══════════════════════════════════════════════════════════

/**
 * Map payload keys MONTH / YEAR (from sustain-dashboard) onto the exact
 * header names present in row 1 of the Mill sheet (e.g. Month, Year).
 * Also handles legacy QUARTER field for backward compatibility.
 */
function resolveMillQuarterYearKeys_(data, headers) {
  if (!data || typeof data !== 'object' || !Array.isArray(headers)) return;
  var list = headers.map(function(x) { return String(x || '').trim(); });
  function findMonthCol() {
    for (var i = 0; i < list.length; i++) {
      if (/^month$/i.test(list[i]) || /^bulan$/i.test(list[i])) return list[i];
    }
    return null;
  }
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
  var mCol = findMonthCol();
  var qCol = findQuarterCol();
  var yCol = findYearCol();
  // Write MONTH to the Month column
  if (mCol && data['MONTH'] !== undefined && String(data['MONTH']).trim() !== '' &&
      (data[mCol] === undefined || String(data[mCol]).trim() === '')) {
    data[mCol] = data['MONTH'];
  }
  // Legacy: if sheet still has Quarter column, map MONTH→quarter derivation
  if (qCol && data['QUARTER'] !== undefined && String(data['QUARTER']).trim() !== '' &&
      (data[qCol] === undefined || String(data[qCol]).trim() === '')) {
    data[qCol] = data['QUARTER'];
  }
  if (yCol && data['YEAR'] !== undefined && String(data['YEAR']).trim() !== '' &&
      (data[yCol] === undefined || String(data[yCol]).trim() === '')) {
    data[yCol] = data['YEAR'];
  }
}

/** Expose MONTH / YEAR on getAll('mill') for clients that expect uppercase keys. */
function mirrorMillQuarterYearOnRead_(obj) {
  if (!obj || typeof obj !== 'object') return;
  // Mirror MONTH column
  Object.keys(obj).forEach(function(k) {
    if (k === '_row') return;
    var t = String(k).trim();
    if ((/^month$/i.test(t) || /^bulan$/i.test(t)) && obj['MONTH'] === undefined) obj['MONTH'] = obj[k];
  });
  // Legacy backward compat: expose QUARTER if sheet still has it
  Object.keys(obj).forEach(function(k) {
    if (k === '_row') return;
    if (/^quarter$/i.test(String(k).trim()) && obj['QUARTER'] === undefined) obj['QUARTER'] = obj[k];
  });
  // Derive QUARTER from MONTH if QUARTER absent
  if (obj['MONTH'] !== undefined && String(obj['MONTH']).trim() !== '' && obj['QUARTER'] === undefined) {
    var mn = parseInt(String(obj['MONTH']).trim(), 10);
    if (mn >= 1 && mn <= 12) obj['QUARTER'] = 'Q' + Math.ceil(mn / 3);
  }
  // Legacy: derive MONTH from QUARTER when MONTH column empty (Q1 → months 1-3 represented as 1)
  if ((!obj['MONTH'] || String(obj['MONTH']).trim() === '') && obj['QUARTER'] !== undefined && String(obj['QUARTER']).trim() !== '') {
    var qRaw = String(obj['QUARTER']).trim().toUpperCase().replace(/\s+/g, '');
    var qm = qRaw.match(/^Q?([1-4])$/);
    if (qm) obj['MONTH'] = String((parseInt(qm[1], 10) - 1) * 3 + 1);
  }
  Object.keys(obj).forEach(function(k) {
    if (k === '_row') return;
    var t = String(k).trim();
    if ((t === 'Year' || t === 'YEAR') && obj['YEAR'] === undefined) obj['YEAR'] = obj[k];
  });
}

function detectNblHeaderRow_(sheet) {
  return detectHeaderRowByKeys_(
    sheet,
    ['Riser', 'Group Name NBL', 'Company Name NBL'],
    12
  ) || detectHeaderRowByKeys_(
    sheet,
    ['Raiser', 'Group Name', 'Company Name'],
    12
  );
}

/** Normalize NBL row keys to canonical Riser / Group Name NBL / Company Name NBL / SOURCE. */
function mirrorNblFieldsOnRead_(obj) {
  if (!obj || typeof obj !== 'object') return;
  Object.keys(obj).forEach(function(k) {
    if (k === '_row') return;
    var nk = String(k).replace(/\s+/g, ' ').trim().toLowerCase();
    var v = obj[k];
    if (v === undefined || v === null || String(v).trim() === '') return;
    var s = String(v).trim();
    if ((nk === 'riser' || nk === 'raiser') && !String(obj['Riser'] || '').trim()) {
      obj['Riser'] = s;
    }
    if ((/group\s*name(\s*nbl)?/.test(nk) || nk === 'group') && !String(obj['Group Name NBL'] || '').trim()) {
      obj['Group Name NBL'] = s;
    }
    if ((/company\s*name(\s*nbl)?/.test(nk) || nk === 'company') && !String(obj['Company Name NBL'] || '').trim()) {
      obj['Company Name NBL'] = s;
    }
    if (nk === 'source' && !String(obj['SOURCE'] || '').trim()) {
      obj['SOURCE'] = s;
    }
  });
}

function normalizeGrievanceHeaderCell_(h) {
  return String(h || '').replace(/\s+/g, ' ').trim();
}

function grievanceSheetHasHeader_(headers, name) {
  var want = normalizeGrievanceHeaderCell_(name).toLowerCase();
  for (var i = 0; i < (headers || []).length; i++) {
    var t = normalizeGrievanceHeaderCell_(headers[i]).toLowerCase();
    if (!t) continue;
    if (t === want) return true;
    if (want === 'consequence' && t === 'consequense') return true;
  }
  return false;
}

/**
 * Ensure risk-classification columns exist on Grievance Monitoring (row 1).
 * Inserts after "Grievance Description" when present, else before "Verification Findings", else appends.
 */
function ensureGrievanceRiskHeaders_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = SHEETS.grievance;
  if (!name) return null;
  var ws = ss.getSheetByName(name);
  if (!ws) {
    ws = ss.insertSheet(name);
    ws.getRange(1, 1, 1, GRIEVANCE_RISK_HEADERS.length).setValues([GRIEVANCE_RISK_HEADERS]);
    return ws;
  }

  var lastCol = Math.max(ws.getLastColumn(), 1);
  if (ws.getLastRow() === 0) {
    ws.getRange(1, 1, 1, GRIEVANCE_RISK_HEADERS.length).setValues([GRIEVANCE_RISK_HEADERS]);
    return ws;
  }

  var headers = ws.getRange(1, 1, 1, lastCol).getValues()[0].map(normalizeGrievanceHeaderCell_);

  for (var c = 0; c < headers.length; c++) {
    if (/^consequense$/i.test(headers[c]) && !grievanceSheetHasHeader_(headers, 'Consequence')) {
      ws.getRange(1, c + 1).setValue('Consequence');
      headers[c] = 'Consequence';
    }
  }

  var missing = GRIEVANCE_RISK_HEADERS.filter(function(h) {
    return !grievanceSheetHasHeader_(headers, h);
  });
  if (!missing.length) return ws;

  var descIdx = -1;
  var verIdx = -1;
  for (var i = 0; i < headers.length; i++) {
    var u = headers[i].toLowerCase();
    if (descIdx < 0 && u === 'grievance description') descIdx = i;
    if (verIdx < 0 && u === 'verification findings') verIdx = i;
  }
  var insertAt1 = descIdx >= 0 ? descIdx + 2 : (verIdx >= 0 ? verIdx + 1 : headers.length + 1);

  ws.insertColumnsBefore(insertAt1, missing.length);
  ws.getRange(1, insertAt1, 1, missing.length).setValues([missing]);
  return ws;
}

/** Grievance sheet col J — header may be "Group" or "Grievance Subject Group". */
function grievanceGroupHeader_(headers) {
  for (var i = 0; i < (headers || []).length; i++) {
    var h = String(headers[i] || '').replace(/\s+/g, ' ').trim();
    if (/^group$/i.test(h) || /^grievance subject group$/i.test(h)) return h;
  }
  return '';
}

function mirrorGrvFieldsOnRead_(obj, headers, row) {
  if (!obj || typeof obj !== 'object') return;
  var groupHeader = grievanceGroupHeader_(headers);
  var val = '';
  if (groupHeader && obj[groupHeader] != null && String(obj[groupHeader]).trim() !== '') {
    val = String(obj[groupHeader]).trim();
  } else if (row && row.length > 9) {
    var pos = formatApiCellValue_(row[9]);
    if (pos != null && String(pos).trim() !== '') val = String(pos).trim();
  }
  Object.keys(obj).forEach(function(k) {
    if (k === '_row') return;
    var nk = String(k).replace(/\s+/g, ' ').trim().toLowerCase();
    if ((nk === 'group' || nk === 'grievance subject group') && !val) {
      var v = obj[k];
      if (v != null && String(v).trim() !== '') val = String(v).trim();
    }
  });
  if (!val) return;
  obj['Grievance Subject Group'] = val;
  obj['Group'] = val;
  if (groupHeader) obj[groupHeader] = val;

  if (obj['Consequense'] != null && String(obj['Consequense']).trim() !== ''
      && !String(obj['Consequence'] || '').trim()) {
    obj['Consequence'] = String(obj['Consequense']).trim();
  }
}

/** Map form field "Grievance Subject Group" to actual sheet header on write. */
function resolveGrvGroupFieldKeys_(data, headers) {
  if (!data || typeof data !== 'object') return;
  var groupHeader = grievanceGroupHeader_(headers);
  var val = String(
    data['Grievance Subject Group'] || data['Group'] || (groupHeader ? data[groupHeader] : '') || ''
  ).trim();
  if (!groupHeader) return;
  data[groupHeader] = val;
  data['Grievance Subject Group'] = val;
  data['Group'] = val;
}

/** Positional fallback: NBL sheet columns A–D = Riser, Group, Company, Source. */
function mirrorNblFieldsByPosition_(obj, row) {
  if (!obj || !row || !row.length) return;
  function cell(i) {
    var v = row[i];
    if (v === undefined || v === null) return '';
    var s = String(v).trim();
    return s && s !== '—' && s !== '-' ? s : '';
  }
  if (!String(obj['Riser'] || '').trim()) {
    var r = cell(0);
    if (r) obj['Riser'] = r;
  }
  if (!String(obj['Group Name NBL'] || '').trim()) {
    var g = cell(1);
    if (g) obj['Group Name NBL'] = g;
  }
  if (!String(obj['Company Name NBL'] || '').trim()) {
    var c = cell(2);
    if (c) obj['Company Name NBL'] = c;
  }
  if (!String(obj['SOURCE'] || '').trim()) {
    var src = cell(3);
    if (src) obj['SOURCE'] = src;
  }
}

/** Sheet header sometimes truncates "GROUP NAME" → " NAME"; mirror to GROUP NAME on read. */
function mirrorGroupNameOnRead_(obj) {
  if (!obj || typeof obj !== 'object') return;
  if (obj['GROUP NAME'] !== undefined && String(obj['GROUP NAME']).trim() !== '') return;
  var candidates = [obj[' NAME'], obj['Group Name'], obj['Grup Name'], obj['COMPANY GROUP NAME']];
  for (var i = 0; i < candidates.length; i++) {
    var s = String(candidates[i] != null ? candidates[i] : '').trim();
    if (s && s !== '—' && s !== '-') {
      obj['GROUP NAME'] = s;
      return;
    }
  }
  Object.keys(obj).forEach(function(k) {
    if (k === '_row') return;
    var t = String(k).trim();
    if (t === 'NAME' || /^group\s*name$/i.test(t)) {
      var v = String(obj[k] != null ? obj[k] : '').trim();
      if (v && v !== '—' && v !== '-') obj['GROUP NAME'] = v;
    }
  });
}

function mirrorMillGroupNameOnRead_(obj) {
  mirrorGroupNameOnRead_(obj);
}

function mirrorMillCompanyNameOnRead_(obj) {
  if (!obj || typeof obj !== 'object') return;
  var val = String(obj['COMPANY NAME'] || obj['Company Name'] || '').trim();
  if (!val) {
    Object.keys(obj).forEach(function(k) {
      if (k === '_row') return;
      var nk = String(k).replace(/\s+/g, ' ').trim();
      if (!/^company(\s*name)?$/i.test(nk)) return;
      if (/group|trader|code|facility|supply|nbl|profile|owner|tml|ffb/i.test(nk)) return;
      var v = String(obj[k] || '').trim();
      if (v && v !== '—' && v !== '-') val = v;
    });
  }
  if (!val) {
    Object.keys(obj).forEach(function(k) {
      if (k === '_row') return;
      var nk = String(k).replace(/\s+/g, ' ').trim().toLowerCase();
      if (nk.indexOf('company') === -1 || nk.indexOf('name') === -1) return;
      if (/group|trader|code|facility|supply|nbl|profile|owner|tml|ffb|mill/i.test(nk)) return;
      var v = String(obj[k] || '').trim();
      if (v && v !== '—' && v !== '-') val = v;
    });
  }
  if (!val) return;
  obj['COMPANY NAME'] = val;
  obj['Company Name'] = val;
}

/** Header-aware TTP mirror — never map Year/Quarter into GROUP NAME. */
function mirrorTtpFieldsByPosition_(obj, row, headers) {
  if (!obj || !row || !row.length || !headers || !headers.length) return;
  function cell(i) {
    var v = row[i];
    if (v === undefined || v === null) return '';
    var s = String(v).trim();
    return s && s !== '—' && s !== '-' ? s : '';
  }
  headers.forEach(function(h, j) {
    var uk = String(h || '').replace(/\s+/g, ' ').trim().toUpperCase();
    var v = cell(j);
    if (!v) return;
    if (uk === 'GROUP NAME' && !String(obj['GROUP NAME'] || '').trim() && looksLikeTtpGroupNameValue_(v)) {
      obj['GROUP NAME'] = v;
    }
    if (uk === 'COMPANY NAME' && !String(obj['COMPANY NAME'] || '').trim()) {
      obj['COMPANY NAME'] = v;
    }
    if (uk === 'MILL NAME' && !String(obj['MILL NAME'] || '').trim()) {
      obj['MILL NAME'] = v;
    }
    if (uk === 'FFB SUPPLIER GROUP NAME' && !String(obj['FFB SUPPLIER GROUP NAME'] || '').trim()) {
      obj['FFB SUPPLIER GROUP NAME'] = v;
    }
    if (uk === 'FFB SUPPLIER NAME' && !String(obj['FFB SUPPLIER NAME'] || '').trim()) {
      obj['FFB SUPPLIER NAME'] = v;
    }
  });
}

/** SUPPLY CPO/PK: use stored numeric value (getValues), not display string — avoids "66,000" vs "66.000" ambiguity. */
function millSupplyUsesRawNumber_(header) {
  var u = String(header || '').trim().toUpperCase();
  return u === 'SUPPLY CPO' || u === 'SUPPLY PK'
    || u === 'DEFORESTATION WIDTH' || u === 'BURN AREA WIDTH' || u === 'PEAT WIDTH';
}

/** Date-only cells from Sheets → stable yyyy-MM-dd for API (avoids UTC off-by-one in JSON). */
function formatApiCellValue_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return value;
}

/** ISO date strings from the dashboard → local midnight Date for Sheets. */
function coerceSheetDateValue_(value) {
  if (value === undefined || value === null || value === '') return value;
  var s = String(value).trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return value;
  var parts = s.split('-');
  return new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
}

function getData(sheetKey) {
  if (sheetKey === 'grievance') ensureGrievanceRiskHeaders_();
  const sheet = getSheet(sheetKey);
  const range = sheet.getDataRange();
  const rows  = range.getValues();
  if (!rows.length) return [];
  let headerRowNum = 1;
  let headers = rows[0];
  let dataRows = rows.slice(1);

  // Monitoring TTP/TTM may keep title/meta rows above headers.
  if (sheetKey === 'ttp') {
    const hdr = detectTtpHeaderRow_(sheet);
    if (hdr && hdr.headers && hdr.headers.length) {
      headerRowNum = hdr.headerRow;
      headers = hdr.headers;
      dataRows = rows.slice(headerRowNum);
    }
  }

  if (sheetKey === 'blMonitoring') {
    const hdr = getBlHeaderInfo_(sheet);
    if (hdr && hdr.headers && hdr.headers.length) {
      headerRowNum = hdr.headerRow;
      headers = hdr.headers;
      dataRows = rows.slice(headerRowNum);
    }
  }

  if (sheetKey === 'nbl') {
    const hdr = detectNblHeaderRow_(sheet);
    if (hdr && hdr.headers && hdr.headers.length) {
      headerRowNum = hdr.headerRow;
      headers = hdr.headers;
      dataRows = rows.slice(headerRowNum);
    }
    headers = headers.map(function(h) { return String(h || '').replace(/\s+/g, ' ').trim(); });
  }

  // Preserve display formatting (e.g. volume "66.000") for Mill & BL Monitoring.
  const dispRows = (sheetKey === 'mill' || sheetKey === 'blMonitoring')
    ? range.getDisplayValues()
    : null;
  const dispDataRows = dispRows ? dispRows.slice(headerRowNum) : null;

  return dataRows.map(function(row, i) {
    const sourceRow = dispDataRows ? dispDataRows[i] : row;
    let obj;
    if (sheetKey === 'blMonitoring') {
      obj = blCanonicalizeRow_(headers, sourceRow);
      obj._row = headerRowNum + i + 1;
      return obj;
    }
    obj = { _row: headerRowNum + i + 1 };
    headers.forEach(function(h, j) {
      var rawVal = row[j];
      if (sheetKey === 'mill' && millSupplyUsesRawNumber_(h)
          && typeof rawVal === 'number' && !isNaN(rawVal)) {
        obj[h] = rawVal;
      } else {
        obj[h] = formatApiCellValue_(sourceRow[j]);
      }
    });
    if (sheetKey === 'mill') mirrorMillQuarterYearOnRead_(obj);
    if (sheetKey === 'mill') mirrorMillGroupNameOnRead_(obj);
    if (sheetKey === 'mill') mirrorMillCompanyNameOnRead_(obj);
    if (sheetKey === 'ttp') {
      mirrorGroupNameOnRead_(obj);
      if (looksLikeTtpYearOrQuarterValue_(obj['GROUP NAME'])) delete obj['GROUP NAME'];
      mirrorGroupNameOnRead_(obj);
      mirrorMillCompanyNameOnRead_(obj);
      mirrorTtpFieldsByPosition_(obj, row, headers);
    }
    if (sheetKey === 'nbl') {
      mirrorNblFieldsOnRead_(obj);
      mirrorNblFieldsByPosition_(obj, row);
    }
    if (sheetKey === 'grievance') {
      mirrorGrvFieldsOnRead_(obj, headers, row);
    }
    return obj;
  }).filter(function(obj) {
    if (sheetKey === 'facilityProfile') return facilityProfileRowHasContent_(obj);
    if (sheetKey === 'mill') return millRowHasContent_(obj);
    if (sheetKey === 'ttp') return ttpRowHasContent_(obj);
    if (sheetKey !== 'blMonitoring') return true;
    return blRowHasContent_(obj);
  });
}

function addRow(sheetKey, data) {
  const sheet   = getSheet(sheetKey);
  let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (sheetKey === 'ttp') {
    const hdr = detectTtpHeaderRow_(sheet);
    if (hdr && hdr.headers && hdr.headers.length) headers = hdr.headers;
  }
  if (sheetKey === 'blMonitoring') {
    ensureBlMonitoringHeaders_();
    const hdr = getBlHeaderInfo_(sheet);
    headers = hdr.headers;
    const newRow = blExpandCanonicalToRow_(headers, data || {}, null);
    const targetRow = Math.max(sheet.getLastRow(), hdr.headerRow) + 1;
    sheet.getRange(targetRow, 1, 1, newRow.length).setValues([newRow]);
    return { success: true };
  }
  if (sheetKey === 'contactSupplier') {
    const now = nowIso_();
    const user = callerEmail_();
    if (!data.updated_at) data.updated_at = now;
    if (!data.updated_by) data.updated_by = user;
    if (!data.approved_at) data.approved_at = now.slice(0, 10);
    if (!data.statusSDD) data.statusSDD = 'Manual';
  }
  if (sheetKey === 'questionnaireMonitoring') {
    ensureQuestionnaireMonitoringHeaders_();
    headers = QUESTIONNAIRE_MONITORING_HEADERS;
    stampQmRowMeta_(data, callerEmail_(), nowIso_());
    const newRow = headers.map(function(h) { return data[h] !== undefined ? data[h] : ''; });
    sheet.appendRow(newRow);
    return { success: true };
  }
  if (sheetKey === 'eudrPotential') {
    ensureEudrPotentialHeaders_();
    headers = EUDR_POTENTIAL_HEADERS;
    EUDR_SHEET_PRESERVE_FIELDS.forEach(function(h) {
      if (data[h] === undefined) data[h] = '';
    });
    const newRow = headers.map(function(h) { return data[h] !== undefined ? data[h] : ''; });
    sheet.appendRow(newRow);
    return { success: true };
  }
  if (sheetKey === 'mill') {
    resolveMillQuarterYearKeys_(data, headers);
    const newRow = headers.map(function(h) { return data[h] !== undefined ? data[h] : ''; });
    forceCoordStrings_(headers, newRow);
    const targetRow = millFindNextAppendRow_(sheet, headers);
    safeInsertRowAt_(sheet, headers, newRow, targetRow);
    return { success: true, row: targetRow };
  }
  if (sheetKey === 'ttp') {
    const hdr = detectTtpHeaderRow_(sheet);
    if (hdr && hdr.headers && hdr.headers.length) headers = hdr.headers;
    const headerRow = hdr && hdr.headerRow ? hdr.headerRow : 1;
    const newRow = headers.map(function(h) { return data[h] !== undefined ? data[h] : ''; });
    forceCoordStrings_(headers, newRow);
    const targetRow = ttpFindNextAppendRow_(sheet, headers, headerRow);
    safeInsertRowAt_(sheet, headers, newRow, targetRow);
    return { success: true, row: targetRow };
  }
  if (sheetKey === 'grievance') {
    ensureGrievanceRiskHeaders_();
    headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    resolveGrvGroupFieldKeys_(data, headers);
  }
  const newRow  = headers.map(function(h) {
    return coerceSheetDateValue_(data[h] !== undefined ? data[h] : '');
  });
  sheet.appendRow(newRow);
  return { success: true };
}

function updateRow(sheetKey, rowNum, data) {
  const sheet = getSheet(sheetKey);
  const r     = Number(rowNum);
  if (!r || r < 2) throw new Error('Invalid row number for update: ' + rowNum);
  if (sheetKey === 'grievance') ensureGrievanceRiskHeaders_();
  let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (sheetKey === 'ttp') {
    const hdr = detectTtpHeaderRow_(sheet);
    if (hdr && hdr.headers && hdr.headers.length) headers = hdr.headers;
  }
  if (sheetKey === 'blMonitoring') {
    const hdr = getBlHeaderInfo_(sheet);
    headers = hdr.headers;
    const current = sheet.getRange(r, 1, 1, headers.length).getValues()[0];
    const updated = blExpandCanonicalToRow_(headers, data || {}, current);
    sheet.getRange(r, 1, 1, updated.length).setValues([updated]);
    return { success: true };
  }
  if (sheetKey === 'contactSupplier') {
    data.updated_at = nowIso_();
    data.updated_by = callerEmail_();
  }
  if (sheetKey === 'questionnaireMonitoring') {
    ensureQuestionnaireMonitoringHeaders_();
    headers = QUESTIONNAIRE_MONITORING_HEADERS;
    stampQmRowMeta_(data, callerEmail_(), nowIso_());
  }
  if (sheetKey === 'eudrPotential') {
    ensureEudrPotentialHeaders_();
    headers = EUDR_POTENTIAL_HEADERS;
    stampEudrRowMeta_(data, callerEmail_(), nowIso_());
  }
  if (sheetKey === 'mill') resolveMillQuarterYearKeys_(data, headers);
  if (sheetKey === 'grievance') resolveGrvGroupFieldKeys_(data, headers);
  const current = sheet.getRange(r, 1, 1, headers.length).getValues()[0];
  const updated = headers.map(function(h, j) {
    var v = data[h] !== undefined ? data[h] : current[j];
    return coerceSheetDateValue_(v);
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

function findSheetTabFuzzy_(ss, tabName) {
  const want = String(tabName || '').trim().toLowerCase();
  if (!want) return null;
  const sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (String(sheets[i].getName()).trim().toLowerCase() === want) return sheets[i];
  }
  return null;
}

function getSheet(sheetKey) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const name = resolveSheetTabName_(sheetKey);
  if (!name) {
    throw new Error('Sheet key not found: ' + sheetKey + '. Ensure the "' + BL_MONITORING_TAB + '" tab exists in the spreadsheet, then redeploy Apps Script (v3-bl-monitoring).');
  }
  let sheet = ss.getSheetByName(name) || findSheetTabFuzzy_(ss, name);
  if (!sheet) {
    throw new Error('Tab not found: "' + name + '". Buat tab dengan nama persis: ' + BL_MONITORING_TAB);
  }
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

/** Physically remove sheet rows (bottom-up so indices stay valid). */
function hardDeleteSheetRows_(ws, sheetRows) {
  if (!ws || !sheetRows || !sheetRows.length) return 0;
  const sorted = sheetRows.slice().sort(function(a, b) { return b - a; });
  const seen = {};
  let deleted = 0;
  sorted.forEach(function(r) {
    const rowNum = parseInt(r, 10);
    if (isNaN(rowNum) || rowNum < 2 || seen[rowNum]) return;
    seen[rowNum] = true;
    ws.deleteRow(rowNum);
    deleted++;
  });
  return deleted;
}

function collectRelSheetRowsForSubmission_(sheetKey, sid) {
  const result = readRelSheet_(sheetKey, true);
  const sheetRows = [];
  result.rows.forEach(function(r) {
    if (String(r.obj['submission_id'] || '').trim() === sid) {
      sheetRows.push(r._sheetRow);
    }
  });
  return { ws: result.ws, sheetRows: sheetRows, rows: result.rows };
}

/** Remove matching rows from legacy flat SDD sheet (if any). */
function deleteLegacySddRowsForSubmission_(mainObj, millObjs, ffbObjs, sid) {
  let legacyDeleted = 0;
  try {
    const sheet = getSheet('sdd');
    const raw = sheet.getDataRange().getValues();
    if (!raw.length) return 0;

    const headers = raw[0];
    const headerIndex = indexByHeader_(headers);
    const rowsToDelete = [];
    const seen = {};

    function addRowNum(n) {
      const rowNum = parseInt(n, 10);
      if (isNaN(rowNum) || rowNum < 2 || seen[rowNum]) return;
      seen[rowNum] = true;
      rowsToDelete.push(rowNum);
    }

    const sidCol = headerIndex['submission_id'];
    if (sidCol !== undefined && sid) {
      for (let i = 1; i < raw.length; i++) {
        if (String(raw[i][sidCol] || '').trim() === sid) addRowNum(i + 1);
      }
    }

    if (!rowsToDelete.length) {
      const payloads = [mainObj].concat(millObjs || []).concat(ffbObjs || []);
      payloads.forEach(function(data) {
        if (!data || typeof data !== 'object') return;
        const idx = findMatchingSddRowIndex_(raw, headers, headerIndex, data);
        if (idx > 0) addRowNum(idx + 1);
      });
    }

    legacyDeleted = hardDeleteSheetRows_(sheet, rowsToDelete);
  } catch (err) {
    console.warn('[deleteLegacySddRowsForSubmission_]', err.message);
  }
  return legacyDeleted;
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

    const wantsMillTtpSync = payload.mill_ttp_sync && typeof payload.mill_ttp_sync === 'object';
    const wantsMillAdded = payload.main
      && String(payload.main.mill_added || '').trim().toLowerCase() === 'true';
    const tmlLineId = String(payload.mill_added_line || '').trim();
    const supplierTypeEarly = String(mainHit.obj['supplier_type'] || mainHit.obj['Supplier Type'] || '').trim().toUpperCase();
    const tmlRowCount = countMeaningfulTmlRowsForSid_(sid);
    const isTraderTtpMirror = wantsMillTtpSync && tmlLineId && supplierTypeEarly === 'TRADER';
    const isMillKcpTtpPerLine = wantsMillTtpSync && tmlLineId
      && (supplierTypeEarly === 'MILL' || supplierTypeEarly === 'KCP');

    if (wantsMillTtpSync && !wantsMillAdded && !tmlLineId) {
      throw new Error('mill_ttp_sync requires mill_added=true or mill_added_line');
    }

    if (supplierTypeEarly === 'TRADER' && wantsMillAdded) {
      throw new Error('TRADER cannot set mill_added=true directly; complete all mills via Task List save');
    }

    if ((supplierTypeEarly === 'MILL' || supplierTypeEarly === 'KCP')
        && wantsMillAdded && tmlRowCount > 0) {
      throw new Error('MILL/KCP cannot set mill_added=true directly when mill list exists; complete each mill via Task List save');
    }

    if (supplierTypeEarly === 'TRADER' && tmlLineId && !wantsMillTtpSync) {
      throw new Error('mill_added_line for TRADER requires mill_ttp_sync in the same request');
    }

    if ((supplierTypeEarly === 'MILL' || supplierTypeEarly === 'KCP')
        && tmlLineId && !wantsMillTtpSync) {
      throw new Error('mill_added_line for MILL/KCP requires mill_ttp_sync in the same request');
    }

    let ttpSync = null;
    let millAddedLineResult = null;

    if (isTraderTtpMirror) {
      assertTraderTmlLineKnown_(sid, tmlLineId);
      const mirrorIdentity = sanitizeMillTtpMirrorFromOnboarding_(payload.mill_ttp_sync);
      ttpSync = syncTtpMirrorTraderMillFromOnboarding_(
        sid, tmlLineId, mirrorIdentity, mainHit.obj, user, now
      );
      assertMillTtpSyncSucceeded_(ttpSync);
      auditLog_('POST', 'mill_ttp_trader_mirror', 'ttp', user, JSON.stringify({
        submission_id: sid,
        tml_line_id: tmlLineId,
        inserted: ttpSync.inserted,
        updated: ttpSync.updated,
      }));
    } else if (isMillKcpTtpPerLine) {
      assertTraderTmlLineKnown_(sid, tmlLineId);
      const identity = sanitizeMillTtpIdentity_(payload.mill_ttp_sync);
      ttpSync = syncTtpFromMillOnboardingForTmlLine_(
        sid, tmlLineId, identity, mainHit.obj, user, now
      );
      assertMillTtpSyncSucceeded_(ttpSync);
      auditLog_('POST', 'mill_ttp_mill_tml_sync', 'ttp', user, JSON.stringify({
        submission_id: sid,
        tml_line_id: tmlLineId,
        inserted: ttpSync.inserted,
        updated: ttpSync.updated,
        total_ffb: ttpSync.total_ffb,
      }));
    } else if (wantsMillTtpSync) {
      const identity = sanitizeMillTtpIdentity_(payload.mill_ttp_sync);
      ttpSync = syncTtpFromMillOnboarding_(sid, identity, mainHit.obj, user, now);
      assertMillTtpSyncSucceeded_(ttpSync);
      auditLog_('POST', 'mill_ttp_sync', 'ttp', user, JSON.stringify({
        submission_id: sid,
        inserted: ttpSync.inserted,
        updated: ttpSync.updated,
        total_ffb: ttpSync.total_ffb,
        skipped: ttpSync.skipped || false,
        reason: ttpSync.reason || '',
      }));
    }

    if (payload.mill_added_line) {
      if (supplierTypeEarly === 'TRADER'
          || supplierTypeEarly === 'MILL'
          || supplierTypeEarly === 'KCP') {
        assertTraderTmlLineKnown_(sid, payload.mill_added_line);
      }
      millAddedLineResult = applyMillAddedLine_(sid, payload.mill_added_line, mainHit, user, now);
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
      }
    }

    const supplierType = String(mainHit.obj['supplier_type'] || 'MILL');

    const millStats = payload.mills !== undefined
      ? upsertChildSheet_('sddMill', sid, payload.mills,    supplierType, user, now)
      : { upserted: 0, inserted: 0, deleted: 0 };

    const ffbStats = payload.ffb_rows !== undefined
      ? upsertChildSheet_('sddFfb',  sid, payload.ffb_rows, supplierType, user, now)
      : { upserted: 0, inserted: 0, deleted: 0 };

    return {
      success       : true,
      submission_id : sid,
      mills         : millStats,
      ffb           : ffbStats,
      ttp_sync      : ttpSync,
      mill_added_line: millAddedLineResult,
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
  const user = callerEmail_();

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    ensureAllRelationalHeaders_();

    const mainPack = collectRelSheetRowsForSubmission_('sddMain', sid);
    if (!mainPack.sheetRows.length) {
      throw new Error('submission_id not found: ' + sid);
    }

    const mainObj = mainPack.rows.filter(function(r) {
      return String(r.obj['submission_id'] || '').trim() === sid;
    }).map(function(r) { return r.obj; })[0] || {};

    const millPack = collectRelSheetRowsForSubmission_('sddMill', sid);
    const ffbPack  = collectRelSheetRowsForSubmission_('sddFfb', sid);
    const millObjs = millPack.rows
      .filter(function(r) { return String(r.obj['submission_id'] || '').trim() === sid; })
      .map(function(r) { return r.obj; });
    const ffbObjs = ffbPack.rows
      .filter(function(r) { return String(r.obj['submission_id'] || '').trim() === sid; })
      .map(function(r) { return r.obj; });

    const millsDeleted = hardDeleteSheetRows_(millPack.ws, millPack.sheetRows);
    const ffbDeleted   = hardDeleteSheetRows_(ffbPack.ws, ffbPack.sheetRows);
    const mainDeleted  = hardDeleteSheetRows_(mainPack.ws, mainPack.sheetRows);
    const legacyDeleted = deleteLegacySddRowsForSubmission_(mainObj, millObjs, ffbObjs, sid);

    return {
      success         : true,
      submission_id   : sid,
      main_deleted    : mainDeleted,
      mills_deleted   : millsDeleted,
      ffb_deleted     : ffbDeleted,
      legacy_deleted  : legacyDeleted,
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

function supplyNormKey_(s) {
  return String(s || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function supplyMillQuarterTok_(row) {
  var q = row['QUARTER'] || row['Quarter'] || row['quarter'] || '';
  q = String(q || '').trim().replace(/^Q/i, '');
  return q;
}

function supplyMillYearTok_(row) {
  return String(row['YEAR'] || row['Year'] || row['year'] || '').trim();
}

function supplySplitFacilitiesGs_(raw) {
  return String(raw || '').trim().split(/[,;\/]+/).map(function(f) {
    return f.trim();
  }).filter(function(f) { return f && !/^total$/i.test(f); });
}

function supplyNormFacilityKeyGs_(s) {
  return supplyNormKey_(s).replace(/[^a-z0-9]/g, '');
}

function supplyNameMatchesGs_(a, b) {
  var na = supplyNormKey_(a);
  var nb = supplyNormKey_(b);
  if (!na || !nb) return false;
  return na === nb;
}

function supplyIdentityMatchesGs_(excelCo, excelMill, sheetCo, sheetMill) {
  if (supplyNameMatchesGs_(excelCo, sheetCo) || supplyNameMatchesGs_(excelCo, sheetMill)) return true;
  if (excelMill && (supplyNameMatchesGs_(excelMill, sheetMill) || supplyNameMatchesGs_(excelMill, sheetCo))) return true;
  return false;
}

function supplyProductSupplyTokensGs_(raw) {
  var ps = String(raw || '').trim().toUpperCase().replace(/\s+/g, '').replace(/[;&/+]+/g, ',').replace(/,+/g, ',').replace(/^,|,$/g, '');
  if (!ps) return [];
  return ps.split(',').filter(Boolean);
}

function supplyProductSupplyAllowsGs_(rawPs, supplyKind) {
  var tokens = supplyProductSupplyTokensGs_(rawPs);
  if (!tokens.length) return true;
  var hasCpo = tokens.indexOf('CPO') !== -1;
  var hasPk = tokens.indexOf('PK') !== -1;
  if (supplyKind === 'PK') return hasPk;
  return hasCpo;
}

function supplyFacilityMatchesGs_(plant, facRaw) {
  var want = supplyNormFacilityKeyGs_(plant);
  if (!want) return false;
  var list = supplySplitFacilitiesGs_(facRaw);
  if (!list.length) return false;
  for (var i = 0; i < list.length; i++) {
    if (supplyNormFacilityKeyGs_(list[i]) === want) return true;
  }
  return false;
}

/**
 * Find Mill Onboarding Profile row by COMPANY NAME only.
 */
function findMillRowForSupplySubmit_(millData, millHeaders, row) {
  var targetRow = Number(row.target_mill_row || row._mill_row || 0);
  if (targetRow >= 2) return targetRow;

  var excelCo = String(row['COMPANY NAME'] || '').trim();
  if (!excelCo) return 0;

  var coCol = millHeaders.indexOf('COMPANY NAME');
  if (coCol < 0) return 0;

  for (var r = 1; r < millData.length; r++) {
    var sheetCo = String(millData[r][coCol] || '').trim();
    if (supplyNameMatchesGs_(excelCo, sheetCo)) return r + 1;
  }
  return 0;
}

function submitSupplyDraft_(batchId, rows) {
  if (!batchId) throw new Error('batch_id required');
  ensureSupplyDraftHeaders_();

  var millSheet   = getSheet('mill');
  var millData    = millSheet.getDataRange().getValues();
  var millHeaders = millData[0].map(function(h) { return String(h || '').trim(); });

  var pctColCpo = millHeaders.indexOf('PERCENTAGE SUPPLY CPO');
  var pctColPk  = millHeaders.indexOf('PERCENTAGE SUPPLY PK');
  if (pctColCpo < 0) {
    millSheet.insertColumnAfter(millSheet.getLastColumn());
    pctColCpo = millSheet.getLastColumn() - 1;
    millSheet.getRange(1, pctColCpo + 1).setValue('PERCENTAGE SUPPLY CPO');
    millHeaders = millSheet.getDataRange().getValues()[0].map(function(h) { return String(h || '').trim(); });
    pctColCpo = millHeaders.indexOf('PERCENTAGE SUPPLY CPO');
  }
  if (pctColPk < 0) {
    millSheet.insertColumnAfter(millSheet.getLastColumn());
    millSheet.getRange(1, millSheet.getLastColumn()).setValue('PERCENTAGE SUPPLY PK');
    millHeaders = millSheet.getDataRange().getValues()[0].map(function(h) { return String(h || '').trim(); });
    pctColPk = millHeaders.indexOf('PERCENTAGE SUPPLY PK');
  }

  var draftSheet   = getSheet('supplyDraft');
  var draftData    = draftSheet.getDataRange().getValues();
  var draftHeaders = draftData[0].map(function(h) { return String(h || '').trim(); });
  var draftIdColD  = draftHeaders.indexOf('draft_id');
  var statusColD   = draftHeaders.indexOf('status');
  var now          = nowIso_();
  var submitted    = 0;
  var errors       = [];

  rows.forEach(function(row) {
    var matchStatus = String(row.match_status || '').trim().toLowerCase();
    if (matchStatus !== 'matched') {
      errors.push((row['COMPANY NAME'] || 'row') + ': bukan status Matched');
      return;
    }

    var sheetRowNum = findMillRowForSupplySubmit_(millData, millHeaders, row);
    if (!sheetRowNum) {
      errors.push((row['COMPANY NAME'] || '') + ' / ' + (row['MILL NAME'] || '') + ': tidak ditemukan di Mill Onboarding Profile');
      return;
    }

    var supplyType = String(row.supply_type || row.SUPPLY_TYPE || 'CPO').trim().toUpperCase();
    var pctCpo = row['PERCENTAGE SUPPLY CPO'];
    var pctPk  = row['PERCENTAGE SUPPLY PK'];
    var hasCpo = pctCpo !== undefined && pctCpo !== null && String(pctCpo).trim() !== '';
    var hasPk  = pctPk !== undefined && pctPk !== null && String(pctPk).trim() !== '';
    if (!hasCpo && supplyType.indexOf('CPO') >= 0) {
      pctCpo = row.SUPPLY_PERCENTAGE || '';
      hasCpo = String(pctCpo).trim() !== '';
    }
    if (!hasPk && supplyType.indexOf('PK') >= 0) {
      pctPk = row.SUPPLY_PERCENTAGE || '';
      hasPk = String(pctPk).trim() !== '';
    }

    var patch = {};
    if (hasCpo || supplyType === 'CPO') {
      patch['PERCENTAGE SUPPLY CPO'] = hasCpo ? pctCpo : '';
      if (row['FACILITY NAME CPO']) patch['FACILITY NAME CPO'] = row['FACILITY NAME CPO'];
    }
    if (hasPk || supplyType === 'PK') {
      patch['PERCENTAGE SUPPLY PK'] = hasPk ? pctPk : '';
      if (row['FACILITY NAME PK']) patch['FACILITY NAME PK'] = row['FACILITY NAME PK'];
    }
    if (!hasCpo && !hasPk) {
      var pctField = supplyType === 'PK' ? 'PERCENTAGE SUPPLY PK' : 'PERCENTAGE SUPPLY CPO';
      var pctVal = row[pctField];
      if (pctVal === undefined || pctVal === null || String(pctVal).trim() === '') {
        pctVal = row.SUPPLY_PERCENTAGE || '';
      }
      patch[pctField] = pctVal;
      if (supplyType === 'PK' && row['FACILITY NAME PK']) {
        patch['FACILITY NAME PK'] = row['FACILITY NAME PK'];
      }
      if (supplyType === 'CPO' && row['FACILITY NAME CPO']) {
        patch['FACILITY NAME CPO'] = row['FACILITY NAME CPO'];
      }
    }

    var psTokens = [];
    if (hasCpo || supplyType.indexOf('CPO') >= 0) psTokens.push('CPO');
    if (hasPk || supplyType.indexOf('PK') >= 0) psTokens.push('PK');
    if (psTokens.length) {
      patch['PRODUCT SUPPLY'] = psTokens.length > 1 ? 'CPO, PK' : psTokens[0];
    } else if (row['PRODUCT SUPPLY']) {
      patch['PRODUCT SUPPLY'] = row['PRODUCT SUPPLY'];
    }

    var qtyCpo = row['SUPPLY CPO'];
    var qtyPk  = row['SUPPLY PK'];
    if ((qtyCpo === undefined || qtyCpo === null || String(qtyCpo).trim() === '') && row.SUPPLY_QTY && (hasCpo || supplyType.indexOf('CPO') >= 0)) {
      qtyCpo = row.SUPPLY_QTY;
    }
    if ((qtyPk === undefined || qtyPk === null || String(qtyPk).trim() === '') && row.SUPPLY_QTY && (hasPk || supplyType.indexOf('PK') >= 0)) {
      qtyPk = row.SUPPLY_QTY;
    }
    if (qtyCpo !== undefined && qtyCpo !== null && String(qtyCpo).trim() !== '' && (hasCpo || supplyType.indexOf('CPO') >= 0)) {
      patch['SUPPLY CPO'] = qtyCpo;
    }
    if (qtyPk !== undefined && qtyPk !== null && String(qtyPk).trim() !== '' && (hasPk || supplyType.indexOf('PK') >= 0)) {
      patch['SUPPLY PK'] = qtyPk;
    }

    try {
      updateRow('mill', sheetRowNum, patch);
    } catch (err) {
      errors.push((row['COMPANY NAME'] || '') + ': ' + err.message);
      return;
    }

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

  if (errors.length && !submitted) {
    throw new Error(errors.slice(0, 5).join('; '));
  }
  return { success: true, submitted: submitted, batch_id: batchId, errors: errors };
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

// ═══════════════════════════════════════════════════════════
//  SUPPLIED CPO — dynamic sheet detection
// ═══════════════════════════════════════════════════════════

/**
 * Returns names of all sheets whose name starts with "SUPPLIED CPO"
 * (case-insensitive), sorted alphabetically.
 * Example returns: ["SUPPLIED CPO 2025", "SUPPLIED CPO Q1 2026", "SUPPLIED CPO Q2 2026"]
 */
function listSuppliedCpoSheets_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheets()
    .map(function(s) { return s.getName(); })
    .filter(function(n) { return n.trim().toUpperCase().startsWith('SUPPLIED CPO'); })
    .sort();
}

/**
 * Parse a 2D values array looking for PLANT / SELLER / SUM of Qty Kg columns.
 * Skips subtotal rows (PLANT or SELLER contains "total").
 * Applies fill-down for PLANT: if a row has a blank PLANT cell, inherits
 * the last non-blank PLANT value (common in grouped spreadsheet layouts).
 */
function parseSuppliedCpoValues_(values, sheetLabel, out) {
  if (!values || !values.length) return;

  var headerRow = 0;
  var headers = [];
  var scanMax = Math.min(15, values.length);
  for (var r = 0; r < scanMax; r++) {
    var row = values[r].map(function(c) { return String(c).trim(); });
    var hasPlant = row.some(function(h) { return /^plant$/i.test(h); });
    var hasSeller = row.some(function(h) { return /^seller$/i.test(h); });
    if (hasPlant && hasSeller) {
      headerRow = r;
      headers = row;
      break;
    }
  }
  if (!headers.length) {
    headers = values[0].map(function(h) { return String(h).trim(); });
  }

  var plantCol = -1;
  var sellerCol = -1;
  var qtyCol = -1;
  headers.forEach(function(h, j) {
    // Use FIRST PLANT column only (column A). Sheet also has a summary PLANT in col F.
    if (plantCol < 0 && /^plant$/i.test(h)) plantCol = j;
    if (sellerCol < 0 && /^seller$/i.test(h)) sellerCol = j;
    if (qtyCol < 0 && (/sum.*qty.*kg/i.test(h) || (/qty/i.test(h) && /kg/i.test(h)))) qtyCol = j;
  });

  var lastPlant = '';
  for (var i = headerRow + 1; i < values.length; i++) {
    var row = values[i];
    var plant = plantCol >= 0 ? String(row[plantCol] || '').trim() : '';
    var seller = sellerCol >= 0 ? String(row[sellerCol] || '').trim() : '';

    // Fill-down: blank PLANT inherits last non-blank PLANT
    if (plant) {
      lastPlant = plant;
    } else {
      plant = lastPlant;
    }

    if (!plant || !seller) continue;
    if (/total/i.test(plant) || /total/i.test(seller)) continue;

    var qty = qtyCol >= 0 ? row[qtyCol] : '';
    out.push({
      PLANT: plant,
      SELLER: seller,
      'SUM of Qty Kg': qty,
      _sheet: sheetLabel
    });
  }
}

/**
 * Pre-fill merged cells in a mutable 2D values array.
 * Google Sheets getValues() returns the cell value only for the first cell
 * of a merged range; all subsequent merged cells return ''. This function
 * propagates the first-cell value to every cell in the same merge, so
 * downstream parsers don't need to guess the PLANT for blank continuation rows.
 */
function fillMergedCells_(sheet, values) {
  try {
    var mergedRanges = sheet.getMergedRanges();
    mergedRanges.forEach(function(range) {
      var rStart = range.getRow() - 1;      // 0-indexed
      var rEnd   = range.getLastRow() - 1;
      var cStart = range.getColumn() - 1;
      var cEnd   = range.getLastColumn() - 1;
      // Only handle single-column merges spanning multiple rows
      if (cStart === cEnd && rEnd > rStart && rStart < values.length) {
        var fillVal = values[rStart][cStart];
        for (var r = rStart + 1; r <= rEnd && r < values.length; r++) {
          values[r][cStart] = fillVal;
        }
      }
    });
  } catch (e) {
    // getMergedRanges can fail on some sheet types — safe to ignore
  }
  return values;
}

/**
 * Returns all data rows from the requested SUPPLIED CPO sheet(s).
 * Uses getDisplayValues() so merged PLANT cells are read correctly —
 * every cell in a merged range returns the displayed value, not just the first.
 */
function getSuppliedCpoData_(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets;
  if (sheetName && sheetName !== '*') {
    var s = ss.getSheetByName(sheetName);
    if (!s) return { error: 'Sheet not found: ' + sheetName };
    sheets = [s];
  } else {
    sheets = ss.getSheets().filter(function(s) {
      return s.getName().trim().toUpperCase().startsWith('SUPPLIED CPO');
    });
  }

  var allRows = [];
  sheets.forEach(function(sheet) {
    var sName  = sheet.getName();
    var values = readSuppliedSheetValues_(sheet);
    parseSuppliedCpoValues_(values, sName, allRows);
  });
  return allRows;
}

// ═══════════════════════════════════════════════════════════
//  SUPPLIED PK — dynamic sheet detection (mirrors CPO logic)
// ═══════════════════════════════════════════════════════════

/**
 * Returns names of all sheets whose name starts with "SUPPLIED PK"
 * (case-insensitive), sorted alphabetically.
 * Example: ["Supplied PK 2025", "Supplied PK Q1 2026"]
 */
function listSuppliedPkSheets_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheets()
    .map(function(s) { return s.getName(); })
    .filter(function(n) { return n.trim().toUpperCase().startsWith('SUPPLIED PK'); })
    .sort();
}

/**
 * Parse a 2D values array for PK sheets: same columns as CPO
 * (PLANT / SELLER / SUM of Qty Kg). Skips subtotal / empty rows.
 * The GROUP column is intentionally ignored.
 * Applies fill-down for PLANT: blank PLANT cells inherit the last
 * non-blank PLANT value (common grouped spreadsheet layout).
 */
function parseSuppliedPkValues_(values, sheetLabel, out) {
  if (!values || !values.length) return;

  var headerRow = 0;
  var headers = [];
  var scanMax = Math.min(15, values.length);
  for (var r = 0; r < scanMax; r++) {
    var row = values[r].map(function(c) { return String(c).trim(); });
    var hasPlant  = row.some(function(h) { return /^plant$/i.test(h); });
    var hasSeller = row.some(function(h) { return /^seller$/i.test(h); });
    if (hasPlant && hasSeller) { headerRow = r; headers = row; break; }
  }
  if (!headers.length) {
    headers = values[0].map(function(h) { return String(h).trim(); });
  }

  var plantCol = -1, sellerCol = -1, qtyCol = -1;
  headers.forEach(function(h, j) {
    // Use FIRST PLANT column only (column A). Summary table repeats PLANT in col F.
    if (plantCol < 0 && /^plant$/i.test(h)) plantCol = j;
    if (sellerCol < 0 && /^seller$/i.test(h)) sellerCol = j;
    if (qtyCol < 0 && (/sum.*qty.*kg/i.test(h) || (/qty/i.test(h) && /kg/i.test(h)))) qtyCol = j;
  });

  var lastPlant = '';
  for (var i = headerRow + 1; i < values.length; i++) {
    var row = values[i];
    var plant  = plantCol  >= 0 ? String(row[plantCol]  || '').trim() : '';
    var seller = sellerCol >= 0 ? String(row[sellerCol] || '').trim() : '';

    // Fill-down: blank PLANT inherits last non-blank PLANT
    if (plant) {
      lastPlant = plant;
    } else {
      plant = lastPlant;
    }

    if (!plant || !seller) continue;
    if (/total/i.test(plant) || /total/i.test(seller)) continue;

    var qty = qtyCol >= 0 ? row[qtyCol] : '';
    out.push({ PLANT: plant, SELLER: seller, 'SUM of Qty Kg': qty, _sheet: sheetLabel });
  }
}

/**
 * Read columns A–D only (PLANT, GROUP, SELLER, Qty) — ignores summary block in F+.
 */
function readSuppliedSheetValues_(sheet) {
  var numRows = sheet.getLastRow();
  var numCols = Math.min(4, sheet.getLastColumn());
  if (numRows < 2 || numCols < 2) return [];
  var display = sheet.getRange(1, 1, numRows, numCols).getDisplayValues();
  var numeric = sheet.getRange(1, 1, numRows, numCols).getValues();
  return display.map(function(dRow, r) {
    return dRow.map(function(cell, c) {
      var num = numeric[r][c];
      return (typeof num === 'number') ? num : cell;
    });
  });
}

/**
 * Returns all data rows from the requested SUPPLIED PK sheet(s).
 * Uses getDisplayValues() so merged PLANT cells are read correctly.
 */
function getSuppliedPkData_(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets;
  if (sheetName && sheetName !== '*') {
    var s = ss.getSheetByName(sheetName);
    if (!s) return { error: 'Sheet not found: ' + sheetName };
    sheets = [s];
  } else {
    sheets = ss.getSheets().filter(function(s) {
      return s.getName().trim().toUpperCase().startsWith('SUPPLIED PK');
    });
  }

  var allRows = [];
  sheets.forEach(function(sheet) {
    var sName  = sheet.getName();
    var values = readSuppliedSheetValues_(sheet);
    parseSuppliedPkValues_(values, sName, allRows);
  });
  return allRows;
}

/**
 * Proxy static map image for Facility Performance PDF export.
 * Fetches server-side to avoid browser CORS limits.
 */
function getFacilityMapImage_(params) {
  params = params || {};
  var lat = parseFloat(String(params.lat || '').replace(',', '.'));
  var lng = parseFloat(String(params.lng || '').replace(',', '.'));
  var zoom = parseInt(params.zoom || 13, 10);
  var w = parseInt(params.w || 640, 10);
  var h = parseInt(params.h || 320, 10);
  if (isNaN(zoom) || zoom < 1) zoom = 13;
  if (zoom > 18) zoom = 18;
  if (isNaN(w) || w < 200) w = 640;
  if (w > 1280) w = 1280;
  if (isNaN(h) || h < 150) h = 320;
  if (h > 720) h = 720;
  if (isNaN(lat) || isNaN(lng) || Math.abs(lat) > 90 || Math.abs(lng) > 180) {
    return { success: false, error: 'Invalid coordinates' };
  }

  var latStr = lat.toFixed(6);
  var lngStr = lng.toFixed(6);
  var urls = [
    'https://staticmap.openstreetmap.de/staticmap.php?center=' + latStr + ',' + lngStr
      + '&zoom=' + zoom + '&size=' + w + 'x' + h
      + '&maptype=mapnik&markers=' + latStr + ',' + lngStr + ',red-pushpin'
  ];

  var lastErr = '';
  for (var i = 0; i < urls.length; i++) {
    try {
      var resp = UrlFetchApp.fetch(urls[i], { muteHttpExceptions: true, followRedirects: true });
      var code = resp.getResponseCode();
      if (code !== 200) {
        lastErr = 'HTTP ' + code;
        continue;
      }
      var ct = String(resp.getHeaders()['Content-Type'] || '');
      if (ct.indexOf('image') === -1) {
        lastErr = 'Response is not an image';
        continue;
      }
      var b64 = Utilities.base64Encode(resp.getBlob().getBytes());
      return { success: true, mime: 'image/png', dataUrl: 'data:image/png;base64,' + b64 };
    } catch (err) {
      lastErr = String(err && err.message ? err.message : err);
    }
  }
  return { success: false, error: lastErr || 'Map fetch failed' };
}
