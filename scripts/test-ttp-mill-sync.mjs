/**
 * Unit tests for SDD → Mill Onboarding → Traceability Data sync logic.
 * Mirrors pure functions from GoogleAppsScript-backend-v3-full.gs.
 * Run: node scripts/test-ttp-mill-sync.mjs
 */
import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const ROOT = join(__dirname, '..');

const TTP_MILL_HEADER_LINE_ID = '__mill_header__';
const TTP_TRADER_TML_LINE_PREFIX = 'trader_tml_';

function ttpTraderMirrorLineId_(tmlLineId) {
  return TTP_TRADER_TML_LINE_PREFIX + String(tmlLineId || '').trim();
}

function normalizeSddDecisionLabel_(raw) {
  const s = String(raw || '').trim().toLowerCase();
  if (s === 'approve' || s === 'approved') return 'APPROVED';
  if (s === 'hold' || s === 'on hold') return 'ON HOLD';
  if (s === 'reject' || s === 'rejected') return 'REJECTED';
  return String(raw || '').trim().toUpperCase();
}

function isMeaningfulSddFfbRow_(ffb) {
  if (!ffb || typeof ffb !== 'object') return false;
  if (String(ffb['is_deleted'] || '') === '1') return false;
  const name = String(ffb['FFB - Supplier Name'] || '').trim();
  const grp = String(ffb['FFB - Supplier Group Name'] || '').trim();
  return !!(name || grp);
}

function isBlankTtpCell_(v) {
  const s = String(v === undefined || v === null ? '' : v).trim();
  return !s || s === '—' || s === '-';
}

function buildTtpPatchFromSddFfb_(main, ffb, sid, user, now, millIdentity) {
  const identity = millIdentity || {};
  const supplierType = String(main['supplier_type'] || main['Supplier Type'] || '').trim();
  let millName = String(identity['MILL NAME'] || main['Mill Name'] || '').trim();
  if (!millName && /trader/i.test(supplierType)) {
    millName = String(identity['COMPANY NAME'] || main['Company Name'] || '').trim();
  }
  let lat = String(ffb['FFB - Latitude'] || '').trim();
  let lng = String(ffb['FFB - Longitude'] || '').trim();
  if (isBlankTtpCell_(lat)) lat = String(main['Latitude'] || '').trim();
  if (isBlankTtpCell_(lng)) lng = String(main['Longitude'] || '').trim();
  return {
    'GROUP NAME': String(identity['GROUP NAME'] || main['Group Name'] || '').trim(),
    'COMPANY NAME': String(identity['COMPANY NAME'] || main['Company Name'] || '').trim(),
    'MILL NAME': millName,
    'UML ID': String(identity['UML ID'] || '').trim(),
    'FFB SUPPLIER GROUP NAME': String(ffb['FFB - Supplier Group Name'] || '').trim(),
    'FFB SUPPLIER NAME': String(ffb['FFB - Supplier Name'] || '').trim(),
    'CATEGORY': String(ffb['FFB - Supplier Category'] || '').trim(),
    'VILLAGE': String(ffb['FFB - Village'] || '').trim(),
    'submission_id': sid,
    'ffb_line_id': String(ffb['line_id'] || '').trim(),
    'supplier_type': supplierType,
    'synced_at': now,
    'synced_by': user,
  };
}

function buildTtpMillHeaderPatch_(main, millIdentity, sid, user, now) {
  const identity = millIdentity || {};
  const supplierType = String(main['supplier_type'] || main['Supplier Type'] || '').trim();
  let millName = String(identity['MILL NAME'] || main['Mill Name'] || '').trim();
  if (!millName && /trader/i.test(supplierType)) {
    millName = String(identity['COMPANY NAME'] || main['Company Name'] || '').trim();
  }
  return {
    'GROUP NAME': String(identity['GROUP NAME'] || main['Group Name'] || '').trim(),
    'COMPANY NAME': String(identity['COMPANY NAME'] || main['Company Name'] || '').trim(),
    'MILL NAME': millName,
    'UML ID': String(identity['UML ID'] || '').trim(),
    'submission_id': sid,
    'ffb_line_id': TTP_MILL_HEADER_LINE_ID,
    'supplier_type': supplierType,
    'synced_at': now,
    'synced_by': user,
  };
}

function shouldSyncTtp_(mainObj, millIdentity) {
  const decision = normalizeSddDecisionLabel_(
    mainObj['statusSDD'] || mainObj['statusBossDecision'] || ''
  );
  if (decision !== 'APPROVED') {
    return { ok: false, reason: 'not_approved' };
  }
  const scrSt = String(mainObj['SCR - Screening Status'] || '').trim().toLowerCase();
  if (scrSt !== 'submitted') {
    return { ok: false, reason: 'not_submitted' };
  }
  const supplierType = String(mainObj['supplier_type'] || '').trim().toUpperCase();
  if (supplierType !== 'MILL' && supplierType !== 'KCP') {
    return { ok: false, reason: 'supplier_type_not_mill_or_kcp' };
  }
  try {
    sanitizeMillTtpIdentity_(millIdentity);
  } catch (e) {
    return { ok: false, reason: e.message };
  }
  return { ok: true };
}

function sanitizeMillTtpIdentity_(raw) {
  if (!raw || typeof raw !== 'object') throw new Error('mill_ttp_sync identity is required');
  const MAX = 500;
  const clip = (v) => String(v === undefined || v === null ? '' : v).trim().slice(0, MAX);
  const out = {
    'GROUP NAME': clip(raw['GROUP NAME']),
    'COMPANY NAME': clip(raw['COMPANY NAME']),
    'MILL NAME': clip(raw['MILL NAME']),
    'UML ID': clip(raw['UML ID']),
  };
  if (!out['GROUP NAME']) throw new Error('GROUP NAME is required for Traceability sync');
  if (!out['COMPANY NAME']) throw new Error('COMPANY NAME is required for Traceability sync');
  if (!out['MILL NAME']) throw new Error('MILL NAME is required for Traceability sync');
  if (!out['UML ID']) throw new Error('UML ID is required for Traceability sync');
  return out;
}

function sanitizeMillTtpMirrorFromOnboarding_(raw) {
  const base = sanitizeMillTtpIdentity_(raw);
  const MAX = 500;
  const clip = (v) => String(v === undefined || v === null ? '' : v).trim().slice(0, MAX);
  const out = Object.assign({}, base);
  const lat = clip(raw['LAT'] || raw['Latitude']);
  const lng = clip(raw['LONG'] || raw['Longitude']);
  if (lat) out['LAT'] = lat;
  if (lng) out['LONG'] = lng;
  return out;
}

function buildTtpTraderMillMirrorPatch_(millIdentity, sid, tmlLineId, user, now) {
  const identity = millIdentity || {};
  return {
    'GROUP NAME': String(identity['GROUP NAME'] || '').trim(),
    'COMPANY NAME': String(identity['COMPANY NAME'] || '').trim(),
    'MILL NAME': String(identity['MILL NAME'] || '').trim(),
    'UML ID': String(identity['UML ID'] || '').trim(),
    'LAT': String(identity['LAT'] || '').trim(),
    'LONG': String(identity['LONG'] || '').trim(),
    'submission_id': sid,
    'ffb_line_id': ttpTraderMirrorLineId_(tmlLineId),
    'supplier_type': 'TRADER',
    'synced_at': now,
    'synced_by': user,
  };
}

function shouldSyncTraderTtpMirror_(mainObj, millIdentity, tmlLineId) {
  const decision = normalizeSddDecisionLabel_(
    mainObj['statusSDD'] || mainObj['statusBossDecision'] || ''
  );
  if (decision !== 'APPROVED') return { ok: false, reason: 'not_approved' };
  const scrSt = String(mainObj['SCR - Screening Status'] || '').trim().toLowerCase();
  if (scrSt !== 'submitted') return { ok: false, reason: 'not_submitted' };
  const supplierType = String(mainObj['supplier_type'] || '').trim().toUpperCase();
  if (supplierType !== 'TRADER') return { ok: false, reason: 'not_trader' };
  if (!String(tmlLineId || '').trim()) return { ok: false, reason: 'missing_tml_line_id' };
  try {
    sanitizeMillTtpMirrorFromOnboarding_(millIdentity);
  } catch (e) {
    return { ok: false, reason: e.message };
  }
  return { ok: true };
}

// ── Test runner ─────────────────────────────────────────────
let passed = 0;
let failed = 0;

function assert(cond, msg) {
  if (cond) {
    passed++;
    console.log('  ✓ ' + msg);
  } else {
    failed++;
    console.error('  ✗ ' + msg);
  }
}

function assertEq(actual, expected, msg) {
  const a = JSON.stringify(actual);
  const e = JSON.stringify(expected);
  assert(a === e, msg + (a !== e ? ' — got ' + a + ' expected ' + e : ''));
}

console.log('\n=== TTP Mill Onboarding Sync Tests ===\n');

const mainKcp = {
  'Group Name': 'SDD Group Old',
  'Company Name': 'SDD Company Old',
  'Mill Name': 'SEGATI KCP',
  supplier_type: 'KCP',
  statusSDD: 'Approved',
  'SCR - Screening Status': 'submitted',
};
const millIdentity = {
  'GROUP NAME': 'AAA',
  'COMPANY NAME': 'SARI DUMAI SEJATI',
  'MILL NAME': 'SEGATI KCP',
  'UML ID': 'UML-12345',
};
const ffb1 = {
  'line_id': 'ffb-001',
  'FFB - Supplier Group Name': 'TOGAR',
  'FFB - Supplier Name': 'BATU AMPAR',
  'FFB - Supplier Category': 'Own Estate',
  'FFB - Village': 'LANGGAM',
  'FFB - Sub District': 'PELALAWAN',
};

console.log('1. FFB row patch uses Mill Onboarding identity (header fields)');
const ffbPatch = buildTtpPatchFromSddFfb_(mainKcp, ffb1, 'SID-1', 'tester', '2026-06-15', millIdentity);
assertEq(ffbPatch['GROUP NAME'], 'AAA', 'GROUP NAME from mill onboarding');
assertEq(ffbPatch['COMPANY NAME'], 'SARI DUMAI SEJATI', 'COMPANY NAME from mill onboarding');
assertEq(ffbPatch['MILL NAME'], 'SEGATI KCP', 'MILL NAME from mill onboarding');
assertEq(ffbPatch['UML ID'], 'UML-12345', 'UML ID from mill onboarding (manual)');
assertEq(ffbPatch['FFB SUPPLIER GROUP NAME'], 'TOGAR', 'FFB group from SDD');
assertEq(ffbPatch['FFB SUPPLIER NAME'], 'BATU AMPAR', 'FFB name from SDD');
assertEq(ffbPatch['VILLAGE'], 'LANGGAM', 'FFB village from SDD');
assert(ffbPatch['ffb_line_id'] === 'ffb-001', 'FFB line id preserved');

console.log('\n2. Header-only patch when no FFB suppliers');
const headerPatch = buildTtpMillHeaderPatch_(mainKcp, millIdentity, 'SID-1', 'tester', '2026-06-15');
assertEq(headerPatch['GROUP NAME'], 'AAA', 'header GROUP NAME');
assertEq(headerPatch['COMPANY NAME'], 'SARI DUMAI SEJATI', 'header COMPANY NAME');
assertEq(headerPatch['MILL NAME'], 'SEGATI KCP', 'header MILL NAME');
assertEq(headerPatch['UML ID'], 'UML-12345', 'header UML ID');
assertEq(headerPatch['ffb_line_id'], TTP_MILL_HEADER_LINE_ID, 'header stub line id');
assert(!headerPatch['FFB SUPPLIER NAME'], 'no FFB supplier on header stub');

console.log('\n3. Sync gate: approved + submitted + MILL/KCP + required identity');
assert(shouldSyncTtp_(mainKcp, millIdentity).ok, 'KCP approved → sync allowed');
assert(shouldSyncTtp_({ ...mainKcp, supplier_type: 'MILL' }, millIdentity).ok, 'MILL approved → sync allowed');
assert(!shouldSyncTtp_({ ...mainKcp, supplier_type: 'TRADER' }, millIdentity).ok, 'TRADER → skip');
assert(!shouldSyncTtp_({ ...mainKcp, statusSDD: 'Hold' }, millIdentity).ok, 'not approved → skip');
assert(!shouldSyncTtp_({ ...mainKcp, 'SCR - Screening Status': 'Draft' }, millIdentity).ok, 'not submitted → skip');

console.log('\n3c. TRADER mirror patch (identity from Mill Onboarding, one row per TML line)');
const mainTrader = {
  supplier_type: 'TRADER',
  statusSDD: 'Approved',
  'SCR - Screening Status': 'submitted',
};
const traderMirrorIdentity = {
  ...millIdentity,
  'LAT': '-1.234',
  'LONG': '101.567',
};
const traderPatch = buildTtpTraderMillMirrorPatch_(traderMirrorIdentity, 'SID-T1', 'tml-42', 'tester', '2026-06-15');
assertEq(traderPatch['GROUP NAME'], 'AAA', 'trader mirror GROUP NAME');
assertEq(traderPatch['MILL NAME'], 'SEGATI KCP', 'trader mirror MILL NAME');
assertEq(traderPatch['ffb_line_id'], 'trader_tml_tml-42', 'trader mirror line id prefix');
assertEq(traderPatch['supplier_type'], 'TRADER', 'trader mirror supplier type');
assertEq(traderPatch['LAT'], '-1.234', 'trader mirror LAT');
assert(!traderPatch['FFB SUPPLIER NAME'], 'trader mirror has no FFB fields');
assert(shouldSyncTraderTtpMirror_(mainTrader, traderMirrorIdentity, 'tml-42').ok, 'TRADER mirror allowed');
assert(!shouldSyncTraderTtpMirror_(mainKcp, traderMirrorIdentity, 'tml-42').ok, 'KCP cannot use trader mirror gate');

console.log('\n3b. Identity validation (required fields + max length)');
try {
  sanitizeMillTtpIdentity_({ 'GROUP NAME': '', 'COMPANY NAME': 'X', 'MILL NAME': 'Y', 'UML ID': 'Z' });
  assert(false, 'empty GROUP NAME should throw');
} catch (e) {
  assert(e.message.includes('GROUP NAME'), 'rejects empty GROUP NAME');
}
try {
  sanitizeMillTtpIdentity_({ 'GROUP NAME': 'A', 'COMPANY NAME': 'B', 'MILL NAME': 'C' });
  assert(false, 'missing UML ID should throw');
} catch (e) {
  assert(e.message.includes('UML ID'), 'rejects missing UML ID');
}
assert(sanitizeMillTtpIdentity_(millIdentity)['UML ID'] === 'UML-12345', 'sanitized identity preserved');

console.log('\n4. Meaningful FFB filter');
assert(isMeaningfulSddFfbRow_(ffb1), 'FFB with supplier name is meaningful');
assert(!isMeaningfulSddFfbRow_({ 'FFB - Supplier Name': '', 'FFB - Supplier Group Name': '' }), 'empty FFB skipped');

console.log('\n5. Source file static checks');
const gs = readFileSync(join(ROOT, 'scripts/GoogleAppsScript-backend-v3-full.gs'), 'utf8');
const mainJs = readFileSync(join(ROOT, 'src/main.js'), 'utf8');

assert(!gs.includes('syncTtpFromApprovedSubmission_'), 'approve-time TTP sync removed from backend');
assert(gs.includes('syncTtpFromMillOnboarding_'), 'mill onboarding TTP sync present in backend');
assert(gs.includes('syncTtpMirrorTraderMillFromOnboarding_'), 'trader mirror TTP sync present in backend');
assert(gs.includes('payload.mill_ttp_sync'), 'updateSubmission accepts mill_ttp_sync');
assert(!gs.includes('syncTtpFromApprovedSubmission_(sid'), 'no call to old approve sync');

assert(mainJs.includes("mill_ttp_sync: ttpMirror"), 'frontend sends mill_ttp_sync on trader task list save');
assert(mainJs.includes("mill_ttp_sync: ttpIdentity"), 'frontend sends mill_ttp_sync on MILL/KCP task list save');
assert(mainJs.includes("'UML ID': String(data['UML ID']"), 'frontend passes UML ID from mill modal');
assert(mainJs.includes("'GROUP NAME': String(data['GROUP NAME']"), 'frontend passes GROUP NAME');
assert(mainJs.includes("'COMPANY NAME': String(data['COMPANY NAME']"), 'frontend passes COMPANY NAME');
assert(mainJs.includes("'MILL NAME': String(data['MILL NAME']"), 'frontend passes MILL NAME');
assert(gs.includes('sanitizeMillTtpIdentity_'), 'backend sanitizes mill identity');
assert(gs.includes('sanitizeMillTtpMirrorFromOnboarding_'), 'backend sanitizes trader mirror payload');
assert(gs.includes('assertMillTtpSyncSucceeded_'), 'backend asserts sync success before mill_added');
assert(gs.includes('mill_ttp_sync requires mill_added=true or mill_added_line'), 'mill_ttp_sync gated by mill_added or line');
assert(gs.includes('trader_tml_'), 'trader mirror line id prefix in backend');
assert(gs.includes('not_submitted'), 'backend checks SCR submitted status');
assert(mainJs.includes("Fill Group Name, Company Name, Mill Name, and UML ID"), 'frontend validates before save');
assert(mainJs.includes('Mill saved but Traceability sync failed'), 'frontend shows sync failure to user');
assert(mainJs.includes('mirrored to Traceability Data'), 'frontend trader success toast mentions TTP');

console.log('\n6. Mill modal includes UML ID field');
assert(mainJs.includes("'UML ID'") && mainJs.includes('MILL_FIELDS'), 'UML ID in mill onboarding form fields');

console.log('\n7. QUARTER/YEAR from SDD Date Imported');
function parseSddImportDate_(raw) {
  const s = String(raw === undefined || raw === null ? '' : raw).trim();
  if (!s) return null;
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
  m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (m) return new Date(parseInt(m[3], 10), parseInt(m[2], 10) - 1, parseInt(m[1], 10));
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}
function quarterYearFromSddImportDate_(raw) {
  const d = parseSddImportDate_(raw);
  if (!d) return { quarter: '', year: '' };
  const q = Math.floor(d.getMonth() / 3) + 1;
  return { quarter: 'Q' + q, year: String(d.getFullYear()) };
}
assertEq(quarterYearFromSddImportDate_('2026-02-15'), { quarter: 'Q1', year: '2026' }, 'Feb → Q1');
assertEq(quarterYearFromSddImportDate_('2026-05-01'), { quarter: 'Q2', year: '2026' }, 'May → Q2');
assertEq(quarterYearFromSddImportDate_('2026-08-20'), { quarter: 'Q3', year: '2026' }, 'Aug → Q3');
assertEq(quarterYearFromSddImportDate_('2026-11-30'), { quarter: 'Q4', year: '2026' }, 'Nov → Q4');
assert(mainJs.includes('resolveSddPeriodFromMainRow_'), 'resolve period from SDD main row');
assert(mainJs.includes("data['QUARTER'] = period.quarter"), 'save forces QUARTER from SDD import date');
assert(mainJs.includes('applySddPeriodToMillPayload_'), 'prefill applies SDD period');

console.log('\n=== Results: ' + passed + ' passed, ' + failed + ' failed ===\n');
process.exit(failed > 0 ? 1 : 0);
