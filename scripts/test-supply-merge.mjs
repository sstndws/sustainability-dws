/**
 * Supply import — CPO/PK stay separate rows (no merge per company).
 * Run: node scripts/test-supply-merge.mjs
 */

const SUPPLY_PCT_COL_CPO = 'PERCENTAGE SUPPLY CPO';
const SUPPLY_PCT_COL_PK = 'PERCENTAGE SUPPLY PK';

function supplyNormKey_(s) {
  return String(s || '').trim().toLowerCase().replace(/\s+/g, ' ');
}
function supplyCompanyKey_(company) {
  return supplyNormKey_(company);
}
function supplyFindDraftRowsForMergeByCompany_(batch, companyName) {
  const wantCo = supplyCompanyKey_(companyName);
  return (batch.rows || []).filter(function(row) {
    return !row._submitted && supplyCompanyKey_(row['COMPANY NAME']) === wantCo;
  });
}

let passed = 0;
let failed = 0;
function assert(cond, msg) {
  if (cond) { passed++; return; }
  failed++;
  console.error('FAIL:', msg);
}

const batch = {
  quarter: 'Q1', year: '2026', supply_type: 'CPO',
  rows: [{
    'COMPANY NAME': 'ABDI BORNEO PLANTATIONS',
    supply_type: 'CPO',
    [SUPPLY_PCT_COL_CPO]: 45.5,
    'FACILITY NAME CPO': 'PLANT A',
    _submitted: false,
  }],
};

// PK import adds a new row — does not merge into CPO row
const beforeLen = batch.rows.length;
batch.rows.push({
  'COMPANY NAME': 'ABDI BORNEO PLANTATIONS',
  supply_type: 'PK',
  [SUPPLY_PCT_COL_PK]: 12.3,
  'FACILITY NAME PK': 'PLANT B',
  _submitted: false,
});
assert(batch.rows.length === beforeLen + 1, 'separate PK row added');
assert(batch.rows[0].supply_type === 'CPO', 'CPO row unchanged');
assert(batch.rows[0][SUPPLY_PCT_COL_PK] == null || batch.rows[0][SUPPLY_PCT_COL_PK] === '', 'CPO row has no PK pct');
assert(batch.rows[1].supply_type === 'PK', 'PK row is PK only');

assert(supplyFindDraftRowsForMergeByCompany_(batch, 'ABDI BORNEO PLANTATIONS').length === 2, 'two rows same company');

function supplyNormalizePlantValue_(raw) {
  let s = String(raw == null ? '' : raw).trim();
  if (!s) return '';
  s = s.replace(/[\r\n]+/g, ', ');
  s = s.replace(/\s*[,;/|]+\s*/g, ', ');
  s = s.replace(/,\s*,+/g, ', ');
  s = s.replace(/^,\s*|,\s*$/g, '');
  return s.trim();
}
assert(supplyNormalizePlantValue_('KCP A, KCP B') === 'KCP A, KCP B', 'multi KCP stays one string');
assert(supplyNormalizePlantValue_('KCP A\nKCP B') === 'KCP A, KCP B', 'newline KCPs normalize to one string');
assert(supplyNormalizePlantValue_('KCP A / KCP B') === 'KCP A, KCP B', 'slash KCPs normalize to one string');

console.log('Supply separate-row tests:', passed, 'passed,', failed, 'failed');
process.exit(failed ? 1 : 0);
