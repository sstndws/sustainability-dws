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

// CPO+PK dual row: one PLANT must not fill both facility columns
function supplyRowIsDualSupplyKind_(kind) {
  const k = String(kind || '').trim().toUpperCase();
  return k === 'CPO+PK' || k === 'BOTH' || (k.indexOf('CPO') >= 0 && k.indexOf('PK') >= 0);
}
function supplyApplyPlantToDraftFacility_(draft, plant, kind) {
  if (!draft) return;
  const k = String(kind || draft.supply_type || draft.SUPPLY_TYPE || 'CPO').toUpperCase();
  const p = supplyNormalizePlantValue_(plant != null ? plant : draft.PLANT || '');
  if (!p) return;
  if (supplyRowIsDualSupplyKind_(k)) return;
  draft.PLANT = p;
  const facField = (k === 'PK' || (k.indexOf('PK') >= 0 && k.indexOf('CPO') < 0)) ? 'FACILITY NAME PK' : 'FACILITY NAME CPO';
  draft[facField] = p;
}
function supplyClearDualRowPlant_(row) {
  if (!row || String(row.supply_type || '').toUpperCase() !== 'CPO+PK') return;
  delete row.PLANT;
}
function supplyMergeDraftRows_(target, source) {
  if (!target || !source) return;
  if (source['FACILITY NAME CPO']) target['FACILITY NAME CPO'] = source['FACILITY NAME CPO'];
  if (source['FACILITY NAME PK']) target['FACILITY NAME PK'] = source['FACILITY NAME PK'];
  target.supply_type = 'CPO+PK';
  target.SUPPLY_TYPE = 'CPO+PK';
  supplyClearDualRowPlant_(target);
}

const cpoRow = {
  supply_type: 'CPO',
  'FACILITY NAME CPO': 'EUP - BTG',
  PLANT: 'EUP - BTG',
};
supplyApplyPlantToDraftFacility_(cpoRow, 'KCP Bontang', 'PK');
assert(cpoRow['FACILITY NAME CPO'] === 'EUP - BTG', 'CPO facility kept after PK plant apply');
assert(cpoRow['FACILITY NAME PK'] === 'KCP Bontang', 'PK facility set from PK import');
cpoRow.supply_type = 'CPO+PK';
cpoRow.SUPPLY_TYPE = 'CPO+PK';
supplyClearDualRowPlant_(cpoRow);
supplyApplyPlantToDraftFacility_(cpoRow, 'KCP Bontang', 'CPO+PK');
assert(cpoRow['FACILITY NAME CPO'] === 'EUP - BTG', 'dual row: CPO facility not overwritten by PLANT');
assert(cpoRow['FACILITY NAME PK'] === 'KCP Bontang', 'dual row: PK facility kept');
assert(cpoRow.PLANT === undefined, 'dual row: legacy PLANT cleared');

const mergeTarget = { supply_type: 'CPO', 'FACILITY NAME CPO': 'EUP - BTG', PLANT: 'EUP - BTG' };
const mergeSource = { supply_type: 'PK', 'FACILITY NAME PK': 'KCP Bontang', PLANT: 'KCP Bontang' };
supplyMergeDraftRows_(mergeTarget, mergeSource);
assert(mergeTarget['FACILITY NAME CPO'] === 'EUP - BTG', 'merge keeps CPO facility');
assert(mergeTarget['FACILITY NAME PK'] === 'KCP Bontang', 'merge copies PK facility');
assert(mergeTarget.PLANT === undefined, 'merge clears single PLANT on dual row');

console.log('Supply separate-row tests:', passed, 'passed,', failed, 'failed');
process.exit(failed ? 1 : 0);
