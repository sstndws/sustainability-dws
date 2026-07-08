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

function supplyPlantIsKcp_(name) {
  return String(name == null ? '' : name).toUpperCase().indexOf('KCP') !== -1;
}
function supplyRouteFacilityNames_(cpoRaw, pkRaw) {
  const cpoOut = [];
  const pkOut = [];
  const seenCpo = {};
  const seenPk = {};
  function add_(list, seen, tok) {
    const t = String(tok || '').trim();
    if (!t) return;
    const key = t.toUpperCase();
    if (seen[key]) return;
    seen[key] = true;
    list.push(t);
  }
  [cpoRaw, pkRaw].forEach(function(raw) {
    String(raw == null ? '' : raw).split(',').forEach(function(tok) {
      const t = String(tok || '').trim();
      if (!t) return;
      if (supplyPlantIsKcp_(t)) add_(pkOut, seenPk, t);
      else add_(cpoOut, seenCpo, t);
    });
  });
  return { cpo: cpoOut.join(', '), pk: pkOut.join(', ') };
}
function supplyApplyPlantToDraftFacility_(draft, plant, kind) {
  if (!draft) return;
  const k = String(kind || draft.supply_type || 'CPO').toUpperCase();
  const p = supplyNormalizePlantValue_(plant != null ? plant : draft.PLANT || '');
  if (p) draft.PLANT = p;
  const isDual = k === 'CPO+PK' || k === 'BOTH' || (k.indexOf('CPO') >= 0 && k.indexOf('PK') >= 0);
  const isPkOnly = !isDual && (k === 'PK' || (k.indexOf('PK') >= 0 && k.indexOf('CPO') < 0));
  const isCpoOnly = !isDual && !isPkOnly;
  if (isPkOnly) {
    if (p) draft['FACILITY NAME PK'] = p;
    return;
  }
  if (isCpoOnly) {
    if (!p) return;
    const split = supplyRouteFacilityNames_(p, '');
    draft['FACILITY NAME CPO'] = split.cpo || (!supplyPlantIsKcp_(p) ? p : '');
    if (split.pk) draft['FACILITY NAME PK'] = split.pk;
    return;
  }
  if (!p) return;
  const dualSplit = supplyRouteFacilityNames_(p, draft['FACILITY NAME PK'] || '');
  if (dualSplit.cpo) draft['FACILITY NAME CPO'] = dualSplit.cpo;
  if (dualSplit.pk) draft['FACILITY NAME PK'] = dualSplit.pk;
}

const pkDraft = { supply_type: 'PK' };
supplyApplyPlantToDraftFacility_(pkDraft, 'KCP PURA', 'PK');
assert(pkDraft['FACILITY NAME PK'] === 'KCP PURA', 'PK import → FACILITY NAME PK');
assert(!pkDraft['FACILITY NAME CPO'], 'PK import must not set FACILITY NAME CPO');

const cpoDraft = { supply_type: 'CPO' };
supplyApplyPlantToDraftFacility_(cpoDraft, 'CRC', 'CPO');
assert(cpoDraft['FACILITY NAME CPO'] === 'CRC', 'CPO import → FACILITY NAME CPO');
assert(!cpoDraft['FACILITY NAME PK'], 'CRC alone must not set FACILITY NAME PK');

const mixedDraft = { supply_type: 'CPO' };
supplyApplyPlantToDraftFacility_(mixedDraft, 'CRC, KCP PURA', 'CPO');
assert(mixedDraft['FACILITY NAME CPO'] === 'CRC', 'mixed cell: CRC → CPO column');
assert(mixedDraft['FACILITY NAME PK'] === 'KCP PURA', 'mixed cell: KCP → PK column');

console.log('Supply separate-row tests:', passed, 'passed,', failed, 'failed');
process.exit(failed ? 1 : 0);
