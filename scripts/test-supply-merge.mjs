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

// ── POME ISCC + POME INS merge (Task List import) ─────────────────────────
const SUPPLY_PCT_COL_ISCC = 'PERCENTAGE SUPPLY ISCC';
const SUPPLY_PCT_COL_INS = 'PERCENTAGE SUPPLY INS';

function supplyImportIsWaste_(kind) {
  const k = String(kind || '').trim().toUpperCase();
  if (k === 'POME_ISCC' || k === 'POME INS' || k === 'POME_INS' || k === 'SHELL_GGL') return true;
  return k.indexOf('POME_ISCC') >= 0 || k.indexOf('POME_INS') >= 0 || k.indexOf('SHELL_GGL') >= 0;
}
function supplyCombineSupplyTypes_(a, b) {
  function wasteTokens(raw) {
    const out = [];
    function add(u) {
      if (!u) return;
      if (u === 'POME ISCC') u = 'POME_ISCC';
      if (u === 'POME INS') u = 'POME_INS';
      if (u === 'SHELL GGL') u = 'SHELL_GGL';
      if ((u === 'POME_ISCC' || u === 'POME_INS' || u === 'SHELL_GGL') && out.indexOf(u) === -1) out.push(u);
    }
    String(raw || '').trim().toUpperCase().split('+').forEach(function(part) { add(String(part || '').trim()); });
    return out;
  }
  const wa = wasteTokens(a);
  const wb = wasteTokens(b);
  if (wa.length || wb.length) {
    const merged = wa.slice();
    wb.forEach(function(t) { if (merged.indexOf(t) === -1) merged.push(t); });
    const order = ['POME_ISCC', 'POME_INS', 'SHELL_GGL'];
    merged.sort(function(x, y) { return order.indexOf(x) - order.indexOf(y); });
    return merged.join('+');
  }
  return 'CPO';
}
function supplyWasteQtyKindsOnRow_(row) {
  const out = [];
  if (row['SUPPLY ISCC'] != null && String(row['SUPPLY ISCC']).trim() !== '') out.push('POME_ISCC');
  if (row['SUPPLY INS'] != null && String(row['SUPPLY INS']).trim() !== '') out.push('POME_INS');
  if (row['SUPPLY SHELL'] != null && String(row['SUPPLY SHELL']).trim() !== '') out.push('SHELL_GGL');
  return out;
}
function supplyStampWasteSupplyType_(row) {
  const kinds = supplyWasteQtyKindsOnRow_(row);
  if (kinds.length) {
    row.supply_type = kinds.join('+');
    row.SUPPLY_TYPE = row.supply_type;
  }
  row['PRODUCT SUPPLY'] = kinds.map(function(k) {
    return k === 'POME_ISCC' ? 'POME ISCC' : (k === 'POME_INS' ? 'POME INS' : 'SHELL GGL');
  }).join('; ');
}
function supplyMergeWasteDraftRows_(target, source) {
  [
    [SUPPLY_PCT_COL_ISCC, 'SUPPLY ISCC', 'FACILITY NAME ISCC'],
    [SUPPLY_PCT_COL_INS, 'SUPPLY INS', 'FACILITY NAME INS'],
  ].forEach(function(triple) {
    if (source[triple[0]] != null && String(source[triple[0]]).trim() !== '') target[triple[0]] = source[triple[0]];
    if (source[triple[1]] != null && String(source[triple[1]]).trim() !== '') target[triple[1]] = source[triple[1]];
    if (source[triple[2]]) target[triple[2]] = source[triple[2]];
  });
  supplyStampWasteSupplyType_(target);
}
function supplyConsolidateWasteBatches_(batchList) {
  const openByPeriod = {};
  const result = [];
  (batchList || []).forEach(function(b) {
    const kindKey = String(b.supply_type || '').trim().toUpperCase();
    const pk = (b.month || '') + '|' + (b.year || '') + (supplyImportIsWaste_(kindKey) ? '|WASTE' : '');
    if (!openByPeriod[pk]) {
      openByPeriod[pk] = b;
      result.push(b);
      return;
    }
    const target = openByPeriod[pk];
    (b.rows || []).forEach(function(r) { target.rows.push(r); });
    target.supply_type = supplyCombineSupplyTypes_(target.supply_type, b.supply_type);
  });
  return result;
}

const wasteBatchIscc = {
  month: '2', year: '2026', supply_type: 'POME_ISCC',
  rows: [{ 'COMPANY NAME': 'GUNUNG RIJUAN', supply_type: 'POME_ISCC', 'SUPPLY ISCC': 800, month: '2', year: '2026' }],
};
const wasteBatchIns = {
  month: '2', year: '2026', supply_type: 'POME_INS',
  rows: [{ 'COMPANY NAME': 'GUNUNG RIJUAN', supply_type: 'POME_INS', 'SUPPLY INS': 450, month: '2', year: '2026' }],
};
const consolidated = supplyConsolidateWasteBatches_([wasteBatchIscc, wasteBatchIns]);
assert(consolidated.length === 1, 'ISCC + INS batches merge to one waste batch');
assert(consolidated[0].supply_type === 'POME_ISCC+POME_INS', 'batch type combines ISCC+INS');
assert(consolidated[0].rows.length === 2, 'both rows kept until pair merge');

const mergedRow = Object.assign({}, consolidated[0].rows[0]);
supplyMergeWasteDraftRows_(mergedRow, consolidated[0].rows[1]);
assert(mergedRow['SUPPLY ISCC'] === 800, 'merged row keeps ISCC qty');
assert(mergedRow['SUPPLY INS'] === 450, 'merged row keeps INS qty');
assert(mergedRow.supply_type === 'POME_ISCC+POME_INS', 'merged row type ISCC+INS');
assert(mergedRow['PRODUCT SUPPLY'].indexOf('POME ISCC') >= 0, 'product includes POME ISCC');
assert(mergedRow['PRODUCT SUPPLY'].indexOf('POME INS') >= 0, 'product includes POME INS');

console.log('Supply waste-merge tests:', passed, 'passed,', failed, 'failed');
process.exit(failed ? 1 : 0);
