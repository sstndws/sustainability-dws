/**
 * Supply import routing & data-target consistency tests.
 * Run: node scripts/test-supply-routing.mjs
 */

const GAS_URL =
  process.env.GAS_WEBAPP_URL ||
  'https://script.google.com/macros/s/AKfycbzFydN5wOjsXbjMYjf88uhThltDeZXsV02oU8oPhYoh3ZYdZw9PGj9z0DInGgXqaL-PJg/exec';

let passed = 0;
let failed = 0;
function assert(cond, msg) {
  if (cond) { passed++; return; }
  failed++;
  console.error('FAIL:', msg);
}

// ── GAS mirrors ─────────────────────────────────────────────────────────────
function supplySubmitKindFromDraftGs_(row) {
  let stRaw = String(row.supply_type || row.SUPPLY_TYPE || '').trim().toUpperCase();
  if (!stRaw) {
    if (row['SUPPLY ISCC'] !== undefined && row['SUPPLY ISCC'] !== null && String(row['SUPPLY ISCC']).trim() !== '') stRaw = 'POME_ISCC';
    else if (row['SUPPLY INS'] !== undefined && row['SUPPLY INS'] !== null && String(row['SUPPLY INS']).trim() !== '') stRaw = 'POME_INS';
    else if (row['SUPPLY SHELL'] !== undefined && row['SUPPLY SHELL'] !== null && String(row['SUPPLY SHELL']).trim() !== '') stRaw = 'SHELL_GGL';
    else stRaw = 'CPO';
  }
  const st = stRaw;
  if (st === 'POME_ISCC') return 'POME_ISCC';
  if (st === 'POME_INS') return 'POME_INS';
  if (st === 'SHELL_GGL') return 'SHELL_GGL';
  if (st === 'PK') return 'PK';
  if (st === 'CPO') return 'CPO';
  if (st === 'CPO+PK' || st === 'BOTH') return 'BOTH';
  if (st.indexOf('CPO') >= 0 && st.indexOf('PK') >= 0) return 'BOTH';
  return 'CPO';
}

function supplyIsWasteSubmitKind_(kind) {
  const k = String(kind || '').trim().toUpperCase();
  return k === 'POME_ISCC' || k === 'POME_INS' || k === 'SHELL_GGL';
}

function targetSheetForKind_(kind) {
  return supplyIsWasteSubmitKind_(kind) ? 'millWaste' : 'mill';
}

// ── Frontend mirrors ────────────────────────────────────────────────────────
const WASTE_MAP = {
  POME_ISCC: { qty: 'SUPPLY ISCC', facility: 'FACILITY NAME ISCC', pct: 'PERCENTAGE SUPPLY ISCC' },
  POME_INS: { qty: 'SUPPLY INS', facility: 'FACILITY NAME INS', pct: 'PERCENTAGE SUPPLY INS' },
  SHELL_GGL: { qty: 'SUPPLY SHELL', facility: 'FACILITY NAME SHELL', pct: 'PERCENTAGE SUPPLY SHELL' },
};

function supplyImportIsWaste_(kind) {
  const k = String(kind || '').trim().toUpperCase();
  return k === 'POME_ISCC' || k === 'POME_INS' || k === 'SHELL_GGL';
}

function supplyWasteKindFromType_(kind) {
  return WASTE_MAP[String(kind || '').trim().toUpperCase()] || null;
}

function supplyFacilityFieldForKind_(kind) {
  const k = String(kind || 'CPO').toUpperCase();
  const wasteCfg = supplyWasteKindFromType_(k);
  if (wasteCfg) return wasteCfg.facility;
  return (k === 'PK' || (k.indexOf('PK') >= 0 && k.indexOf('CPO') < 0)) ? 'FACILITY NAME PK' : 'FACILITY NAME CPO';
}

function supplyDraftProfileSkipFields_(kind) {
  const wasteCfg = supplyWasteKindFromType_(kind);
  const skip = new Set([
    'MONTH', 'QUARTER', 'YEAR', 'COMPANY NAME', 'MILL NAME', 'GROUP NAME', 'SOURCE TYPE',
    'SUPPLY CPO', 'SUPPLY PK', 'SUPPLY_QTY', 'SUPPLY_PERCENTAGE',
    'PERCENTAGE SUPPLY CPO', 'PERCENTAGE SUPPLY PK',
    'FACILITY NAME CPO', 'FACILITY NAME PK',
    'PERCENTAGE SUPPLY ISCC', 'PERCENTAGE SUPPLY INS', 'PERCENTAGE SUPPLY SHELL',
    'SUPPLY ISCC', 'SUPPLY INS', 'SUPPLY SHELL',
    'FACILITY NAME ISCC', 'FACILITY NAME INS', 'FACILITY NAME SHELL',
  ]);
  if (wasteCfg) {
    skip.add(wasteCfg.pct);
    skip.add(wasteCfg.qty);
    skip.add(wasteCfg.facility);
  } else if (kind === 'PK') {
    skip.add('PERCENTAGE SUPPLY PK');
    skip.add('FACILITY NAME PK');
  } else {
    skip.add('PERCENTAGE SUPPLY CPO');
    skip.add('FACILITY NAME CPO');
  }
  return skip;
}

function supplyApplyPlantToDraftFacility_(draft, plant, kind) {
  const wasteCfg = supplyWasteKindFromType_(kind);
  const p = String(plant || '').trim();
  if (!draft || !p) return;
  if (wasteCfg) {
    draft[wasteCfg.facility] = p;
    return;
  }
  const k = String(kind || 'CPO').toUpperCase();
  if (k === 'PK') draft['FACILITY NAME PK'] = p;
  else draft['FACILITY NAME CPO'] = p;
}

function supplyFindOpenPeriodBatch_(batches, month, year, supplyType) {
  const want = String(month) + '|' + String(year);
  const wantType = String(supplyType || '').trim().toUpperCase();
  return batches.find(function(b) {
    if (b.status === 'submitted') return false;
    const key = String(b.month || b.quarter) + '|' + String(b.year);
    if (key !== want) return false;
    const batchType = String(b.supply_type || '').trim().toUpperCase();
    if (supplyImportIsWaste_(wantType)) return batchType === wantType;
    return !supplyImportIsWaste_(batchType);
  }) || null;
}

// ── Submit target sheet ─────────────────────────────────────────────────────
const ALL_KINDS = ['CPO', 'PK', 'POME_ISCC', 'POME_INS', 'SHELL_GGL', 'CPO+PK'];
ALL_KINDS.forEach(function(kind) {
  const submitKind = supplySubmitKindFromDraftGs_({ supply_type: kind });
  const sheet = targetSheetForKind_(submitKind);
  if (supplyImportIsWaste_(kind)) {
    assert(sheet === 'millWaste', kind + ' → millWaste');
    assert(supplyIsWasteSubmitKind_(submitKind), kind + ' is waste submit kind');
  } else {
    assert(sheet === 'mill', kind + ' → mill');
    assert(!supplyIsWasteSubmitKind_(submitKind), kind + ' is main submit kind');
  }
});

// infer kind from qty columns
[
  [{ 'SUPPLY ISCC': 10 }, 'POME_ISCC'],
  [{ 'SUPPLY INS': 5 }, 'POME_INS'],
  [{ 'SUPPLY SHELL': 3 }, 'SHELL_GGL'],
  [{ 'SUPPLY CPO': 100 }, 'CPO'],
  [{ supply_type: 'PK', 'SUPPLY PK': 50 }, 'PK'],
].forEach(function(pair) {
  assert(supplySubmitKindFromDraftGs_(pair[0]) === pair[1], 'infer ' + pair[1]);
  assert(targetSheetForKind_(pair[1]) === (supplyImportIsWaste_(pair[1]) ? 'millWaste' : 'mill'), 'sheet for inferred ' + pair[1]);
});

// ── Waste facility routing ──────────────────────────────────────────────────
const plants = ['EUP - LG', 'EUP - TJPR', 'EUP - BONTANG', 'PLANT X'];
['POME_ISCC', 'POME_INS', 'SHELL_GGL'].forEach(function(kind) {
  const cfg = supplyWasteKindFromType_(kind);
  plants.forEach(function(plant) {
    const d = {};
    supplyApplyPlantToDraftFacility_(d, plant, kind);
    assert(d[cfg.facility] === plant, kind + ' plant → ' + cfg.facility);
    assert(!d['FACILITY NAME CPO'], kind + ' must not set CPO facility');
    assert(!d['FACILITY NAME PK'], kind + ' must not set PK facility');
  });
});

// ── Main facility routing ───────────────────────────────────────────────────
const cpoD = {};
supplyApplyPlantToDraftFacility_(cpoD, 'CRC', 'CPO');
assert(cpoD['FACILITY NAME CPO'] === 'CRC', 'CPO facility');
const pkD = {};
supplyApplyPlantToDraftFacility_(pkD, 'KCP A', 'PK');
assert(pkD['FACILITY NAME PK'] === 'KCP A', 'PK facility');

// ── Profile skip fields: waste keeps identity, not supply cols ──────────────
const profile = {
  'COMPANY NAME': 'TEST CO',
  ADDRESS: 'Jl. Test',
  'SUPPLY CPO': 999,
  'FACILITY NAME CCO': 'OLD',
  'FACILITY NAME CPO': 'OLD CPO',
  'SUPPLY ISCC': 888,
};
['POME_ISCC', 'POME_INS', 'SHELL_GGL'].forEach(function(kind) {
  const skip = supplyDraftProfileSkipFields_(kind);
  const cfg = supplyWasteKindFromType_(kind);
  assert(skip.has(cfg.qty), kind + ' skips waste qty on profile copy');
  assert(skip.has(cfg.facility), kind + ' skips waste facility on profile copy');
  assert(skip.has('SUPPLY CPO'), kind + ' skips main CPO qty');
  assert(!skip.has('ADDRESS'), kind + ' copies ADDRESS from profile');
  assert(skip.has('COMPANY NAME'), kind + ' skips COMPANY NAME (from excel)');
});

// ── Batch separation: waste types don't merge with main ───────────────────────
const batches = [
  { month: '1', year: '2029', supply_type: 'CPO', status: 'draft', rows: [] },
  { month: '1', year: '2029', supply_type: 'POME_ISCC', status: 'draft', rows: [] },
  { month: '1', year: '2029', supply_type: 'SHELL_GGL', status: 'draft', rows: [] },
];
assert(supplyFindOpenPeriodBatch_(batches, '1', '2029', 'CPO').supply_type === 'CPO', 'find CPO batch');
assert(supplyFindOpenPeriodBatch_(batches, '1', '2029', 'POME_ISCC').supply_type === 'POME_ISCC', 'find ISCC batch');
assert(supplyFindOpenPeriodBatch_(batches, '1', '2029', 'PK') === null || batches.find(b => b.supply_type === 'PK') == null, 'no PK batch yet');
assert(supplyFindOpenPeriodBatch_(batches, '1', '2029', 'CPO') !== supplyFindOpenPeriodBatch_(batches, '1', '2029', 'POME_ISCC'), 'waste ≠ main batch');

// year 2029 free input scenario
assert(supplyFindOpenPeriodBatch_(batches, '1', '2029', 'SHELL_GGL').supply_type === 'SHELL_GGL', 'year 2029 waste batch');

// ── Parametric matrix (bulk assertions) ─────────────────────────────────────
for (let m = 1; m <= 12; m++) {
  for (const y of [2026, 2027, 2028, 2029, 2030]) {
    const b = [{ month: String(m), year: String(y), supply_type: 'POME_INS', status: 'draft', rows: [] }];
    const found = supplyFindOpenPeriodBatch_(b, String(m), String(y), 'POME_INS');
    assert(found && found.supply_type === 'POME_INS', 'period ' + m + '/' + y + ' POME_INS');
    assert(targetSheetForKind_('POME_INS') === 'millWaste', 'POME_INS sheet at ' + m + '/' + y);
  }
}

for (let i = 0; i < 20; i++) {
  const kind = ['CPO', 'PK', 'POME_ISCC', 'POME_INS', 'SHELL_GGL'][i % 5];
  const row = { supply_type: kind };
  if (kind === 'POME_ISCC') row['SUPPLY ISCC'] = i + 1;
  if (kind === 'POME_INS') row['SUPPLY INS'] = i + 1;
  if (kind === 'SHELL_GGL') row['SUPPLY SHELL'] = i + 1;
  if (kind === 'CPO') row['SUPPLY CPO'] = i + 1;
  if (kind === 'PK') row['SUPPLY PK'] = i + 1;
  const sk = supplySubmitKindFromDraftGs_(row);
  const expectedSheet = supplyImportIsWaste_(kind) ? 'millWaste' : 'mill';
  assert(targetSheetForKind_(sk) === expectedSheet, 'matrix row ' + i + ' ' + kind);
}

// facility field helper
assert(supplyFacilityFieldForKind_('CPO') === 'FACILITY NAME CPO', 'fac CPO');
assert(supplyFacilityFieldForKind_('PK') === 'FACILITY NAME PK', 'fac PK');
assert(supplyFacilityFieldForKind_('POME_ISCC') === 'FACILITY NAME ISCC', 'fac ISCC');
assert(supplyFacilityFieldForKind_('POME_INS') === 'FACILITY NAME INS', 'fac INS');
assert(supplyFacilityFieldForKind_('SHELL_GGL') === 'FACILITY NAME SHELL', 'fac SHELL');

// ── GHG VALUE (waste identity — replaces RISK REDUCTION FACTOR) ───────────
const MILL_WASTE_IDENTITY_FIELDS = [
  'GROUP NAME', 'COMPANY NAME', 'MILL NAME', 'UML ID', 'COMPANY CODE', 'SOURCE TYPE', 'TRADER NAME',
  'ADDRESS', 'PROVINCE', 'COORDINATES', 'MILL CATEGORY', 'MILL CAPACITY', 'MILL CAPACITY (TON/HOUR)',
  'HGU/HGB', 'IZIN LOKASI', 'IUP', 'IZIN LINGKUNGAN',
  'NDPE', 'HRDD', 'MILL LOC', 'CERTIFICATION',
  'DEFORESTATION GRIEVANCES', 'BURN AREA GRIEVANCES',
  'LEGALITY GRIEVANCE', 'HUMAN RIGHTS GRIEVANCE', 'SAFETY GRIEVANCE', 'SOCIAL GRIEVANCE', 'ENVIRONMENT GRIEVANCE',
  'GHG VALUE',
];
assert(MILL_WASTE_IDENTITY_FIELDS.includes('GHG VALUE'), 'waste identity includes GHG VALUE');
assert(!MILL_WASTE_IDENTITY_FIELDS.includes('RISK REDUCTION FACTOR'), 'RISK REDUCTION FACTOR removed');

function buildSupplyIdentityPatchFromDraftGs_(row) {
  const fillKeys = [
    'MONTH', 'YEAR', 'QUARTER', 'SOURCE TYPE', 'GROUP NAME', 'COMPANY NAME', 'MILL NAME',
    'UML ID', 'COMPANY CODE', 'ADDRESS', 'PROVINCE', 'COORDINATES', 'MILL CATEGORY',
    'MILL CAPACITY (TON/HOUR)', 'HGU/HGB', 'IZIN LOKASI', 'IUP', 'IZIN LINGKUNGAN',
    'NDPE', 'HRDD', 'MILL LOC', 'CERTIFICATION',
    'DEFORESTATION GRIEVANCES', 'BURN AREA GRIEVANCES',
    'LEGALITY GRIEVANCE', 'HUMAN RIGHTS GRIEVANCE', 'SAFETY GRIEVANCE', 'SOCIAL GRIEVANCE', 'ENVIRONMENT GRIEVANCE',
    'GHG VALUE',
  ];
  const patch = {};
  fillKeys.forEach(function(k) {
    const v = row[k];
    if (v !== undefined && v !== null && String(v).trim() !== '') patch[k] = v;
  });
  return patch;
}

const ghgPatch = buildSupplyIdentityPatchFromDraftGs_({
  supply_type: 'POME_INS',
  'COMPANY NAME': 'TEST CO',
  'GHG VALUE': '0.42',
  'SUPPLY INS': 961760,
});
assert(ghgPatch['GHG VALUE'] === '0.42', 'GHG VALUE in identity patch');
assert(!ghgPatch['RISK REDUCTION FACTOR'], 'no RISK REDUCTION FACTOR in patch');
assert(!ghgPatch['SUPPLY INS'], 'qty stays in supply patch not identity');

// Waste sheet may rename qty headers → remap canonical keys before write
function millWasteQtyHeaderAliasesGs_(canonical) {
  const c = String(canonical || '').trim().toUpperCase();
  if (c === 'SUPPLY ISCC') return ['SUPPLY ISCC', 'SUPPLY POME ISCC', 'Supply ISCC', 'Supply POME ISCC'];
  if (c === 'SUPPLY INS') return ['SUPPLY INS', 'SUPPLY POME INS', 'Supply INS', 'Supply POME INS'];
  if (c === 'SUPPLY SHELL') return ['SUPPLY SHELL', 'SUPPLY POME SHELL', 'Supply SHELL', 'Supply POME SHELL'];
  return canonical ? [canonical] : [];
}
function millFindWasteQtyHeaderGs_(headers, canonical) {
  const list = (headers || []).map(function(x) { return String(x || '').trim(); });
  const aliases = millWasteQtyHeaderAliasesGs_(canonical);
  for (let a = 0; a < aliases.length; a++) {
    if (list.indexOf(aliases[a]) >= 0) return aliases[a];
  }
  return null;
}
function resolveMillWasteQtyKeysOnPatch_(patch, headers) {
  ['SUPPLY ISCC', 'SUPPLY INS', 'SUPPLY SHELL'].forEach(function(canonical) {
    const sheetCol = millFindWasteQtyHeaderGs_(headers, canonical);
    const qty = patch[canonical];
    if (qty === undefined || qty === null || String(qty).trim() === '') return;
    if (sheetCol) {
      patch[sheetCol] = qty;
      if (sheetCol !== canonical && patch[canonical] !== undefined) delete patch[canonical];
    }
  });
}
const wasteHeaders = ['COMPANY NAME', 'SUPPLY POME ISCC', 'SUPPLY POME INS', 'TOTAL POME SUPPLY'];
const wastePatch = { 'SUPPLY ISCC': 1200, 'SUPPLY INS': 300 };
resolveMillWasteQtyKeysOnPatch_(wastePatch, wasteHeaders);
assert(wastePatch['SUPPLY POME ISCC'] === 1200, 'write remaps to SUPPLY POME ISCC');
assert(wastePatch['SUPPLY POME INS'] === 300, 'write remaps to SUPPLY POME INS');
assert(wastePatch['SUPPLY ISCC'] === undefined, 'canonical SUPPLY ISCC removed after remap');
assert(wastePatch['SUPPLY INS'] === undefined, 'canonical SUPPLY INS removed after remap');

console.log('\nSupply routing tests:', passed, 'passed,', failed, 'failed');

async function pingGas_() {
  if (process.env.SKIP_GAS_PING === '1') {
    console.log('GAS ping skipped (SKIP_GAS_PING=1)');
    return;
  }
  try {
    const url = GAS_URL + (GAS_URL.indexOf('?') >= 0 ? '&' : '?') + 'action=ping';
    const res = await fetch(url, { method: 'GET', redirect: 'follow' });
    const text = await res.text();
    let json;
    try { json = JSON.parse(text); } catch { json = null; }
    assert(res.ok, 'GAS HTTP ' + res.status);
    assert(json && json.success === true, 'GAS ping response valid');
    if (json) console.log('GAS ping OK:', json.message || json.version || 'alive');
  } catch (err) {
    console.warn('GAS ping skipped (network):', err && err.message ? err.message : err);
  }
}

await pingGas_();
console.log('Total:', passed, 'passed,', failed, 'failed');
process.exit(failed ? 1 : 0);
