/**
 * Mill registry General / Main / Waste product view merge tests.
 * Run: node scripts/test-mill-registry-views.mjs
 */

let passed = 0;
let failed = 0;
function assert(cond, msg) {
  if (cond) { passed++; return; }
  failed++;
  console.error('FAIL:', msg);
}

function pickMillCompanyName_(row) {
  return row['COMPANY NAME'] || row['Company Name'] || '';
}

function millMonthVal(row) {
  return row.MONTH || row['MONTH'] || row.month || '';
}

function millYearVal(row) {
  return row.YEAR || row['YEAR'] || row.year || '';
}

function millParseSupplyQty_(raw) {
  if (raw == null || raw === '') return 0;
  if (typeof raw === 'number' && isFinite(raw)) return raw < 0 ? 0 : raw;
  const s = String(raw).trim().replace(/\./g, '').replace(',', '.');
  const n = parseFloat(s);
  return isNaN(n) || n < 0 ? 0 : n;
}

function millFormatSupplyQtyDisplay_(raw) {
  if (raw == null || raw === '') return '';
  if (typeof raw === 'number' && isFinite(raw)) {
    if (Number.isInteger(raw) && Math.abs(raw) >= 1000) {
      return String(raw).replace(/\B(?=(\d{3})+(?!\d))/g, '.');
    }
    return String(raw);
  }
  return String(raw).trim();
}

function millNormalizeProductSupply_(row) {
  const raw = row['PRODUCT SUPPLY'] || '';
  if (!raw) return '';
  return String(raw).trim().toUpperCase().replace(/\s+/g, '').replace(/[;&/+]+/g, ',').replace(/,+/g, ',').replace(/^,|,$/g, '');
}

function millNormalizeWasteQtyAliasesOnRow_(row) {
  if (!row || typeof row !== 'object') return row;
  function fillCanonical_(canonical, aliases) {
    const cur = row[canonical];
    if (cur != null && String(cur).trim() !== '' && String(cur).trim() !== '—') return;
    for (let i = 0; i < aliases.length; i++) {
      const raw = row[aliases[i]];
      if (raw != null && String(raw).trim() !== '' && String(raw).trim() !== '—') {
        row[canonical] = raw;
        return;
      }
    }
  }
  fillCanonical_('SUPPLY ISCC', ['SUPPLY POME ISCC', 'Supply POME ISCC', 'SUPPLY ISCC POME']);
  fillCanonical_('SUPPLY INS', ['SUPPLY POME INS', 'Supply POME INS', 'SUPPLY INS POME']);
  fillCanonical_('SUPPLY SHELL', ['SUPPLY POME SHELL', 'Supply POME SHELL']);
  return row;
}

function millCollectProductSupplyTokens_(row) {
  millNormalizeWasteQtyAliasesOnRow_(row);
  const seen = new Set();
  const out = [];
  function add(tok) {
    const t = String(tok || '').trim();
    if (!t) return;
    const k = t.toUpperCase();
    if (seen.has(k)) return;
    seen.add(k);
    out.push(t);
  }
  const ps = millNormalizeProductSupply_(row);
  if (ps) ps.split(/[,;/]+/).forEach(function(part) { add(part.trim()); });
  if (millParseSupplyQty_(row['SUPPLY CPO']) > 0) add('CPO');
  if (millParseSupplyQty_(row['SUPPLY PK']) > 0) add('PK');
  if (millParseSupplyQty_(row['SUPPLY ISCC']) > 0) add('POME ISCC');
  if (millParseSupplyQty_(row['SUPPLY INS']) > 0) add('POME INS');
  if (millParseSupplyQty_(row['SUPPLY SHELL']) > 0) add('SHELL GGL');
  return out;
}

function millJoinProductSupplyTokens_(rows) {
  const seen = new Set();
  const out = [];
  (rows || []).forEach(function(r) {
    millCollectProductSupplyTokens_(r).forEach(function(tok) {
      const k = tok.toUpperCase();
      if (seen.has(k)) return;
      seen.add(k);
      out.push(tok);
    });
  });
  return out.join('; ');
}

function millBuildQtySummaryFromRow_(row) {
  if (!row) return '';
  millNormalizeWasteQtyAliasesOnRow_(row);
  const parts = [];
  function push(label, field) {
    const raw = row[field];
    const q = millParseSupplyQty_(raw);
    if (q > 0) parts.push(label + ': ' + millFormatSupplyQtyDisplay_(raw != null && String(raw).trim() !== '' ? raw : q));
  }
  push('CPO', 'SUPPLY CPO');
  push('PK', 'SUPPLY PK');
  push('POME ISCC', 'SUPPLY ISCC');
  push('POME INS', 'SUPPLY INS');
  push('SHELL GGL', 'SUPPLY SHELL');
  return parts.join('; ');
}

function millBuildQtySummaryFromRows_(rows) {
  const parts = [];
  const seen = new Set();
  (rows || []).forEach(function(r) {
    const s = millBuildQtySummaryFromRow_(r);
    if (!s || seen.has(s)) return;
    seen.add(s);
    parts.push(s);
  });
  return parts.join(' · ');
}

function millGeneralMergeKey_(row) {
  return [
    String(pickMillCompanyName_(row) || '').trim().toLowerCase(),
    String(millMonthVal(row) || '').trim().toLowerCase(),
    String(millYearVal(row) || '').trim().toLowerCase(),
  ].join('\u0001');
}

function millMergeGeneralRegistryRows_(mainRows, wasteRows) {
  const map = new Map();
  const order = [];
  function ingest(r, isMain) {
    const key = millGeneralMergeKey_(r);
    if (!pickMillCompanyName_(r)) return;
    if (!map.has(key)) {
      map.set(key, { primary: r, members: [r], hasMain: !!isMain });
      order.push(key);
      return;
    }
    const g = map.get(key);
    g.members.push(r);
    if (isMain) {
      g.primary = r;
      g.hasMain = true;
    } else if (!g.hasMain) {
      g.primary = r;
    }
  }
  (mainRows || []).forEach(function(r) { ingest(r, true); });
  (wasteRows || []).forEach(function(r) { ingest(r, false); });
  return order.map(function(key) {
    const g = map.get(key);
    const merged = Object.assign({}, g.primary);
    merged['PRODUCT SUPPLY'] = millJoinProductSupplyTokens_(g.members);
    merged._millQtySummary = millBuildQtySummaryFromRows_(g.members);
    merged._millGeneralMerged = g.members.length > 1;
    merged._millGeneralMergeCount = g.members.length;
    return merged;
  });
}

// ── General merge: same company + month + year ─────────────────────────────
const mainRows = [{
  'COMPANY NAME': 'ABDI BORNEO PLANTATIONS',
  MONTH: '2', YEAR: '2026',
  'GROUP NAME': 'SAMLING',
  'RESULT RISK LEVEL': 'HIGH',
  'PRODUCT SUPPLY': 'CPO',
  'SUPPLY CPO': 6052,
  PROVINCE: 'NORTH KALIMANTAN',
}];
const wasteRows = [{
  'COMPANY NAME': 'ABDI BORNEO PLANTATIONS',
  MONTH: '2', YEAR: '2026',
  'GROUP NAME': 'OTHER GROUP',
  'PRODUCT SUPPLY': 'POME INS',
  'SUPPLY INS': 1200,
  PROVINCE: 'SHOULD NOT WIN',
}];

const merged = millMergeGeneralRegistryRows_(mainRows, wasteRows);
assert(merged.length === 1, 'one merged row');
assert(merged[0]['RESULT RISK LEVEL'] === 'HIGH', 'identity from main row');
assert(merged[0].PROVINCE === 'NORTH KALIMANTAN', 'province from main row');
assert(merged[0]['PRODUCT SUPPLY'].indexOf('CPO') >= 0, 'product includes CPO');
assert(merged[0]['PRODUCT SUPPLY'].indexOf('POME INS') >= 0, 'product includes POME INS');
assert(merged[0]._millQtySummary.indexOf('CPO:') >= 0, 'qty includes CPO');
assert(merged[0]._millQtySummary.indexOf('POME INS:') >= 0, 'qty includes POME INS');
assert(merged[0]._millGeneralMerged === true, 'flagged as general merge');

// different period → separate rows
const main2 = [{ 'COMPANY NAME': 'CO A', MONTH: '1', YEAR: '2026', 'SUPPLY CPO': 100 }];
const waste2 = [{ 'COMPANY NAME': 'CO A', MONTH: '2', YEAR: '2026', 'SUPPLY INS': 50 }];
const merged2 = millMergeGeneralRegistryRows_(main2, waste2);
assert(merged2.length === 2, 'different month stays separate');

// waste-only company
const wasteOnly = [{ 'COMPANY NAME': 'WASTE ONLY CO', MONTH: '3', YEAR: '2026', 'SUPPLY ISCC': 500 }];
const merged3 = millMergeGeneralRegistryRows_([], wasteOnly);
assert(merged3.length === 1, 'waste-only row kept');
assert(merged3[0]['PRODUCT SUPPLY'].indexOf('POME ISCC') >= 0, 'waste-only product');

// renamed waste qty headers (SUPPLY POME ISCC / INS) still drive product + qty display
const pomeAliasRow = { 'COMPANY NAME': 'POME ALIAS CO', MONTH: '4', YEAR: '2026', 'SUPPLY POME ISCC': 777, 'SUPPLY POME INS': 88 };
assert(millCollectProductSupplyTokens_(pomeAliasRow).indexOf('POME ISCC') >= 0, 'POME ISCC from SUPPLY POME ISCC');
assert(millCollectProductSupplyTokens_(pomeAliasRow).indexOf('POME INS') >= 0, 'POME INS from SUPPLY POME INS');
assert(String(pomeAliasRow['SUPPLY ISCC']) === '777', 'canonical SUPPLY ISCC filled from alias');
assert(millBuildQtySummaryFromRow_(pomeAliasRow).indexOf('POME ISCC:') >= 0, 'qty summary from POME alias');

// supply CPO/PK display helpers
function millSupplyCpoCellText_(row) {
  const q = millParseSupplyQty_(row['SUPPLY CPO']);
  if (q <= 0) return '—';
  return millFormatSupplyQtyDisplay_(row['SUPPLY CPO']);
}
function millSupplyPkCellText_(row) {
  const q = millParseSupplyQty_(row['SUPPLY PK']);
  if (q <= 0) return '—';
  return millFormatSupplyQtyDisplay_(row['SUPPLY PK']);
}
assert(millSupplyCpoCellText_(mainRows[0]) === '6.052', 'CPO qty formatted');
assert(millSupplyPkCellText_(mainRows[0]) === '—', 'empty PK shows dash');

// Newest period with empty supply should still carry product/qty from older row
function millIsBlankSupplyCell_(v) {
  if (v == null) return true;
  const s = String(v).trim();
  return !s || s === '—' || s === '-' || /^no\s*data$/i.test(s);
}
function millFillEmptySupplyFields_(target, source) {
  ['PRODUCT SUPPLY', 'SUPPLY ISCC', 'SUPPLY INS', 'SUPPLY SHELL'].forEach(function(k) {
    if (!millIsBlankSupplyCell_(target[k])) return;
    if (!millIsBlankSupplyCell_(source[k])) target[k] = source[k];
  });
  return target;
}
function millRowPeriodSortKey_(r) {
  return (parseInt(r.YEAR, 10) || 0) * 100 + (parseInt(r.MONTH, 10) || 0);
}
function millPickNewestCarryForward_(existing, incoming) {
  const sk = millRowPeriodSortKey_(incoming);
  const skOld = millRowPeriodSortKey_(existing);
  if (sk > skOld) {
    const newer = Object.assign({}, incoming);
    millFillEmptySupplyFields_(newer, existing);
    if (millIsBlankSupplyCell_(newer['PRODUCT SUPPLY']) && !millIsBlankSupplyCell_(existing['PRODUCT SUPPLY'])) {
      newer['PRODUCT SUPPLY'] = existing['PRODUCT SUPPLY'];
    }
    return newer;
  }
  if (sk < skOld) {
    const newer = Object.assign({}, existing);
    millFillEmptySupplyFields_(newer, incoming);
    return newer;
  }
  return incoming;
}
const olderWaste = {
  'COMPANY NAME': 'GUNUNG RIJUAN', MONTH: '2', YEAR: '2026',
  'PRODUCT SUPPLY': 'POME INS', 'SUPPLY INS': 1200, _millSheetSource: 'waste',
};
const newerEmpty = {
  'COMPANY NAME': 'GUNUNG RIJUAN', MONTH: '3', YEAR: '2026',
  'PRODUCT SUPPLY': '', 'SUPPLY INS': '', _millSheetSource: 'waste',
};
const carried = millPickNewestCarryForward_(olderWaste, newerEmpty);
assert(carried.MONTH === '3', 'newest month kept');
assert(carried['PRODUCT SUPPLY'] === 'POME INS', 'product supply carried from older period');
assert(String(carried['SUPPLY INS']) === '1200', 'qty carried from older period');

// COMPLIMENT/NOT COMPLIMENT sheet codes → display labels
function millFormatComplimentLabel_(raw) {
  if (raw === undefined || raw === null) return '';
  const s = String(raw).trim();
  if (!s) return '';
  const u = s.toUpperCase().replace(/\s+/g, ' ');
  if (u === 'C' || u === 'COMPLIMENT') return 'Compliment';
  if (u === 'NC' || u === 'NOT COMPLIMENT' || u === 'NON COMPLIMENT' || u === 'NOT COMPLETE') {
    return 'Not Compliment';
  }
  if (/^not\s*compliment$/i.test(s)) return 'Not Compliment';
  if (/^compliment$/i.test(s)) return 'Compliment';
  return s;
}
assert(millFormatComplimentLabel_('C') === 'Compliment', 'C → Compliment');
assert(millFormatComplimentLabel_('NC') === 'Not Compliment', 'NC → Not Compliment');
assert(millFormatComplimentLabel_('compliment') === 'Compliment', 'compliment passthrough');
assert(millFormatComplimentLabel_('') === '', 'empty compliment');

// parametric: 50 merge scenarios
for (let i = 0; i < 50; i++) {
  const co = 'COMPANY ' + (i % 10);
  const m = String((i % 12) + 1);
  const y = String(2026 + (i % 3));
  const mRow = { 'COMPANY NAME': co, MONTH: m, YEAR: y, 'SUPPLY CPO': i + 1 };
  const wRow = { 'COMPANY NAME': co, MONTH: m, YEAR: y, 'SUPPLY INS': i + 2 };
  const out = millMergeGeneralRegistryRows_([mRow], [wRow]);
  assert(out.length === 1, 'parametric merge ' + i);
  assert(out[0]._millGeneralMerged === true, 'parametric merged flag ' + i);
}

console.log('\nMill registry view tests:', passed, 'passed,', failed, 'failed');
process.exit(failed ? 1 : 0);
