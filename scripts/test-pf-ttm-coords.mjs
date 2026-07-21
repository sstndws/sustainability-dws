/**
 * Facility Performance TTM coordinate resolution (year-only / dedupe).
 * Run: node scripts/test-pf-ttm-coords.mjs
 */
import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const mainJs = readFileSync(join(__dirname, '..', 'src', 'main.js'), 'utf8');

let passed = 0;
let failed = 0;
function assert(cond, msg) {
  if (cond) { passed++; return; }
  failed++;
  console.error('FAIL:', msg);
}

function millYearVal(r) { return r.YEAR || r['YEAR'] || ''; }
function millMonthVal(r) { return r.MONTH || r['MONTH'] || ''; }

function parseMillYearSort(v) {
  const s = String(v || '').trim();
  const m = s.match(/(19|20)\d{2}/);
  if (m) return parseInt(m[0], 10);
  const n = parseInt(s, 10);
  return isNaN(n) ? 0 : n;
}

function parseMillMonthSort(v) {
  const n = parseInt(String(v || '').trim(), 10);
  if (!isNaN(n) && n >= 1 && n <= 12) return n;
  return 0;
}

function pfMillRowPeriodSortKey_(r) {
  const y = parseMillYearSort(millYearVal(r));
  if (!y) return 0;
  const m = parseMillMonthSort(millMonthVal(r));
  return y * 100 + (m || 0);
}

function pfMillRowIdentityKey_(r) {
  const id = String(r['UML ID'] || '').trim().toUpperCase();
  if (id) return 'id:' + id;
  const co = String(r['COMPANY NAME'] || '').trim().toUpperCase();
  const mill = String(r['MILL NAME'] || '').trim().toUpperCase();
  return 'cm:' + co + '\x1f' + mill;
}

function pfCompanyCoordValue_(row) {
  if (!row) return '';
  const raw = row['COORDINATES'] || row['Coordinates'] || row['Coordinate'] || row['COORDINATE'] || '';
  return String(raw).trim();
}

function pfIsValidCoord_(coord) {
  const s = String(coord || '').trim();
  if (!s || s === '—') return false;
  if (/^(no\s*data|n\/a|na|none|null|-+)$/i.test(s)) return false;
  const nums = s.match(/-?\d+(?:[.,]\d+)?/g);
  if (!nums || nums.length < 2) return false;
  const lat = parseFloat(String(nums[0]).replace(',', '.'));
  const lng = parseFloat(String(nums[1]).replace(',', '.'));
  if (isNaN(lat) || isNaN(lng)) return false;
  if (Math.abs(lat) > 90 || Math.abs(lng) > 180) return false;
  if (lat === 0 && lng === 0) return false;
  return true;
}

function pfRowHasTraceableCoord_(row) {
  return pfIsValidCoord_(pfCompanyCoordValue_(row));
}

function pfMergeCoordOntoMillRow_(target, source) {
  if (!target || !source || pfRowHasTraceableCoord_(target)) return target;
  const coord = pfCompanyCoordValue_(source);
  if (!pfIsValidCoord_(coord)) return target;
  return Object.assign({}, target, {
    COORDINATES: coord,
    COORDINATE: coord,
  });
}

function pfDedupeMillRowsLatest_(rows) {
  const byId = new Map();
  rows.forEach(function(r) {
    const key = pfMillRowIdentityKey_(r);
    const existing = byId.get(key);
    if (!existing) {
      byId.set(key, r);
      return;
    }
    const skNew = pfMillRowPeriodSortKey_(r);
    const skOld = pfMillRowPeriodSortKey_(existing);
    let winner = existing;
    let other = r;
    if (skNew > skOld) {
      winner = r;
      other = existing;
    } else if (skNew === skOld && (r._row || 0) > (existing._row || 0)) {
      winner = r;
      other = existing;
    }
    byId.set(key, pfMergeCoordOntoMillRow_(winner, other));
  });
  return Array.from(byId.values());
}

function pfSellerMatchesCompany_(sellerUpper, companyName) {
  const co = String(companyName || '').trim().toUpperCase();
  return co === sellerUpper;
}

function pfFindBestMillCoordForSeller_(sellerUpper, periodRows, registryRows) {
  const candidates = [];
  function collect(rows) {
    (rows || []).forEach(function(r) {
      if (!pfSellerMatchesCompany_(sellerUpper, r['COMPANY NAME'])) return;
      if (!pfRowHasTraceableCoord_(r)) return;
      const coord = pfCompanyCoordValue_(r);
      if (!pfIsValidCoord_(coord)) return;
      candidates.push({ coord, sortKey: pfMillRowPeriodSortKey_(r), rowNum: r._row || 0 });
    });
  }
  collect(periodRows);
  if (!candidates.length) collect(registryRows);
  if (!candidates.length) return '';
  candidates.sort(function(a, b) {
    if (b.sortKey !== a.sortKey) return b.sortKey - a.sortKey;
    return b.rowNum - a.rowNum;
  });
  return candidates[0].coord;
}

// Latest month row lacks coords; earlier month in same year has coords
const yearRows = [
  { 'COMPANY NAME': 'AHANA PLANTATION AND PRODUCT', 'UML ID': 'PO1000013891', MONTH: '2', YEAR: '2026', COORDINATE: '0.01184, 111.55973', _row: 1 },
  { 'COMPANY NAME': 'AHANA PLANTATION AND PRODUCT', 'UML ID': 'PO1000013891', MONTH: '3', YEAR: '2026', COORDINATE: '', _row: 2 },
];
const deduped = pfDedupeMillRowsLatest_(yearRows);
assert(deduped.length === 1, 'dedupe keeps one row per mill');
assert(deduped[0].MONTH === '3', 'dedupe keeps latest month');
assert(pfIsValidCoord_(pfCompanyCoordValue_(deduped[0])), 'dedupe carries coords from earlier month');

const period2026 = yearRows.filter(function(r) { return r.YEAR === '2026'; });
const best = pfFindBestMillCoordForSeller_('AHANA PLANTATION AND PRODUCT', period2026, []);
assert(best === '0.01184, 111.55973', 'best coord from any month in period');

// Registry fallback when period rows have no coords
const registry = [
  { 'COMPANY NAME': 'PENITI SUNGAI PURUN', MONTH: '12', YEAR: '2025', COORDINATE: '-0.123, 109.456', _row: 10 },
];
const best2 = pfFindBestMillCoordForSeller_('PENITI SUNGAI PURUN', [], registry);
assert(best2 === '-0.123, 109.456', 'registry fallback for coordinates');

// COORDINATE singular header (Mill Onboarding sheet column L)
const singular = { 'COMPANY NAME': 'TEST MILL', COORDINATE: '0.318318, 111.1825' };
assert(pfIsValidCoord_(pfCompanyCoordValue_(singular)), 'COORDINATE singular header accepted');

// MILL NAME match when SELLER uses mill name not company name
const millNameRow = {
  'COMPANY NAME': 'AHANA PLANTATION AND PRODUCT',
  'MILL NAME': 'AHANA MILL UNIT 1',
  MONTH: '3', YEAR: '2026', COORDINATE: '0.01184, 111.55973', _row: 608,
};
function pfSellerMatchesMillRow_(sellerUpper, row) {
  function norm(name) {
    let s = String(name || '').trim().toUpperCase().replace(/[^A-Z0-9]/g, '').toLowerCase();
    return s.replace(/^(pt|pd|cv|tbk|persero)/, '');
  }
  function match(a, b) {
    const na = norm(a); const nb = norm(b);
    return na && nb && na === nb;
  }
  if (match(sellerUpper, row['COMPANY NAME'])) return true;
  if (match(sellerUpper, row['MILL NAME'])) return true;
  return false;
}
assert(pfSellerMatchesMillRow_('AHANA MILL UNIT 1', millNameRow), 'SELLER matches MILL NAME');
assert(pfSellerMatchesMillRow_('AHANA PLANTATION AND PRODUCT', millNameRow), 'SELLER matches COMPANY NAME');

// PT prefix stripped
function pfNormalizeSellerKey_(name) {
  let s = String(name || '').trim().toUpperCase().replace(/[^A-Z0-9]/g, '').toLowerCase();
  return s.replace(/^(pt|pd|cv|tbk|persero)/, '');
}
assert(
  pfNormalizeSellerKey_('PT PENITI SUNGAI PURUN') === pfNormalizeSellerKey_('PENITI SUNGAI PURUN'),
  'PT prefix normalization'
);

// Source wiring
// Mill Onboarding supply lookup (2026+)
function pfShouldUseLegacySuppliedSheets_(y) {
  const yr = parseInt(String(y || ''), 10);
  return !yr || yr < 2026;
}
function pfBuildMillSuppliedLookup_(millRows, productKind) {
  const byPlant = {};
  const isPk = productKind === 'pk';
  (millRows || []).forEach(function(r) {
    const qtyCol = isPk ? 'SUPPLY PK' : 'SUPPLY CPO';
    const facCol = isPk ? 'FACILITY NAME PK' : 'FACILITY NAME CPO';
    let qty = parseFloat(String(r[qtyCol] || '').replace(/,/g, ''));
    if (isNaN(qty) || qty <= 0) return;
    const qtyKg = qty * 1000;
    const seller = String(r['COMPANY NAME'] || '').trim().toUpperCase();
    const plant = String(r[facCol] || '').trim().toUpperCase();
    if (!seller || !plant) return;
    if (!byPlant[plant]) byPlant[plant] = { sellers: {}, totalQty: 0 };
    if (!byPlant[plant].sellers[seller]) byPlant[plant].sellers[seller] = 0;
    byPlant[plant].sellers[seller] += qtyKg;
    byPlant[plant].totalQty += qtyKg;
  });
  return byPlant;
}
assert(!pfShouldUseLegacySuppliedSheets_('2026'), '2026 skips legacy Supplied CPO sheets');
assert(pfShouldUseLegacySuppliedSheets_('2025'), '2025 still uses legacy sheets');
const millRows2026 = [{
  'COMPANY NAME': 'AHANA PLANTATION AND PRODUCT',
  'FACILITY NAME CPO': 'MEDCO',
  'SUPPLY CPO': 500,
  MONTH: '3', YEAR: '2026',
  COORDINATE: '0.01184, 111.55973',
}];
const millLookup = pfBuildMillSuppliedLookup_(millRows2026, 'cpo');
assert(millLookup.MEDCO && millLookup.MEDCO.sellers['AHANA PLANTATION AND PRODUCT'] === 500000, 'mill SUPPLY CPO → kg lookup');

assert(mainJs.includes('pfShouldUseLegacySuppliedSheets_'), 'legacy sheet gate');
assert(mainJs.includes('pfBuildMillSuppliedLookup_'), 'mill supply lookup');
assert(mainJs.includes('pfResolveSuppliedCpoLookup_'), 'resolved CPO lookup');
assert(mainJs.includes('if (!pfShouldUseLegacySuppliedSheets_()) return [];'), 'skip legacy fetch for 2026');

console.log('\nPF TTM coordinate tests: ' + passed + ' passed, ' + failed + ' failed');
process.exit(failed > 0 ? 1 : 0);
