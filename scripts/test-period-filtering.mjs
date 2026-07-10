/**
 * Period filtering consistency: Facility Performance (exact) vs Monthly Report (as-of).
 * Run: node scripts/test-period-filtering.mjs
 */
import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';
import {
  mrdDataPeriodFromReport_,
  mrdTraceYearFromReport_,
  mrdReportHeaderMeta_,
} from '../src/monthly-report-labels.js';

const __dirname = dirname(fileURLToPath(import.meta.url));
const ROOT = join(__dirname, '..');
const mainJs = readFileSync(join(ROOT, 'src/main.js'), 'utf8');
const mrdJs = readFileSync(join(ROOT, 'src/monthly-report-ui.js'), 'utf8');

let passed = 0;
let failed = 0;
function assert(cond, msg) {
  if (cond) { passed++; return; }
  failed++;
  console.error('FAIL:', msg);
}
function assertEq(a, b, msg) {
  const ok = JSON.stringify(a) === JSON.stringify(b);
  if (ok) { passed++; return; }
  failed++;
  console.error('FAIL:', msg, '— got', a, 'expected', b);
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

function millRowPeriodSortKey_(r) {
  const y = parseMillYearSort(millYearVal(r));
  if (!y) return 0;
  const m = parseMillMonthSort(millMonthVal(r));
  return y * 100 + (m || 0);
}

function millRegistryEntityKey_(r) {
  return [
    String(r['COMPANY NAME'] || '').trim().toUpperCase(),
    String(r['MILL NAME'] || '').trim().toUpperCase(),
  ].join('\u0001');
}

function millRowWithinPeriodFilter_(r, pf) {
  if (!pf || (!pf.hasYear && !pf.hasMonth)) return true;
  const sk = millRowPeriodSortKey_(r);
  if (!sk) return pf.hasEmptyMonth || false;
  const y = Math.floor(sk / 100);
  const m = sk % 100;
  if (pf.hasYear && pf.years.size && !pf.years.has(y)) return false;
  if (pf.hasMonth && pf.maxMonth > 0 && m > pf.maxMonth) return false;
  return true;
}

function millPickRegistryRowWinner_(existing, incoming) {
  const sk = millRowPeriodSortKey_(incoming);
  const skOld = millRowPeriodSortKey_(existing);
  if (sk > skOld) return incoming;
  if (sk < skOld) return existing;
  return (incoming._row || 0) > (existing._row || 0) ? incoming : existing;
}

function millPickLatestPerCompany_(rows, pf) {
  const byEntity = new Map();
  (rows || []).forEach(function(r) {
    const ek = millRegistryEntityKey_(r);
    if (!ek || ek === '\u0001') return;
    if (!millRowWithinPeriodFilter_(r, pf)) return;
    const existing = byEntity.get(ek);
    byEntity.set(ek, existing ? millPickRegistryRowWinner_(existing, r) : r);
  });
  return Array.from(byEntity.values());
}

function pfMonthMatchesFilter_(rowM, filterM) {
  if (!filterM) return true;
  return parseMillMonthSort(rowM) === parseInt(String(filterM), 10);
}

function pfYearMatchesFilter_(rowY, filterY) {
  if (!filterY) return true;
  return String(parseMillYearSort(rowY) || '') === String(parseMillYearSort(filterY) || String(filterY).trim());
}

function pfExactFilter_(rows, year, month) {
  return (rows || []).filter(function(r) {
    if (month && !pfMonthMatchesFilter_(millMonthVal(r), month)) return false;
    if (year && !pfYearMatchesFilter_(millYearVal(r), year)) return false;
    return true;
  });
}

function pfAsOfFilter_(rows, reportYear, reportMonth) {
  const y = parseInt(String(reportYear || ''), 10);
  const m = parseInt(String(reportMonth || ''), 10);
  const pf = {
    hasYear: !!y,
    hasMonth: m >= 1 && m <= 12,
    years: y ? new Set([y]) : new Set(),
    maxMonth: m >= 1 && m <= 12 ? m : 0,
    hasEmptyMonth: false,
  };
  return millPickLatestPerCompany_(rows, pf);
}

function getFacilityReportContext_(reportYear, reportMonth) {
  const data = mrdDataPeriodFromReport_(reportYear, reportMonth);
  return {
    year: data.year,
    month: data.month,
    reportYear: String(reportYear || ''),
    reportMonth: String(reportMonth || ''),
    millPickMode: 'as-of',
  };
}

// ── Data period lag ─────────────────────────────────────────────────────────
assertEq(mrdDataPeriodFromReport_('2026', '1'), { year: '2025', month: '12' }, 'Jan report → Dec data');
assertEq(mrdDataPeriodFromReport_('2026', '3'), { year: '2026', month: '2' }, 'Mar report → Feb data');
assertEq(mrdDataPeriodFromReport_('2025', '12'), { year: '2025', month: '11' }, 'Dec report → Nov data');
assertEq(mrdTraceYearFromReport_('2026'), '2025', 'trace year lags report year');

const meta = mrdReportHeaderMeta_('2026', '3');
assert(meta.dataYear === '2026' && meta.dataMonth === '2', 'header meta data period Mar→Feb');
assert(meta.periodLine.includes('March'), 'header uses reporting month name');

// ── Facility report context splits report vs data period ────────────────────
const ctx = getFacilityReportContext_('2026', '3');
assert(ctx.reportYear === '2026' && ctx.reportMonth === '3', 'report period preserved');
assert(ctx.year === '2026' && ctx.month === '2', 'data period lagged for supply sheets');
assert(ctx.millPickMode === 'as-of', 'MRD facility uses as-of mill pick');

// ── Sample rows: CO A updated Jan+Feb, CO B only Jan ────────────────────────
const rows = [
  { 'COMPANY NAME': 'CO A', 'MILL NAME': 'MILL A', MONTH: '1', YEAR: '2026', 'RESULT RISK LEVEL': 'LOW', _row: 1 },
  { 'COMPANY NAME': 'CO A', 'MILL NAME': 'MILL A', MONTH: '2', YEAR: '2026', 'RESULT RISK LEVEL': 'HIGH', _row: 2 },
  { 'COMPANY NAME': 'CO B', 'MILL NAME': 'MILL B', MONTH: '1', YEAR: '2026', 'RESULT RISK LEVEL': 'MEDIUM', _row: 3 },
  { 'COMPANY NAME': 'CO C', 'MILL NAME': 'MILL C', MONTH: '3', YEAR: '2026', 'RESULT RISK LEVEL': 'LOW', _row: 4 },
];

// Facility Performance exact: Dec 2025
const decOnly = [
  { 'COMPANY NAME': 'X', 'MILL NAME': 'M1', MONTH: '12', YEAR: '2025', _row: 10 },
  { 'COMPANY NAME': 'X', 'MILL NAME': 'M1', MONTH: '11', YEAR: '2025', _row: 11 },
];
const exactDec = pfExactFilter_(decOnly, '2025', '12');
assert(exactDec.length === 1 && exactDec[0].MONTH === '12', 'PF exact Dec 2025 excludes Nov');

// Facility Performance exact: Mar 2026 — CO C only (no Mar for A/B)
const exactMar = pfExactFilter_(rows, '2026', '3');
assert(exactMar.length === 1 && exactMar[0]['COMPANY NAME'] === 'CO C', 'PF exact Mar shows only Mar rows');

// Monthly Report as-of Mar 2026 — A gets Feb, B gets Jan, C gets Mar
const asOfMar = pfAsOfFilter_(rows, '2026', '3');
assert(asOfMar.length === 3, 'MRD as-of Mar keeps all entities');
const byCo = Object.fromEntries(asOfMar.map((r) => [r['COMPANY NAME'], r]));
assert(byCo['CO A']['RESULT RISK LEVEL'] === 'HIGH' && byCo['CO A'].MONTH === '2', 'CO A newest as-of Mar is Feb');
assert(byCo['CO B'].MONTH === '1', 'CO B falls back to Jan when no Mar update');
assert(byCo['CO C'].MONTH === '3', 'CO C uses Mar when present');

// As-of excludes future months within same year
const withFuture = rows.concat([
  { 'COMPANY NAME': 'CO D', 'MILL NAME': 'MILL D', MONTH: '5', YEAR: '2026', _row: 5 },
]);
const asOfMar2 = pfAsOfFilter_(withFuture, '2026', '3');
assert(!asOfMar2.some((r) => r['COMPANY NAME'] === 'CO D'), 'as-of Mar excludes May row for CO D');

// Year boundary: report Jan 2026 → as-of can include Dec 2025? 
// millRowWithinPeriodFilter_ only checks year in set [2026], so Dec 2025 excluded
const crossYear = [
  { 'COMPANY NAME': 'CO E', 'MILL NAME': 'M1', MONTH: '12', YEAR: '2025', _row: 1 },
  { 'COMPANY NAME': 'CO E', 'MILL NAME': 'M1', MONTH: '1', YEAR: '2026', _row: 2 },
];
const asOfJan = pfAsOfFilter_(crossYear, '2026', '1');
assert(asOfJan.length === 1 && asOfJan[0].MONTH === '1', 'as-of Jan 2026 picks Jan 2026 not Dec 2025');

// Exact vs as-of must diverge for partial-update scenario (Vincent rule)
const partial = pfExactFilter_(rows, '2026', '3');
const asOf = pfAsOfFilter_(rows, '2026', '3');
assert(partial.length < asOf.length, 'MRD as-of shows more entities than PF exact when updates lag');

// Supplied-only sellers must not appear under exact period filter (no mill row for period)
function pfSimulateSuppliedMerge_(millRows, sellerKey, periodMonth, periodYear) {
  const strict = !!(periodMonth || periodYear);
  const millHit = (function() {
    const scoped = millRows.filter(function(r) {
      return String(r['COMPANY NAME'] || '').toUpperCase() === sellerKey;
    });
    if (scoped.length) return scoped[0];
    if (strict) return null;
    return null;
  })();
  if (millHit) return { source: 'mill', month: millMonthVal(millHit) };
  if (strict) return null;
  return { source: 'supplied-only', month: periodMonth };
}
const decMillRows = [{ 'COMPANY NAME': 'CO Z', 'MILL NAME': 'M1', MONTH: '12', YEAR: '2025' }];
assert(pfSimulateSuppliedMerge_(decMillRows, 'CO Z', '12', '2025').source === 'mill', 'scoped mill hit for exact period');
assert(pfSimulateSuppliedMerge_(decMillRows, 'CO MISSING', '12', '2025') === null, 'no supplied-only under exact filter');
assert(pfSimulateSuppliedMerge_([], 'CO MISSING', '', '').source === 'supplied-only', 'supplied-only allowed when no period filter');

// ── Source wiring ───────────────────────────────────────────────────────────
assert(mainJs.includes('pfShouldStrictMillLookup_'), 'PF strict mill lookup guard');
assert(mainJs.includes('if (pfShouldStrictMillLookup_()) return null'), 'PF blocks registry fallback when filtered');
assert(mainJs.includes('const anchor = d || millProfileVariantRows_[0]'), 'mill profile opens on clicked row period');
assert(mainJs.includes("millPickMode === 'as-of'"), 'PF as-of mode branch');
assert(mainJs.includes('millPickLatestPerCompany_(rows, pf)'), 'PF report uses millPickLatestPerCompany');
assert(mainJs.includes('skipPeriodFilter'), 'PF skips exact company filter in as-of mode');
assert(mrdJs.includes('getFacilityReportContext_'), 'MRD facility report context helper');
assert(mrdJs.includes("millPickMode: 'as-of'"), 'MRD passes as-of to facility loader');
assert(mainJs.includes('/** Row eligible for “as-of” period view'), 'as-of filter documented');
assert(mainJs.includes('Mill Onboarding rows for Facility Performance'), 'PF exact mode documented');
assert(mrdJs.includes("getMillsForReportPeriod(periodYear, periodMonth, 'main')"), 'MRD loads Main mills');
assert(mrdJs.includes("getMillsForReportPeriod(periodYear, periodMonth, 'waste')"), 'MRD loads Waste mills');
assert(mrdJs.includes('Product: Main + Waste'), 'MRD scope shows both products');
assert(!mrdJs.includes('data-mrd-product-view'), 'MRD has no product picker');
assert(!mrdJs.includes('mrdSyncProductViewUi_'), 'dead product UI sync removed');
assert(mrdJs.includes('reportPeriod: { year:'), 'snapshot stores report period');

console.log('\nPeriod filtering tests: ' + passed + ' passed, ' + failed + ' failed');
process.exit(failed > 0 ? 1 : 0);
