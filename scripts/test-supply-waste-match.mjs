/**
 * Waste Task List profile match — month N should match month N-1 waste / task list.
 * Run: node scripts/test-supply-waste-match.mjs
 */

let passed = 0;
let failed = 0;
function assert(cond, msg) {
  if (cond) { passed++; return; }
  failed++;
  console.error('FAIL:', msg);
}

function supplyNormKey_(s) {
  return String(s || '').trim().toLowerCase().replace(/\s+/g, ' ');
}
function supplyCompanyKey_(company) {
  return supplyNormKey_(company);
}
function supplyImportIsWaste_(kind) {
  const k = String(kind || '').trim().toUpperCase();
  if (k === 'POME_ISCC' || k === 'POME_INS' || k === 'SHELL_GGL') return true;
  return k.indexOf('POME_ISCC') >= 0 || k.indexOf('POME_INS') >= 0 || k.indexOf('SHELL_GGL') >= 0;
}
function parseMillYearSort(v) {
  const m = String(v || '').trim().match(/(19|20)\d{2}/);
  return m ? parseInt(m[0], 10) : (parseInt(String(v || ''), 10) || 0);
}
function parseMillMonthSort(v) {
  const n = parseInt(String(v || '').trim(), 10);
  return (!isNaN(n) && n >= 1 && n <= 12) ? n : 0;
}
function millRowPeriodSortKey_(r) {
  const y = parseMillYearSort(r.YEAR || r.year);
  if (!y) return 0;
  const m = parseMillMonthSort(r.MONTH || r.month);
  return y * 100 + (m || 0);
}
function supplyNameMatches_(a, b) {
  return supplyNormKey_(a) === supplyNormKey_(b);
}
function supplyCompanyMatchesProfile_(excelCompany, profileRow) {
  const ec = String(excelCompany || '').trim();
  const pc = String(profileRow['COMPANY NAME'] || '').trim();
  if (!ec || !pc) return false;
  return supplyNameMatches_(ec, pc);
}
function supplyMillRowIsEmptySlot_(profile) {
  const keys = ['COMPANY NAME', 'MILL NAME', 'GROUP NAME', 'UML ID'];
  return !keys.some(function(k) {
    const v = String(profile[k] || '').trim();
    return v && v !== '—';
  });
}
function supplyOnboardingPoolForBatch_(batch, pools) {
  if (supplyImportIsWaste_(batch && batch.supply_type)) return pools.waste;
  return pools.main;
}
function supplyOnboardingPoolForKind_(supplyKind, pools) {
  return supplyOnboardingPoolForBatch_({ supply_type: supplyKind }, pools);
}
function supplyPickBestProfileFromPool_(pool, company) {
  const want = String(company || '').trim();
  if (!want || !pool || !pool.length) return null;
  let best = null;
  let bestSk = 0;
  pool.forEach(function(d) {
    if (!supplyCompanyMatchesProfile_(want, d)) return;
    if (supplyMillRowIsEmptySlot_(d)) return;
    const sk = millRowPeriodSortKey_(d);
    if (!best || sk > bestSk) { best = d; bestSk = sk; }
  });
  return best;
}
function supplyTaskListRowIsProfileReference_(row) {
  if (!row) return false;
  if (row._submitted || row.status === 'submitted' || row.profile_draft_saved === 'true') return true;
  if (String(row.match_status || '').toLowerCase() === 'matched') return true;
  return !!(String(row['GROUP NAME'] || '').trim());
}
function supplyWasteProfileMatchPriority_(candidate, supplyKind) {
  if (!supplyImportIsWaste_(supplyKind)) return candidate && candidate._row ? 2 : 1;
  if (candidate && candidate._millSheetSource === 'waste') return 4;
  if (candidate && supplyTaskListRowIsProfileReference_(candidate) && !candidate._millSheetSource) return 3;
  return candidate && candidate._row ? 2 : 1;
}
function supplyPickLatestTaskListProfileForCompany_(company, supplyKind, batches, opts) {
  opts = opts || {};
  const wantCo = supplyCompanyKey_(company);
  if (!wantCo) return null;
  const kindIsWaste = supplyImportIsWaste_(supplyKind);
  let best = null;
  let bestSk = 0;
  let bestPri = 0;
  (batches || []).forEach(function(b) {
    if (opts.excludeBatchId && b.batch_id === opts.excludeBatchId) return;
    if (supplyImportIsWaste_(b.supply_type) !== kindIsWaste) return;
    (b.rows || []).forEach(function(row) {
      if (supplyCompanyKey_(row['COMPANY NAME']) !== wantCo) return;
      if (!supplyTaskListRowIsProfileReference_(row)) return;
      const sk = millRowPeriodSortKey_(row);
      const pri = supplyWasteProfileMatchPriority_(row, supplyKind);
      if (!best || sk > bestSk || (sk === bestSk && pri > bestPri)) { best = row; bestSk = sk; bestPri = pri; }
    });
  });
  return best;
}
function supplyPickLatestMillProfileForCompany_(company, supplyKind, pools, batches, opts) {
  opts = opts || {};
  const kind = String(supplyKind || 'CPO').trim().toUpperCase();
  const want = String(company || '').trim();
  if (!want) return null;
  const candidates = [];
  const onboardingBest = supplyPickBestProfileFromPool_(supplyOnboardingPoolForKind_(kind, pools), want);
  if (onboardingBest) candidates.push(onboardingBest);
  if (supplyImportIsWaste_(kind)) {
    const taskBest = supplyPickLatestTaskListProfileForCompany_(want, kind, batches, opts);
    if (taskBest) candidates.push(taskBest);
    if (!onboardingBest && !taskBest) {
      const mainBest = supplyPickBestProfileFromPool_(pools.main, want);
      if (mainBest) candidates.push(mainBest);
    }
  }
  let best = null;
  let bestSk = 0;
  let bestPri = 0;
  candidates.forEach(function(c) {
    const sk = millRowPeriodSortKey_(c);
    const pri = supplyWasteProfileMatchPriority_(c, kind);
    if (!best || sk > bestSk || (sk === bestSk && pri > bestPri)) { best = c; bestSk = sk; bestPri = pri; }
  });
  return best;
}
function supplyFindMillProfileMatch_(excelRow, supplyKind, pools, batches, opts) {
  const company = String(excelRow.COMPANY_NAME || excelRow.company || '').trim();
  if (!company) return { status: 'new', row: null };
  const ref = supplyPickLatestMillProfileForCompany_(company, supplyKind, pools, batches, opts);
  return ref ? { status: 'matched', row: ref } : { status: 'new', row: null };
}

const pools = {
  main: [
    { 'COMPANY NAME': 'WASTE CO', MONTH: '1', YEAR: '2026', 'SUPPLY CPO': 100, _row: 10 },
    { 'COMPANY NAME': 'DUAL CO', MONTH: '1', YEAR: '2026', 'SUPPLY CPO': 100, _row: 11 },
  ],
  waste: [
    { 'COMPANY NAME': 'WASTE CO', MONTH: '1', YEAR: '2026', 'SUPPLY ISCC': 50, _row: 20, _millSheetSource: 'waste' },
    { 'COMPANY NAME': 'POME ONLY', MONTH: '1', YEAR: '2026', 'SUPPLY ISCC': 30, _row: 21, _millSheetSource: 'waste' },
  ],
};
const batches = [{
  batch_id: 'b-jan', month: '1', year: '2026', supply_type: 'POME_ISCC',
  rows: [{
    'COMPANY NAME': 'OTHER WASTE', MONTH: '1', YEAR: '2026',
    'GROUP NAME': 'GRP A', profile_draft_saved: 'true', match_status: 'matched',
  }, {
    'COMPANY NAME': 'POME ONLY', MONTH: '1', YEAR: '2026',
    'GROUP NAME': 'GRP B', status: 'submitted', match_status: 'matched',
  }],
}];

// CPO import still matches main sheet only
const cpoMatch = supplyFindMillProfileMatch_({ COMPANY_NAME: 'WASTE CO' }, 'CPO', pools, batches);
assert(cpoMatch.status === 'matched', 'CPO matches main onboarding');
assert(cpoMatch.row._row === 10, 'CPO profile from main sheet');

// POME month 2 import matches month-1 waste onboarding (not CPO main)
const pomeMatch = supplyFindMillProfileMatch_({ COMPANY_NAME: 'WASTE CO' }, 'POME_ISCC', pools, batches);
assert(pomeMatch.status === 'matched', 'POME matches waste onboarding');
assert(pomeMatch.row._row === 20, 'POME profile from waste sheet not CPO row');

// Company with CPO on main + waste row — POME prefers waste sheet
const dualMatch = supplyFindMillProfileMatch_({ COMPANY_NAME: 'DUAL CO' }, 'POME_ISCC', {
  main: pools.main,
  waste: [{ 'COMPANY NAME': 'DUAL CO', MONTH: '1', YEAR: '2026', 'SUPPLY ISCC': 10, _row: 22, _millSheetSource: 'waste' }],
}, batches);
assert(dualMatch.status === 'matched', 'Dual company POME matches waste not CPO');
assert(dualMatch.row._millSheetSource === 'waste', 'Dual company uses waste profile');

// POME-only company matches via prior submitted task list when waste pool empty
const pomeOnlyPools = { main: pools.main, waste: [] };
const pomeOnlyMatch = supplyFindMillProfileMatch_({ COMPANY_NAME: 'POME ONLY' }, 'POME_ISCC', pomeOnlyPools, batches);
assert(pomeOnlyMatch.status === 'matched', 'POME-only matches January submitted task list');
assert(pomeOnlyMatch.row['GROUP NAME'] === 'GRP B', 'Profile from submitted task list row');

// Shell import matches prior task list when no onboarding row
const shellMatch = supplyFindMillProfileMatch_({ COMPANY_NAME: 'OTHER WASTE' }, 'SHELL_GGL', pools, batches);
assert(shellMatch.status === 'matched', 'Shell matches previous month task list');
assert(shellMatch.row['GROUP NAME'] === 'GRP A', 'Profile copied from task list row');

console.log('\nSupply waste match tests:', passed, 'passed,', failed, 'failed');
process.exit(failed ? 1 : 0);
