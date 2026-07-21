/**
 * Simulates EUDR group-name backfill for mills where the newest period row lacks GROUP NAME.
 */
import assert from 'node:assert/strict';

function eudrEntityKey_(group, company, mill) {
  return [group, company, mill].map(function(s) {
    return String(s || '').trim().toLowerCase();
  }).join('|');
}

function pickMillGroupName_(row) {
  const s = String(row['GROUP NAME'] || '').trim();
  return s && s !== '—' && s !== '-' ? s : '';
}

function pickMillCompanyName_(row) {
  return String(row['COMPANY NAME'] || '').trim();
}

function millPickField_(row, keys) {
  for (const k of keys) {
    const v = row[k];
    if (v != null && String(v).trim()) return String(v).trim();
  }
  return '';
}

function millYearVal(r) { return String(r.YEAR || r['YEAR'] || '').trim(); }
function millMonthVal(r) { return String(r.MONTH || r['MONTH'] || '').trim(); }

function eudrCompanyMillKey_(company, mill) {
  return eudrEntityKey_('', company, mill);
}

function eudrBestGroupNameForMill_(allData, company, mill) {
  let best = '';
  (allData || []).forEach(function(r) {
    if (pickMillCompanyName_(r) !== company) return;
    if (millPickField_(r, ['MILL NAME', 'Mill Name']) !== mill) return;
    const g = pickMillGroupName_(r);
    if (g) best = g;
  });
  return best;
}

function eudrBuildMillRegistry_(allData) {
  const seen = {};
  const groups = {};
  const dedupArr = [];
  (allData || []).forEach(function(r) {
    const company = pickMillCompanyName_(r);
    const mill = millPickField_(r, ['MILL NAME', 'Mill Name']);
    const key = eudrCompanyMillKey_(company, mill);
    if (!company || !mill || !key || key === '|') return;
    const group = pickMillGroupName_(r);
    if (group) groups[key] = group;
    if (seen[key] === undefined) {
      seen[key] = dedupArr.length;
      dedupArr.push(r);
    } else {
      const prev = dedupArr[seen[key]];
      const prevYear = parseInt(millYearVal(prev), 10) || 0;
      const curYear = parseInt(millYearVal(r), 10) || 0;
      const prevMonth = parseInt(millMonthVal(prev), 10) || 0;
      const curMonth = parseInt(millMonthVal(r), 10) || 0;
      if (curYear > prevYear || (curYear === prevYear && curMonth > prevMonth)) {
        dedupArr[seen[key]] = r;
      }
    }
  });
  return dedupArr.map(function(r) {
    const company = pickMillCompanyName_(r);
    const mill = millPickField_(r, ['MILL NAME', 'Mill Name']);
    const cmKey = eudrCompanyMillKey_(company, mill);
    const group = pickMillGroupName_(r) || groups[cmKey] || eudrBestGroupNameForMill_(allData, company, mill);
    return { 'GROUP NAME': group, 'COMPANY NAME': company, 'MILL NAME': mill };
  });
}

// CITRA MAHKOTA scenario: newest row missing group, older row has KPN PLANTATION
const allData = [
  { YEAR: '2024', MONTH: '6', 'GROUP NAME': 'KPN PLANTATION', 'COMPANY NAME': 'CITRA MAHKOTA', 'MILL NAME': 'CITRA MAHKOTA' },
  { YEAR: '2025', MONTH: '12', 'GROUP NAME': '', 'COMPANY NAME': 'CITRA MAHKOTA', 'MILL NAME': 'CITRA MAHKOTA' },
];

const out = eudrBuildMillRegistry_(allData);
assert.equal(out.length, 1, 'one mill entry');
assert.equal(out[0]['GROUP NAME'], 'KPN PLANTATION', 'group backfilled from older row');
assert.equal(out[0]['COMPANY NAME'], 'CITRA MAHKOTA');

// Old bug: dedup by group+company+mill would produce two entries
function oldBuild(allData) {
  const seen = {};
  const dedupArr = [];
  allData.forEach(function(r) {
    const group = pickMillGroupName_(r);
    const company = pickMillCompanyName_(r);
    const mill = millPickField_(r, ['MILL NAME']);
    const key = eudrEntityKey_(group, company, mill);
    if (seen[key] === undefined) {
      seen[key] = dedupArr.length;
      dedupArr.push(r);
    }
  });
  return dedupArr.map(function(r) {
    return { 'GROUP NAME': pickMillGroupName_(r), 'COMPANY NAME': pickMillCompanyName_(r) };
  });
}

const oldOut = oldBuild(allData);
assert.equal(oldOut.length, 2, 'old logic duplicated mill');
assert.ok(oldOut.some(function(r) { return !r['GROUP NAME']; }), 'old logic had empty group row');

console.log('test-eudr-group-backfill: OK');
