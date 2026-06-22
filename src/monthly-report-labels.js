/** Column labels & formatters aligned with dashboard web panels. */

export function normalizeSddCategory(v) {
  const s = String(v || '').trim().toUpperCase();
  if (s === 'MILL' || s === 'MIL') return 'MILL';
  if (s === 'KCP') return 'KCP';
  if (s === 'TRADER' || s === 'TRD') return 'TRADER';
  return String(v || '').trim() || '—';
}

function normalizeSddDecision(raw) {
  if (!raw) return '';
  const s = String(raw).trim().toLowerCase();
  if (s === 'approve' || s === 'approved') return 'APPROVED';
  if (s === 'hold' || s === 'on hold') return 'ON HOLD';
  if (s === 'reject' || s === 'rejected') return 'REJECTED';
  return String(raw).trim().toUpperCase();
}

export function sddStatusText(r) {
  const scrStatus = String(r['SCR - Screening Status'] || '').trim().toLowerCase();
  if (scrStatus !== 'submitted') return 'Draft';
  const decRaw = String(
    r.statusSDD || r.statusSdd || r['Status SDD'] ||
    r.statusBossDecision || r['Status Boss Decision'] || ''
  ).trim();
  if (!decRaw) return 'Submitted';
  return 'Submitted · ' + normalizeSddDecision(decRaw);
}

export function sddCompanyName(r) {
  return String(
    r['Company Name'] || r['Group Name'] || r['Grup Name'] || r['Mill Name'] || '—'
  ).trim() || '—';
}

export function sddDateImport(r) {
  const raw = r['Date Imported'] || r['DATE IMPORTED'] || r.date_imported || '';
  return formatDashboardDate_(raw);
}

export function sddLastUpdate(r) {
  const raw = r.updated_at || r['SCR - Last Updated'] || '';
  return formatDashboardDate_(raw);
}

function formatDashboardDate_(raw) {
  const s = String(raw || '').trim();
  if (!s) return '—';
  const d = new Date(s);
  if (isNaN(d.getTime())) return s.slice(0, 10) || '—';
  return d.toLocaleDateString('id-ID', { day: '2-digit', month: 'short', year: 'numeric' });
}

export const MRD_SDD_TABLE_TITLE = 'Saved screening (Draft / Submitted)';

export const MRD_SDD_COLS = ['Date Import', 'Company Name', 'Category', 'Status', 'Last Update'];

export const MRD_MILL_SUMMARY_COLS = [
  'Result Risk Level', 'Group Name', 'Company Name', 'Mill Name', 'Province', 'No Buy List',
];

export const MRD_MILL_FULL_COLS = [
  'Result Risk Level', 'Group Name', 'Company Name', 'Mill Name', 'Province', 'No Buy List',
  'Supplier Status', 'Certification', 'Total Grievances', 'Facility Name CPO', 'Facility Name PK',
];

export const MRD_GRV_SUMMARY_COLS = [
  'Date Received', 'Category', 'Complainant', 'Group', 'Grievance Subject',
  'Risk Classification', 'Grievance Status',
];

export const MRD_GRV_DETAIL_COLS = [
  'Grievance ID', 'Date Received', 'Category', 'Complainant', 'Group', 'Grievance Subject',
  'Risk Classification', 'Grievance Status', 'Grievance Description', 'Verification Findings',
  'Corrective Action', 'Preventive Action',
];

export function grvGroupName_(r) {
  const row = (r && r.row) || r || {};
  return String(
    row['Grievance Subject Group'] || row['Group'] || row['GRIEVANCE SUBJECT GROUP'] || ''
  ).trim() || '—';
}

export const MRD_EUDR_COLS = [
  'Group Name', 'Company Name', 'Mill Name', 'Province', 'CPO / PK Supply',
];

/** CPO + PK facility supply in one cell (from EUDR / mill onboarding). */
export function eudrCombinedSupply_(row) {
  const r = row || {};
  const cpo = String(r['SUPPLY TO'] || r['SUPPLY TO CPO'] || '').trim();
  const pk = String(r['SUPPLY TO PK'] || '').trim();
  const parts = [];
  if (cpo) parts.push('CPO: ' + cpo);
  if (pk) parts.push('PK: ' + pk);
  return parts.length ? parts.join(' · ') : '—';
}

export const MRD_TRACE_DETAIL_COLS = [
  'Group Name', 'Company Name', 'Mill Name', 'Province',
  '% TTM CPO', '% TTM PK', '% TTP CPO', '% TTP PK', 'Supplier Data',
];

export const MRD_FACILITY_COMPANY_COLS = [
  'Group Name', 'Company Name', 'Certification', 'No Buy List', 'Result Risk Level',
  'Total Grievance', '% TTM TRACEABLE', '% TTP TRACEABLE',
];

export function facilityPctColLabel(isPk) {
  return isPk ? '% TTP PK TRACEABLE' : '% TTP TRACEABLE';
}

export function facilityTtmColLabel(isPk) {
  return isPk ? '% TTM PK TRACEABLE' : '% TTM TRACEABLE';
}

export function facilitySummaryColLabels(isPk) {
  const ttm = facilityTtmColLabel(isPk);
  const pct = facilityPctColLabel(isPk);
  return ['Facility', 'Companies', 'No Buy List', 'High Risk', 'Total Grievance', 'Estimated ISPO Supply %', ttm, pct];
}

/** Short / wrapped PDF header labels so columns do not overlap on A4. */
const PDF_HEAD_SHORT = {
  'Result Risk Level': 'Risk\nLevel',
  'Group Name': 'Group\nName',
  'Company Name': 'Company\nName',
  'Mill Name': 'Mill\nName',
  'No Buy List': 'No Buy\nList',
  'Supplier Status': 'Supplier\nStatus',
  'Total Grievances': 'Total\nGriev.',
  'Facility Name CPO': 'Facility\nCPO',
  'Facility Name PK': 'Facility\nPK',
  'Date Received': 'Date\nReceived',
  'Grievance Subject': 'Grievance\nSubject',
  'Risk Classification': 'Risk\nClass.',
  'Grievance Status': 'Grievance\nStatus',
  'Grievance Description': 'Description',
  'Verification Findings': 'Verification',
  'Corrective Action': 'Corrective',
  'Preventive Action': 'Preventive',
  'Estimated ISPO Supply %': 'Estimated\nISPO Supply %',
  '% TRACEABLE': '%\nTRACEABLE',
  '% PK TRACEABLE': '% PK\nTRACEABLE',
  '% TTP TRACEABLE': '% TTP\nTRACEABLE',
  '% TTP PK TRACEABLE': '% TTP\nPK',
  '% TTM TRACEABLE': '% TTM\nTRACEABLE',
  '% TTM PK TRACEABLE': '% TTM\nPK',
  '% TTM CPO': '% TTM\nCPO',
  '% TTM PK': '% TTM\nPK',
  '% TTP CPO': '% TTP\nCPO',
  '% TTP PK': '% TTP\nPK',
  'Supplier Data': 'Supplier\nData',
  'Company Group Name': 'Company\nGroup',
  'CPO / PK Supply': 'CPO /\nPK Supply',
  'NBL Riser': 'NBL\nRiser',
  'Last Update': 'Last\nUpdate',
  'Date Import': 'Date\nImport',
  'Company Name NBL': 'Company\nName',
};

export function pdfHeadRow(labels) {
  return labels.map(function(label) {
    if (PDF_HEAD_SHORT[label]) return PDF_HEAD_SHORT[label];
    const s = String(label || '');
    if (s.length <= 12) return s;
    const parts = s.split(' ');
    if (parts.length > 1) {
      const mid = Math.ceil(parts.length / 2);
      return parts.slice(0, mid).join(' ') + '\n' + parts.slice(mid).join(' ');
    }
    if (s.length > 14) return s.slice(0, 14) + '\n' + s.slice(14);
    return s;
  });
}

export function pdfTableHead(labels) {
  return [pdfHeadRow(labels)];
}

export function pdfCellTrim(val, maxLen) {
  const s = pdfSanitizeText_(val);
  const n = maxLen || 42;
  if (s.length <= n) return s;
  return s.slice(0, Math.max(0, n - 1)) + '…';
}

function pdfSanitizeText_(v) {
  return String(v == null ? '' : v).replace(/\s+/g, ' ').trim() || '—';
}

const MRD_MONTH_FULL_ = [
  '', 'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
];
export const MRD_MONTH_SHORT_ = [
  '', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
  'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec',
];

/** Full month name from 1–12. */
export function mrdMonthFullName_(month) {
  const m = parseInt(month, 10);
  return MRD_MONTH_FULL_[m] || '';
}

/** Business week of month (days 1–7 → Wk 1, 8–14 → Wk 2, 15–21 → Wk 3, …). */
export function mrdWeekOfMonth_(date) {
  const d = date instanceof Date ? date : new Date(date);
  if (isNaN(d.getTime())) return 1;
  return Math.min(5, Math.ceil(d.getDate() / 7));
}

/**
 * Official monthly reporting cutoff — always the 15th.
 * Month selected → 15th of that month; year only → 15 Dec; else 15th of export month.
 */
export function mrdReportCutoffDate_(year, month, fallbackDate) {
  const y = parseInt(year, 10);
  const m = parseInt(month, 10);
  if (y && m >= 1 && m <= 12) return new Date(y, m - 1, 15);
  if (y) return new Date(y, 11, 15);
  const fb = fallbackDate instanceof Date ? fallbackDate : new Date(fallbackDate || Date.now());
  return new Date(fb.getFullYear(), fb.getMonth(), 15);
}

/** Human-readable report period for headers. */
export function mrdPeriodDisplayLabel_(year, month) {
  const y = String(year || '').trim();
  const m = parseInt(month, 10);
  if (y && m >= 1 && m <= 12) return mrdMonthFullName_(m) + ' ' + y;
  if (y) return 'Full year ' + y;
  return 'All periods';
}

/**
 * Data lags one month behind the reporting period.
 * e.g. January 2026 report → December 2025 data; February 2026 → January 2026.
 */
export function mrdDataPeriodFromReport_(reportYear, reportMonth) {
  const ry = parseInt(reportYear, 10);
  const rm = parseInt(reportMonth, 10);
  if (!ry) return { year: '', month: '' };
  if (!rm || rm < 1 || rm > 12) {
    return { year: String(ry - 1), month: '' };
  }
  let dm = rm - 1;
  let dy = ry;
  if (dm < 1) {
    dm = 12;
    dy = ry - 1;
  }
  return { year: String(dy), month: String(dm) };
}

export function mrdFormatReportDate_(date) {
  const d = date instanceof Date ? date : new Date(date);
  if (isNaN(d.getTime())) return '—';
  return d.toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' });
}

/**
 * Monthly Report header copy — period, 15th cutoff, week-of-month (boss feedback).
 * @returns {{ period, periodLine, cutoffLine, cutoffCompact, weekOfMonth, cutoffFormatted, exportedAt }}
 */
export function mrdReportHeaderMeta_(year, month, exportDate) {
  const exportDt = exportDate instanceof Date ? exportDate : new Date(exportDate || Date.now());
  const cutoff = mrdReportCutoffDate_(year, month, exportDt);
  const week = mrdWeekOfMonth_(cutoff);
  const cutoffMonthFull = mrdMonthFullName_(cutoff.getMonth() + 1);
  const cutoffMonthShort = MRD_MONTH_SHORT_[cutoff.getMonth() + 1] || '';
  const period = mrdPeriodDisplayLabel_(year, month);
  const dataPeriod = mrdDataPeriodFromReport_(year, month);
  const dataLabel = mrdPeriodDisplayLabel_(dataPeriod.year, dataPeriod.month);
  const cutoffFormatted = mrdFormatReportDate_(cutoff);
  const cutoffCompact = '15 ' + cutoffMonthShort + ' ' + cutoff.getFullYear();
  return {
    period: period,
    dataPeriod: dataLabel,
    periodLine: 'Reporting period: ' + period,
    dataPeriodLine: 'Data period: ' + dataLabel,
    cutoffLine: 'Reporting cutoff: ' + cutoffFormatted + ' · Week ' + week + ' of ' + cutoffMonthFull,
    cutoffCompact: 'Cutoff ' + cutoffCompact + ' · Week ' + week,
    weekOfMonth: week,
    cutoffFormatted: cutoffFormatted,
    exportedAt: exportDt.toLocaleString('en-GB', { hour12: false }),
    dataYear: dataPeriod.year,
    dataMonth: dataPeriod.month,
  };
}

/** Case-insensitive A–Z compare for monthly report lists. */
export function mrdLocaleCompare_(a, b) {
  return String(a || '').localeCompare(String(b || ''), undefined, { sensitivity: 'base' });
}

export function mrdSortSddRows_(rows) {
  return rows.slice().sort(function(a, b) {
    return mrdLocaleCompare_(sddCompanyName(a), sddCompanyName(b));
  });
}

export function mrdSortMillItems_(items) {
  return items.slice().sort(function(a, b) {
    const ra = a.row || a;
    const rb = b.row || b;
    const g = mrdLocaleCompare_(ra['GROUP NAME'], rb['GROUP NAME']);
    if (g !== 0) return g;
    const c = mrdLocaleCompare_(ra['COMPANY NAME'], rb['COMPANY NAME']);
    if (c !== 0) return c;
    return mrdLocaleCompare_(ra['MILL NAME'], rb['MILL NAME']);
  });
}

/** Grievance stays newest-first by date received. */
export function mrdSortGrvItemsByDateDesc_(items) {
  return items.slice().sort(function(a, b) {
    const da = String((a.row || a)['Date Received'] || '');
    const db = String((b.row || b)['Date Received'] || '');
    return db.localeCompare(da);
  });
}

export function mrdSortEudrItems_(items) {
  return items.slice().sort(function(a, b) {
    const ra = a.row || a;
    const rb = b.row || b;
    const g = mrdLocaleCompare_(ra['GROUP NAME'], rb['GROUP NAME']);
    if (g !== 0) return g;
    const c = mrdLocaleCompare_(ra['COMPANY NAME'], rb['COMPANY NAME']);
    if (c !== 0) return c;
    return mrdLocaleCompare_(ra['MILL NAME'], rb['MILL NAME']);
  });
}


export function mrdSortFacilityCompanies_(companies) {
  return companies || [];
}

export function mrdSortFacilityBundles_(bundles) {
  return bundles || [];
}

export function mrdSortBundlesByFacility_(bundles) {
  return bundles.slice().sort(function(a, b) {
    return mrdLocaleCompare_(a.facility, b.facility);
  });
}

export function mrdFormatNblRisers_(item) {
  if (!item) return '—';

  // Split a riser cell value that may pack multiple names separated by commas
  // or slashes (e.g. "CARGILL, APICAL" or "ADM/Wilmar") and return individual parts.
  function splitRiserCell_(v) {
    return String(v || '').split(/[,\/]/).map(function(s) { return s.trim(); }).filter(Boolean);
  }

  // Stable dedup key: normalize dashes (incl. en/em-dash) + surrounding spaces,
  // then lowercase. "GAR - Sinarmas", "GAR-SINARMAS", "GAR – Sinarmas" → "gar-sinarmas".
  function riserKey_(v) {
    return String(v || '').trim()
      .replace(/\s*[-\u2013\u2014]\s*/g, '-')
      .replace(/\s+/g, ' ')
      .toLowerCase();
  }

  function dedupeRiserList_(parts) {
    const out = [];
    const seen = {};
    (parts || []).forEach(function(r) {
      const s = String(r || '').trim();
      if (!s) return;
      const k = riserKey_(s);
      if (seen[k]) return;
      seen[k] = true;
      out.push(s);
    });
    return out.sort(function(a, b) {
      return a.localeCompare(b, undefined, { sensitivity: 'base' });
    });
  }

  const matches = item.nblMatches || [];
  if (matches.length) {
    // Flatten: each m.riser may itself contain "A, B" — split before dedup.
    const allParts = [];
    matches.forEach(function(m) {
      if (m && m.riser) splitRiserCell_(m.riser).forEach(function(p) { allParts.push(p); });
    });
    const risers = dedupeRiserList_(allParts);
    if (risers.length) return risers.join(', ');
  }

  const by = String(item.nblBy || '').trim();
  if (!by) return '—';
  if (/source unresolved/i.test(by)) return '—';
  const stripped = by.replace(/^NBL by\s+/i, '').trim();
  if (!stripped) return '—';
  // nblBy is already a comma string — split then dedup.
  return dedupeRiserList_(stripped.split(',').map(function(s) { return s.trim(); })).join(', ') || '—';
}

/** Normalized group|company|mill key for TTP ↔ mill onboarding joins. */
export function mrdNormalizeMillKey_(group, company, mill) {
  return [group, company, mill].map(function(x) {
    return String(x || '').trim().toLowerCase();
  }).join('|');
}

export function mrdResolveTtpSupplierCol_(ttpFields) {
  const fields = ttpFields || [];
  if (!fields.length) return 'FFB SUPPLIER NAME';
  const exact = fields.find(function(h) {
    const u = String(h || '').toUpperCase().replace(/\s+/g, ' ');
    return u === 'FFB SUPPLIER NAME';
  });
  if (exact) return exact;
  const loose = fields.find(function(h) {
    const u = String(h || '').toUpperCase();
    return /FFB/.test(u) && /SUPPLIER/.test(u) && /NAME/.test(u) && !/GROUP/.test(u);
  });
  return loose || 'FFB SUPPLIER NAME';
}

export function mrdTtpRowHasSupplier_(row, supplierCol) {
  if (!row || typeof row !== 'object') return false;
  const col = supplierCol || 'FFB SUPPLIER NAME';
  const direct = String(row[col] || '').trim();
  if (direct && direct !== '—' && direct !== '-') return true;
  for (const k of Object.keys(row)) {
    if (k === '_row' || k === '_sddSearchBlob' || (String(k).length && String(k)[0] === '_')) continue;
    const u = String(k).toUpperCase().replace(/\s+/g, ' ');
    if (u === 'FFB SUPPLIER NAME' || (u.includes('FFB') && u.includes('SUPPLIER') && u.includes('NAME') && !u.includes('GROUP'))) {
      const v = String(row[k] || '').trim();
      if (v && v !== '—' && v !== '-') return true;
    }
  }
  return false;
}

export function mrdBuildTtpByMillMaps_(ttpRows, millCol, groupCol, companyCol) {
  const byComposite = new Map();
  const byMillName = new Map();
  (ttpRows || []).forEach(function(r) {
    const mill = String(r[millCol] || r['MILL NAME'] || '').trim();
    const composite = mrdNormalizeMillKey_(
      r[groupCol] || r['GROUP NAME'],
      r[companyCol] || r['COMPANY NAME'],
      mill
    );
    if (!byComposite.has(composite)) byComposite.set(composite, []);
    byComposite.get(composite).push(r);
    if (mill) {
      if (!byMillName.has(mill)) byMillName.set(mill, []);
      byMillName.get(mill).push(r);
    }
  });
  return { byComposite: byComposite, byMillName: byMillName };
}

export function mrdTtpRowsForMill_(maps, millRow) {
  if (!maps || !millRow) return [];
  const composite = mrdNormalizeMillKey_(millRow['GROUP NAME'], millRow['COMPANY NAME'], millRow['MILL NAME']);
  const fromComposite = maps.byComposite.get(composite);
  if (fromComposite && fromComposite.length) return fromComposite;
  const millName = String(millRow['MILL NAME'] || '').trim();
  return (millName && maps.byMillName.get(millName)) || [];
}

export function mrdSortEmptyMillItems_(items) {
  return items.slice().sort(function(a, b) {
    const ra = a.millRow || a.row || a;
    const rb = b.millRow || b.row || b;
    const g = mrdLocaleCompare_(ra['GROUP NAME'], rb['GROUP NAME']);
    if (g !== 0) return g;
    const c = mrdLocaleCompare_(ra['COMPANY NAME'], rb['COMPANY NAME']);
    if (c !== 0) return c;
    return mrdLocaleCompare_(ra['MILL NAME'], rb['MILL NAME']);
  });
}
