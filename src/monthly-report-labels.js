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

export const MRD_EUDR_COLS = [
  'Group Name', 'Company Name', 'Mill Name', 'Province', 'Supply To', 'Status',
];

export const MRD_FACILITY_COMPANY_COLS = [
  'Group Name', 'Company Name', 'Certification', 'No Buy List', 'Result Risk Level',
  'Total Grievance', '% TRACEABLE',
];

export function facilityPctColLabel(isPk) {
  return isPk ? '% PK TRACEABLE' : '% TRACEABLE';
}

export function facilitySummaryColLabels(isPk) {
  const pct = facilityPctColLabel(isPk);
  return ['Facility', 'Companies', 'No Buy List', 'High Risk', 'Total Grievance', 'Est. ISPO Supply %', pct];
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
  'Est. ISPO Supply %': 'ISPO\nSupply %',
  '% TRACEABLE': '%\nTRACEABLE',
  '% PK TRACEABLE': '% PK\nTRACEABLE',
  'Company Group Name': 'Company\nGroup',
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
