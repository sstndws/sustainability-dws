/**
 * No Buy List — PDF & Excel export helpers.
 */
import { buildBrandedExcelSheet_ } from './excel-brand-header.js';

export const NBL_REGISTRY_FIELDS = ['Riser', 'Group Name NBL', 'Company Name NBL', 'SOURCE'];
export const UNILEVER_NBL_FIELDS = [
  'Riser', 'UML ID', 'COMPANY NAME', 'MILL NAME', 'COUNTRY', 'PROVINCE',
  'DISTRICT / REGENCY', 'LAT.', 'LONG.',
];

function nblStamp_() {
  return new Date().toISOString().slice(0, 10);
}

export function nblExportTitle_(source) {
  return source === 'unilever' ? 'Unilever No Buy List' : 'No Buy List — Registry';
}

export function nblExportFilename_(source, ext) {
  const base = source === 'unilever' ? 'No Buy List - Unilever NBL' : 'No Buy List - NBL Registry';
  return base + ' - ' + nblStamp_() + '.' + ext;
}

export function nblFieldValue_(row, field, source) {
  if (source === 'unilever') {
    if (field === 'Riser') return String(row._nblRiser || row._nblNo || '');
    if (field === 'UML ID') return String(row._nblUml || '');
    if (field === 'COMPANY NAME') return String(row._nblCompany || '');
    if (field === 'MILL NAME') return String(row._nblMill || '');
    if (field === 'COUNTRY') return String(row._nblCountry || '');
    if (field === 'PROVINCE') return String(row._nblProvince || '');
    if (field === 'DISTRICT / REGENCY') return String(row._nblDistrict || '');
    if (field === 'LAT.') return String(row._nblLat || '');
    if (field === 'LONG.') return String(row._nblLong || '');
    return String(row[field] || '');
  }
  if (field === 'Riser') return String(row._nblRiser || '');
  if (field === 'Group Name NBL') return String(row._nblGroup || '');
  if (field === 'Company Name NBL') return String(row._nblCompany || '');
  if (field === 'SOURCE') return String(row._nblSource || '');
  return String(row[field] || '');
}

export function nblFilterExportRows_(allRows, searchQuery) {
  const q = String(searchQuery || '').trim().toLowerCase();
  return (allRows || []).filter(function(d) {
    return !q || (d._nblSearchBlob || '').includes(q);
  });
}

export function nblExportToExcel_(rows, source, fields) {
  if (typeof XLSX === 'undefined') {
    throw new Error('Excel library not loaded. Refresh the page and try again.');
  }
  const headers = fields || (source === 'unilever' ? UNILEVER_NBL_FIELDS : NBL_REGISTRY_FIELDS);
  const body = (rows || []).map(function(row) {
    return headers.map(function(f) { return nblFieldValue_(row, f, source); });
  });
  const ws = buildBrandedExcelSheet_(XLSX, headers, body, {
    headerFill: '8B1A1A',
    includeCompanyInfo: false,
  });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, source === 'unilever' ? 'Unilever NBL' : 'NBL Registry');
  XLSX.writeFile(wb, nblExportFilename_(source, 'xlsx'));
}

export async function nblExportToPdf_(rows, source, fields, getJsPDF) {
  const JsPDFLib = typeof getJsPDF === 'function' ? getJsPDF() : null;
  if (!JsPDFLib) throw new Error('PDF library not loaded. Refresh the page and try again.');

  const headers = fields || (source === 'unilever' ? UNILEVER_NBL_FIELDS : NBL_REGISTRY_FIELDS);
  const body = (rows || []).map(function(row) {
    return headers.map(function(f) {
      return String(nblFieldValue_(row, f, source) || '').replace(/\r?\n/g, ' ');
    });
  });

  const landscape = source === 'unilever';
  const doc = new JsPDFLib({
    unit: 'mm',
    format: 'a4',
    orientation: landscape ? 'landscape' : 'portrait',
  });
  const RED = [139, 26, 26];
  const WHITE = [255, 255, 255];
  const GRY = [110, 96, 96];

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(13);
  doc.setTextColor.apply(doc, RED);
  doc.text(nblExportTitle_(source), 14, 14);
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(9);
  doc.setTextColor.apply(doc, GRY);
  const stamp = new Date().toLocaleString('en-GB', {
    day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit',
  });
  doc.text('Exported: ' + stamp + ' · ' + body.length + ' row(s)', 14, 20);

  doc.autoTable({
    head: [headers],
    body: body,
    startY: 26,
    margin: { left: 12, right: 12 },
    styles: { fontSize: landscape ? 6.5 : 8, cellPadding: 2, textColor: [26, 10, 10], overflow: 'linebreak' },
    headStyles: { fillColor: RED, textColor: WHITE, fontStyle: 'bold', fontSize: landscape ? 7 : 8.5 },
    alternateRowStyles: { fillColor: [253, 250, 250] },
    theme: 'striped',
  });

  doc.save(nblExportFilename_(source, 'pdf'));
}
