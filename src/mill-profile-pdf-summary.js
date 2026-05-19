/**
 * Mill profile PDF — Summary section layouts.
 * Set MILL_PROFILE_PDF_SUMMARY_VARIANT to 'backup' for the alternate design.
 */
export const MILL_PROFILE_PDF_SUMMARY_VARIANT = 'primary'; // 'primary' | 'backup'

const BRAND = [139, 26, 26];
const BRAND_SOFT = [252, 246, 246];
const INK = [26, 26, 26];
const MUTED = [106, 96, 96];

function norm_(v) {
  return String(v == null ? '' : v).trim();
}

function titleCase_(s) {
  return norm_(s)
    .toLowerCase()
    .replace(/\b\w/g, function(c) { return c.toUpperCase(); });
}

function display_(v) {
  const s = norm_(v);
  return s || '—';
}

function formatNbl_(raw) {
  const v = norm_(raw).toLowerCase();
  if (!v) return '';
  if (v === 'yes' || v.includes('nbl') || v.includes('no buy')) return 'Yes';
  if (v === 'no') return 'No';
  return titleCase_(raw);
}

function nblStyle_(raw) {
  const v = norm_(raw).toLowerCase();
  const isNbl = v === 'yes' || v.includes('nbl') || v.includes('no buy');
  if (!v || v === '—') return { fill: [248, 246, 246], text: MUTED, bold: false };
  if (isNbl) return { fill: [255, 235, 232], text: [183, 28, 28], bold: true };
  return { fill: [232, 248, 238], text: [21, 128, 61], bold: true };
}

function riskStyle_(raw) {
  const v = norm_(raw).toLowerCase();
  if (!v || v === '—') return { fill: [248, 246, 246], text: MUTED, bold: false };
  if (v.includes('high') || v.includes('tinggi')) return { fill: [255, 221, 218], text: [164, 22, 18], bold: true };
  if (v.includes('med') || v.includes('sedang')) return { fill: [255, 243, 205], text: [146, 112, 0], bold: true };
  if (v.includes('low') || v.includes('rendah')) return { fill: [220, 245, 232], text: [21, 128, 61], bold: true };
  return { fill: [242, 238, 238], text: [74, 63, 63], bold: true };
}

function statusStyle_(raw) {
  const v = norm_(raw).toLowerCase();
  if (!v || v === '—') return { fill: [248, 246, 246], text: MUTED, bold: false };
  if (['active', 'compliant'].some(function(k) { return v.includes(k); })) {
    return { fill: [232, 248, 238], text: [21, 128, 61], bold: true };
  }
  if (['review', 'pending', 'conditional'].some(function(k) { return v.includes(k); })) {
    return { fill: [255, 243, 205], text: [146, 112, 0], bold: true };
  }
  return { fill: BRAND_SOFT, text: BRAND, bold: true };
}

function summaryValues_(row, getValue) {
  const statusRaw = getValue('SUPPLIER STATUS');
  const nblRaw = getValue('BUYER NO BUY LIST');
  const riskRaw = getValue('RESULT RISK LEVEL');
  return {
    status: { label: 'Status', text: display_(titleCase_(statusRaw)), style: statusStyle_(statusRaw) },
    nbl: { label: 'No Buy List', text: display_(formatNbl_(nblRaw)), style: nblStyle_(nblRaw) },
    risk: { label: 'Result Risk Level', text: display_(norm_(riskRaw).toUpperCase()), style: riskStyle_(riskRaw) },
  };
}

/** Primary: 3-column KPI strip — executive summary at a glance. */
export function renderMillProfileSummaryPdfPrimary(doc, row, startY, getValue) {
  const v = summaryValues_(row, getValue);
  const bodyRow = [v.status.text, v.nbl.text, v.risk.text];
  const styles = [v.status.style, v.nbl.style, v.risk.style];

  doc.autoTable({
    head: [
      [{ content: 'Summary', colSpan: 3, styles: { fillColor: BRAND, textColor: [255, 255, 255], fontStyle: 'bold', fontSize: 10, cellPadding: { top: 3.5, bottom: 3.5, left: 4, right: 4 } } }],
      [
        { content: v.status.label, styles: { fillColor: [245, 236, 236], textColor: BRAND, fontStyle: 'bold', fontSize: 8, halign: 'center' } },
        { content: v.nbl.label, styles: { fillColor: [245, 236, 236], textColor: BRAND, fontStyle: 'bold', fontSize: 8, halign: 'center' } },
        { content: v.risk.label, styles: { fillColor: [245, 236, 236], textColor: BRAND, fontStyle: 'bold', fontSize: 8, halign: 'center' } },
      ],
    ],
    body: [bodyRow],
    startY: startY,
    margin: { left: 14, right: 14 },
    theme: 'plain',
    styles: {
      fontSize: 11,
      cellPadding: { top: 5, bottom: 5, left: 4, right: 4 },
      lineColor: [225, 210, 210],
      lineWidth: 0.25,
      halign: 'center',
      valign: 'middle',
      minCellHeight: 16,
    },
    columnStyles: {
      0: { cellWidth: 58 },
      1: { cellWidth: 58 },
      2: { cellWidth: 'auto' },
    },
    didParseCell: function(data) {
      if (data.section !== 'body' || data.row.index !== 0) return;
      const st = styles[data.column.index];
      if (!st) return;
      data.cell.styles.fillColor = st.fill;
      data.cell.styles.textColor = st.text;
      data.cell.styles.fontStyle = st.bold ? 'bold' : 'normal';
      data.cell.styles.fontSize = 11.5;
    },
    didDrawCell: function(data) {
      if (data.section !== 'body' || data.row.index !== 0) return;
      doc.setDrawColor.apply(doc, BRAND);
      doc.setLineWidth(0.6);
      doc.line(data.cell.x, data.cell.y, data.cell.x + data.cell.width, data.cell.y);
    },
  });

  return doc.lastAutoTable && doc.lastAutoTable.finalY != null ? doc.lastAutoTable.finalY : startY + 28;
}

/** Backup: 2-column table with left accent bar and tinted value cells. */
export function renderMillProfileSummaryPdfBackup(doc, row, startY, getValue) {
  const v = summaryValues_(row, getValue);
  const rows = [
    [v.status.label, v.status.text, v.status.style],
    [v.nbl.label, v.nbl.text, v.nbl.style],
    [v.risk.label, v.risk.text, v.risk.style],
  ];

  doc.autoTable({
    head: [['Summary', 'Value']],
    body: rows.map(function(r) { return [r[0], r[1]]; }),
    startY: startY,
    margin: { left: 14, right: 14 },
    theme: 'grid',
    styles: { fontSize: 9.5, cellPadding: 3.2, lineColor: [225, 210, 210], lineWidth: 0.2 },
    headStyles: { fillColor: BRAND, textColor: [255, 255, 255], fontStyle: 'bold' },
    columnStyles: { 0: { cellWidth: 58, fontStyle: 'bold', textColor: INK }, 1: { cellWidth: 'auto' } },
    didParseCell: function(data) {
      if (data.section !== 'body') return;
      const st = rows[data.row.index][2];
      if (data.column.index === 0) {
        data.cell.styles.fillColor = BRAND_SOFT;
        data.cell.styles.fontStyle = 'bold';
        data.cell.styles.textColor = BRAND;
      } else if (st) {
        data.cell.styles.fillColor = st.fill;
        data.cell.styles.textColor = st.text;
        data.cell.styles.fontStyle = st.bold ? 'bold' : 'normal';
        data.cell.styles.fontSize = 10.5;
        data.cell.styles.halign = 'center';
      }
    },
    didDrawCell: function(data) {
      if (data.section !== 'body' || data.column.index !== 0) return;
      doc.setFillColor.apply(doc, BRAND);
      doc.rect(data.cell.x, data.cell.y, 1.4, data.cell.height, 'F');
    },
  });

  return doc.lastAutoTable && doc.lastAutoTable.finalY != null ? doc.lastAutoTable.finalY : startY + 32;
}

export function renderMillProfileSummaryPdf(doc, row, startY, getValue) {
  const variant = (typeof window !== 'undefined' && window.MILL_PROFILE_PDF_SUMMARY_VARIANT) ||
    MILL_PROFILE_PDF_SUMMARY_VARIANT;
  if (variant === 'backup') {
    return renderMillProfileSummaryPdfBackup(doc, row, startY, getValue);
  }
  return renderMillProfileSummaryPdfPrimary(doc, row, startY, getValue);
}
