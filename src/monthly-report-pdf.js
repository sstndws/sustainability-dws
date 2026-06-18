/**
 * Monthly Report (Detail) — PDF export without TOC; continuous layout, no blank pages.
 */

import {
  MRD_SDD_TABLE_TITLE,
  MRD_SDD_COLS,
  MRD_MILL_SUMMARY_COLS,
  MRD_MILL_FULL_COLS,
  MRD_GRV_SUMMARY_COLS,
  MRD_GRV_DETAIL_COLS,
  MRD_EUDR_COLS,
  MRD_FACILITY_COMPANY_COLS,
  facilityPctColLabel,
  facilitySummaryColLabels,
  normalizeSddCategory,
  sddStatusText,
  sddCompanyName,
  sddDateImport,
  sddLastUpdate,
} from './monthly-report-labels.js';

const BRAND = [139, 26, 26];
const PK_GREEN = [13, 110, 70];
const TRACE_ORANGE = [230, 81, 0];
const TTM_CPO = [139, 26, 26];
const TTM_PK = [46, 125, 50];
const TTP_CPO = [230, 81, 0];
const TTP_PK = [21, 101, 192];
const GRV_PURPLE = [106, 27, 154];
const NBL_RED = [192, 57, 43];
const EUDR_TEAL = [0, 131, 143];
const INK = [26, 10, 10];
const INK_MUTED = [110, 96, 96];
const INK_LIGHT = [140, 120, 120];
const WHITE = [255, 255, 255];
const BG_SOFT = [252, 250, 250];
const BORDER = [230, 220, 220];

const DEFAULT_SECTIONS = ['kpi', 'sdd', 'mill', 'trace', 'grv', 'nbl', 'facility', 'eudr'];
const MIN_CONTENT_Y = 24;

function pdfSanitize(v) {
  const s = String(v == null ? '' : v).replace(/\s+/g, ' ').trim();
  return s || '—';
}

function monthLabel(month) {
  const names = ['', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  return month ? names[parseInt(month, 10)] || month : 'All months';
}

function periodLabel(year, month) {
  return (year || 'All years') + ' · ' + monthLabel(month);
}

function isNblYes_(val) {
  return /yes|nbl|no buy/i.test(String(val || ''));
}

function pctFmt_(n) {
  if (isNaN(n)) return '—';
  return (Math.round(n * 10) / 10) + '%';
}

function colWidths_(ratios, totalW) {
  const sum = ratios.reduce(function(a, b) { return a + b; }, 0);
  const widths = ratios.map(function(r) { return (totalW * r) / sum; });
  const used = widths.slice(0, -1).reduce(function(a, b) { return a + b; }, 0);
  widths[widths.length - 1] = totalW - used;
  return widths;
}

function drawFooters_(doc, pageW, pageH, mL, mR, mFoot) {
  const total = doc.internal.getNumberOfPages();
  for (let p = 1; p <= total; p++) {
    doc.setPage(p);
    doc.setDrawColor.apply(doc, BORDER);
    doc.setLineWidth(0.15);
    doc.line(mL, pageH - mFoot, pageW - mR, pageH - mFoot);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(6.5);
    doc.setTextColor.apply(doc, INK_LIGHT);
    doc.text('Sustainability Dashboard · Monthly Report', mL, pageH - 5);
    doc.text('Page ' + p + ' of ' + total, pageW - mR, pageH - 5, { align: 'right' });
  }
}

function createPdfContext_(jsPDFLib, opts) {
  opts = opts || {};
  const doc = new jsPDFLib({ orientation: opts.orientation || 'portrait', unit: 'mm', format: 'a4' });
  const pageW = doc.internal.pageSize.getWidth();
  const pageH = doc.internal.pageSize.getHeight();
  const mL = 14;
  const mR = 14;
  const mFoot = 14;
  const cW = pageW - mL - mR;
  let y = 0;
  let sectionStarted = false;
  const pageBottomY = {};
  const pageTopY = {};

  const ctx = {
    doc: doc,
    pageW: pageW,
    pageH: pageH,
    mL: mL,
    mR: mR,
    mFoot: mFoot,
    cW: cW,
    periodText: '',
    generatedAt: '',
    reportTitle: 'Monthly Report',
    reportSubtitle: 'Compliance snapshot — SDD, Mill, Traceability, Grievance, NBL, Facility, EUDR',
    detailLevel: 'summary',
  };

  function markContent_(yPos) {
    if (yPos < MIN_CONTENT_Y) return;
    const p = doc.internal.getNumberOfPages();
    pageTopY[p] = Math.min(pageTopY[p] || yPos, yPos);
    pageBottomY[p] = Math.max(pageBottomY[p] || 0, yPos);
  }

  function markPageSpan_(pageNum, topY, bottomY) {
    if (topY < MIN_CONTENT_Y && bottomY < MIN_CONTENT_Y) return;
    pageTopY[pageNum] = Math.min(pageTopY[pageNum] || topY, topY);
    pageBottomY[pageNum] = Math.max(pageBottomY[pageNum] || 0, bottomY);
  }

  function drawCompactHeader_() {
    doc.setFillColor.apply(doc, BRAND);
    doc.rect(0, 0, pageW, 15, 'F');
    doc.setTextColor.apply(doc, WHITE);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(9.5);
    doc.text('Monthly Report', mL, 6.5);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(7);
    doc.text(ctx.periodText + '   ·   ' + ctx.generatedAt, mL, 11.5);
    return 18;
  }

  function ensureSpace(needed) {
    if (y + needed <= pageH - mFoot - 6) return;
    doc.addPage();
    y = drawCompactHeader_();
  }

  function remainingSpace_() {
    return pageH - mFoot - 6 - y;
  }

  function drawSectionBar_(title, accent, fontSize) {
    ensureSpace(14);
    doc.setFillColor.apply(doc, accent || BRAND);
    doc.roundedRect(mL, y, cW, 8, 1.2, 1.2, 'F');
    doc.setTextColor.apply(doc, WHITE);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(fontSize || 9);
    doc.text(title, mL + 3, y + 5.5);
    y += 12;
    markContent_(y);
  }

  function beginSection_(title, accent) {
    if (sectionStarted && remainingSpace_() < 30) {
      doc.addPage();
      y = drawCompactHeader_();
    } else if (sectionStarted) {
      y += 3;
    }
    sectionStarted = true;
    drawSectionBar_(title, accent);
  }

  function beginSubsection_(title, accent) {
    if (remainingSpace_() < 22) {
      doc.addPage();
      y = drawCompactHeader_();
    } else {
      y += 4;
    }
    drawSectionBar_(title, accent, 8.5);
  }

  function beginFacilityBlock_(title, accent, newPage) {
    if (newPage) {
      doc.addPage();
      y = drawCompactHeader_();
    } else if (sectionStarted && remainingSpace_() < 50) {
      doc.addPage();
      y = drawCompactHeader_();
    } else if (sectionStarted) {
      y += 6;
    }
    sectionStarted = true;
    drawSectionBar_(title, accent);
  }

  function drawAutoTable_(head, body, accent, colStyles, opts) {
    opts = opts || {};
    if (!body.length) return;

    const minStartH = 20;
    const headH = 9;
    const minRowH = 7;
    if (y + headH + minRowH > pageH - mFoot - 8) {
      doc.addPage();
      y = drawCompactHeader_();
    }

    const startPage = doc.internal.getNumberOfPages();
    const tableTop = y;
    const base = {
      theme: 'grid',
      margin: { left: mL, right: mR, bottom: mFoot + 12, top: 18 },
      tableWidth: cW,
      styles: {
        fontSize: opts.fontSize || 7.5,
        cellPadding: opts.cellPadding || 2.5,
        textColor: INK,
        lineColor: BORDER,
        lineWidth: 0.15,
        overflow: 'linebreak',
        valign: 'middle',
        minCellHeight: 5,
      },
      headStyles: {
        fillColor: accent || BRAND,
        textColor: WHITE,
        fontStyle: 'bold',
        fontSize: opts.fontSize || 7.5,
        halign: 'center',
      },
      alternateRowStyles: { fillColor: [255, 253, 253] },
      rowPageBreak: 'avoid',
      showHead: 'firstPage',
    };
    doc.autoTable(Object.assign({}, base, {
      head: head,
      body: body,
      startY: tableTop,
      columnStyles: colStyles || {},
      willDrawPage: function(data) {
        if (data.pageNumber > startPage) {
          doc.setPage(data.pageNumber);
          drawCompactHeader_();
        }
      },
      didDrawPage: function(data) {
        const p = data.pageNumber;
        const cursor = data.cursor || {};
        if (cursor.y != null) {
          const top = p === startPage ? tableTop : 18;
          markPageSpan_(p, top, cursor.y + 6);
        }
      },
    }));

    const table = doc.lastAutoTable;
    if (table && table.finalY) {
      const endPage = doc.internal.getNumberOfPages();
      markPageSpan_(endPage, endPage === startPage ? tableTop : 18, table.finalY);
      y = table.finalY + (opts.gapAfter == null ? 4 : opts.gapAfter);
      markContent_(y);
    }
  }

  function drawKvBlock_(heading, pairs, accent) {
    if (!pairs.length) return;
    if (remainingSpace_() < 18) {
      doc.addPage();
      y = drawCompactHeader_();
    }
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(7.5);
    doc.setTextColor.apply(doc, accent || BRAND);
    doc.text(heading, mL, y + 2);
    y += 5;
    markContent_(y);
    drawAutoTable_(
      [['Field', 'Value']],
      pairs.map(function(p) { return [p[0], pdfSanitize(p[1])]; }),
      accent,
      {
        0: { cellWidth: 40, fontStyle: 'bold', fillColor: BG_SOFT, textColor: INK_MUTED, fontSize: 7 },
        1: { cellWidth: cW - 40, fontSize: 7 },
      },
      { fontSize: 7, cellPadding: 2, gapAfter: 4 }
    );
  }

  function drawMainHeader_() {
    doc.setFillColor.apply(doc, BRAND);
    doc.rect(0, 0, pageW, 30, 'F');
    doc.setFillColor(180, 40, 40);
    doc.rect(0, 28.5, pageW, 1.5, 'F');
    doc.setTextColor.apply(doc, WHITE);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(17);
    doc.text(ctx.reportTitle || 'Monthly Report', mL, 13);
    const sub = String(ctx.reportSubtitle || '').trim();
    doc.setFont('helvetica', 'normal');
    if (sub) {
      doc.setFontSize(8.5);
      doc.text(sub, mL, 20);
      doc.setFontSize(8);
      doc.text(ctx.periodText + '   ·   Generated: ' + ctx.generatedAt, mL, 26);
      markContent_(36);
      return 36;
    }
    doc.setFontSize(8);
    doc.text(ctx.periodText + '   ·   Generated: ' + ctx.generatedAt, mL, 20);
    markContent_(28);
    return 28;
  }

  function startBodyAfterCover_(coverEndY) {
    if (coverEndY + 28 <= pageH - mFoot - 6) {
      y = coverEndY + 4;
    } else {
      doc.addPage();
      y = drawCompactHeader_();
    }
    sectionStarted = false;
  }

  function pruneBlankPages_() {
    let total = doc.internal.getNumberOfPages();
    for (let p = total; p >= 2; p--) {
      const bottom = pageBottomY[p] || 0;
      const top = pageTopY[p] || bottom;
      const span = bottom - top;
      const isEmpty = bottom < MIN_CONTENT_Y;
      const isOrphanStrip = span > 0 && span < 14;
      if (isEmpty || isOrphanStrip) {
        doc.deletePage(p);
      }
    }
  }

  ctx.beginSection_ = beginSection_;
  ctx.beginSubsection_ = beginSubsection_;
  ctx.beginFacilityBlock_ = beginFacilityBlock_;
  ctx.drawAutoTable_ = drawAutoTable_;
  ctx.drawKvBlock_ = drawKvBlock_;
  ctx.drawMainHeader_ = drawMainHeader_;
  ctx.startBodyAfterCover_ = startBodyAfterCover_;
  ctx.pruneBlankPages_ = pruneBlankPages_;
  ctx.markContent_ = markContent_;
  ctx.setY = function(v) { y = v; markContent_(v); };
  ctx.getY = function() { return y; };
  ctx.ensureSpace_ = function(needed) { ensureSpace(needed); };

  return ctx;
}

function kpiCardItems_(stats) {
  const s = stats || {};
  return [
    { label: 'SDD Requested', value: String(s.sddRequested != null ? s.sddRequested : (s.sddTotal || 0)) },
    { label: 'SDD Done', value: String(s.sddDone != null ? s.sddDone : (s.sddSubmitted || 0)) },
    { label: 'Total Mills', value: String(s.totalMills != null ? s.totalMills : 0) },
    { label: 'Groups', value: String(s.totalGroups != null ? s.totalGroups : 0) },
    { label: 'Untraceable Mills', value: String(s.emptyTraceMills != null ? s.emptyTraceMills : 0) },
    { label: 'EUDR Potential', value: String(s.eudrPotential != null ? s.eudrPotential : 0) },
  ];
}

const MRD_SECTION_LABELS = {
  kpi: 'Overview · Key metrics',
  sdd: 'Supplier Due Diligence',
  mill: 'Mill Onboarding',
  trace: 'Traceability Data',
  grv: 'Grievance Monitoring',
  nbl: 'Active NBL Mills',
  facility: 'Facility Performance',
  eudr: 'EUDR Potential',
};

/** Cover block on page 1 — returns Y position after last drawn element. */
function drawCoverPage_(ctx, stats, sections, full) {
  const doc = ctx.doc;
  doc.setPage(1);
  const mL = ctx.mL;
  const cW = ctx.cW;
  let y = 40;

  if (sections.indexOf('kpi') !== -1) {
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(12);
    doc.setTextColor.apply(doc, INK);
    doc.text('At a glance', mL, y);
    y += 10;

    const items = kpiCardItems_(stats);
    const cols = 2;
    const gap = 6;
    const cardW = (cW - gap) / cols;
    const cardH = 18;

    items.forEach(function(item, i) {
      const col = i % cols;
      const row = Math.floor(i / cols);
      const x = mL + col * (cardW + gap);
      const cy = y + row * (cardH + gap);

      doc.setDrawColor.apply(doc, BORDER);
      doc.setFillColor.apply(doc, WHITE);
      doc.setLineWidth(0.2);
      doc.roundedRect(x, cy, cardW, cardH, 2, 2, 'FD');
      doc.setFillColor.apply(doc, BRAND);
      doc.rect(x + 2, cy + 2, cardW - 4, 0.8, 'F');

      doc.setFont('helvetica', 'normal');
      doc.setFontSize(6.5);
      doc.setTextColor.apply(doc, INK_MUTED);
      doc.text(item.label.toUpperCase(), x + cardW / 2, cy + 7, { align: 'center' });

      doc.setFont('helvetica', 'bold');
      doc.setFontSize(12);
      doc.setTextColor.apply(doc, INK);
      doc.text(item.value, x + cardW / 2, cy + 14, { align: 'center' });
    });

    y += Math.ceil(items.length / cols) * (cardH + gap) + 8;
    ctx.markContent_(y);
  }

  doc.setFillColor.apply(doc, BG_SOFT);
  doc.setDrawColor.apply(doc, BORDER);
  doc.setLineWidth(0.2);
  doc.roundedRect(mL, y, cW, 26, 2, 2, 'FD');
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(8);
  doc.setTextColor.apply(doc, INK);
  doc.text('Export scope', mL + 4, y + 6);
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(7.5);
  doc.setTextColor.apply(doc, INK_MUTED);
  doc.text(
    'Detail level: ' + (full ? 'Full' : 'Summary'),
    mL + 4,
    y + 12
  );
  const sectionLabels = sections.map(function(id) { return MRD_SECTION_LABELS[id] || id; });
  const secLines = doc.splitTextToSize(sectionLabels.join(' · '), cW - 8);
  doc.text(secLines, mL + 4, y + 18);
  y += 30;
  ctx.markContent_(y);
  return y;
}

function drawSddSection_(ctx, rows, noHeader) {
  if (!noHeader) ctx.beginSection_('01 · Supplier Due Diligence', BRAND);
  ctx.beginSubsection_(MRD_SDD_TABLE_TITLE, BRAND);
  const w = colWidths_([14, 28, 12, 22, 14], ctx.cW);
  const body = (rows || []).length
    ? rows.map(function(r) {
      return [
        pdfSanitize(sddDateImport(r)),
        pdfSanitize(sddCompanyName(r)),
        pdfSanitize(normalizeSddCategory(r['Supplier Type'] || r['SUPPLIER_TYPE'] || r.supplier_type)),
        pdfSanitize(sddStatusText(r)),
        pdfSanitize(sddLastUpdate(r)),
      ];
    })
    : [['—', 'No SDD records for this period.', '—', '—', '—']];
  ctx.drawAutoTable_(
    [MRD_SDD_COLS],
    body,
    BRAND,
    {
      0: { cellWidth: w[0] }, 1: { cellWidth: w[1] }, 2: { cellWidth: w[2] },
      3: { cellWidth: w[3], fontSize: 6.5 }, 4: { cellWidth: w[4] },
    }
  );
}

function millDetailCells_(item) {
  const r = item.row;
  return {
    supplierStatus: pdfSanitize(r['SUPPLIER STATUS']),
    certification: pdfSanitize(r['CERTIFICATION']),
    grievances: pdfSanitize(r['TOTAL GRIEVANCES']),
    facilityCpo: pdfSanitize(r['FACILITY NAME CPO']),
    facilityPk: pdfSanitize(r['FACILITY NAME PK']),
    nblBy: pdfSanitize(item.nblBy),
  };
}

function drawMillSection_(ctx, data, full) {
  const mills = data.mills || [];
  if (!mills.length && !(data.highRiskMills || []).length) return;
  ctx.beginSection_('02 · Mill Onboarding', BRAND);

  if (mills.length) {
    if (full) {
      const w = colWidths_([11, 12, 14, 14, 10, 8, 11, 10, 8, 12, 12], ctx.cW);
      ctx.drawAutoTable_(
        [MRD_MILL_FULL_COLS],
        mills.map(function(item) {
          const r = item.row;
          const d = millDetailCells_(item);
          return [
            pdfSanitize(item.risk),
            pdfSanitize(r['GROUP NAME']),
            pdfSanitize(r['COMPANY NAME']),
            pdfSanitize(r['MILL NAME']),
            pdfSanitize(r['PROVINCE']),
            pdfSanitize(isNblYes_(item.nbl) ? 'Yes' : item.nbl),
            d.supplierStatus,
            d.certification,
            d.grievances,
            d.facilityCpo,
            d.facilityPk,
          ];
        }),
        BRAND,
        {
          0: { cellWidth: w[0] }, 1: { cellWidth: w[1] }, 2: { cellWidth: w[2] },
          3: { cellWidth: w[3] }, 4: { cellWidth: w[4] }, 5: { cellWidth: w[5] },
          6: { cellWidth: w[6], fontSize: 6.5 }, 7: { cellWidth: w[7], fontSize: 6.5 },
          8: { cellWidth: w[8] }, 9: { cellWidth: w[9], fontSize: 6.5 }, 10: { cellWidth: w[10], fontSize: 6.5 },
        },
        { fontSize: 7, cellPadding: 2.2 }
      );
    } else {
      const w = colWidths_([12, 16, 18, 18, 12, 10], ctx.cW);
      ctx.drawAutoTable_(
        [MRD_MILL_SUMMARY_COLS],
        mills.map(function(item) {
          const r = item.row;
          return [
            pdfSanitize(item.risk),
            pdfSanitize(r['GROUP NAME']),
            pdfSanitize(r['COMPANY NAME']),
            pdfSanitize(r['MILL NAME']),
            pdfSanitize(r['PROVINCE']),
            pdfSanitize(isNblYes_(item.nbl) ? 'Yes' : item.nbl),
          ];
        }),
        BRAND,
        {
          0: { cellWidth: w[0] }, 1: { cellWidth: w[1] }, 2: { cellWidth: w[2] },
          3: { cellWidth: w[3] }, 4: { cellWidth: w[4] }, 5: { cellWidth: w[5] },
        }
      );
    }
  }

  const highRisk = data.highRiskMills || [];
  if (highRisk.length) {
    ctx.beginSubsection_('High Risk Suppliers', BRAND);
    const hw = colWidths_([18, 22, 22, 14, 14], ctx.cW);
    ctx.drawAutoTable_(
      [['Group Name', 'Company Name', 'Mill Name', 'Result Risk Level', 'Province']],
      highRisk.map(function(item) {
        const r = item.row;
        return [
          pdfSanitize(r['GROUP NAME']), pdfSanitize(r['COMPANY NAME']), pdfSanitize(r['MILL NAME']),
          pdfSanitize(item.risk), pdfSanitize(r['PROVINCE']),
        ];
      }),
      BRAND,
      {
        0: { cellWidth: hw[0] }, 1: { cellWidth: hw[1] }, 2: { cellWidth: hw[2] },
        3: { cellWidth: hw[3] }, 4: { cellWidth: hw[4] },
      }
    );
  }
}

function traceMetricCards_(t) {
  t = t || {};
  return [
    { label: 'TTM CPO %', value: pdfSanitize(t.ttmCpoFmt || '—'), accent: TTM_CPO, valueColor: TTM_CPO },
    { label: 'TTM PK %', value: pdfSanitize(t.ttmPkFmt || '—'), accent: TTM_PK, valueColor: TTM_PK },
    { label: 'TTP CPO %', value: pdfSanitize(t.ttpCpoFmt || '—'), accent: TTP_CPO, valueColor: TTP_CPO },
    { label: 'TTP PK %', value: pdfSanitize(t.ttpPkFmt || '—'), accent: TTP_PK, valueColor: TTP_PK },
  ];
}

function drawTraceTotalsSection_(ctx, totals, year, noHeader) {
  const t = totals || {};
  if (!noHeader) ctx.beginSection_('03 · Traceability Data ' + pdfSanitize(year), TRACE_ORANGE);
  drawMetricCardGrid_(ctx, traceMetricCards_(t), { cols: 4, cardH: 24, gapAfter: 3 });
}

function drawGrvSection_(ctx, rows, full, noHeader) {
  if (!rows.length) return;
  if (!noHeader) ctx.beginSection_('04 · Grievance Monitoring', GRV_PURPLE);
  if (!full) {
    const w = colWidths_([12, 12, 14, 12, 16, 12, 12], ctx.cW);
    ctx.drawAutoTable_(
      [MRD_GRV_SUMMARY_COLS],
      rows.map(function(item) {
        const r = item.row;
        return [
          pdfSanitize(r['Date Received']),
          pdfSanitize(r['Grievance Category']),
          pdfSanitize(r['Complainant']),
          pdfSanitize(r['Grievance Subject Group']),
          pdfSanitize(r['Grievance Subject'] || r['Subject']),
          pdfSanitize(r['Risk Classification']),
          pdfSanitize(r['Grievance Status']),
        ];
      }),
      GRV_PURPLE,
      {
        0: { cellWidth: w[0] }, 1: { cellWidth: w[1] }, 2: { cellWidth: w[2] },
        3: { cellWidth: w[3] }, 4: { cellWidth: w[4] }, 5: { cellWidth: w[5] }, 6: { cellWidth: w[6] },
      }
    );
    return;
  }

  const w = colWidths_([10, 10, 10, 12, 10, 12, 10, 10, 14, 12, 12, 12], ctx.cW);
  ctx.drawAutoTable_(
    [MRD_GRV_DETAIL_COLS],
    rows.map(function(item) {
      const r = item.row;
      return [
        pdfSanitize(r['Grievance ID']),
        pdfSanitize(r['Date Received']),
        pdfSanitize(r['Grievance Category']),
        pdfSanitize(r['Complainant']),
        pdfSanitize(r['Grievance Subject Group']),
        pdfSanitize(r['Grievance Subject'] || r['Subject']),
        pdfSanitize(r['Risk Classification']),
        pdfSanitize(r['Grievance Status']),
        pdfSanitize(r['Grievance Description']),
        pdfSanitize(r['Verification Findings']),
        pdfSanitize(r['Corrective Action']),
        pdfSanitize(r['Preventive Action']),
      ];
    }),
    GRV_PURPLE,
    {
      0: { cellWidth: w[0] }, 1: { cellWidth: w[1] }, 2: { cellWidth: w[2] },
      3: { cellWidth: w[3] }, 4: { cellWidth: w[4] }, 5: { cellWidth: w[5] }, 6: { cellWidth: w[6] },
      7: { cellWidth: w[7], fontSize: 6.5 }, 8: { cellWidth: w[8], fontSize: 6.5 },
      9: { cellWidth: w[9], fontSize: 6.5 }, 10: { cellWidth: w[10], fontSize: 6.5 },
      11: { cellWidth: w[11], fontSize: 6.5 },
    },
    { fontSize: 7, cellPadding: 2.2 }
  );
}

function drawNblSection_(ctx, rows) {
  if (!rows.length) return;
  ctx.beginSection_('05 · Active NBL Mills', NBL_RED);
  const w = colWidths_([18, 22, 24, 14, 12, 8], ctx.cW);
  ctx.drawAutoTable_(
    [['Group Name', 'Company Name', 'Mill Name', 'Province', 'Result Risk Level', 'No Buy List']],
    rows.map(function(item) {
      const r = item.row || item;
      return [
        pdfSanitize(r['GROUP NAME']),
        pdfSanitize(r['COMPANY NAME']),
        pdfSanitize(r['MILL NAME']),
        pdfSanitize(r['PROVINCE']),
        pdfSanitize(item.risk || r['RESULT RISK LEVEL']),
        'Yes',
      ];
    }),
    NBL_RED,
    { 0: { cellWidth: w[0] }, 1: { cellWidth: w[1] }, 2: { cellWidth: w[2] }, 3: { cellWidth: w[3] }, 4: { cellWidth: w[4] }, 5: { cellWidth: w[5] } }
  );
}

/** Summary PDF — Active NBL list (Company Group Name · Company Name). */
function drawNblSummaryList_(ctx, rows) {
  if (!rows.length) return;
  const w = colWidths_([38, 62], ctx.cW);
  ctx.drawAutoTable_(
    [['Company Group Name', 'Company Name']],
    rows.map(function(item) {
      const r = item.row || item;
      return [
        pdfSanitize(r['GROUP NAME'] || r.group || ''),
        pdfSanitize(r['COMPANY NAME'] || r.company || ''),
      ];
    }),
    NBL_RED,
    { 0: { cellWidth: w[0] }, 1: { cellWidth: w[1] } },
    { fontSize: 7.5, cellPadding: 2.5, gapAfter: 3 }
  );
}

function drawFacilityHero_(ctx, bundle, accent, pctLabel, facilityPct, companiesCount) {
  const doc = ctx.doc;
  const mL = ctx.mL;
  const mR = ctx.mR;
  const cW = ctx.cW;
  const pageW = ctx.pageW;
  const badge = bundle.type === 'pk' ? 'PK' : 'CPO';
  const y0 = ctx.getY();

  doc.setDrawColor.apply(doc, BORDER);
  doc.setFillColor.apply(doc, BG_SOFT);
  doc.setLineWidth(0.25);
  doc.roundedRect(mL, y0, cW, 18, 2, 2, 'FD');

  doc.setFillColor.apply(doc, accent);
  doc.roundedRect(mL + 4, y0 + 5, 20, 8, 1.5, 1.5, 'F');
  doc.setTextColor.apply(doc, WHITE);
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(8);
  doc.text(badge, mL + 14, y0 + 10.5, { align: 'center' });

  doc.setTextColor.apply(doc, INK);
  doc.setFontSize(14);
  doc.text(pdfSanitize(bundle.facility), mL + 28, y0 + 11);

  doc.setFont('helvetica', 'normal');
  doc.setFontSize(8.5);
  doc.setTextColor.apply(doc, INK_MUTED);
  doc.text(
    pctLabel + ': ' + pdfSanitize(facilityPct) + '   ·   ' + companiesCount + ' companies',
    pageW - mR - 4,
    y0 + 11,
    { align: 'right' }
  );
  ctx.setY(y0 + 22);
}

function drawFacilityKpiCards_(ctx, sum, facilityPct, pctLabel, accent) {
  ctx.ensureSpace_(18);
  const doc = ctx.doc;
  const mL = ctx.mL;
  const cW = ctx.cW;
  const y0 = ctx.getY();
  const items = [
    { label: 'No Buy List', value: sum.nblYes > 0 ? String(sum.nblYes) + ' Yes' : '0' },
    { label: 'High Risk', value: String(sum.highRisk || 0) },
    { label: 'Total Grievance', value: String(sum.grievanceSum || 0) },
    { label: 'Est. ISPO Supply %', value: pdfSanitize(sum.ispoPct || '—') },
    { label: pctLabel, value: pdfSanitize(facilityPct) },
  ];
  const gap = 3;
  const cardW = (cW - gap * (items.length - 1)) / items.length;
  const cardH = 16;

  items.forEach(function(item, i) {
    const x = mL + i * (cardW + gap);
    doc.setDrawColor.apply(doc, BORDER);
    doc.setFillColor.apply(doc, WHITE);
    doc.setLineWidth(0.25);
    doc.roundedRect(x, y0, cardW, cardH, 2, 2, 'FD');
    doc.setFillColor.apply(doc, accent);
    doc.rect(x + 2, y0 + 2, cardW - 4, 0.9, 'F');

    doc.setFont('helvetica', 'normal');
    doc.setFontSize(6.5);
    doc.setTextColor.apply(doc, INK_MUTED);
    doc.text(item.label.toUpperCase(), x + cardW / 2, y0 + 6.5, { align: 'center' });

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(10.5);
    doc.setTextColor.apply(doc, INK);
    doc.text(item.value, x + cardW / 2, y0 + 12.5, { align: 'center' });
  });
  ctx.setY(y0 + cardH + 6);
}

function drawFacilityCompanyTable_(ctx, companies, pctLabel, accent) {
  const doc = ctx.doc;
  if (ctx.pageH - ctx.mFoot - 6 - ctx.getY() < 16) {
    ctx.ensureSpace_(18);
  }
  const y0 = ctx.getY();
  const cw = colWidths_([14, 18, 14, 10, 12, 10, 14], ctx.cW);
  const headers = MRD_FACILITY_COMPANY_COLS.slice();
  headers[headers.length - 1] = pctLabel;

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(8);
  doc.setTextColor.apply(doc, INK);
  doc.text('Company breakdown', ctx.mL, y0 + 2);
  ctx.setY(y0 + 5);

  ctx.drawAutoTable_(
    [headers],
    companies.map(function(c) {
      return [
        pdfSanitize(c.group),
        pdfSanitize(c.company),
        pdfSanitize(c.certification),
        pdfSanitize(c.nbl),
        pdfSanitize(c.riskLevel),
        pdfSanitize(c.grievance),
        pctFmt_(c.ttpPctNum),
      ];
    }),
    accent,
    {
      0: { cellWidth: cw[0] }, 1: { cellWidth: cw[1] }, 2: { cellWidth: cw[2], fontSize: 6.8 },
      3: { cellWidth: cw[3], halign: 'center' }, 4: { cellWidth: cw[4], halign: 'center' },
      5: { cellWidth: cw[5], halign: 'right' }, 6: { cellWidth: cw[6], halign: 'right', fontStyle: 'bold' },
    },
    { fontSize: 7.5, cellPadding: 2.2, gapAfter: 2 }
  );
}

function drawFacilityBlock_(ctx, bundle, full, idx) {
  const isPk = bundle.type === 'pk';
  const accent = isPk ? PK_GREEN : BRAND;
  const pctLabel = facilityPctColLabel(isPk);
  const sum = bundle.summary || {};
  const facilityPct = isPk ? (sum.avgPk || '—') : (sum.avgCpo || '—');
  const companies = (bundle.companies || []).slice().sort(function(a, b) {
    return String(a.company).localeCompare(String(b.company), undefined, { sensitivity: 'base' });
  });
  if (!companies.length) return;

  const title = (isPk ? 'PK' : 'CPO') + ' · ' + pdfSanitize(bundle.facility);
  ctx.beginFacilityBlock_(title, accent, idx > 0);

  drawFacilityHero_(ctx, bundle, accent, pctLabel, facilityPct, companies.length);
  drawFacilityKpiCards_(ctx, sum, facilityPct, pctLabel, accent);

  if (full) {
    const profiles = (bundle.profiles || []).filter(function(p) {
      return ['_cplSite', '_cplAddress', '_cplCapacity', '_cplCoordinate', '_cplFacility', '_cplCert'].some(function(k) {
        return pdfSanitize(p[k]) !== '—';
      });
    });
    if (profiles.length) {
      const pw = colWidths_([14, 14, 14, 14, 14, 14], ctx.cW);
      ctx.drawAutoTable_(
        [['Company', 'Site', 'Address', 'Capacity', 'Coordinate', 'Certification']],
        profiles.map(function(p) {
          return [
            pdfSanitize(p._cplCompany),
            pdfSanitize(p._cplSite),
            pdfSanitize(p._cplAddress),
            pdfSanitize(p._cplCapacity),
            pdfSanitize(p._cplCoordinate),
            pdfSanitize(p._cplCert),
          ];
        }),
        accent,
        {
          0: { cellWidth: pw[0] }, 1: { cellWidth: pw[1] }, 2: { cellWidth: pw[2], fontSize: 6.5 },
          3: { cellWidth: pw[3] }, 4: { cellWidth: pw[4], fontSize: 6.5 }, 5: { cellWidth: pw[5], fontSize: 6.5 },
        },
        { fontSize: 7, cellPadding: 2, gapAfter: 4 }
      );
    }
  }

  drawFacilityCompanyTable_(ctx, companies, pctLabel, accent);
}

function drawFacilitySection_(ctx, bundles, full) {
  const active = (bundles || []).filter(function(b) { return (b.companies || []).length > 0; });
  if (!active.length) return;
  active.forEach(function(bundle, idx) {
    drawFacilityBlock_(ctx, bundle, full, idx);
  });
}

function drawEudrSection_(ctx, rows, noHeader) {
  if (!rows.length) return;
  if (!noHeader) ctx.beginSection_('07 · EUDR Potential', EUDR_TEAL);
  const w = colWidths_([12, 18, 20, 20, 14, 16], ctx.cW);
  ctx.drawAutoTable_(
    [MRD_EUDR_COLS],
    rows.map(function(item) {
      const r = item.row;
      return [
        pdfSanitize(r['GROUP NAME']), pdfSanitize(r['COMPANY NAME']), pdfSanitize(r['MILL NAME']),
        pdfSanitize(r['PROVINCE']), pdfSanitize(r['SUPPLY TO']),
        'Potential',
      ];
    }),
    EUDR_TEAL,
    {
      0: { cellWidth: w[0] }, 1: { cellWidth: w[1] }, 2: { cellWidth: w[2] },
      3: { cellWidth: w[3] }, 4: { cellWidth: w[4] },
      5: { cellWidth: w[5], fontSize: 6.5 },
    },
    { fontSize: 7, cellPadding: 2.2 }
  );
}

function facilityBundleSummaryRow_(bundle) {
  const sum = bundle.summary || {};
  const isPk = bundle.type === 'pk';
  const pct = (bundle.traceCalc && bundle.traceCalc.formatted)
    ? bundle.traceCalc.formatted
    : (isPk ? (sum.avgPk || '—') : (sum.avgCpo || '—'));
  return [
    pdfSanitize(bundle.facility),
    pdfSanitize((bundle.companies || []).length),
    pdfSanitize(sum.nblYes != null ? sum.nblYes : 0),
    pdfSanitize(sum.highRisk != null ? sum.highRisk : 0),
    pdfSanitize(sum.grievanceSum != null ? sum.grievanceSum : 0),
    pdfSanitize(sum.ispoPct || '—'),
    pdfSanitize(pct),
  ];
}

function drawFacilitySummaryFromBundles_(ctx, bundles) {
  const active = (bundles || []).filter(function(b) {
    return (b.companies || []).length > 0 && pdfSanitize(b.facility) !== '—';
  });
  if (!active.length) return;

  const cpo = active.filter(function(b) { return b.type !== 'pk'; });
  const pk = active.filter(function(b) { return b.type === 'pk'; });
  const w7 = colWidths_([24, 10, 10, 10, 12, 14, 12], ctx.cW);
  const colStyles = {
    0: { cellWidth: w7[0] }, 1: { cellWidth: w7[1], halign: 'center' },
    2: { cellWidth: w7[2], halign: 'center' }, 3: { cellWidth: w7[3], halign: 'center' },
    4: { cellWidth: w7[4], halign: 'center' }, 5: { cellWidth: w7[5], halign: 'right' },
    6: { cellWidth: w7[6], halign: 'right', fontStyle: 'bold' },
  };

  if (cpo.length) {
    ctx.beginSubsection_('CPO Facility Performance', BRAND);
    ctx.drawAutoTable_(
      [facilitySummaryColLabels(false)],
      cpo.map(facilityBundleSummaryRow_),
      BRAND,
      colStyles,
      { fontSize: 7.5, cellPadding: 2.2, gapAfter: 3 }
    );
  }
  if (pk.length) {
    ctx.beginSubsection_('PK Facility Performance', PK_GREEN);
    ctx.drawAutoTable_(
      [facilitySummaryColLabels(true)],
      pk.map(facilityBundleSummaryRow_),
      PK_GREEN,
      colStyles,
      { fontSize: 7.5, cellPadding: 2.2, gapAfter: 3 }
    );
  }
}

/** @deprecated Use drawFacilitySummaryFromBundles_ */
function drawFacilitySummaryTables_(ctx, facility) {
  drawFacilitySummaryFromBundles_(ctx, []);
}

function websiteKpiItems_(stats) {
  const s = stats || {};
  return [
    { label: 'SDD Requested', value: String(s.sddRequested != null ? s.sddRequested : (s.sddTotal || 0)), sub: (s.sddDone != null ? s.sddDone : (s.sddSubmitted || 0)) + ' done' },
    { label: 'Total Mills', value: String(s.totalMills != null ? s.totalMills : 0), sub: (s.totalGroups || 0) + ' groups' },
    { label: 'Untraceable Mills', value: String(s.emptyTraceMills != null ? s.emptyTraceMills : 0), sub: 'mills without suppliers', hot: (s.emptyTraceMills || 0) > 0 },
    { label: 'EUDR Potential', value: String(s.eudrPotential != null ? s.eudrPotential : 0), sub: 'by formula' },
  ];
}

function drawMetricCardGrid_(ctx, items, opts) {
  opts = opts || {};
  if (!items.length) return;
  const doc = ctx.doc;
  const mL = ctx.mL;
  const cW = ctx.cW;
  const cols = opts.cols || 3;
  const gap = opts.gap != null ? opts.gap : 4;
  const cardH = opts.cardH || 20;
  const cardW = (cW - gap * (cols - 1)) / cols;
  const rowCount = Math.ceil(items.length / cols);
  const blockH = rowCount * cardH + Math.max(0, rowCount - 1) * gap;
  ctx.ensureSpace_(blockH + 2);
  const y0 = ctx.getY();

  items.forEach(function(item, i) {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const x = mL + col * (cardW + gap);
    const cy = y0 + row * (cardH + gap);
    const accent = item.hot ? NBL_RED : (item.accent || opts.accent || BRAND);
    const valueColor = item.hot ? NBL_RED : (item.valueColor || accent || INK);

    doc.setDrawColor.apply(doc, BORDER);
    doc.setFillColor.apply(doc, WHITE);
    doc.setLineWidth(0.2);
    doc.roundedRect(x, cy, cardW, cardH, 2, 2, 'FD');
    doc.setFillColor.apply(doc, accent);
    doc.rect(x + 2, cy + 2, cardW - 4, 0.8, 'F');

    doc.setFont('helvetica', 'normal');
    doc.setFontSize(6.2);
    doc.setTextColor.apply(doc, INK_MUTED);
    doc.text(item.label.toUpperCase(), x + cardW / 2, cy + 6.5, { align: 'center' });

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(item.value.length > 6 ? 10 : 12);
    doc.setTextColor.apply(doc, valueColor);
    doc.text(item.value, x + cardW / 2, cy + 12.5, { align: 'center' });

    if (item.sub) {
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(5.8);
      doc.setTextColor.apply(doc, INK_LIGHT);
      doc.text(String(item.sub), x + cardW / 2, cy + 16.5, { align: 'center' });
    }
  });

  ctx.setY(y0 + blockH + (opts.gapAfter == null ? 4 : opts.gapAfter));
}

function sectionSummaryConfig_(id, stats, data, year) {
  const s = stats || {};
  const grvOpen = (data.grv || []).filter(function(item) {
    return String(item.row['Grievance Status'] || '').toLowerCase().includes('open');
  }).length;
  const configs = {
    sdd: {
      num: '01', title: 'Supplier Due Diligence', accent: BRAND,
      desc: (s.sddRequested || s.sddTotal || 0) + ' requested · ' + (s.sddDone || s.sddSubmitted || 0) + ' done',
      metrics: [
        { label: 'Requested', value: String(s.sddRequested || s.sddTotal || 0) },
        { label: 'Done', value: String(s.sddDone || s.sddSubmitted || 0) },
        { label: 'Draft', value: String(s.sddDraft || 0) },
      ],
    },
    mill: {
      num: '02', title: 'Mill Onboarding', accent: BRAND,
      desc: (s.totalMills || 0) + ' mills · ' + (s.highRisk || 0) + ' high risk',
      metrics: [
        { label: 'Total Mills', value: String(s.totalMills || 0), sub: (s.totalGroups || 0) + ' groups' },
        { label: 'High Risk', value: String(s.highRisk || 0), hot: true },
        { label: 'Active NBL Mills', value: String(s.nblMills || 0) },
        { label: 'Groups', value: String(s.totalGroups || 0) },
      ],
    },
    trace: {
      num: '03', title: 'Traceability Data ' + pdfSanitize(year), accent: TRACE_ORANGE,
      desc: 'TTM (mill coordinates) · TTP (supplier traceability) · ' + (s.emptyTraceMills || 0) + ' mills without supplier',
      metrics: (function() {
        const t = data.traceTotals || {};
        const cards = traceMetricCards_(t);
        const fallbacks = [
          t.ttmCpoFmt || s.ttmCpoPct,
          t.ttmPkFmt || s.ttmPkPct,
          t.ttpCpoFmt || s.ttpCpoPct,
          t.ttpPkFmt || s.ttpPkPct,
        ];
        return cards.map(function(card, i) {
          return {
            label: card.label,
            value: pdfSanitize(fallbacks[i] || '—'),
            accent: card.accent,
            valueColor: card.valueColor,
          };
        });
      })(),
    },
    grv: {
      num: '04', title: 'Grievance Monitoring', accent: GRV_PURPLE,
      desc: (s.grievances || 0) + ' in period · ' + grvOpen + ' open',
      metrics: [
        { label: 'Total', value: String(s.grievances || 0) },
        { label: 'Open', value: String(grvOpen), hot: grvOpen > 0 },
        { label: 'Closed', value: String(Math.max(0, (s.grievances || 0) - grvOpen)) },
      ],
    },
    nbl: {
      num: '05', title: 'Active NBL Mills', accent: NBL_RED,
      desc: (s.nblMills || 0) + ' mills on No Buy List',
      metrics: [
        { label: 'Active NBL Mills', value: String(s.nblMills || 0), hot: (s.nblMills || 0) > 0 },
        { label: 'Total Mills', value: String(s.totalMills || 0) },
      ],
    },
    facility: {
      num: '06', title: 'Facility Performance', accent: PK_GREEN,
      desc: 'CPO & PK summary',
      metrics: [
        { label: 'CPO Facilities', value: String((data.facility && data.facility.cpo || []).length) },
        { label: 'PK Facilities', value: String((data.facility && data.facility.pk || []).length) },
        { label: 'Total Facilities', value: String(s.facilities || 0) },
      ],
    },
    eudr: {
      num: '07', title: 'EUDR Potential', accent: EUDR_TEAL,
      desc: (s.eudrPotential || 0) + ' potential mills',
      metrics: [
        { label: 'Potential', value: String(s.eudrPotential || 0) },
        { label: 'Total Mills', value: String(s.totalMills || 0) },
      ],
    },
  };
  return configs[id] || null;
}

function drawSectionSummaryBlock_(ctx, cfg) {
  if (!cfg) return;
  ctx.beginSection_(cfg.num + ' · ' + cfg.title, cfg.accent);
  const doc = ctx.doc;
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(7.5);
  doc.setTextColor.apply(doc, INK_MUTED);
  doc.text(cfg.desc, ctx.mL, ctx.getY());
  ctx.setY(ctx.getY() + 5);
  drawMetricCardGrid_(ctx, cfg.metrics, { cols: Math.min(4, cfg.metrics.length), accent: cfg.accent, cardH: 19, gapAfter: 3 });
}

function drawSummaryReportBody_(ctx, data, sections, stats, year) {
  const bodySections = sections.filter(function(id) { return id !== 'kpi'; });
  bodySections.forEach(function(id) {
    const cfg = sectionSummaryConfig_(id, stats, data, year);
    if (!cfg) return;
    drawSectionSummaryBlock_(ctx, cfg);
    if (id === 'sdd') drawSddSection_(ctx, data.sdd || [], true);
    else if (id === 'trace') { /* totals shown in metric cards above */ }
    else if (id === 'grv') drawGrvSection_(ctx, data.grv || [], false, true);
    else if (id === 'nbl') drawNblSummaryList_(ctx, data.nblAll || []);
    else if (id === 'facility') drawFacilitySummaryFromBundles_(ctx, data.facilityBundles || []);
    else if (id === 'eudr') drawEudrSection_(ctx, data.eudrPotential || [], true);
  });
}

function drawDetailReportBody_(ctx, data, sections, year) {
  if (sections.indexOf('sdd') !== -1) drawSddSection_(ctx, data.sdd || []);
  if (sections.indexOf('mill') !== -1) drawMillSection_(ctx, data, true);
  if (sections.indexOf('trace') !== -1) drawTraceTotalsSection_(ctx, data.traceTotals || {}, year);
  if (sections.indexOf('grv') !== -1) drawGrvSection_(ctx, data.grv || [], true);
  if (sections.indexOf('nbl') !== -1) drawNblSection_(ctx, data.nblAll || []);
  if (sections.indexOf('facility') !== -1) drawFacilitySection_(ctx, data.facilityBundles || [], true);
  if (sections.indexOf('eudr') !== -1) drawEudrSection_(ctx, data.eudrPotential || []);
}

function pdfFileLabel_(year, month) {
  return (year || 'all') + (month ? '_' + monthLabel(month) : '');
}

function saveMonthlyReportPdf_(doc, variant, year, month) {
  const tag = variant === 'summary' ? 'Summary' : 'Detail';
  const fileLabel = pdfFileLabel_(year, month).replace(/\s+/g, '-');
  doc.save('Monthly-Report-' + tag + '-' + fileLabel + '.pdf');
}

function finalizePdf_(ctx) {
  ctx.pruneBlankPages_();
  drawFooters_(ctx.doc, ctx.pageW, ctx.pageH, ctx.mL, ctx.mR, ctx.mFoot);
  return ctx.doc;
}

function buildSummaryPdfDoc_(jsPDFLib, opts) {
  const data = opts.data || {};
  const stats = data.stats || {};
  const sections = (opts.sections && opts.sections.length) ? opts.sections : DEFAULT_SECTIONS.slice();
  const ctx = createPdfContext_(jsPDFLib);
  ctx.periodText = 'Period: ' + periodLabel(opts.year, opts.month);
  ctx.generatedAt = new Date().toLocaleString('en-GB', { hour12: false });
  ctx.reportTitle = 'Monthly Report — Summary';
  ctx.reportSubtitle = '';

  const headerEnd = ctx.drawMainHeader_();
  ctx.setY(headerEnd + 2);

  if (sections.indexOf('kpi') !== -1) {
    docSetSubhead_(ctx, 'Overview');
    drawMetricCardGrid_(ctx, websiteKpiItems_(stats), { cols: 4, cardH: 22, gapAfter: 6 });
  }

  drawSummaryReportBody_(ctx, data, sections, stats, opts.year);
  return finalizePdf_(ctx);
}

function docSetSubhead_(ctx, text) {
  const doc = ctx.doc;
  ctx.ensureSpace_(10);
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(10);
  doc.setTextColor.apply(doc, INK);
  doc.text(text, ctx.mL, ctx.getY() + 2);
  ctx.setY(ctx.getY() + 8);
}

function buildDetailPdfDoc_(jsPDFLib, opts) {
  const data = opts.data || {};
  const sections = (opts.sections && opts.sections.length) ? opts.sections : DEFAULT_SECTIONS.slice();
  const detailSections = sections.filter(function(id) { return id !== 'kpi'; });
  const ctx = createPdfContext_(jsPDFLib, { orientation: 'landscape' });
  ctx.periodText = 'Period: ' + periodLabel(opts.year, opts.month);
  ctx.generatedAt = new Date().toLocaleString('en-GB', { hour12: false });

  ctx.setY(drawCompactHeaderForExport_(ctx));
  if (!detailSections.length) {
    const doc = ctx.doc;
    ctx.ensureSpace_(12);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(9);
    doc.setTextColor.apply(doc, INK_MUTED);
    doc.text('No detail sections selected. Include at least one section besides Overview.', ctx.mL, ctx.getY() + 4);
    ctx.markContent_(ctx.getY() + 8);
  } else {
    drawDetailReportBody_(ctx, data, detailSections, opts.year);
  }
  return finalizePdf_(ctx);
}

function drawCompactHeaderForExport_(ctx) {
  const doc = ctx.doc;
  doc.setFillColor.apply(doc, BRAND);
  doc.rect(0, 0, ctx.pageW, 22, 'F');
  doc.setTextColor.apply(doc, WHITE);
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(11);
  doc.text('Monthly Report — Detail', ctx.mL, 9);
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(7);
  doc.text(ctx.periodText + '   ·   ' + ctx.generatedAt, ctx.mL, 15.5);
  ctx.markContent_(20);
  return 22;
}

/** @deprecated Use buildMonthlyReportPdfPair_ */
export async function buildMonthlyReportPdf_(opts) {
  return buildMonthlyReportPdfPair_(opts);
}

export async function buildMonthlyReportPdfPair_(opts) {
  const jsPDFLib = opts.getJsPDF();
  if (!jsPDFLib) throw new Error('PDF library is not ready.');

  const summaryDoc = buildSummaryPdfDoc_(jsPDFLib, opts);
  saveMonthlyReportPdf_(summaryDoc, 'summary', opts.year, opts.month);

  await new Promise(function(resolve) { setTimeout(resolve, 350); });

  const detailDoc = buildDetailPdfDoc_(jsPDFLib, opts);
  saveMonthlyReportPdf_(detailDoc, 'detail', opts.year, opts.month);
}

export const MRD_PDF_SECTIONS = [
  { id: 'kpi', label: 'Overview · KPI cards' },
  { id: 'sdd', label: MRD_SECTION_LABELS.sdd },
  { id: 'mill', label: MRD_SECTION_LABELS.mill },
  { id: 'trace', label: MRD_SECTION_LABELS.trace },
  { id: 'grv', label: MRD_SECTION_LABELS.grv },
  { id: 'nbl', label: MRD_SECTION_LABELS.nbl },
  { id: 'facility', label: MRD_SECTION_LABELS.facility },
  { id: 'eudr', label: MRD_SECTION_LABELS.eudr },
];
