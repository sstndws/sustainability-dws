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
  facilityTtmColLabel,
  facilitySummaryColLabels,
  pdfTableHead,
  pdfHeadRow,
  pdfCellTrim,
  normalizeSddCategory,
  sddStatusText,
  sddCompanyName,
  sddDateImport,
  sddLastUpdate,
  mrdSortSddRows_,
  mrdSortMillItems_,
  mrdSortGrvItemsByDateDesc_,
  mrdSortEudrItems_,
  mrdSortFacilityBundles_,
  mrdSortBundlesByFacility_,
  mrdSortFacilityCompanies_,
  mrdFormatNblRisers_,
  mrdReportHeaderMeta_,
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

/** Vertical rhythm — tight within sections; sectionGap only between major points. */
const PDF_LAYOUT = {
  bodyGap: 6,
  sectionGap: 8,
  afterSectionBar: 2,
  subsectionGap: 4,
  sectionDescGap: 1,
  descToCards: 3,
  descLineH: 3.4,
  cardsToContent: 4,
  cardGap: 3,
  headerMainH: 44,
  headerDetailH: 32,
  headerCompactH: 16,
  compactBodyGap: 4,
  autoTableTopMargin: 22,
};

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

function applyReportHeaderMeta_(ctx, year, month) {
  const meta = mrdReportHeaderMeta_(year, month);
  ctx.periodText = meta.periodLine;
  ctx.dataPeriodText = meta.dataPeriodLine;
  ctx.cutoffText = meta.cutoffLine;
  ctx.cutoffCompact = meta.period + ' · ' + meta.dataPeriod + ' · ' + meta.cutoffCompact;
  ctx.exportedAt = meta.exportedAt;
  return meta;
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

function drawFooters_(doc, pageW, pageH, mL, mR, mFoot, exportedAt) {
  const total = doc.internal.getNumberOfPages();
  for (let p = 1; p <= total; p++) {
    doc.setPage(p);
    doc.setDrawColor.apply(doc, BORDER);
    doc.setLineWidth(0.15);
    doc.line(mL, pageH - mFoot, pageW - mR, pageH - mFoot);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(6.5);
    doc.setTextColor.apply(doc, INK_LIGHT);
    const left = exportedAt
      ? 'Sustainability Dashboard · Monthly Report · Exported ' + exportedAt
      : 'Sustainability Dashboard · Monthly Report';
    doc.text(left, mL, pageH - 5);
    doc.text('Page ' + p + ' of ' + total, pageW - mR, pageH - 5, { align: 'right' });
  }
}

function monthlyReportFileName_(variant, year, month) {
  const tag = variant === 'summary' ? 'Summary' : 'Detail';
  const fileLabel = pdfFileLabel_(year, month).replace(/\s+/g, '-');
  return 'Monthly-Report-' + tag + '-' + fileLabel + '.pdf';
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
    dataPeriodText: '',
    cutoffText: '',
    cutoffCompact: '',
    exportedAt: '',
    reportTitle: 'Monthly Report',
    reportSubtitle: 'Compliance snapshot — SDD, Mill, Traceability, Grievance, NBL, Facility, EUDR',
    detailLevel: 'summary',
    pdfMode: opts.pdfMode || 'detail',
    isLandscape: opts.orientation === 'landscape',
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

  function compactHeaderBandH_() {
    return ctx.isLandscape ? PDF_LAYOUT.headerDetailH : PDF_LAYOUT.headerCompactH;
  }

  function autoTableTopMargin_() {
    return compactHeaderBandH_() + PDF_LAYOUT.compactBodyGap + 2;
  }

  function drawCompactHeader_() {
    const bandH = compactHeaderBandH_();
    doc.setFillColor.apply(doc, BRAND);
    doc.rect(0, 0, pageW, bandH, 'F');
    doc.setTextColor.apply(doc, WHITE);
    doc.setFont('helvetica', 'bold');
    if (ctx.isLandscape) {
      doc.setFontSize(11);
      doc.text('Monthly Report — Detail', mL, 9);
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(6.5);
      doc.text(ctx.periodText || '', mL, 14.5);
      doc.text(ctx.dataPeriodText || '', mL, 18.5);
      doc.text(ctx.cutoffText || '', mL, 22.5);
      if (ctx.exportedAt) {
        doc.text('Generated on ' + ctx.exportedAt, pageW - mR, 22.5, { align: 'right' });
      }
    } else {
      doc.setFontSize(9.5);
      doc.text('Monthly Report', mL, 6.5);
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(6.5);
      doc.text(ctx.cutoffCompact || ctx.periodText || '', mL, 11.5);
      if (ctx.exportedAt) {
        doc.text('Generated on ' + ctx.exportedAt, pageW - mR, 11.5, { align: 'right' });
      }
    }
    return bandH + PDF_LAYOUT.compactBodyGap;
  }

  function ensureSpace(needed) {
    if (y + needed <= pageH - mFoot - 6) return;
    doc.addPage();
    y = drawCompactHeader_();
  }

  function remainingSpace_() {
    return pageH - mFoot - 6 - y;
  }

  function ensureBlockFits_(needed) {
    const extra = sectionStarted ? PDF_LAYOUT.sectionGap : 0;
    if (remainingSpace_() < needed + extra) {
      doc.addPage();
      y = drawCompactHeader_();
    }
  }

  function drawSectionBar_(title, accent, fontSize) {
    const fs = fontSize || 9;
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(fs);
    const lines = doc.splitTextToSize(String(title || ''), cW - 6);
    const barH = Math.max(8, lines.length * 3.6 + 3);
    ensureSpace(barH + 4);
    const barY = y;
    doc.setFillColor.apply(doc, accent || BRAND);
    doc.roundedRect(mL, barY, cW, barH, 1.2, 1.2, 'F');
    doc.setTextColor.apply(doc, WHITE);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(fs);
    doc.text(lines, mL + 3, barY + 4.5 + (lines.length > 1 ? 0 : 0.5));
    y = barY + barH + PDF_LAYOUT.afterSectionBar;
    markContent_(y);
  }

  function beginSection_(title, accent) {
    if (sectionStarted && remainingSpace_() < 36) {
      doc.addPage();
      y = drawCompactHeader_();
    } else if (sectionStarted) {
      y += PDF_LAYOUT.sectionGap;
    }
    sectionStarted = true;
    drawSectionBar_(title, accent);
  }

  function beginSubsection_(title, accent) {
    if (remainingSpace_() < 28) {
      doc.addPage();
      y = drawCompactHeader_();
    } else {
      y += PDF_LAYOUT.subsectionGap;
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
      y += PDF_LAYOUT.subsectionGap;
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
    const tableTopMargin = autoTableTopMargin_();
    const base = {
      theme: 'grid',
      margin: { left: mL, right: mR, bottom: mFoot + 12, top: tableTopMargin },
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
        fontSize: opts.headFontSize != null ? opts.headFontSize : 5.8,
        halign: 'center',
        valign: 'middle',
        overflow: 'linebreak',
        minCellHeight: opts.headMinHeight != null ? opts.headMinHeight : 11,
        cellPadding: { top: 1.2, right: 1, bottom: 1.2, left: 1 },
      },
      alternateRowStyles: { fillColor: [255, 253, 253] },
      rowPageBreak: opts.rowPageBreak || 'avoid',
      showHead: opts.showHeadEveryPage ? 'everyPage' : 'firstPage',
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
          const top = p === startPage ? tableTop : tableTopMargin;
          markPageSpan_(p, top, cursor.y + 6);
        }
      },
    }));

    const table = doc.lastAutoTable;
    if (table && table.finalY) {
      const endPage = doc.internal.getNumberOfPages();
      markPageSpan_(endPage, endPage === startPage ? tableTop : tableTopMargin, table.finalY);
      y = table.finalY + (opts.gapAfter == null ? 2 : opts.gapAfter);
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
    const bandH = PDF_LAYOUT.headerMainH;
    doc.setFillColor.apply(doc, BRAND);
    doc.rect(0, 0, pageW, bandH, 'F');
    doc.setFillColor(180, 40, 40);
    doc.rect(0, bandH - 1.5, pageW, 1.5, 'F');
    doc.setTextColor.apply(doc, WHITE);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(17);
    doc.text(ctx.reportTitle || 'Monthly Report', mL, 11.5);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(7.8);
    doc.text(ctx.periodText || '', mL, 18.5);
    doc.setFontSize(7.5);
    doc.text(ctx.dataPeriodText || '', mL, 23.5);
    doc.text(ctx.cutoffText || '', mL, 28.5);
    if (ctx.exportedAt) {
      doc.setFontSize(6.5);
      doc.text('Generated on ' + ctx.exportedAt, pageW - mR, 28.5, { align: 'right' });
    }
    markContent_(bandH);
    return bandH;
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
      if (bottom < MIN_CONTENT_Y) {
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
  ctx.ensureBlockFits_ = ensureBlockFits_;

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
  rows = mrdSortSddRows_(rows || []);
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
    pdfTableHead(MRD_SDD_COLS),
    body,
    BRAND,
    {
      0: { cellWidth: w[0] }, 1: { cellWidth: w[1] }, 2: { cellWidth: w[2] },
      3: { cellWidth: w[3], fontSize: 6.5 }, 4: { cellWidth: w[4] },
    },
    { fontSize: 6.5, cellPadding: 2, headFontSize: 5.6, headMinHeight: 11 }
  );
}

function millDetailCells_(item) {
  const r = item.row;
  return {
    supplierStatus: pdfSanitize(r['SUPPLIER STATUS']),
    certification: pdfCellTrim(r['CERTIFICATION'], 36),
    grievances: pdfSanitize(r['TOTAL GRIEVANCES']),
    facilityCpo: pdfCellTrim(r['FACILITY NAME CPO'], 44),
    facilityPk: pdfCellTrim(r['FACILITY NAME PK'], 44),
    nblBy: pdfSanitize(item.nblBy),
  };
}

function drawMillSection_(ctx, data, full) {
  const mills = mrdSortMillItems_(data.mills || []);
  if (!mills.length && !(data.highRiskMills || []).length) return;
  ctx.beginSection_('02 · Mill Onboarding', BRAND);

  if (mills.length) {
    if (full) {
      const w = colWidths_([9, 13, 16, 15, 9, 7, 10, 9, 7, 15, 15], ctx.cW);
      ctx.drawAutoTable_(
        pdfTableHead(MRD_MILL_FULL_COLS),
        mills.map(function(item) {
          const r = item.row;
          const d = millDetailCells_(item);
          return [
            pdfSanitize(item.risk),
            pdfCellTrim(r['GROUP NAME'], 28),
            pdfCellTrim(r['COMPANY NAME'], 32),
            pdfCellTrim(r['MILL NAME'], 28),
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
        { fontSize: 6.5, cellPadding: 2, headFontSize: 5.6, headMinHeight: 12 }
      );
    } else {
      const w = colWidths_([10, 15, 18, 18, 11, 10], ctx.cW);
      ctx.drawAutoTable_(
        pdfTableHead(MRD_MILL_SUMMARY_COLS),
        mills.map(function(item) {
          const r = item.row;
          return [
            pdfSanitize(item.risk),
            pdfCellTrim(r['GROUP NAME'], 28),
            pdfCellTrim(r['COMPANY NAME'], 32),
            pdfCellTrim(r['MILL NAME'], 28),
            pdfSanitize(r['PROVINCE']),
            pdfSanitize(isNblYes_(item.nbl) ? 'Yes' : item.nbl),
          ];
        }),
        BRAND,
        {
          0: { cellWidth: w[0] }, 1: { cellWidth: w[1] }, 2: { cellWidth: w[2] },
          3: { cellWidth: w[3] }, 4: { cellWidth: w[4] }, 5: { cellWidth: w[5] },
        },
        { headFontSize: 5.6, headMinHeight: 11 }
      );
    }
  }

  const highRisk = mrdSortMillItems_(data.highRiskMills || []);
  if (highRisk.length) {
    ctx.beginSubsection_('High Risk Suppliers', BRAND);
    const hw = colWidths_([18, 22, 22, 14, 14], ctx.cW);
    ctx.drawAutoTable_(
      pdfTableHead(['Group Name', 'Company Name', 'Mill Name', 'Result Risk Level', 'Province']),
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
  drawMetricCardGrid_(ctx, traceMetricCards_(t), { cols: 4, cardH: 24, gapAfter: 0 });
}

function grvStatusCounts_(rows) {
  let open = 0;
  let closed = 0;
  let invalid = 0;
  (rows || []).forEach(function(item) {
    const st = String((item.row || item)['Grievance Status'] || '').toLowerCase();
    if (st.includes('open')) open += 1;
    else if (st.includes('invalid')) invalid += 1;
    else if (st.includes('closed')) closed += 1;
  });
  return { open: open, closed: closed, invalid: invalid };
}

function drawGrvSection_(ctx, rows, full, noHeader) {
  rows = mrdSortGrvItemsByDateDesc_(rows);
  if (!rows.length) return;
  if (!noHeader) ctx.beginSection_('04 · Grievance Monitoring', GRV_PURPLE);
  if (!full) {
    const w = colWidths_([12, 12, 14, 12, 16, 12, 12], ctx.cW);
    ctx.drawAutoTable_(
      pdfTableHead(MRD_GRV_SUMMARY_COLS),
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
      },
      { fontSize: 6.5, cellPadding: 2, headFontSize: 5.5, headMinHeight: 11 }
    );
    const docSum = ctx.doc;
    ctx.ensureSpace_(6);
    docSum.setFont('helvetica', 'italic');
    docSum.setFontSize(6.5);
    docSum.setTextColor.apply(docSum, INK_LIGHT);
    docSum.text(rows.length + ' record' + (rows.length === 1 ? '' : 's') + ' in this period', ctx.mL, ctx.getY() + 3);
    ctx.setY(ctx.getY() + 6);
    return;
  }

  const w = colWidths_([9, 9, 9, 10, 9, 11, 9, 9, 14, 11, 11, 11], ctx.cW);
  ctx.drawAutoTable_(
    pdfTableHead(MRD_GRV_DETAIL_COLS),
    rows.map(function(item) {
      const r = item.row;
      return [
        pdfSanitize(r['Grievance ID']),
        pdfSanitize(r['Date Received']),
        pdfSanitize(r['Grievance Category']),
        pdfSanitize(r['Complainant']),
        pdfSanitize(r['Grievance Subject Group']),
        pdfCellTrim(r['Grievance Subject'] || r['Subject'], 48),
        pdfSanitize(r['Risk Classification']),
        pdfSanitize(r['Grievance Status']),
        pdfCellTrim(r['Grievance Description'], 320),
        pdfCellTrim(r['Verification Findings'], 120),
        pdfCellTrim(r['Corrective Action'], 120),
        pdfCellTrim(r['Preventive Action'], 120),
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
    { fontSize: 6.5, cellPadding: 2, headFontSize: 5.4, headMinHeight: 12, rowPageBreak: 'auto', showHeadEveryPage: true }
  );
  const doc = ctx.doc;
  ctx.ensureSpace_(6);
  doc.setFont('helvetica', 'italic');
  doc.setFontSize(6.5);
  doc.setTextColor.apply(doc, INK_LIGHT);
  doc.text(rows.length + ' record' + (rows.length === 1 ? '' : 's') + ' in this period', ctx.mL, ctx.getY() + 3);
  ctx.setY(ctx.getY() + 6);
}

function drawNblSection_(ctx, rows) {
  rows = mrdSortMillItems_(rows);
  if (!rows.length) return;
  ctx.beginSection_('05 · Active NBL Mills', NBL_RED);
  const w = colWidths_([14, 18, 18, 10, 10, 8, 14], ctx.cW);
  ctx.drawAutoTable_(
    pdfTableHead(['Group Name', 'Company Name', 'Mill Name', 'Province', 'Result Risk Level', 'No Buy List', 'NBL Riser']),
    rows.map(function(item) {
      const r = item.row || item;
      return [
        pdfSanitize(r['GROUP NAME']),
        pdfSanitize(r['COMPANY NAME']),
        pdfSanitize(r['MILL NAME']),
        pdfSanitize(r['PROVINCE']),
        pdfSanitize(item.risk || r['RESULT RISK LEVEL']),
        'Yes',
        pdfSanitize(mrdFormatNblRisers_(item)),
      ];
    }),
    NBL_RED,
    {
      0: { cellWidth: w[0] }, 1: { cellWidth: w[1] }, 2: { cellWidth: w[2] }, 3: { cellWidth: w[3] },
      4: { cellWidth: w[4] }, 5: { cellWidth: w[5], halign: 'center' }, 6: { cellWidth: w[6], fontSize: 6.5 },
    },
    { fontSize: 6.5, cellPadding: 2, headFontSize: 5.4, headMinHeight: 11 }
  );
}

/** Summary PDF — Active NBL list (Company Group Name · Company Name · NBL Riser). */
function drawNblSummaryList_(ctx, rows) {
  rows = mrdSortMillItems_(rows);
  if (!rows.length) return;
  const w = colWidths_([28, 36, 36], ctx.cW);
  ctx.drawAutoTable_(
    pdfTableHead(['Company Group Name', 'Company Name', 'NBL Riser']),
    rows.map(function(item) {
      const r = item.row || item;
      return [
        pdfSanitize(r['GROUP NAME'] || r.group || ''),
        pdfSanitize(r['COMPANY NAME'] || r.company || ''),
        pdfSanitize(mrdFormatNblRisers_(item)),
      ];
    }),
    NBL_RED,
    { 0: { cellWidth: w[0] }, 1: { cellWidth: w[1] }, 2: { cellWidth: w[2], fontSize: 6.5 } },
    { fontSize: 7.5, cellPadding: 2.5, gapAfter: 3, headFontSize: 5.4, headMinHeight: 11 }
  );
}

function drawFacilityHero_(ctx, bundle, accent, pctLabel, facilityPct, ttmLabel, facilityTtm, companiesCount) {
  const doc = ctx.doc;
  const mL = ctx.mL;
  const cW = ctx.cW;
  const badge = bundle.type === 'pk' ? 'PK' : 'CPO';
  const y0 = ctx.getY();

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(12);
  const nameLines = doc.splitTextToSize(pdfSanitize(bundle.facility), cW - 36);
  const boxH = Math.max(22, 14 + nameLines.length * 3.6 + 6);

  doc.setDrawColor.apply(doc, BORDER);
  doc.setFillColor.apply(doc, BG_SOFT);
  doc.setLineWidth(0.25);
  doc.roundedRect(mL, y0, cW, boxH, 2, 2, 'FD');

  doc.setFillColor.apply(doc, accent);
  doc.roundedRect(mL + 4, y0 + 5, 20, 8, 1.5, 1.5, 'F');
  doc.setTextColor.apply(doc, WHITE);
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(8);
  doc.text(badge, mL + 14, y0 + 10.5, { align: 'center' });

  doc.setTextColor.apply(doc, INK);
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(12);
  doc.text(nameLines, mL + 28, y0 + 9);

  doc.setFont('helvetica', 'normal');
  doc.setFontSize(8);
  doc.setTextColor.apply(doc, INK_MUTED);
  const metaY = y0 + 9 + nameLines.length * 3.6 + 1;
  doc.text(
    ttmLabel + ': ' + pdfSanitize(facilityTtm) + '   ·   '
      + pctLabel + ': ' + pdfSanitize(facilityPct) + '   ·   '
      + companiesCount + ' companies',
    mL + 28,
    metaY
  );
  ctx.setY(y0 + boxH + 2);
}

function drawFacilityKpiCards_(ctx, sum, facilityPct, pctLabel, facilityTtm, ttmLabel, accent) {
  ctx.ensureSpace_(18);
  const doc = ctx.doc;
  const mL = ctx.mL;
  const cW = ctx.cW;
  const y0 = ctx.getY();
  const items = [
    { label: 'No Buy List', value: sum.nblYes > 0 ? String(sum.nblYes) + ' Yes' : '0' },
    { label: 'High Risk', value: String(sum.highRisk || 0) },
    { label: 'Total Grievance', value: String(sum.grievanceSum || 0) },
    { label: 'Estimated ISPO Supply %', value: pdfSanitize(sum.ispoPct || '—') },
    { label: ttmLabel, value: pdfSanitize(facilityTtm) },
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

function drawFacilityCompanyTable_(ctx, companies, isPk, accent) {
  const doc = ctx.doc;
  if (ctx.pageH - ctx.mFoot - 6 - ctx.getY() < 16) {
    ctx.ensureSpace_(18);
  }
  const y0 = ctx.getY();
  const cw = colWidths_([12, 16, 12, 9, 10, 9, 10, 10], ctx.cW);
  const ttpLabel = facilityPctColLabel(isPk);
  const ttmLabel = facilityTtmColLabel(isPk);

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(8);
  doc.setTextColor.apply(doc, INK);
  doc.text('Company breakdown', ctx.mL, y0 + 2);
  ctx.setY(y0 + 5);

  ctx.drawAutoTable_(
    pdfTableHead([
      'Group Name', 'Company Name', 'Certification', 'No Buy List', 'Result Risk Level',
      'Total Grievance', ttmLabel, ttpLabel,
    ]),
    companies.map(function(c) {
      return [
        pdfSanitize(c.group || c['GROUP NAME'] || ''),
        pdfSanitize(c.company),
        pdfSanitize(c.certification),
        pdfSanitize(c.nbl),
        pdfSanitize(c.riskLevel),
        pdfSanitize(c.grievance),
        pctFmt_(c.ttmPctNum != null ? c.ttmPctNum : 0),
        pctFmt_(c.ttpPctNum),
      ];
    }),
    accent,
    {
      0: { cellWidth: cw[0] }, 1: { cellWidth: cw[1] }, 2: { cellWidth: cw[2], fontSize: 6.8 },
      3: { cellWidth: cw[3], halign: 'center' }, 4: { cellWidth: cw[4], halign: 'center' },
      5: { cellWidth: cw[5], halign: 'right' }, 6: { cellWidth: cw[6], halign: 'right', fontStyle: 'bold' },
      7: { cellWidth: cw[7], halign: 'right', fontStyle: 'bold' },
    },
    { fontSize: 6.5, cellPadding: 2, gapAfter: 2, headFontSize: 5.5, headMinHeight: 11 }
  );
}

function drawFacilityBlock_(ctx, bundle, full, idx) {
  const isPk = bundle.type === 'pk';
  const accent = isPk ? PK_GREEN : BRAND;
  const pctLabel = facilityPctColLabel(isPk);
  const ttmLabel = facilityTtmColLabel(isPk);
  const sum = bundle.summary || {};
  const facilityPct = (bundle.traceCalc && bundle.traceCalc.formatted)
    ? bundle.traceCalc.formatted
    : (isPk ? (sum.avgPk || '—') : (sum.avgCpo || '—'));
  const facilityTtm = (bundle.ttmCalc && bundle.ttmCalc.formatted)
    ? bundle.ttmCalc.formatted
    : (sum.avgTtm || '0%');
  const companies = mrdSortFacilityCompanies_(bundle.companies || []);
  if (!companies.length) return;

  const title = (isPk ? 'PK' : 'CPO') + ' · ' + pdfSanitize(bundle.facility);
  ctx.beginFacilityBlock_(title, accent, idx > 0);

  drawFacilityHero_(ctx, bundle, accent, pctLabel, facilityPct, ttmLabel, facilityTtm, companies.length);
  drawFacilityKpiCards_(ctx, sum, facilityPct, pctLabel, facilityTtm, ttmLabel, accent);

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

  drawFacilityCompanyTable_(ctx, companies, isPk, accent);
}

function drawFacilitySection_(ctx, bundles, full) {
  const active = mrdSortFacilityBundles_((bundles || []).filter(function(b) { return (b.companies || []).length > 0; }));
  if (!active.length) return;
  ctx.beginSection_('06 · Facility Performance', PK_GREEN);
  active.forEach(function(bundle, idx) {
    drawFacilityBlock_(ctx, bundle, full, idx);
  });
}

function drawEudrSection_(ctx, rows, noHeader) {
  rows = mrdSortEudrItems_(rows);
  if (!rows.length) return;
  if (!noHeader) ctx.beginSection_('07 · EUDR Potential', EUDR_TEAL);
  const w = colWidths_([12, 18, 20, 20, 14, 16], ctx.cW);
  ctx.drawAutoTable_(
    pdfTableHead(MRD_EUDR_COLS),
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
    { fontSize: 6.5, cellPadding: 2, headFontSize: 5.6, headMinHeight: 11 }
  );
}

function facilityBundleSummaryRow_(bundle) {
  const sum = bundle.summary || {};
  const isPk = bundle.type === 'pk';
  const pct = (bundle.traceCalc && bundle.traceCalc.formatted)
    ? bundle.traceCalc.formatted
    : (isPk ? (sum.avgPk || '—') : (sum.avgCpo || '—'));
  const ttm = (bundle.ttmCalc && bundle.ttmCalc.formatted)
    ? bundle.ttmCalc.formatted
    : (sum.avgTtm || '0%');
  return [
    pdfSanitize(bundle.facility),
    pdfSanitize((bundle.companies || []).length),
    pdfSanitize(sum.nblYes != null ? sum.nblYes : 0),
    pdfSanitize(sum.highRisk != null ? sum.highRisk : 0),
    pdfSanitize(sum.grievanceSum != null ? sum.grievanceSum : 0),
    pdfSanitize(sum.ispoPct || '—'),
    pdfSanitize(ttm),
    pdfSanitize(pct),
  ];
}

function drawFacilitySummaryFromBundles_(ctx, bundles) {
  const active = mrdSortFacilityBundles_((bundles || []).filter(function(b) {
    return (b.companies || []).length > 0 && pdfSanitize(b.facility) !== '—';
  }));
  if (!active.length) return;

  const cpo = mrdSortBundlesByFacility_(active.filter(function(b) { return b.type !== 'pk'; }));
  const pk = mrdSortBundlesByFacility_(active.filter(function(b) { return b.type === 'pk'; }));
  const w7 = colWidths_([22, 9, 9, 9, 11, 12, 11, 11], ctx.cW);
  const colStyles = {
    0: { cellWidth: w7[0] }, 1: { cellWidth: w7[1], halign: 'center' },
    2: { cellWidth: w7[2], halign: 'center' }, 3: { cellWidth: w7[3], halign: 'center' },
    4: { cellWidth: w7[4], halign: 'center' }, 5: { cellWidth: w7[5], halign: 'right' },
    6: { cellWidth: w7[6], halign: 'right', fontStyle: 'bold' },
    7: { cellWidth: w7[7], halign: 'right', fontStyle: 'bold' },
  };

  if (cpo.length) {
    ctx.beginSubsection_('CPO Facility Performance', BRAND);
    ctx.drawAutoTable_(
      pdfTableHead(facilitySummaryColLabels(false)),
      cpo.map(facilityBundleSummaryRow_),
      BRAND,
      colStyles,
      { fontSize: 7, cellPadding: 2, gapAfter: 3, headFontSize: 5.5, headMinHeight: 11 }
    );
  }
  if (pk.length) {
    ctx.beginSubsection_('PK Facility Performance', PK_GREEN);
    ctx.drawAutoTable_(
      pdfTableHead(facilitySummaryColLabels(true)),
      pk.map(facilityBundleSummaryRow_),
      PK_GREEN,
      colStyles,
      { fontSize: 7, cellPadding: 2, gapAfter: 3, headFontSize: 5.5, headMinHeight: 11 }
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
  const gap = opts.gap != null ? opts.gap : PDF_LAYOUT.cardGap;
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

  ctx.setY(y0 + blockH + (opts.gapAfter == null ? 2 : opts.gapAfter));
}

function sectionSummaryConfig_(id, stats, data, year) {
  const s = stats || {};
  const grvCounts = grvStatusCounts_(data.grv || []);
  const grvOpen = grvCounts.open;
  const grvClosed = grvCounts.closed;
  const grvInvalid = grvCounts.invalid;
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
      desc: (s.grievances || 0) + ' in period · ' + grvOpen + ' open · ' + grvClosed + ' closed'
        + (grvInvalid ? ' · ' + grvInvalid + ' invalid' : ''),
      metrics: [
        { label: 'Total', value: String(s.grievances || 0) },
        { label: 'Open', value: String(grvOpen), hot: grvOpen > 0 },
        { label: 'Closed', value: String(grvClosed) },
        { label: 'Invalid', value: String(grvInvalid) },
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

function drawSectionSummaryBlock_(ctx, sectionId, cfg) {
  if (!cfg) return;
  const cols = Math.min(4, cfg.metrics.length);
  const cardRows = Math.ceil(cfg.metrics.length / cols);
  const cardH = 17;
  const gridH = cardRows * cardH + Math.max(0, cardRows - 1) * PDF_LAYOUT.cardGap;
  const doc = ctx.doc;
  const descLines = doc.splitTextToSize(String(cfg.desc || ''), ctx.cW);
  const descH = descLines.length * PDF_LAYOUT.descLineH;
  const blockNeed = 10 + PDF_LAYOUT.afterSectionBar + PDF_LAYOUT.sectionDescGap
    + descH + PDF_LAYOUT.descToCards + gridH + PDF_LAYOUT.cardsToContent;
  ctx.ensureBlockFits_(blockNeed);
  ctx.beginSection_(cfg.num + ' · ' + cfg.title, cfg.accent);
  const descY = ctx.getY() + PDF_LAYOUT.sectionDescGap;
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(7.5);
  doc.setTextColor.apply(doc, INK_MUTED);
  doc.text(descLines, ctx.mL, descY);
  ctx.setY(descY + descH + PDF_LAYOUT.descToCards);
  drawMetricCardGrid_(ctx, cfg.metrics, {
    cols: cols,
    accent: cfg.accent,
    cardH: cardH,
    gap: PDF_LAYOUT.cardGap,
    gapAfter: PDF_LAYOUT.cardsToContent,
  });
}

function drawSummaryReportBody_(ctx, data, sections, stats, year) {
  const bodySections = sections.filter(function(id) { return id !== 'kpi'; });
  bodySections.forEach(function(id) {
    const cfg = sectionSummaryConfig_(id, stats, data, year);
    if (!cfg) return;
    drawSectionSummaryBlock_(ctx, id, cfg);
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
  doc.save(monthlyReportFileName_(variant, year, month));
}

function finalizePdf_(ctx) {
  ctx.pruneBlankPages_();
  drawFooters_(ctx.doc, ctx.pageW, ctx.pageH, ctx.mL, ctx.mR, ctx.mFoot, ctx.exportedAt);
  return ctx.doc;
}

function buildSummaryPdfDoc_(jsPDFLib, opts) {
  const data = opts.data || {};
  const stats = data.stats || {};
  const sections = (opts.sections && opts.sections.length) ? opts.sections : DEFAULT_SECTIONS.slice();
  const ctx = createPdfContext_(jsPDFLib, { pdfMode: 'summary' });
  const headerMeta = applyReportHeaderMeta_(ctx, opts.year, opts.month);
  ctx.reportTitle = 'Monthly Report — Summary';
  ctx.reportSubtitle = '';

  const headerEnd = ctx.drawMainHeader_();
  ctx.setY(headerEnd + PDF_LAYOUT.bodyGap);

  if (sections.indexOf('kpi') !== -1) {
    docSetSubhead_(ctx, 'Overview');
    drawMetricCardGrid_(ctx, websiteKpiItems_(stats), {
      cols: 4,
      cardH: 18,
      gap: PDF_LAYOUT.cardGap,
      gapAfter: 0,
    });
    ctx.setY(ctx.getY() + PDF_LAYOUT.sectionGap);
  }

  drawSummaryReportBody_(ctx, data, sections, stats, headerMeta.dataYear || opts.year);
  return finalizePdf_(ctx);
}

function docSetSubhead_(ctx, text) {
  const doc = ctx.doc;
  ctx.ensureSpace_(10);
  const textY = ctx.getY() + 1;
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(10);
  doc.setTextColor.apply(doc, INK);
  doc.text(text, ctx.mL, textY);
  ctx.setY(textY + 5);
}

function buildDetailPdfDoc_(jsPDFLib, opts) {
  const data = opts.data || {};
  const sections = (opts.sections && opts.sections.length) ? opts.sections : DEFAULT_SECTIONS.slice();
  const detailSections = sections.filter(function(id) { return id !== 'kpi'; });
  const ctx = createPdfContext_(jsPDFLib, { orientation: 'landscape', pdfMode: 'detail' });
  const headerMeta = applyReportHeaderMeta_(ctx, opts.year, opts.month);

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
    drawDetailReportBody_(ctx, data, detailSections, headerMeta.dataYear || opts.year);
  }
  return finalizePdf_(ctx);
}

function drawCompactHeaderForExport_(ctx) {
  const doc = ctx.doc;
  const bandH = PDF_LAYOUT.headerDetailH;
  doc.setFillColor.apply(doc, BRAND);
  doc.rect(0, 0, ctx.pageW, bandH, 'F');
  doc.setTextColor.apply(doc, WHITE);
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(11);
  doc.text('Monthly Report — Detail', ctx.mL, 9);
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(6.5);
  doc.text(ctx.periodText || '', ctx.mL, 14.5);
  doc.text(ctx.dataPeriodText || '', ctx.mL, 18.5);
  doc.text(ctx.cutoffText || '', ctx.mL, 22.5);
  if (ctx.exportedAt) {
    doc.text('Generated on ' + ctx.exportedAt, ctx.pageW - ctx.mR, 22.5, { align: 'right' });
  }
  ctx.markContent_(bandH);
  return bandH + PDF_LAYOUT.bodyGap;
}

/** @deprecated Use buildMonthlyReportPdfPair_ */
export async function buildMonthlyReportPdf_(opts) {
  return buildMonthlyReportPdfPair_(opts);
}

export async function buildMonthlyReportPdfPair_(opts) {
  const jsPDFLib = opts.getJsPDF();
  if (!jsPDFLib) throw new Error('PDF library is not ready.');

  const summaryDoc = buildSummaryPdfDoc_(jsPDFLib, opts);
  const detailDoc = buildDetailPdfDoc_(jsPDFLib, opts);

  saveMonthlyReportPdf_(summaryDoc, 'summary', opts.year, opts.month);
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
