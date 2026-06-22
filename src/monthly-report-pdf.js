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
  mrdSortEmptyMillItems_,
  mrdFormatNblRisers_,
  mrdReportHeaderMeta_,
  grvGroupName_,
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

/** Vertical rhythm — compact; sections flow on same page when space allows. */
const PDF_LAYOUT = {
  bodyGap: 4,
  sectionGap: 5,
  afterSectionBar: 1.5,
  subsectionGap: 3,
  sectionDescGap: 1,
  descToCards: 2,
  descLineH: 3.2,
  cardsToContent: 2,
  cardGap: 2.5,
  headerMainH: 36,
  headerDetailH: 26,
  headerCompactH: 19,
  compactBodyGap: 3,
  autoTableTopMargin: 22,
  sectionMinRemain: 42,
};

const DEFAULT_SECTIONS = ['kpi', 'sdd', 'highRisk', 'mill', 'trace', 'grv', 'nbl', 'facility', 'eudr'];
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

/** PDF Mill Onboarding export — HIGH RISK only (matches website Result Risk Level). */
function isHighRiskMillPdf_(item) {
  const risk = String(item && item.risk || '').trim().toUpperCase();
  if (risk === 'HIGH') return true;
  const r = (item && item.row) || {};
  const direct = String(r['RESULT RISK LEVEL'] || r['RISK LEVEL'] || r['Risk Level'] || '').trim().toUpperCase();
  return direct === 'HIGH';
}

function pdfMillOnboardingRows_(data) {
  return mrdSortMillItems_((data.mills || []).filter(isHighRiskMillPdf_));
}

/** Summary PDF mill table includes Result Risk Level (HIGH). */
const MRD_MILL_SUMMARY_PDF_COLS = MRD_MILL_SUMMARY_COLS;

function pdfMillRiskCell_(item) {
  const risk = String(item && item.risk || '').trim();
  if (risk) return pdfSanitize(risk);
  const r = (item && item.row) || {};
  return pdfSanitize(r['RESULT RISK LEVEL'] || r['RISK LEVEL'] || 'HIGH');
}

/** Skip 02A when 02 Mill Onboarding is exported — both lists are identical (HIGH RISK only). */
export function mrdPdfSectionsNoDupHighRisk_(sections) {
  const hasMill = sections.indexOf('mill') !== -1;
  if (!hasMill) return sections;
  return sections.filter(function(id) { return id !== 'highRisk'; });
}

function drawUntraceableMillsTable_(ctx, emptyMills) {
  const rows = mrdSortEmptyMillItems_(emptyMills || []);
  const count = rows.length;
  ctx.beginSection_('03 · Untraceable Mills (' + count + ')', NBL_RED);
  const doc = ctx.doc;
  const note = count
    ? count + ' mill' + (count === 1 ? '' : 's')
      + ' without FFB supplier data in TTP — full list below (count shown in Summary Overview).'
    : 'All mills have supplier traceability data for this period.';
  const noteLines = doc.splitTextToSize(note, ctx.cW);
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(7);
  doc.setTextColor.apply(doc, INK_MUTED);
  doc.text(noteLines, ctx.mL, ctx.getY() + 2);
  ctx.setY(ctx.getY() + noteLines.length * PDF_LAYOUT.descLineH + 4);
  if (!count) return;

  const w = colWidths_([22, 26, 26, 14], ctx.cW);
  ctx.drawFlowingTableChunks_(
    pdfTableHead(['Group Name', 'Company Name', 'Mill Name', 'Province']),
    rows.map(function(item) {
      const r = item.millRow || item.row || item;
      return [
        pdfSanitize(r['GROUP NAME']),
        pdfSanitize(r['COMPANY NAME']),
        pdfSanitize(r['MILL NAME']),
        pdfSanitize(r['PROVINCE']),
      ];
    }),
    NBL_RED,
    {
      0: { cellWidth: w[0] }, 1: { cellWidth: w[1] },
      2: { cellWidth: w[2] }, 3: { cellWidth: w[3] },
    },
    {
      fontSize: 6.5,
      cellPadding: 1.6,
      headFontSize: 5.5,
      headMinHeight: 10,
      chunkSize: 15,
    }
  );
}

function drawTraceSection_(ctx, data, year, full, noHeader) {
  const totals = data.traceTotals || {};
  const cardH = 18;
  if (full) {
    drawUntraceableMillsTable_(ctx, data.emptyMills);
    const traceBlockNeed = PDF_LAYOUT.subsectionGap + 11 + PDF_LAYOUT.afterSectionBar + cardH + 2;
    ctx.ensureBlockFits_(traceBlockNeed);
    ctx.beginSubsection_('Traceability Data · ' + pdfSanitize(year), TRACE_ORANGE);
  } else if (!noHeader) {
    ctx.beginSection_('03 · Traceability Data ' + pdfSanitize(year), TRACE_ORANGE);
  }
  drawMetricCardGrid_(ctx, traceMetricCards_(totals), {
    cols: 4,
    cardH: cardH,
    gapAfter: 0,
    keepWithHeader: full,
  });
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
  const protectedPages = new Set();

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

  function summaryTopMargin_() {
    return 28;
  }

  function afterPageBreak_() {
    if (ctx.pdfMode === 'summary') {
      y = summaryTopMargin_();
    } else {
      y = drawCompactHeader_();
    }
    markContent_(y);
    return y;
  }

  function autoTableTopMargin_() {
    if (ctx.pdfMode === 'summary') return summaryTopMargin_();
    return compactHeaderBandH_() + PDF_LAYOUT.compactBodyGap + 2;
  }

  function compactHeaderBandH_() {
    return ctx.isLandscape ? PDF_LAYOUT.headerDetailH : PDF_LAYOUT.headerCompactH;
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
      if (ctx.exportedAt) {
        doc.text('Generated on ' + ctx.exportedAt, mL, 19);
      }
    } else {
      doc.setFontSize(9.5);
      doc.text('Monthly Report', mL, 6.5);
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(6.5);
      doc.text(ctx.periodText || '', mL, 11.5);
      if (ctx.exportedAt) {
        doc.text('Generated on ' + ctx.exportedAt, mL, 15);
      }
    }
    return bandH + PDF_LAYOUT.compactBodyGap;
  }

  function ensureSpace(needed) {
    if (y + needed <= pageH - mFoot - 6) return;
    doc.addPage();
    afterPageBreak_();
  }

  function remainingSpace_() {
    return pageH - mFoot - 6 - y;
  }

  function ensureBlockFits_(needed) {
    const extra = sectionStarted ? PDF_LAYOUT.sectionGap : 0;
    if (remainingSpace_() < needed + extra) {
      doc.addPage();
      afterPageBreak_();
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
    if (sectionStarted && remainingSpace_() < PDF_LAYOUT.sectionMinRemain) {
      doc.addPage();
      afterPageBreak_();
    } else if (sectionStarted) {
      y += PDF_LAYOUT.sectionGap;
    }
    sectionStarted = true;
    drawSectionBar_(title, accent);
  }

  function beginSubsection_(title, accent) {
    if (remainingSpace_() < 28) {
      doc.addPage();
      afterPageBreak_();
    } else {
      y += PDF_LAYOUT.subsectionGap;
    }
    drawSectionBar_(title, accent, 8.5);
  }

  function beginFacilityBlock_(title, accent, newPage) {
    if (newPage) {
      doc.addPage();
      afterPageBreak_();
    } else if (sectionStarted && remainingSpace_() < 50) {
      doc.addPage();
      afterPageBreak_();
    } else if (sectionStarted) {
      y += PDF_LAYOUT.subsectionGap;
    }
    sectionStarted = true;
    drawSectionBar_(title, accent);
  }

  function drawAutoTable_(head, body, accent, colStyles, opts) {
    opts = opts || {};
    if (!body.length) return;

    resetAutoTableState_();

    const minStartH = 20;
    const headH = 9;
    const minRowH = 7;
    if (y + headH + minRowH > pageH - mFoot - 8) {
      doc.addPage();
      afterPageBreak_();
    }

    const startPage = doc.internal.getNumberOfPages();
    const tableTop = y;
    const continuationTop = autoTableTopMargin_();
    const base = {
      theme: 'grid',
      margin: { left: mL, right: mR, bottom: mFoot + 8, top: continuationTop },
      tableWidth: cW,
      styles: {
        fontSize: opts.fontSize || 7,
        cellPadding: opts.cellPadding != null ? opts.cellPadding : 1.8,
        textColor: INK,
        lineColor: BORDER,
        lineWidth: 0.12,
        overflow: 'linebreak',
        valign: 'middle',
        minCellHeight: 4,
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
      bodyStyles: { valign: 'middle' },
      alternateRowStyles: { fillColor: [255, 253, 253] },
      rowPageBreak: opts.rowPageBreak || 'auto',
      showHead: opts.showHead || (opts.showHeadEveryPage ? 'everyPage' : 'firstPage'),
      pageBreak: 'auto',
    };
    doc.autoTable(Object.assign({}, base, {
      head: head,
      body: body,
      startY: tableTop,
      columnStyles: colStyles || {},
      willDrawPage: function(hookData) {
        if (hookData.pageNumber > startPage) {
          doc.setPage(hookData.pageNumber);
          if (ctx.pdfMode !== 'summary') drawCompactHeader_();
          if (hookData.settings && hookData.settings.margin) {
            hookData.settings.margin.top = continuationTop;
          }
        }
      },
      didDrawPage: function(data) {
        const p = data.pageNumber;
        protectedPages.add(p);
        const cursor = data.cursor || {};
        if (cursor.y != null) {
          const top = p === startPage ? tableTop : continuationTop;
          markPageSpan_(p, top, cursor.y + 6);
        }
      },
    }));

    const table = doc.lastAutoTable;
    if (table && table.finalY) {
      const endPage = doc.internal.getNumberOfPages();
      protectedPages.add(endPage);
      markPageSpan_(endPage, endPage === startPage ? tableTop : continuationTop, table.finalY);
      y = table.finalY + (opts.gapAfter == null ? 2 : opts.gapAfter);
      markContent_(y);
    }
  }

  /** One table = one column header. Large tables split without re-printing the header. */
  function drawTableComplete_(head, allBody, accent, colStyles, opts) {
    opts = opts || {};
    if (!allBody.length) return;
    const maxSingle = opts.maxRowsPerCall != null ? opts.maxRowsPerCall : 200;
    if (allBody.length <= maxSingle) {
      drawAutoTable_(head, allBody, accent, colStyles, Object.assign({}, opts, {
        showHead: 'firstPage',
        rowPageBreak: 'auto',
      }));
      return;
    }
    const chunkSize = opts.chunkSize || 36;
    let idx = 0;
    while (idx < allBody.length) {
      resetAutoTableState_();
      if (idx > 0) newPage_();
      const end = Math.min(idx + chunkSize, allBody.length);
      drawAutoTable_(head, allBody.slice(idx, end), accent, colStyles, Object.assign({}, opts, {
        showHead: idx === 0 ? 'firstPage' : 'never',
        rowPageBreak: 'auto',
        gapAfter: end >= allBody.length ? (opts.gapAfter == null ? 1 : opts.gapAfter) : 0,
      }));
      idx = end;
    }
  }

  /** @deprecated alias */
  function drawAutoTableAllRows_(head, body, accent, colStyles, opts) {
    drawTableComplete_(head, body, accent, colStyles, opts);
  }

  function drawKvBlock_(heading, pairs, accent) {
    if (!pairs.length) return;
    if (remainingSpace_() < 18) {
      doc.addPage();
      afterPageBreak_();
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
    doc.text(ctx.periodText || '', mL, 19);
    if (ctx.exportedAt) {
      doc.setFontSize(6.5);
      doc.text('Generated on ' + ctx.exportedAt, mL, 24.5);
    }
    markContent_(bandH);
    return bandH;
  }

  function startBodyAfterCover_(coverEndY) {
    if (coverEndY + 28 <= pageH - mFoot - 6) {
      y = coverEndY + 4;
    } else {
      doc.addPage();
      afterPageBreak_();
    }
    sectionStarted = false;
  }

  function resetAutoTableState_() {
    try { delete doc.lastAutoTable; } catch (_) { /* ignore */ }
  }

  /** Each numbered report section starts on its own page. */
  function beginSectionPage_() {
    doc.addPage();
    if (ctx.pdfMode === 'summary') {
      y = 28;
    } else {
      y = drawCompactHeader_();
    }
    sectionStarted = false;
    markContent_(y);
    return y;
  }

  function pruneBlankPages_() {
    /* Keep all pages — deleting continuation pages was dropping EUDR / table rows. */
  }

  function newPage_() {
    doc.addPage();
    afterPageBreak_();
    return y;
  }

  /** Large tables only (>100 rows). Smaller tables: one continuous autoTable call. */
  function drawPagedAutoTable_(head, body, accent, colStyles, opts) {
    opts = opts || {};
    if (!body.length) return;
    const maxSingle = opts.maxSingleRows != null ? opts.maxSingleRows : 100;
    if (maxSingle > 0 && body.length <= maxSingle) {
      drawAutoTable_(head, body, accent, colStyles, Object.assign({}, opts, {
        showHead: 'firstPage',
        rowPageBreak: 'auto',
      }));
      return;
    }
    const chunkSize = opts.chunkSize || 36;
    let idx = 0;
    while (idx < body.length) {
      if (idx > 0) newPage_();
      const end = Math.min(idx + chunkSize, body.length);
      const isLast = end >= body.length;
      drawAutoTable_(head, body.slice(idx, end), accent, colStyles, Object.assign({}, opts, {
        showHead: idx === 0 ? 'firstPage' : 'never',
        rowPageBreak: 'auto',
        gapAfter: isLast ? (opts.gapAfter == null ? 2 : opts.gapAfter) : 0,
      }));
      idx = end;
    }
  }

  /** Chunk large tables; continue on same page when space allows (no forced blank pages). */
  function drawFlowingTableChunks_(head, allBody, accent, colStyles, opts) {
    opts = opts || {};
    if (!allBody.length) return;
    const chunkSize = opts.chunkSize || 15;
    const minRowBlock = 14;
    let idx = 0;
    while (idx < allBody.length) {
      resetAutoTableState_();
      let showHead = 'never';
      if (idx === 0) {
        showHead = 'firstPage';
      } else if (y + minRowBlock > pageH - mFoot - 8) {
        doc.addPage();
        afterPageBreak_();
        showHead = 'firstPage';
      }
      const end = Math.min(idx + chunkSize, allBody.length);
      const isLast = end >= allBody.length;
      drawAutoTable_(head, allBody.slice(idx, end), accent, colStyles, Object.assign({}, opts, {
        showHead: showHead,
        rowPageBreak: 'auto',
        gapAfter: isLast ? (opts.gapAfter == null ? 2 : opts.gapAfter) : 0,
      }));
      idx = end;
    }
  }

  ctx.beginSection_ = beginSection_;
  ctx.beginSubsection_ = beginSubsection_;
  ctx.beginFacilityBlock_ = beginFacilityBlock_;
  ctx.drawAutoTable_ = drawAutoTable_;
  ctx.drawAutoTableAllRows_ = drawAutoTableAllRows_;
  ctx.drawTableComplete_ = drawTableComplete_;
  ctx.drawKvBlock_ = drawKvBlock_;
  ctx.drawMainHeader_ = drawMainHeader_;
  ctx.startBodyAfterCover_ = startBodyAfterCover_;
  ctx.beginSectionPage_ = beginSectionPage_;
  ctx.pruneBlankPages_ = pruneBlankPages_;
  ctx.newPage_ = newPage_;
  ctx.drawPagedAutoTable_ = drawPagedAutoTable_;
  ctx.drawFlowingTableChunks_ = drawFlowingTableChunks_;
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
  highRisk: 'High Risk Suppliers',
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
    certification: pdfSanitize(r['CERTIFICATION']),
    grievances: pdfSanitize(r['TOTAL GRIEVANCES']),
    facilityCpo: pdfSanitize(r['FACILITY NAME CPO']),
    facilityPk: pdfSanitize(r['FACILITY NAME PK']),
    nblBy: pdfSanitize(item.nblBy),
  };
}

function drawMillSection_(ctx, data, full, noHeader) {
  const mills = pdfMillOnboardingRows_(data);
  if (!mills.length) return;
  const doc = ctx.doc;

  if (!noHeader) {
    ctx.beginSection_('02 · Mill Onboarding', BRAND);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(7);
    doc.setTextColor.apply(doc, INK_MUTED);
    doc.text(String(mills.length) + ' high-risk mills', ctx.mL, ctx.getY() + 1);
    ctx.setY(ctx.getY() + 6);
    ctx.markContent_(ctx.getY());
  }

  if (full) {
    const w = colWidths_([9, 13, 16, 15, 9, 7, 10, 9, 7, 14, 14], ctx.cW);
    const body = mills.map(function(item) {
      const r = item.row;
      const d = millDetailCells_(item);
      return [
        pdfMillRiskCell_(item),
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
    });
    ctx.drawTableComplete_(
      pdfTableHead(MRD_MILL_FULL_COLS),
      body,
      BRAND,
      {
        0: { cellWidth: w[0], fontStyle: 'bold', textColor: NBL_RED },
        1: { cellWidth: w[1] }, 2: { cellWidth: w[2] }, 3: { cellWidth: w[3] },
        4: { cellWidth: w[4] }, 5: { cellWidth: w[5] },
        6: { cellWidth: w[6], fontSize: 6.5 }, 7: { cellWidth: w[7], fontSize: 6.5 },
        8: { cellWidth: w[8] }, 9: { cellWidth: w[9], fontSize: 6.5 }, 10: { cellWidth: w[10], fontSize: 6.5 },
      },
      { fontSize: 6.5, cellPadding: 1.6, headFontSize: 5.6, headMinHeight: 10 }
    );
  } else {
    const w = colWidths_([10, 15, 18, 18, 11, 10], ctx.cW);
    ctx.drawTableComplete_(
      pdfTableHead(MRD_MILL_SUMMARY_PDF_COLS),
      mills.map(function(item) {
        const r = item.row;
        return [
          pdfMillRiskCell_(item),
          pdfSanitize(r['GROUP NAME']),
          pdfSanitize(r['COMPANY NAME']),
          pdfSanitize(r['MILL NAME']),
          pdfSanitize(r['PROVINCE']),
          pdfSanitize(isNblYes_(item.nbl) ? 'Yes' : item.nbl),
        ];
      }),
      BRAND,
      {
        0: { cellWidth: w[0], fontStyle: 'bold', textColor: NBL_RED },
        1: { cellWidth: w[1] }, 2: { cellWidth: w[2] }, 3: { cellWidth: w[3] },
        4: { cellWidth: w[4] }, 5: { cellWidth: w[5] },
      },
      { headFontSize: 5.6, headMinHeight: 10 }
    );
  }
}

function drawHighRiskSection_(ctx, data, full, noHeader) {
  const rows = mrdSortMillItems_(data.highRiskMills || []);
  if (!rows.length) return;
  const stats = data.stats || {};
  const doc = ctx.doc;

  if (!noHeader) {
    ctx.beginSection_('High Risk Suppliers', NBL_RED);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(8.5);
    doc.setTextColor.apply(doc, NBL_RED);
    doc.text(String(rows.length) + ' mills · Result Risk Level = HIGH', ctx.mL, ctx.getY() + 3);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(7);
    doc.setTextColor.apply(doc, INK_MUTED);
    doc.text('Separate from Mill Onboarding — ' + (stats.totalMills || 0) + ' total mills in period', ctx.mL, ctx.getY() + 8);
    ctx.setY(ctx.getY() + 12);
    ctx.markContent_(ctx.getY());
  }

  if (full) {
    const w = colWidths_([10, 14, 18, 18, 12, 8, 10, 10], ctx.cW);
    const body = rows.map(function(item) {
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
      ];
    });
    ctx.drawTableComplete_(
      pdfTableHead(['Result Risk Level', 'Group Name', 'Company Name', 'Mill Name', 'Province', 'No Buy List', 'Supplier Status', 'Certification']),
      body,
      NBL_RED,
      {
        0: { cellWidth: w[0], fontStyle: 'bold', textColor: NBL_RED },
        1: { cellWidth: w[1] }, 2: { cellWidth: w[2] }, 3: { cellWidth: w[3] },
        4: { cellWidth: w[4] }, 5: { cellWidth: w[5] },
        6: { cellWidth: w[6], fontSize: 6.5 }, 7: { cellWidth: w[7], fontSize: 6.5 },
      },
      { fontSize: 7, cellPadding: 1.8, headFontSize: 6, headMinHeight: 10 }
    );
  } else {
    const w = colWidths_([12, 18, 22, 22, 14], ctx.cW);
    const body = rows.map(function(item) {
      const r = item.row;
      return [
        pdfSanitize(item.risk),
        pdfSanitize(r['GROUP NAME']),
        pdfSanitize(r['COMPANY NAME']),
        pdfSanitize(r['MILL NAME']),
        pdfSanitize(r['PROVINCE']),
      ];
    });
    ctx.drawTableComplete_(
      pdfTableHead(['Result Risk Level', 'Group Name', 'Company Name', 'Mill Name', 'Province']),
      body,
      NBL_RED,
      {
        0: { cellWidth: w[0], fontStyle: 'bold', textColor: NBL_RED },
        1: { cellWidth: w[1] }, 2: { cellWidth: w[2] }, 3: { cellWidth: w[3] }, 4: { cellWidth: w[4] },
      },
      { headFontSize: 5.6, headMinHeight: 10 }
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
  drawTraceSection_(ctx, { traceTotals: totals, traceRows: [] }, year, false, noHeader);
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
          pdfSanitize(grvGroupName_(r)),
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
        pdfSanitize(grvGroupName_(r)),
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
    { fontSize: 6.5, cellPadding: 2, headFontSize: 5.4, headMinHeight: 12, rowPageBreak: 'auto' }
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
  const body = rows.map(function(item) {
    const r = item.row;
    return [
      pdfSanitize(r['GROUP NAME']), pdfSanitize(r['COMPANY NAME']), pdfSanitize(r['MILL NAME']),
      pdfSanitize(r['PROVINCE']), pdfSanitize(r['SUPPLY TO']),
      'Potential',
    ];
  });
  ctx.drawTableComplete_(
    pdfTableHead(MRD_EUDR_COLS),
    body,
    EUDR_TEAL,
    {
      0: { cellWidth: w[0] }, 1: { cellWidth: w[1] }, 2: { cellWidth: w[2] },
      3: { cellWidth: w[3] }, 4: { cellWidth: w[4] },
      5: { cellWidth: w[5], fontSize: 6.5 },
    },
    { fontSize: 6.5, cellPadding: 1.6, headFontSize: 5.6, headMinHeight: 10 }
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
  const untrace = s.emptyTraceMills != null ? s.emptyTraceMills : 0;
  return [
    { label: 'SDD Requested', value: String(s.sddRequested != null ? s.sddRequested : (s.sddTotal || 0)), sub: (s.sddDone != null ? s.sddDone : (s.sddSubmitted || 0)) + ' done' },
    { label: 'Total Mills', value: String(s.totalMills != null ? s.totalMills : 0), sub: (s.totalGroups || 0) + ' groups' },
    {
      label: 'Untraceable Mills',
      value: String(untrace),
      sub: 'mills without supplier data',
      subHint: untrace > 0 ? 'Full list in Detail report' : '',
      hot: untrace > 0,
      cardH: untrace > 0 ? 22 : undefined,
    },
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
  const defaultCardH = opts.cardH || 20;
  const cardW = (cW - gap * (cols - 1)) / cols;
  const rowHeights = [];
  items.forEach(function(item, i) {
    const row = Math.floor(i / cols);
    const h = item.cardH || defaultCardH;
    rowHeights[row] = Math.max(rowHeights[row] || 0, h);
  });
  const blockH = rowHeights.reduce(function(sum, h, ri) {
    return sum + h + (ri > 0 ? gap : 0);
  }, 0);
  if (!opts.keepWithHeader) ctx.ensureSpace_(blockH + 2);
  const y0 = ctx.getY();
  const rowTops = [];
  let accY = y0;
  rowHeights.forEach(function(h, ri) {
    rowTops[ri] = accY;
    accY += h + gap;
  });

  items.forEach(function(item, i) {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const x = mL + col * (cardW + gap);
    const cy = rowTops[row];
    const cardH = item.cardH || defaultCardH;
    const accent = item.hot ? NBL_RED : (item.accent || opts.accent || BRAND);
    const valueColor = item.hot ? NBL_RED : (item.valueColor || accent || INK);
    const hasHint = !!item.subHint;

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
      doc.setFontSize(5.6);
      doc.setTextColor.apply(doc, INK_LIGHT);
      doc.text(String(item.sub), x + cardW / 2, cy + (hasHint ? 16 : 16.5), { align: 'center' });
    }
    if (item.subHint) {
      doc.setFont('helvetica', 'italic');
      doc.setFontSize(5.4);
      doc.setTextColor.apply(doc, item.hot ? [168, 72, 62] : INK_MUTED);
      doc.text(String(item.subHint), x + cardW / 2, cy + 19.5, { align: 'center' });
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
      desc: (s.highRisk || 0) + ' high-risk mills',
      metrics: [
        { label: 'High Risk Mills', value: String(s.highRisk || 0), hot: (s.highRisk || 0) > 0 },
        { label: 'Total Mills', value: String(s.totalMills || 0), sub: (s.totalGroups || 0) + ' groups' },
        { label: 'Active NBL Mills', value: String(s.nblMills || 0) },
      ],
    },
    highRisk: {
      num: '02A', title: 'High Risk Suppliers', accent: NBL_RED,
      desc: (s.highRisk || 0) + ' mills · Result Risk Level = HIGH',
      metrics: [
        { label: 'High Risk', value: String(s.highRisk || 0), hot: true },
        { label: 'Total Mills', value: String(s.totalMills || 0) },
        { label: 'Active NBL', value: String(s.nblMills || 0) },
      ],
    },
    trace: {
      num: '03', title: 'Traceability Data ' + pdfSanitize(year), accent: TRACE_ORANGE,
      desc: 'TTM (mill coordinates) · TTP (supplier traceability) · '
        + (s.emptyTraceMills || 0) + ' untraceable mill' + ((s.emptyTraceMills || 0) === 1 ? '' : 's')
        + ' — full list in Detail report',
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
  const cardH = 15;
  const gridH = cardRows * cardH + Math.max(0, cardRows - 1) * PDF_LAYOUT.cardGap;
  const doc = ctx.doc;
  const descLines = doc.splitTextToSize(String(cfg.desc || ''), ctx.cW);
  const descH = descLines.length * PDF_LAYOUT.descLineH;
  const blockNeed = 8 + PDF_LAYOUT.afterSectionBar + PDF_LAYOUT.sectionDescGap
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
  const bodySections = mrdPdfSectionsNoDupHighRisk_(sections.filter(function(id) { return id !== 'kpi'; }));
  bodySections.forEach(function(id) {
    const cfg = sectionSummaryConfig_(id, stats, data, year);
    if (!cfg) return;
    drawSectionSummaryBlock_(ctx, id, cfg);
    if (id === 'sdd') drawSddSection_(ctx, data.sdd || [], true);
    else if (id === 'highRisk') drawHighRiskSection_(ctx, data, false, true);
    else if (id === 'mill') drawMillSection_(ctx, data, false, true);
    else if (id === 'trace') { /* totals shown in metric cards above */ }
    else if (id === 'grv') drawGrvSection_(ctx, data.grv || [], false, true);
    else if (id === 'nbl') drawNblSummaryList_(ctx, data.nblAll || []);
    else if (id === 'facility') drawFacilitySummaryFromBundles_(ctx, data.facilityBundles || []);
    else if (id === 'eudr') drawEudrSection_(ctx, data.eudrPotential || [], true);
  });
}

function drawDetailReportBody_(ctx, data, sections, year) {
  const ordered = mrdPdfSectionsNoDupHighRisk_(['sdd', 'highRisk', 'mill', 'trace', 'grv', 'nbl', 'facility', 'eudr'])
    .filter(function(id) { return sections.indexOf(id) !== -1; });
  ordered.forEach(function(id) {
    if (id === 'sdd') drawSddSection_(ctx, data.sdd || []);
    else if (id === 'highRisk') drawHighRiskSection_(ctx, data, true);
    else if (id === 'mill') drawMillSection_(ctx, data, true);
    else if (id === 'trace') drawTraceSection_(ctx, data, year, true);
    else if (id === 'grv') drawGrvSection_(ctx, data.grv || [], true);
    else if (id === 'nbl') drawNblSection_(ctx, data.nblAll || []);
    else if (id === 'facility') drawFacilitySection_(ctx, data.facilityBundles || [], true);
    else if (id === 'eudr') drawEudrSection_(ctx, data.eudrPotential || []);
  });
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
      cardH: 16,
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
  if (ctx.exportedAt) {
    doc.text('Generated on ' + ctx.exportedAt, ctx.mL, 19);
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
  { id: 'highRisk', label: MRD_SECTION_LABELS.highRisk },
  { id: 'mill', label: MRD_SECTION_LABELS.mill },
  { id: 'trace', label: MRD_SECTION_LABELS.trace },
  { id: 'grv', label: MRD_SECTION_LABELS.grv },
  { id: 'nbl', label: MRD_SECTION_LABELS.nbl },
  { id: 'facility', label: MRD_SECTION_LABELS.facility },
  { id: 'eudr', label: MRD_SECTION_LABELS.eudr },
];
