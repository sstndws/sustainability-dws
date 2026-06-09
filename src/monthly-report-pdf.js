/**
 * Monthly Report (Detail) — PDF export without TOC; continuous layout, no blank pages.
 */

const BRAND = [139, 26, 26];
const PK_GREEN = [13, 110, 70];
const TRACE_ORANGE = [230, 81, 0];
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

function createPdfContext_(jsPDFLib) {
  const doc = new jsPDFLib({ orientation: 'portrait', unit: 'mm', format: 'a4' });
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
      y += 6;
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
    doc.text('Monthly Report', mL, 13);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(8.5);
    doc.text('Compliance snapshot — SDD, Mill, Traceability, Grievance, NBL, Facility, EUDR', mL, 20);
    doc.setFontSize(8);
    doc.text(ctx.periodText + '   ·   Generated: ' + ctx.generatedAt, mL, 26);
    markContent_(36);
    return 36;
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
    { label: 'SDD Submitted', value: String(s.sddSubmitted != null ? s.sddSubmitted : 0) },
    { label: 'SDD Draft', value: String(s.sddDraft != null ? s.sddDraft : 0) },
    { label: 'Total Mills', value: String(s.totalMills != null ? s.totalMills : 0) },
    { label: 'Groups', value: String(s.totalGroups != null ? s.totalGroups : 0) },
    { label: 'High Risk', value: String(s.highRisk != null ? s.highRisk : 0) },
    { label: 'NBL Mills', value: String(s.nblMills != null ? s.nblMills : 0) },
    { label: 'Empty Traceability', value: String(s.emptyTraceMills != null ? s.emptyTraceMills : 0) },
    { label: 'Grievances', value: String(s.grievances != null ? s.grievances : 0) },
    { label: 'NBL Entries', value: String(s.nblEntries != null ? s.nblEntries : 0) },
    { label: 'EUDR Potential', value: String(s.eudrPotential != null ? s.eudrPotential : 0) },
  ];
}

const MRD_SECTION_LABELS = {
  kpi: 'Overview · Key metrics',
  sdd: 'Supplier Due Diligence',
  mill: 'Mill Onboarding',
  trace: 'Traceability (empty mills)',
  grv: 'Grievance Monitoring',
  nbl: 'No Buy List',
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

function drawSddSection_(ctx, rows) {
  if (!rows.length) return;
  ctx.beginSection_('01 · Supplier Due Diligence', BRAND);
  const w = colWidths_([14, 12, 18, 22, 14], ctx.cW);
  ctx.drawAutoTable_(
    [['Status', 'Type', 'Group', 'Mill', 'Updated']],
    rows.map(function(r) {
      return [
        pdfSanitize(r['SCR - Screening Status']),
        pdfSanitize(r.supplier_type || r['Supplier Type']),
        pdfSanitize(r['Group Name'] || r['Grup Name']),
        pdfSanitize(r['Mill Name']),
        pdfSanitize(r.updated_at || r['SCR - Last Updated']).slice(0, 10),
      ];
    }),
    BRAND,
    {
      0: { cellWidth: w[0] }, 1: { cellWidth: w[1] }, 2: { cellWidth: w[2] },
      3: { cellWidth: w[3] }, 4: { cellWidth: w[4] },
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
      const w = colWidths_([12, 14, 14, 8, 7, 10, 10, 10, 8, 8, 9], ctx.cW);
      ctx.drawAutoTable_(
        [['Group', 'Company', 'Mill', 'Risk', 'NBL', 'Province', 'Supplier Status', 'Certification', 'Grievances', 'Facility CPO', 'Facility PK']],
        mills.map(function(item) {
          const r = item.row;
          const d = millDetailCells_(item);
          return [
            pdfSanitize(r['GROUP NAME']),
            pdfSanitize(r['COMPANY NAME']),
            pdfSanitize(r['MILL NAME']),
            pdfSanitize(item.risk),
            pdfSanitize(isNblYes_(item.nbl) ? 'Yes' : item.nbl),
            pdfSanitize(r['PROVINCE']),
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
      const w = colWidths_([16, 20, 20, 10, 10, 14], ctx.cW);
      ctx.drawAutoTable_(
        [['Group', 'Company', 'Mill', 'Risk', 'NBL', 'Province']],
        mills.map(function(item) {
          const r = item.row;
          return [
            pdfSanitize(r['GROUP NAME']),
            pdfSanitize(r['COMPANY NAME']),
            pdfSanitize(r['MILL NAME']),
            pdfSanitize(item.risk),
            pdfSanitize(isNblYes_(item.nbl) ? 'Yes' : item.nbl),
            pdfSanitize(r['PROVINCE']),
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
    const hw = colWidths_([18, 22, 22, 12, 14], ctx.cW);
    ctx.drawAutoTable_(
      [['Group', 'Company', 'Mill', 'Risk', 'Province']],
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

function drawTraceSection_(ctx, rows, year) {
  if (!rows.length) return;
  ctx.beginSection_('03 · Traceability — Mills Without Supplier (' + pdfSanitize(year) + ')', TRACE_ORANGE);
  const w = colWidths_([22, 26, 30, 12], ctx.cW);
  ctx.drawAutoTable_(
    [['Group', 'Company', 'Mill', 'Status']],
    rows.map(function(item) {
      const r = item.millRow;
      return [pdfSanitize(r['GROUP NAME']), pdfSanitize(r['COMPANY NAME']), pdfSanitize(r['MILL NAME']), 'Empty'];
    }),
    TRACE_ORANGE,
    { 0: { cellWidth: w[0] }, 1: { cellWidth: w[1] }, 2: { cellWidth: w[2] }, 3: { cellWidth: w[3] } }
  );
}

function drawGrvSection_(ctx, rows, full) {
  if (!rows.length) return;
  ctx.beginSection_('04 · Grievance Monitoring', GRV_PURPLE);
  if (!full) {
    const w = colWidths_([14, 14, 18, 16, 18, 10, 10], ctx.cW);
    ctx.drawAutoTable_(
      [['ID', 'Date', 'Complainant', 'Category', 'Subject', 'Risk', 'Status']],
      rows.map(function(item) {
        const r = item.row;
        return [
          pdfSanitize(r['Grievance ID']), pdfSanitize(r['Date Received']), pdfSanitize(r['Complainant']),
          pdfSanitize(r['Grievance Category']), pdfSanitize(r['Subject']), pdfSanitize(r['Risk Classification']),
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

  const w = colWidths_([10, 10, 12, 12, 14, 8, 8, 18, 14, 12, 12, 10], ctx.cW);
  ctx.drawAutoTable_(
    [['ID', 'Date', 'Complainant', 'Category', 'Subject', 'Risk', 'Status', 'Description', 'Verification', 'Corrective', 'Preventive', 'Closed']],
    rows.map(function(item) {
      const r = item.row;
      return [
        pdfSanitize(r['Grievance ID']),
        pdfSanitize(r['Date Received']),
        pdfSanitize(r['Complainant']),
        pdfSanitize(r['Grievance Category']),
        pdfSanitize(r['Subject']),
        pdfSanitize(r['Risk Classification']),
        pdfSanitize(r['Grievance Status']),
        pdfSanitize(r['Grievance Description']),
        pdfSanitize(r['Verification Findings']),
        pdfSanitize(r['Corrective Action']),
        pdfSanitize(r['Preventive Action']),
        pdfSanitize(r['Date Closed']),
      ];
    }),
    GRV_PURPLE,
    {
      0: { cellWidth: w[0] }, 1: { cellWidth: w[1] }, 2: { cellWidth: w[2] },
      3: { cellWidth: w[3] }, 4: { cellWidth: w[4] }, 5: { cellWidth: w[5] }, 6: { cellWidth: w[6] },
      7: { cellWidth: w[7], fontSize: 6.5 }, 8: { cellWidth: w[8], fontSize: 6.5 },
      9: { cellWidth: w[9], fontSize: 6.5 }, 10: { cellWidth: w[10], fontSize: 6.5 }, 11: { cellWidth: w[11] },
    },
    { fontSize: 7, cellPadding: 2.2 }
  );
}

function drawNblSection_(ctx, rows) {
  if (!rows.length) return;
  ctx.beginSection_('05 · No Buy List', NBL_RED);
  const w = colWidths_([18, 18, 24, 30], ctx.cW);
  ctx.drawAutoTable_(
    [['Source', 'Riser', 'Group', 'Company']],
    rows.map(function(r) {
      return [pdfSanitize(r.source), pdfSanitize(r.riser), pdfSanitize(r.group), pdfSanitize(r.company)];
    }),
    NBL_RED,
    { 0: { cellWidth: w[0] }, 1: { cellWidth: w[1] }, 2: { cellWidth: w[2] }, 3: { cellWidth: w[3] } }
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
  const cw = colWidths_([20, 14, 13, 13, 8, 8, 8, 16], ctx.cW);

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(8);
  doc.setTextColor.apply(doc, INK);
  doc.text('Company Performance Breakdown', ctx.mL, y0 + 2);
  ctx.setY(y0 + 5);

  ctx.drawAutoTable_(
    [['Company', 'Group', 'Certification', 'Coordinate', 'NBL', 'Risk', 'Grievance', pctLabel]],
    companies.map(function(c) {
      return [
        pdfSanitize(c.company),
        pdfSanitize(c.group),
        pdfSanitize(c.certification),
        pdfSanitize(c.coordinate),
        pdfSanitize(c.nbl),
        pdfSanitize(c.riskLevel),
        pdfSanitize(c.grievance),
        pctFmt_(c.ttpPctNum),
      ];
    }),
    accent,
    {
      0: { cellWidth: cw[0] }, 1: { cellWidth: cw[1] }, 2: { cellWidth: cw[2], fontSize: 6.8 },
      3: { cellWidth: cw[3], fontSize: 6.5 }, 4: { cellWidth: cw[4], halign: 'center' },
      5: { cellWidth: cw[5], halign: 'center' }, 6: { cellWidth: cw[6], halign: 'right' },
      7: { cellWidth: cw[7], halign: 'right', fontStyle: 'bold' },
    },
    { fontSize: 7.5, cellPadding: 2.2, gapAfter: 2 }
  );
}

function drawFacilityBlock_(ctx, bundle, full, idx) {
  const isPk = bundle.type === 'pk';
  const accent = isPk ? PK_GREEN : BRAND;
  const pctLabel = isPk ? '% PK Traceable' : '% Traceable';
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

function drawEudrSection_(ctx, rows) {
  if (!rows.length) return;
  ctx.beginSection_('07 · EUDR Potential', EUDR_TEAL);
  const w = colWidths_([12, 18, 20, 20, 14, 16], ctx.cW);
  ctx.drawAutoTable_(
    [['Status', 'Group', 'Company', 'Mill', 'Province', 'Supply To']],
    rows.map(function(item) {
      const r = item.row;
      return [
        'Potential',
        pdfSanitize(r['GROUP NAME']), pdfSanitize(r['COMPANY NAME']), pdfSanitize(r['MILL NAME']),
        pdfSanitize(r['PROVINCE']), pdfSanitize(r['SUPPLY TO']),
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

export async function buildMonthlyReportPdf_(opts) {
  const jsPDFLib = opts.getJsPDF();
  if (!jsPDFLib) throw new Error('PDF library is not ready.');

  const data = opts.data || {};
  const sections = (opts.sections && opts.sections.length) ? opts.sections : DEFAULT_SECTIONS.slice();
  const full = opts.detailLevel === 'full';
  const ctx = createPdfContext_(jsPDFLib);
  ctx.periodText = 'Period: ' + periodLabel(opts.year, opts.month);
  ctx.generatedAt = new Date().toLocaleString('en-GB', { hour12: false });
  ctx.detailLevel = full ? 'full' : 'summary';

  ctx.drawMainHeader_();
  const coverEnd = drawCoverPage_(ctx, data.stats || {}, sections, full);
  ctx.startBodyAfterCover_(coverEnd);

  if (sections.indexOf('sdd') !== -1) drawSddSection_(ctx, data.sdd || []);
  if (sections.indexOf('mill') !== -1) drawMillSection_(ctx, data, full);
  if (sections.indexOf('trace') !== -1) drawTraceSection_(ctx, data.emptyMills || [], opts.year);
  if (sections.indexOf('grv') !== -1) drawGrvSection_(ctx, data.grv || [], full);
  if (sections.indexOf('nbl') !== -1) drawNblSection_(ctx, data.nblAll || []);
  if (sections.indexOf('facility') !== -1) drawFacilitySection_(ctx, data.facilityBundles || [], full);
  if (sections.indexOf('eudr') !== -1) drawEudrSection_(ctx, data.eudrPotential || []);

  ctx.pruneBlankPages_();
  drawFooters_(ctx.doc, ctx.pageW, ctx.pageH, ctx.mL, ctx.mR, ctx.mFoot);

  const detailTag = full ? 'Full' : 'Summary';
  const fileLabel = (opts.year || 'all') + (opts.month ? '_' + monthLabel(opts.month) : '');
  const fname = 'Monthly-Report-' + detailTag + '-' + fileLabel.replace(/\s+/g, '-') + '.pdf';
  ctx.doc.save(fname);
}

export const MRD_PDF_SECTIONS = [
  { id: 'kpi', label: MRD_SECTION_LABELS.kpi },
  { id: 'sdd', label: MRD_SECTION_LABELS.sdd },
  { id: 'mill', label: MRD_SECTION_LABELS.mill },
  { id: 'trace', label: MRD_SECTION_LABELS.trace },
  { id: 'grv', label: MRD_SECTION_LABELS.grv },
  { id: 'nbl', label: MRD_SECTION_LABELS.nbl },
  { id: 'facility', label: MRD_SECTION_LABELS.facility },
  { id: 'eudr', label: MRD_SECTION_LABELS.eudr },
];
