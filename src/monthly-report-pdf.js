/**
 * Monthly Report (Detail) — multi-section PDF export.
 */

const BRAND = [139, 26, 26];
const PK_GREEN = [13, 110, 70];
const INK = [26, 10, 10];
const INK_MUTED = [110, 96, 96];
const INK_LIGHT = [140, 120, 120];
const WHITE = [255, 255, 255];
const BG_SOFT = [252, 250, 250];
const BORDER = [230, 220, 220];

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

function tableBase_(doc, mL, mR, cW, mFoot) {
  return {
    theme: 'grid',
    margin: { left: mL, right: mR, bottom: mFoot + 4 },
    tableWidth: cW,
    styles: {
      fontSize: 7.5,
      cellPadding: 2.5,
      textColor: INK,
      lineColor: BORDER,
      lineWidth: 0.15,
      overflow: 'linebreak',
      valign: 'middle',
    },
    headStyles: {
      fillColor: BRAND,
      textColor: WHITE,
      fontStyle: 'bold',
      fontSize: 7.5,
      halign: 'center',
    },
    alternateRowStyles: { fillColor: [255, 253, 253] },
  };
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
    doc.text('Sustainability Dashboard · Monthly Report (Detail)', mL, pageH - 5);
    doc.text('Page ' + p + ' of ' + total, pageW - mR, pageH - 5, { align: 'right' });
  }
}

function createPdfContext_(jsPDFLib) {
  const doc = new jsPDFLib({ orientation: 'portrait', unit: 'mm', format: 'a4' });
  const pageW = doc.internal.pageSize.getWidth();
  const pageH = doc.internal.pageSize.getHeight();
  const mL = 12;
  const mR = 12;
  const mFoot = 12;
  const cW = pageW - mL - mR;
  let y = 0;
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
  };

  function drawCompactHeader_() {
    doc.setFillColor.apply(doc, BRAND);
    doc.rect(0, 0, pageW, 16, 'F');
    doc.setTextColor.apply(doc, WHITE);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(10);
    doc.text('Monthly Report (Detail)', mL, 7);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(7);
    doc.text(ctx.periodText + '   ·   Generated: ' + ctx.generatedAt, mL, 12.5);
    return 20;
  }

  function ensureSpace(needed) {
    if (y + needed <= pageH - mFoot - 4) return;
    doc.addPage();
    y = drawCompactHeader_();
  }

  function drawMainHeader_() {
    doc.setFillColor.apply(doc, BRAND);
    doc.rect(0, 0, pageW, 28, 'F');
    doc.setFillColor(180, 40, 40);
    doc.rect(0, 26.5, pageW, 1.5, 'F');
    doc.setTextColor.apply(doc, WHITE);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(16);
    doc.text('Monthly Report (Detail)', mL, 12);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(8.5);
    doc.text('Compliance snapshot — SDD, Mill, Traceability, Grievance, NBL, Facility, EUDR', mL, 19);
    doc.setFontSize(8);
    doc.text(ctx.periodText + '   ·   Generated: ' + ctx.generatedAt, mL, 25);
    return 34;
  }

  function drawSectionTitle_(title, accent) {
    ensureSpace(14);
    doc.setFillColor.apply(doc, accent || BRAND);
    doc.roundedRect(mL, y, cW, 7, 1.2, 1.2, 'F');
    doc.setTextColor.apply(doc, WHITE);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(8.5);
    doc.text(title, mL + 3, y + 4.8);
    y += 10;
  }

  function drawAutoTable_(head, body, accent, colStyles) {
    if (!body.length) {
      ensureSpace(8);
      doc.setFont('helvetica', 'italic');
      doc.setFontSize(8);
      doc.setTextColor.apply(doc, INK_MUTED);
      doc.text('No data for this section.', mL, y + 2);
      y += 8;
      return;
    }
    ensureSpace(16);
    const base = tableBase_(doc, mL, mR, cW, mFoot);
    if (accent) base.headStyles.fillColor = accent;
    doc.autoTable(Object.assign({}, base, {
      head: head,
      body: body,
      startY: y,
      columnStyles: colStyles || {},
    }));
    y = doc.lastAutoTable.finalY + 5;
  }

  function drawKpiRow_(stats) {
    ensureSpace(22);
    const items = [
      ['SDD Submitted', stats.sddSubmitted],
      ['SDD Draft', stats.sddDraft],
      ['Total Mills', stats.totalMills],
      ['Groups', stats.totalGroups],
      ['High Risk', stats.highRisk],
      ['NBL Mills', stats.nblMills],
      ['Empty Trace', stats.emptyTraceMills],
      ['Grievances', stats.grievances],
      ['NBL Entries', stats.nblEntries],
      ['EUDR Potential', stats.eudrPotential],
    ];
    drawAutoTable_([['Metric', 'Value']], items.map(function(it) { return [it[0], String(it[1])]; }));
  }

  function drawFacilityBlock_(bundle, type) {
    const isPk = type === 'pk';
    const accent = isPk ? PK_GREEN : BRAND;
    const pctLabel = isPk ? '% PK Traceable' : '% Traceable';
    const sum = bundle.summary || {};
    const facilityPct = isPk ? (sum.avgPk || '—') : (sum.avgCpo || '—');

    drawSectionTitle_((isPk ? 'PK' : 'CPO') + ' · ' + pdfSanitize(bundle.facility), accent);

    ensureSpace(10);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(7.5);
    doc.setTextColor.apply(doc, INK_MUTED);
    doc.text(
      pctLabel + ': ' + pdfSanitize(facilityPct)
        + '   ·   NBL: ' + pdfSanitize(sum.nblYes > 0 ? sum.nblYes + ' Yes' : '0')
        + '   ·   High Risk: ' + pdfSanitize(sum.highRisk || 0)
        + '   ·   Grievance: ' + pdfSanitize(sum.grievanceSum || 0)
        + '   ·   ISPO: ' + pdfSanitize(sum.ispoPct || '—'),
      mL,
      y + 2
    );
    y += 6;

    const profiles = bundle.profiles || [];
    if (profiles.length) {
      profiles.forEach(function(p) {
        drawAutoTable_(
          [['Facility Profile', '']],
          [
            ['Company', pdfSanitize(p._cplCompany)],
            ['Site', pdfSanitize(p._cplSite)],
            ['Address', pdfSanitize(p._cplAddress)],
            ['Capacity', pdfSanitize(p._cplCapacity)],
            ['Coordinate', pdfSanitize(p._cplCoordinate)],
            ['Facility type', pdfSanitize(p._cplFacility)],
            ['Certification', pdfSanitize(p._cplCert)],
          ],
          accent,
          {
            0: { cellWidth: 36, fontStyle: 'bold', fillColor: BG_SOFT, textColor: INK_MUTED },
            1: { cellWidth: cW - 36 },
          }
        );
      });
    }

    const companies = bundle.companies || [];
    drawAutoTable_(
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
          !isNaN(c.ttpPctNum) ? (Math.round(c.ttpPctNum * 10) / 10) + '%' : '—',
        ];
      }),
      accent
    );
  }

  ctx.drawMainHeader_ = drawMainHeader_;
  ctx.drawSectionTitle_ = drawSectionTitle_;
  ctx.drawAutoTable_ = drawAutoTable_;
  ctx.drawKpiRow_ = drawKpiRow_;
  ctx.drawFacilityBlock_ = drawFacilityBlock_;
  ctx.setY = function(v) { y = v; };
  ctx.getY = function() { return y; };

  return ctx;
}

/**
 * @param {object} opts
 * @param {Function} opts.getJsPDF
 * @param {object} opts.data — prepared export payload
 * @param {string} opts.year
 * @param {string} opts.month
 */
export async function buildMonthlyReportPdf_(opts) {
  const jsPDFLib = opts.getJsPDF();
  if (!jsPDFLib) throw new Error('Library PDF belum siap.');

  const data = opts.data || {};
  const ctx = createPdfContext_(jsPDFLib);
  ctx.periodText = 'Period: ' + periodLabel(opts.year, opts.month);
  ctx.generatedAt = new Date().toLocaleString('id-ID');

  ctx.setY(ctx.drawMainHeader_());
  ctx.drawKpiRow_(data.stats || {});

  ctx.drawSectionTitle_('01 · Supplier Due Diligence');
  ctx.drawAutoTable_(
    [['Status', 'Type', 'Group', 'Mill', 'Updated']],
    (data.sdd || []).map(function(r) {
      return [
        pdfSanitize(r['SCR - Screening Status']),
        pdfSanitize(r.supplier_type || r['Supplier Type']),
        pdfSanitize(r['Group Name'] || r['Grup Name']),
        pdfSanitize(r['Mill Name']),
        pdfSanitize(r.updated_at || r['SCR - Last Updated']).slice(0, 10),
      ];
    })
  );

  ctx.drawSectionTitle_('02 · Mill Onboarding');
  const millBody = [];
  (data.mills || []).forEach(function(item) {
    const r = item.row;
    millBody.push([
      pdfSanitize(r['GROUP NAME']),
      pdfSanitize(r['COMPANY NAME']),
      pdfSanitize(r['MILL NAME']),
      pdfSanitize(item.risk),
      pdfSanitize(isNblYes_(item.nbl) ? 'Yes' : item.nbl),
      pdfSanitize(r['PROVINCE']),
      pdfSanitize(r['SUPPLIER STATUS']),
      pdfSanitize(r['CERTIFICATION']),
      pdfSanitize(r['TOTAL GRIEVANCES']),
      pdfSanitize(r['FACILITY NAME CPO']),
      pdfSanitize(r['FACILITY NAME PK']),
      pdfSanitize(item.nblBy),
    ]);
    if (item.nblMatches && item.nblMatches.length) {
      item.nblMatches.forEach(function(m) {
        millBody.push([
          '', '', 'NBL source: ' + pdfSanitize(m.source), pdfSanitize(m.target), 'via ' + pdfSanitize(m.by),
          '', '', '', '', '', '', '',
        ]);
      });
    }
  });
  ctx.drawAutoTable_(
    [['Group', 'Company', 'Mill', 'Risk', 'NBL', 'Province', 'Supplier Status', 'Certification', 'Grievances', 'Facility CPO', 'Facility PK', 'NBL By']],
    millBody
  );

  if (data.highRiskMills && data.highRiskMills.length) {
    ctx.drawSectionTitle_('02b · High Risk Suppliers', BRAND);
    ctx.drawAutoTable_(
      [['Group', 'Company', 'Mill', 'Risk', 'Province']],
      data.highRiskMills.map(function(item) {
        const r = item.row;
        return [
          pdfSanitize(r['GROUP NAME']), pdfSanitize(r['COMPANY NAME']), pdfSanitize(r['MILL NAME']),
          pdfSanitize(item.risk), pdfSanitize(r['PROVINCE']),
        ];
      })
    );
  }

  ctx.drawSectionTitle_('03 · Traceability — Mills Without Supplier (' + pdfSanitize(opts.year) + ')', [230, 81, 0]);
  ctx.drawAutoTable_(
    [['Group', 'Company', 'Mill', 'Status']],
    (data.emptyMills || []).map(function(item) {
      const r = item.millRow;
      return [pdfSanitize(r['GROUP NAME']), pdfSanitize(r['COMPANY NAME']), pdfSanitize(r['MILL NAME']), 'Kosong'];
    })
  );

  ctx.drawSectionTitle_('04 · Grievance Monitoring', [106, 27, 154]);
  const grvBody = [];
  (data.grv || []).forEach(function(item) {
    const r = item.row;
    grvBody.push([
      pdfSanitize(r['Grievance ID']), pdfSanitize(r['Date Received']), pdfSanitize(r['Complainant']),
      pdfSanitize(r['Grievance Category']), pdfSanitize(r['Subject']), pdfSanitize(r['Risk Classification']),
      pdfSanitize(r['Grievance Status']),
    ]);
    [
      ['Description', r['Grievance Description']],
      ['Verification', r['Verification Findings']],
      ['Corrective Action', r['Corrective Action']],
      ['Preventive Action', r['Preventive Action']],
      ['Action Taken', r['Action Taken']],
      ['Date Closed', r['Date Closed']],
    ].forEach(function(pair) {
      if (pdfSanitize(pair[1]) !== '—') {
        grvBody.push(['', pair[0], pdfSanitize(pair[1]), '', '', '', '']);
      }
    });
  });
  ctx.drawAutoTable_(
    [['ID', 'Date', 'Complainant', 'Category', 'Subject', 'Risk', 'Status']],
    grvBody
  );

  ctx.drawSectionTitle_('05 · No Buy List', [192, 57, 43]);
  ctx.drawAutoTable_(
    [['Source', 'Riser', 'Group', 'Company']],
    (data.nblAll || []).map(function(r) {
      return [pdfSanitize(r.source), pdfSanitize(r.riser), pdfSanitize(r.group), pdfSanitize(r.company)];
    })
  );

  ctx.drawSectionTitle_('06 · Facility Performance', [46, 125, 50]);
  (data.facilityBundles || []).forEach(function(bundle) {
    ctx.drawFacilityBlock_(bundle, bundle.type);
  });

  ctx.drawSectionTitle_('07 · EUDR Potential', [0, 131, 143]);
  ctx.drawAutoTable_(
    [['Status', 'Group', 'Company', 'Mill', 'Province', 'Supply To']],
    (data.eudrPotential || []).map(function(item) {
      const r = item.row;
      return [
        'Potential',
        pdfSanitize(r['GROUP NAME']), pdfSanitize(r['COMPANY NAME']), pdfSanitize(r['MILL NAME']),
        pdfSanitize(r['PROVINCE']), pdfSanitize(r['SUPPLY TO']),
      ];
    })
  );

  drawFooters_(ctx.doc, ctx.pageW, ctx.pageH, ctx.mL, ctx.mR, ctx.mFoot);

  const fileLabel = (opts.year || 'all') + (opts.month ? '_' + monthLabel(opts.month) : '');
  const fname = 'Monthly-Report-' + fileLabel.replace(/\s+/g, '-') + '.pdf';
  ctx.doc.save(fname);
}

function isNblYes_(val) {
  return /yes|nbl|no buy/i.test(String(val || ''));
}
