/**
 * Monthly Report — Executive PDF (charts mirror Summary PDF section metrics).
 */
import { mrdReportHeaderMeta_ } from './monthly-report-labels.js';
import {
  mrdBuildExecutiveChartBundle_,
  mrdOverviewKpiItems_,
} from './monthly-report-pdf.js';
import { classifyRiskBucket_ } from './mill-executive-report.js';
import { registerPdfFonts_, setPdfFont_, PDF_FONT_SANS } from './pdf-fonts.js';
import { getMillExecutiveBackgroundDataUrl_ } from './mill-executive-bg.js';

const CHART_CARD_FILLS = {
  sec_sdd: [252, 248, 246],
  sec_mill: [252, 248, 246],
  sec_trace: [246, 249, 253],
  sec_grv: [249, 246, 241],
  sec_nbl: [246, 249, 253],
  sec_facility: [246, 249, 253],
  sec_eudr: [252, 248, 246],
};

const PDF_CHART_JPEG_Q = 0.62;
const PIE_SIZE = { w: 380, h: 340 };
const BAR_WIDE = { w: 560, h: 300 };
const BAR_MED = { w: 420, h: 300 };

const PIE_COLORS = {
  Draft: '#D4A017',
  Submitted: '#2E7D32',
  High: '#C03030',
  Medium: '#D4A017',
  Low: '#2E7D32',
};

const CHART_TITLE_OVERRIDE = {
  sdd: 'SDD status',
  mill: 'Mill risk level',
};

const CHART_CAPTION_OVERRIDE = {
  sdd: '01 · Draft vs submitted screenings',
  mill: '02 · Result risk level mix',
  facility: '06 · NBL % per facility (unique companies supplying each facility)',
};

const METRIC_BAR_COLORS = [
  'rgba(139, 26, 26, 0.85)',
  'rgba(46, 125, 50, 0.85)',
  'rgba(230, 81, 0, 0.85)',
  'rgba(21, 101, 192, 0.85)',
  'rgba(0, 131, 143, 0.85)',
];

let chartInstances = {};

function yieldToBrowser_(ms) {
  return new Promise(function(resolve) {
    requestAnimationFrame(function() {
      setTimeout(resolve, ms == null ? 24 : ms);
    });
  });
}

export function destroyMrdExecutiveCharts_() {
  Object.keys(chartInstances).forEach(function(k) {
    try { chartInstances[k].destroy(); } catch (_) { /* noop */ }
  });
  chartInstances = {};
}

function parseMetricNumber_(raw) {
  const s = String(raw != null ? raw : '').replace('%', '').trim().replace(',', '.');
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function sectionChartKey_(id) {
  return 'sec_' + id;
}

function barOptions_(pctMode, canvas) {
  const w = canvas ? canvas.width : 560;
  const h = canvas ? canvas.height : 300;
  return {
    responsive: false,
    animation: false,
    width: w,
    height: h,
    plugins: { legend: { display: false } },
    scales: {
      x: {
        grid: { display: false },
        ticks: { font: { size: Math.round(h * 0.038), family: PDF_FONT_SANS }, maxRotation: 45, minRotation: 0 },
      },
      y: {
        beginAtZero: true,
        suggestedMax: pctMode ? 100 : undefined,
        ticks: {
          font: { size: Math.round(h * 0.036), family: PDF_FONT_SANS },
          callback: pctMode ? function(v) { return v + '%'; } : undefined,
        },
        grid: { color: 'rgba(139, 26, 26, 0.08)' },
      },
    },
  };
}

function hbarOptions_(canvas) {
  const w = canvas ? canvas.width : BAR_WIDE.w;
  const h = canvas ? canvas.height : BAR_WIDE.h;
  return {
    responsive: false,
    animation: false,
    width: w,
    height: h,
    indexAxis: 'y',
    plugins: { legend: { display: false } },
    scales: {
      y: {
        grid: { display: false },
        ticks: {
          font: { size: 11, family: PDF_FONT_SANS },
          autoSkip: false,
        },
      },
      x: {
        beginAtZero: true,
        suggestedMax: 100,
        ticks: {
          font: { size: 11, family: PDF_FONT_SANS },
          callback: function(v) { return v + '%'; },
        },
        grid: { color: 'rgba(139, 26, 26, 0.08)' },
      },
    },
  };
}

/** Same numbers/captions as Summary PDF — one bar chart per section. */
export function buildMrdExecutiveData_(snapshot, opts) {
  opts = opts || {};
  const data = snapshot || {};
  const year = opts.year != null ? opts.year : ((data.reportPeriod && data.reportPeriod.year) || '');
  const sections = opts.sections || [
    'kpi', 'sdd', 'highRisk', 'mill', 'trace', 'grv', 'nbl', 'facility', 'eudr',
  ];
  return {
    stats: data.stats || {},
    overview: mrdOverviewKpiItems_(data.stats),
    sections: mrdBuildExecutiveChartBundle_(data, { year: year, sections: sections }),
    mills: data.mills || [],
    facilityBundles: data.facilityBundles || [],
  };
}

export function mrdExecutiveChartSizes_(data) {
  const sizes = {};
  (data.sections || []).forEach(function(section) {
    const key = sectionChartKey_(section.id);
    if (section.id === 'sdd' || section.id === 'mill') {
      sizes[key] = [PIE_SIZE.w, PIE_SIZE.h];
    } else if (section.id === 'trace') {
      sizes[key] = [BAR_WIDE.w, BAR_WIDE.h];
    } else if (section.id === 'facility') {
      const fac = buildFacilityNblPctChart_(data.facilityBundles);
      const n = Math.max(fac.labels.length, 1);
      sizes[key] = [BAR_WIDE.w, Math.min(760, Math.max(320, n * 34 + 100))];
    } else {
      sizes[key] = [BAR_MED.w, BAR_MED.h];
    }
  });
  return sizes;
}

const PIE_PCT_PLUGIN = {
  id: 'mrdPiePct',
  afterDraw: function(chart) {
    if (chart.config.type !== 'pie' && chart.config.type !== 'doughnut') return;
    const ctx = chart.ctx;
    chart.data.datasets.forEach(function(dataset, di) {
      const meta = chart.getDatasetMeta(di);
      const total = dataset.data.reduce(function(a, b) { return a + b; }, 0);
      if (!total) return;
      meta.data.forEach(function(arc, idx) {
        const val = dataset.data[idx];
        if (!val) return;
        const pct = Math.round((val / total) * 100);
        if (pct < 4) return;
        const pos = arc.tooltipPosition();
        const fontSize = Math.max(11, Math.round(Math.min(chart.width, chart.height) * 0.048));
        ctx.save();
        ctx.font = 'bold ' + fontSize + 'px ' + PDF_FONT_SANS;
        ctx.fillStyle = '#ffffff';
        ctx.shadowColor = 'rgba(0,0,0,0.4)';
        ctx.shadowBlur = 2;
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText(pct + '%', pos.x, pos.y);
        ctx.restore();
      });
    });
  },
};

const PIE_CENTER_TEXT_PLUGIN = {
  id: 'mrdPieCenterText',
  afterDraw: function(chart) {
    const text = chart.config.options && chart.config.options.plugins
      && chart.config.options.plugins.mrdCenterText;
    if (!text) return;
    const ctx = chart.ctx;
    const area = chart.chartArea;
    if (!area) return;
    const cx = (area.left + area.right) / 2;
    const cy = (area.top + area.bottom) / 2;
    ctx.save();
    ctx.font = 'bold 22px ' + PDF_FONT_SANS;
    ctx.fillStyle = '#78716c';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText(text, cx, cy);
    ctx.restore();
  },
};

function pieOptions_(canvas, centerText) {
  const w = canvas ? canvas.width : PIE_SIZE.w;
  const h = canvas ? canvas.height : PIE_SIZE.h;
  return {
    responsive: false,
    animation: false,
    width: w,
    height: h,
    plugins: {
      legend: {
        position: 'bottom',
        labels: {
          boxWidth: 12,
          padding: 10,
          font: { size: Math.round(h * 0.034), family: PDF_FONT_SANS },
        },
      },
      mrdCenterText: centerText || '',
    },
  };
}

function renderPieBuckets_(Chart, key, canvas, buckets, colorMap, emptyCenterText) {
  if (!Chart || !canvas) return;
  const entries = Object.entries(buckets || {}).filter(function(e) {
    return e[1] > 0;
  });
  const total = entries.reduce(function(s, e) { return s + e[1]; }, 0);
  if (!total) {
    chartInstances[key] = new Chart(canvas, {
      type: 'doughnut',
      data: {
        labels: Object.keys(buckets || {}),
        datasets: [{
          data: Object.keys(buckets || {}).map(function() { return 1; }),
          backgroundColor: Object.keys(buckets || {}).map(function() { return '#e7e5e4'; }),
          borderWidth: 2,
          borderColor: '#fff',
        }],
      },
      options: pieOptions_(canvas, emptyCenterText || '0 / 0'),
      plugins: [PIE_CENTER_TEXT_PLUGIN, PIE_PCT_PLUGIN],
    });
    return;
  }
  chartInstances[key] = new Chart(canvas, {
    type: 'pie',
    data: {
      labels: entries.map(function(e) { return e[0]; }),
      datasets: [{
        data: entries.map(function(e) { return e[1]; }),
        backgroundColor: entries.map(function(e) { return colorMap[e[0]] || '#8B7355'; }),
        borderWidth: 2,
        borderColor: '#fff',
      }],
    },
    options: pieOptions_(canvas),
    plugins: [PIE_PCT_PLUGIN],
  });
}

function buildMillRiskBuckets_(mills) {
  const buckets = { High: 0, Medium: 0, Low: 0 };
  (mills || []).forEach(function(item) {
    const bucket = classifyRiskBucket_(item && item.risk);
    if (bucket === 'Other' || bucket === 'Unclassified') return;
    buckets[bucket] = (buckets[bucket] || 0) + 1;
  });
  return buckets;
}

function buildSddStatusBuckets_(stats) {
  const s = stats || {};
  return {
    Draft: s.sddDraft || 0,
    Submitted: s.sddSubmitted != null ? s.sddSubmitted : (s.sddDone || 0),
  };
}

function mrdFacilityIsNblYes_(val) {
  return /yes|nbl|no buy/i.test(String(val || ''));
}

function companyKey_(c) {
  return [
    c.group || c['GROUP NAME'] || '',
    c.company || c['COMPANY NAME'] || '',
  ].map(function(x) { return String(x || '').trim().toLowerCase(); }).join('|');
}

/** Per facility: NBL companies ÷ unique companies × 100. */
function buildFacilityNblPctChart_(bundles) {
  const rows = [];
  (bundles || []).forEach(function(bundle) {
    const name = String(bundle && bundle.facility || '').trim();
    if (!name) return;
    const companies = bundle.companies || [];
    if (!companies.length) return;
    const seen = new Set();
    let total = 0;
    let nbl = 0;
    companies.forEach(function(c) {
      const key = companyKey_(c);
      if (!key || key === '|') return;
      if (seen.has(key)) return;
      seen.add(key);
      total += 1;
      if (mrdFacilityIsNblYes_(c.nbl)) nbl += 1;
    });
    if (!total) return;
    const pct = Math.round((nbl / total) * 1000) / 10;
    const isPk = bundle.type === 'pk';
    let label = name;
    if (label.length > 48) label = label.slice(0, 46) + '…';
    rows.push({
      label: label,
      pct: pct,
      isPk: isPk,
      sortName: name.toLowerCase(),
    });
  });
  rows.sort(function(a, b) {
    return b.pct - a.pct || a.sortName.localeCompare(b.sortName);
  });
  return {
    labels: rows.map(function(r) { return r.label; }),
    values: rows.map(function(r) { return r.pct; }),
    colors: rows.map(function(r) {
      return r.isPk ? 'rgba(46, 125, 50, 0.82)' : 'rgba(139, 26, 26, 0.85)';
    }),
  };
}

function metricsToChart_(section) {
  const isPct = section && section.id === 'trace';
  const labels = [];
  const values = [];
  (section.metrics || []).forEach(function(m) {
    labels.push(String(m.label || '').trim());
    values.push(parseMetricNumber_(m.value));
  });
  return { labels: labels, values: values, isPct: isPct };
}

function renderSectionChart_(Chart, section, canvas, data) {
  if (!Chart || !canvas || !section) return;
  const key = sectionChartKey_(section.id);
  if (section.id === 'sdd') {
    const buckets = buildSddStatusBuckets_(data && data.stats);
    renderPieBuckets_(Chart, key, canvas, buckets, PIE_COLORS, '0 / 0');
    return;
  }
  if (section.id === 'mill') {
    const buckets = buildMillRiskBuckets_(data && data.mills);
    renderPieBuckets_(Chart, key, canvas, buckets, PIE_COLORS);
    return;
  }
  if (section.id === 'facility') {
    const fac = buildFacilityNblPctChart_(data && data.facilityBundles);
    if (!fac.labels.length) return;
    chartInstances[key] = new Chart(canvas, {
      type: 'bar',
      data: {
        labels: fac.labels,
        datasets: [{
          label: 'NBL %',
          data: fac.values,
          backgroundColor: fac.colors,
          borderRadius: 4,
        }],
      },
      options: hbarOptions_(canvas),
    });
    return;
  }
  const chartData = metricsToChart_(section);
  if (!chartData.labels.length) return;
  chartInstances[key] = new Chart(canvas, {
    type: 'bar',
    data: {
      labels: chartData.labels,
      datasets: [{
        data: chartData.values,
        backgroundColor: chartData.labels.map(function(_, i) {
          return METRIC_BAR_COLORS[i % METRIC_BAR_COLORS.length];
        }),
        borderRadius: 6,
      }],
    },
    options: barOptions_(chartData.isPct, canvas),
  });
}

function renderMrdExecutiveChartsBody_(Chart, data, els) {
  if (!Chart || !data || !els) return;
  (data.sections || []).forEach(function(section) {
    const key = sectionChartKey_(section.id);
    if (!els[key]) return;
    renderSectionChart_(Chart, section, els[key], data);
  });
}

export function renderMrdExecutiveCharts_(Chart, data, els) {
  destroyMrdExecutiveCharts_();
  renderMrdExecutiveChartsBody_(Chart, data, els);
}

/** One chart per tick — keeps export from freezing the tab. */
export async function renderMrdExecutiveChartsAsync_(Chart, data, els) {
  destroyMrdExecutiveCharts_();
  if (!Chart || !data || !els) return;
  const sections = data.sections || [];
  for (let i = 0; i < sections.length; i++) {
    const section = sections[i];
    const key = sectionChartKey_(section.id);
    if (!els[key]) continue;
    const partial = {};
    partial[key] = els[key];
    renderMrdExecutiveChartsBody_(Chart, data, partial);
    await yieldToBrowser_(40);
  }
}

function canvasToPdfChartImage_(canvas, cardFill) {
  const fill = cardFill || [246, 249, 253];
  if (!canvas || typeof canvas.toDataURL !== 'function') return '';
  const w = canvas.width;
  const h = canvas.height;
  const flat = document.createElement('canvas');
  flat.width = w;
  flat.height = h;
  const ctx = flat.getContext('2d');
  if (!ctx) return canvas.toDataURL('image/jpeg', PDF_CHART_JPEG_Q);
  ctx.fillStyle = 'rgb(' + fill[0] + ',' + fill[1] + ',' + fill[2] + ')';
  ctx.fillRect(0, 0, w, h);
  ctx.drawImage(canvas, 0, 0);
  return flat.toDataURL('image/jpeg', PDF_CHART_JPEG_Q);
}

export function collectMrdExecutiveChartImages_(els) {
  const out = {};
  if (!els) return out;
  Object.keys(els).forEach(function(key) {
    const canvas = els[key];
    if (!canvas || typeof canvas.toDataURL !== 'function') return;
    try {
      out[key] = canvasToPdfChartImage_(canvas, CHART_CARD_FILLS[key]);
    } catch (_) { /* noop */ }
  });
  return out;
}

export function mrdExecutiveFilename_(year, month) {
  const names = ['', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const m = parseInt(String(month || ''), 10);
  const label = m >= 1 && m <= 12 ? names[m] + ' ' + year : 'Full Year ' + year;
  return 'Monthly Report Executive - ' + label + '.pdf';
}

export function mrdExecutiveHeaderMeta_(year, month) {
  const meta = mrdReportHeaderMeta_(year, month);
  return {
    periodLine: meta.periodLine,
    dataPeriodLine: meta.dataPeriodLine,
    cutoffLine: meta.cutoffLine,
  };
}

function pdfDataUrlFormat_(dataUrl) {
  if (!dataUrl || typeof dataUrl !== 'string') return 'JPEG';
  if (dataUrl.indexOf('data:image/png') === 0) return 'PNG';
  return 'JPEG';
}

function buildMrdExecutiveSummaryLine_(data) {
  const captions = (data.sections || []).map(function(s) { return s.caption; }).filter(Boolean);
  if (captions.length) return captions.join('   ·   ');
  const st = data.stats || {};
  return [
    (st.sddRequested || 0) + ' SDD requested',
    (st.totalMills || 0) + ' mills',
    (st.eudrPotential || 0) + ' EUDR potential',
  ].join('   ·   ');
}

function sectionPanelCfg_(section, data) {
  const id = section.id;
  const wide = id === 'trace' || id === 'facility';
  const isPie = id === 'sdd' || id === 'mill';
  const titleOverride = CHART_TITLE_OVERRIDE[id];
  const capOverride = CHART_CAPTION_OVERRIDE[id];
  let w = isPie ? PIE_SIZE.w : (wide ? BAR_WIDE.w : BAR_MED.w);
  let h = isPie ? PIE_SIZE.h : (wide ? BAR_WIDE.h : BAR_MED.h);
  if (id === 'facility') {
    const fac = buildFacilityNblPctChart_(data && data.facilityBundles);
    const n = Math.max(fac.labels.length, 1);
    h = Math.min(760, Math.max(320, n * 34 + 100));
    w = BAR_WIDE.w;
  }
  return {
    key: sectionChartKey_(id),
    title: titleOverride || ((section.num ? section.num + ' · ' : '') + section.title),
    caption: capOverride || section.caption || '',
    w: w,
    h: h,
  };
}

export async function exportMrdExecutivePdf_(meta, data, chartImages, getJsPDF) {
  const JsPDF = typeof getJsPDF === 'function' ? getJsPDF() : null;
  if (!JsPDF) throw new Error('PDF library not loaded. Refresh the page and try again.');

  const doc = new JsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4', compress: true });
  await registerPdfFonts_(doc);

  const pw = doc.internal.pageSize.getWidth();
  const ph = doc.internal.pageSize.getHeight();
  const M = 14;
  const W = pw - M * 2;
  const stats = data.stats || {};
  const imgs = chartImages || {};
  const panels = (data.sections || []).map(function(section) {
    return sectionPanelCfg_(section, data);
  });

  const pageBg = await getMillExecutiveBackgroundDataUrl_();
  function paintPageBg_() {
    if (pageBg) {
      try {
        doc.addImage(pageBg, pdfDataUrlFormat_(pageBg), 0, 0, pw, ph, undefined, 'FAST');
      } catch (_) {
        doc.setFillColor(250, 247, 242);
        doc.rect(0, 0, pw, ph, 'F');
      }
    } else {
      doc.setFillColor(250, 247, 242);
      doc.rect(0, 0, pw, ph, 'F');
    }
  }
  paintPageBg_();

  const PAL = {
    hero: [252, 250, 246],
    mist: [246, 249, 253],
    blush: [252, 248, 246],
    sand: [249, 246, 241],
  };

  const CARD_ALPHA = 0.86;
  function withCardAlpha_(fn) {
    const canAlpha = typeof doc.setGState === 'function' && typeof doc.GState === 'function';
    if (canAlpha) {
      doc.saveGraphicsState();
      doc.setGState(new doc.GState({ opacity: CARD_ALPHA }));
    }
    fn();
    if (canAlpha) doc.restoreGraphicsState();
  }

  function softCard_(cx, cy, cw, ch, fill, radius) {
    radius = radius || 5;
    withCardAlpha_(function() {
      doc.setFillColor(Math.max(0, fill[0] - 14), Math.max(0, fill[1] - 14), Math.max(0, fill[2] - 14));
      doc.roundedRect(cx + 0.4, cy + 0.6, cw, ch, radius, radius, 'F');
      doc.setFillColor(fill[0], fill[1], fill[2]);
      doc.roundedRect(cx, cy, cw, ch, radius, radius, 'F');
    });
    doc.setDrawColor(255, 255, 255);
    doc.setLineWidth(0.35);
    doc.roundedRect(cx, cy, cw, ch, radius, radius, 'S');
  }

  const FOOTER_H = 5;
  const SUMMARY_H = 14;

  const HERO_Y = 6;
  const HERO_H = 32;
  const HERO_PAD = 6;
  const innerX = M + HERO_PAD;
  softCard_(M, HERO_Y, W, HERO_H, PAL.hero, 6);

  const heroInnerW = W - HERO_PAD * 2;
  const kpiAreaW = heroInnerW * 0.46;
  const kpiAreaX = innerX + heroInnerW - kpiAreaW;
  const kpiBoxGap = 2.5;
  const kpiBoxW = (kpiAreaW - kpiBoxGap * 3) / 4;
  const kpiBoxH = HERO_H - HERO_PAD * 2;
  const kpiBoxY = HERO_Y + HERO_PAD;

  setPdfFont_(doc, 'serif');
  doc.setFontSize(26);
  doc.setTextColor(15, 12, 10);
  doc.text('Monthly Report', innerX, HERO_Y + 12);

  setPdfFont_(doc, 'sans-bold');
  doc.setFontSize(9.5);
  doc.setTextColor(55, 48, 45);
  doc.text('Executive report  ·  ' + (meta.periodLine || '').replace(/^Reporting period:\s*/i, ''), innerX, HERO_Y + 19);

  setPdfFont_(doc, 'sans');
  doc.setFontSize(7);
  doc.setTextColor(120, 100, 100);
  const metaLine = [meta.dataPeriodLine, meta.cutoffLine].filter(Boolean).join('  ·  ');
  if (metaLine) {
    const lines = doc.splitTextToSize(metaLine, heroInnerW - kpiAreaW - 8);
    doc.text(lines.slice(0, 2), innerX, HERO_Y + 24);
  }

  const overview = data.overview || mrdOverviewKpiItems_(stats);
  overview.slice(0, 4).forEach(function(k, i) {
    const boxX = kpiAreaX + i * (kpiBoxW + kpiBoxGap);
    softCard_(boxX, kpiBoxY, kpiBoxW, kpiBoxH, [255, 255, 252], 4);
    setPdfFont_(doc, 'sans-bold');
    doc.setFontSize(16);
    doc.setTextColor(139, 26, 26);
    doc.text(String(k.value != null ? k.value : '—'), boxX + kpiBoxW / 2, kpiBoxY + kpiBoxH / 2 - 1, { align: 'center' });
    setPdfFont_(doc, 'sans');
    doc.setFontSize(6.5);
    doc.setTextColor(87, 83, 78);
    doc.text(String(k.label || ''), boxX + kpiBoxW / 2, kpiBoxY + kpiBoxH / 2 + 5, { align: 'center' });
  });

  const chartsY = HERO_Y + HERO_H + 4;
  const rowGap = 3.5;
  const TITLE_H = 9;
  const CAPTION_H = 5;
  const IMG_PAD = 3;
  const col2Gap = 4;
  const col2W = (W - col2Gap) / 2;
  const chartsMaxY = ph - FOOTER_H - SUMMARY_H - 6;

  async function chartPanel_(cfg, cx, cy, cw, ch) {
    softCard_(cx, cy, cw, ch, PAL.mist, 5);
    setPdfFont_(doc, 'sans-bold');
    doc.setFontSize(8);
    doc.setTextColor(41, 37, 36);
    doc.text(cfg.title, cx + 5, cy + 6);
    setPdfFont_(doc, 'sans');
    doc.setFontSize(6.2);
    doc.setTextColor(120, 100, 100);
    const cap = doc.splitTextToSize(cfg.caption || '', cw - 10);
    doc.text(cap.slice(0, 2), cx + 5, cy + 10.5);

    if (!imgs[cfg.key]) return;
    const bx = cx + IMG_PAD;
    const by = cy + TITLE_H + CAPTION_H;
    const bw = cw - IMG_PAD * 2;
    const bh = ch - TITLE_H - CAPTION_H - IMG_PAD;
    const ratio = cfg.w / cfg.h;
    const boxR = bw / bh;
    let dw; let dh; let dx; let dy;
    if (ratio > boxR) {
      dw = bw; dh = bw / ratio; dx = bx; dy = by + (bh - dh) / 2;
    } else {
      dh = bh; dw = bh * ratio; dy = by; dx = bx + (bw - dw) / 2;
    }
    doc.addImage(imgs[cfg.key], pdfDataUrlFormat_(imgs[cfg.key]), dx, dy, dw, dh, undefined, 'FAST');
    await yieldToBrowser_(8);
  }

  let panelIdx = 0;
  let pageStartY = chartsY;
  const rowH = Math.floor((chartsMaxY - chartsY - rowGap) / 2);

  while (panelIdx < panels.length) {
    if (panelIdx > 0 && panelIdx % 4 === 0) {
      doc.addPage();
      paintPageBg_();
      pageStartY = 12;
    }
    const localIdx = panelIdx % 4;
    const row = Math.floor(localIdx / 2);
    const col = localIdx % 2;
    const cy = pageStartY + row * (rowH + rowGap);
    const cx = M + col * (col2W + col2Gap);
    await chartPanel_(panels[panelIdx], cx, cy, col2W, rowH);
    panelIdx += 1;
  }

  const summaryY = ph - SUMMARY_H - 4;
  setPdfFont_(doc, 'serif');
  doc.setFontSize(9);
  doc.setTextColor(168, 162, 158);
  doc.text('Summary', pw / 2, summaryY, { align: 'center' });
  setPdfFont_(doc, 'sans-bold');
  doc.setFontSize(7.5);
  doc.setTextColor(41, 37, 36);
  const summaryText = buildMrdExecutiveSummaryLine_(data);
  const summaryLines = doc.splitTextToSize(summaryText, W);
  doc.text(summaryLines.slice(0, 3), pw / 2, summaryY + 5, { align: 'center' });

  setPdfFont_(doc, 'sans');
  doc.setFontSize(6);
  doc.setTextColor(168, 162, 158);
  doc.text('Same metrics as Monthly Report Summary PDF  ·  Main + Waste mills', pw / 2, ph - 3, { align: 'center' });

  doc.save(meta.filename);
}

/** @deprecated use mrdExecutiveChartSizes_(execData) */
export const MRD_EXEC_CHART_SIZES = {};
