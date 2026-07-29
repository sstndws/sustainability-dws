import './overview-landing.css';
import { BRAND_TAGLINE } from './brand.js';

let refreshBusy = false;
let welcomeName = 'there';
let overviewHubRoot = null;

const OVERVIEW_CACHE_KEY = 'sustain-dashboard.overviewMetrics.v1';

function currentPeriodKey_() {
  const now = new Date();
  const year = now.getFullYear();
  const quarter = Math.ceil((now.getMonth() + 1) / 3);
  return year + '-Q' + quarter;
}

function readOverviewCache_() {
  try {
    const raw = localStorage.getItem(OVERVIEW_CACHE_KEY);
    if (!raw) return null;
    const entry = JSON.parse(raw);
    if (!entry || !entry.periodKey || !entry.payload) return null;
    return entry;
  } catch (_e) {
    return null;
  }
}

function sanitizePayloadForCache_(payload) {
  if (!payload || typeof payload !== 'object') return payload;
  const out = Object.assign({}, payload);
  if (payload.snapshot && typeof payload.snapshot === 'object') {
    const snap = Object.assign({}, payload.snapshot);
    if (snap.groups && typeof snap.groups.size === 'number' && snap.groupCount == null) {
      snap.groupCount = snap.groups.size;
    }
    delete snap.groups;
    out.snapshot = snap;
  }
  return out;
}

function writeOverviewCache_(periodKey, payload) {
  try {
    localStorage.setItem(
      OVERVIEW_CACHE_KEY,
      JSON.stringify({
        periodKey: periodKey,
        savedAt: Date.now(),
        payload: sanitizePayloadForCache_(payload),
      })
    );
  } catch (e) {
    console.warn('[overview] cache write failed', e);
  }
}

function bindQuarterCacheWatch_() {
  if (bindQuarterCacheWatch_.bound) return;
  bindQuarterCacheWatch_.bound = true;
  document.addEventListener('visibilitychange', function() {
    if (document.visibilityState !== 'visible' || !overviewHubRoot) return;
    refreshHub_(overviewHubRoot, { force: false });
  });
}
bindQuarterCacheWatch_.bound = false;

const MODULES = [
  {
    id: 'ndpe',
    title: 'NDPE & Forest protection',
    desc: 'No deforestation, peat, or exploitation — mills screened against cut-off dates and buyer NDPE policies.',
    icon: 'forest',
    footerKey: 'millsScreened',
  },
  {
    id: 'trace',
    title: 'Traceability to source',
    desc: 'TTM links mills to coordinates; TTP tracks FFB to estates and dealers.',
    icon: 'trace',
    footerKey: 'traceGap',
  },
  {
    id: 'eudr',
    title: 'EUDR readiness',
    desc: 'Due diligence for deforestation-free EU placement — geolocation, risk assessment, and statements.',
    icon: 'eudr',
    footerKey: 'eudrReview',
  },
  {
    id: 'grievance',
    title: 'Grievance & remediation',
    desc: 'Cases logged, triaged, and closed with evidence. Trends inform supplier engagement.',
    icon: 'grievance',
    footerKey: 'grievanceOpen',
  },
  {
    id: 'nbl',
    title: 'No Buy List',
    desc: 'Suppliers or mills on the NBL stay blocked until remediation criteria are met.',
    icon: 'nbl',
    footerKey: 'nblBlocked',
  },
  {
    id: 'cert',
    title: 'Standards & certification',
    desc: 'RSPO, ISPO, and internal criteria tracked with onboarding and facility performance.',
    icon: 'cert',
    footerKey: 'certMills',
  },
  {
    id: 'supply',
    title: 'Responsible supply base',
    desc: 'CPO, PK, and selected waste streams monitored across refinery supply bases.',
    icon: 'supply',
    footerKey: 'suppliersActive',
  },
  {
    id: 'report',
    title: 'Monthly reporting',
    desc: 'Snapshots for mills, SDD, traceability, grievance, NBL, and facility KPIs.',
    icon: 'report',
    footerKey: 'nextReport',
  },
];

const ICONS = {
  forest: '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M12 22V12"/><path d="M7 12L12 3l5 9"/><path d="M5 12h14"/></svg>',
  trace: '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/></svg>',
  eudr: '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><circle cx="12" cy="12" r="10"/><path d="M2 12h20"/></svg>',
  grievance: '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/></svg>',
  nbl: '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><circle cx="12" cy="12" r="10"/><line x1="4.93" y1="4.93" x2="19.07" y2="19.07"/></svg>',
  cert: '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><circle cx="12" cy="8" r="5"/><path d="M8.5 14.5L7 22l5-2.5L17 22l-1.5-7.5"/></svg>',
  supply: '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><rect x="1" y="3" width="15" height="13"/><polygon points="16 8 20 8 23 11 23 16 16 16 16 8"/><circle cx="5.5" cy="18.5" r="2.5"/><circle cx="18.5" cy="18.5" r="2.5"/></svg>',
  report: '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>',
};

function formatWelcomeName(emailRaw) {
  const s = (emailRaw || '').trim();
  if (!s) return 'there';
  const local = s.split('@')[0] || s;
  return local
    .replace(/[._-]+/g, ' ')
    .split(' ')
    .filter(Boolean)
    .map(function(w) { return w.charAt(0).toUpperCase() + w.slice(1).toLowerCase(); })
    .join(' ');
}

function pct_(part, total) {
  const t = Number(total) || 0;
  if (!t) return 0;
  return Math.round((Number(part) / t) * 100);
}

function nextReportLabel_() {
  const d = new Date();
  const next = new Date(d.getFullYear(), d.getMonth() + 1, 1);
  return next.toLocaleDateString('en-GB', { month: 'short', day: 'numeric' });
}

function buildHubModel_(payload) {
  const snap = (payload && payload.snapshot) || {};
  const total = snap.totalMills || 0;
  const nbl = snap.nbl || 0;
  const nonNblPct = total ? Math.round(((total - nbl) / total) * 100) : 0;
  const high = snap.highRisk || 0;
  const ndpePct = Math.max(0, Math.min(100, nonNblPct || (100 - pct_(high, total))));

  let tracePct = 0;
  if (snap.ttpTrace && !isNaN(snap.ttpTrace.pct)) tracePct = Math.round(snap.ttpTrace.pct);
  else {
    const tr = snap.traceability || {};
    const trT = (tr.Traceable || 0) + (tr.Untraceable || 0);
    tracePct = trT ? Math.round((tr.Traceable / trT) * 100) : 0;
  }

  const grv = snap.grievanceProgress || {};
  const grvOpen = grv.open || 0;
  const grvResolved = grv.total ? Math.round(grv.pct || (grv.closed / grv.total) * 100) : 100;

  const eudrP = payload && payload.eudr ? payload.eudr.potential : 0;
  const eudrTotal = (payload && payload.eudr ? payload.eudr.potential + payload.eudr.notPotential : 0) || total;
  const eudrPct = eudrTotal ? Math.round((eudrP / eudrTotal) * 100) : 0;

  const certCount = Object.entries(snap.certification || {})
    .filter(function(e) { return e[0] !== 'None'; })
    .reduce(function(sum, e) { return sum + e[1]; }, 0);

  const supplierCount = Object.keys(snap.facilityQty || {}).filter(Boolean).length;

  const pills = {
    ndpe: ndpePct >= 85 ? { text: 'Compliant', tone: 'green' } : { text: 'Review', tone: 'amber' },
    trace: tracePct >= 80 ? { text: 'On track', tone: 'green' } : { text: 'Gap found', tone: 'amber' },
    eudr: eudrPct >= 50 ? { text: 'Potential', tone: 'blue' } : { text: 'In review', tone: 'blue' },
    grievance: grvOpen > 0 ? { text: grvOpen + ' open', tone: 'amber' } : { text: 'Clear', tone: 'green' },
    nbl: nbl > 0 ? { text: nbl + ' blocked', tone: 'amber' } : { text: 'Clear', tone: 'green' },
    cert: { text: 'On track', tone: 'green' },
    supply: { text: 'Monitored', tone: 'green' },
    report: { text: 'Ready', tone: 'blue' },
  };

  const footers = {
    millsScreened: 'Mills screened: ' + total,
    traceGap: 'Traceable supply: ' + tracePct + '%',
    eudrReview: 'Potential mills: ' + eudrP,
    grievanceOpen: 'Open cases: ' + grvOpen,
    nblBlocked: 'Re-entries: monitoring',
    certMills: 'Certified: ' + certCount + ' mills',
    suppliersActive: 'Suppliers: ' + supplierCount + ' active',
    nextReport: 'Next report: ' + nextReportLabel_(),
  };

  return {
    hero: {
      ndpePct: ndpePct,
      grievances: grvOpen,
      nbl: nbl,
    },
    progress: [
      { label: 'NDPE compliance', pct: ndpePct, tone: 'burgundy' },
      { label: 'Traceability coverage', pct: tracePct, tone: 'teal' },
      { label: 'EUDR readiness', pct: eudrPct, tone: 'blue' },
      { label: 'Grievance resolution', pct: grvResolved, tone: 'amber' },
    ],
    pills: pills,
    footers: footers,
    activity: buildActivity_(payload, tracePct, nbl, grvOpen),
  };
}

function buildActivity_(payload, tracePct, nbl, grvOpen) {
  const now = new Date();
  const items = [
    {
      tone: 'green',
      title: 'Overview refreshed',
      desc: 'Live metrics synced from mill registry and monitoring modules.',
      time: 'Just now',
    },
  ];
  if (tracePct > 0 && tracePct < 80) {
    items.push({
      tone: 'amber',
      title: 'Traceability gap flagged',
      desc: 'TTP traceability at ' + tracePct + '% — review Facility Performance.',
      time: 'Today',
    });
  }
  if (nbl > 0) {
    items.push({
      tone: 'red',
      title: 'NBL status updated',
      desc: nbl + ' supplier(s) on No Buy List require governance review.',
      time: 'Today',
    });
  }
  if (grvOpen > 0) {
    items.push({
      tone: 'blue',
      title: 'Grievance cases open',
      desc: grvOpen + ' active grievance case(s) in the current year.',
      time: 'This week',
    });
  }
  if (payload && payload.periodLabel) {
    items.push({
      tone: 'green',
      title: 'Reporting period',
      desc: 'Snapshot basis: ' + payload.periodLabel,
      time: now.getFullYear().toString(),
    });
  }
  return items.slice(0, 4);
}

function renderHub_(root, model) {
  const welcome = root.querySelector('#ovHubWelcome');
  if (welcome) welcome.textContent = 'Welcome back, ' + welcomeName;

  const heroStats = root.querySelector('#ovHubHeroStats');
  if (heroStats && model.hero) {
    heroStats.innerHTML = [
      { v: model.hero.ndpePct + '%', l: 'NDPE Compliance' },
      { v: String(model.hero.grievances), l: 'Active Grievances' },
      { v: String(model.hero.nbl), l: 'NBL Suppliers' },
    ].map(function(s) {
      return '<div class="ov-hub__glass"><div class="ov-hub__glass-val">' + s.v + '</div><div class="ov-hub__glass-lbl">' + s.l + '</div></div>';
    }).join('');
  }

  const grid = root.querySelector('#ovHubModules');
  if (grid) {
    grid.textContent = '';
    MODULES.forEach(function(m) {
      const pill = model.pills[m.id] || { text: '—', tone: 'blue' };
      const foot = model.footers[m.footerKey] || '';
      const art = document.createElement('article');
      art.className = 'ov-hub__mod';
      art.innerHTML =
        '<div class="ov-hub__mod-head">'
        + '<span class="ov-hub__mod-icon ov-hub__mod-icon--' + m.icon + '">' + (ICONS[m.icon] || '') + '</span>'
        + '<span class="ov-hub__pill ov-hub__pill--' + pill.tone + '">' + pill.text + '</span>'
        + '</div>'
        + '<h3 class="ov-hub__mod-title">' + m.title + '</h3>'
        + '<p class="ov-hub__mod-desc">' + m.desc + '</p>'
        + '<footer class="ov-hub__mod-foot"><span>' + foot + '</span><span class="ov-hub__mod-arrow" aria-hidden="true">→</span></footer>';
      grid.appendChild(art);
    });
  }

  const prog = root.querySelector('#ovHubProgress');
  if (prog) {
    prog.innerHTML = model.progress.map(function(p) {
      return '<div class="ov-hub__prog-row">'
        + '<div class="ov-hub__prog-meta"><span>' + p.label + '</span><strong>' + p.pct + '%</strong></div>'
        + '<div class="ov-hub__prog-track"><div class="ov-hub__prog-fill ov-hub__prog-fill--' + p.tone + '" style="width:' + p.pct + '%"></div></div>'
        + '</div>';
    }).join('');
  }

  const act = root.querySelector('#ovHubActivity');
  if (act) {
    act.innerHTML = model.activity.map(function(a) {
      return '<li class="ov-hub__act-item">'
        + '<span class="ov-hub__act-dot ov-hub__act-dot--' + a.tone + '"></span>'
        + '<div><div class="ov-hub__act-title">' + a.title + '</div>'
        + '<div class="ov-hub__act-desc">' + a.desc + '</div></div>'
        + '<time class="ov-hub__act-time">' + a.time + '</time></li>';
    }).join('');
  }
}

async function refreshHub_(root, options) {
  if (refreshBusy) return;
  const force = !!(options && options.force);
  const periodKey = currentPeriodKey_();
  const btn = root.querySelector('#ovHubRefresh');

  if (!force) {
    const cached = readOverviewCache_();
    if (cached && cached.periodKey === periodKey && cached.payload) {
      renderHub_(root, buildHubModel_(cached.payload));
      return;
    }
  }

  refreshBusy = true;
  if (btn) btn.disabled = true;
  try {
    const fetcher = typeof window.refreshOverviewMetricsData_ === 'function'
      ? window.refreshOverviewMetricsData_
      : null;
    if (!fetcher) {
      const cached = readOverviewCache_();
      const payload = cached && cached.payload ? cached.payload : { snapshot: {} };
      renderHub_(root, buildHubModel_(payload));
      return;
    }
    const payload = await fetcher();
    writeOverviewCache_(periodKey, payload);
    renderHub_(root, buildHubModel_(payload));
  } catch (e) {
    console.warn('[overview]', e);
    const cached = readOverviewCache_();
    if (cached && cached.payload && cached.periodKey === periodKey) {
      renderHub_(root, buildHubModel_(cached.payload));
    } else {
      renderHub_(root, buildHubModel_({ snapshot: {} }));
    }
  } finally {
    refreshBusy = false;
    if (btn) btn.disabled = false;
  }
}

export function updateOverviewWelcomeFromEmail(emailRaw) {
  welcomeName = formatWelcomeName(emailRaw);
  const el = document.getElementById('ovHubWelcome');
  if (el) el.textContent = 'Welcome back, ' + welcomeName;
}

export function mountOverviewLanding(container) {
  if (!container) return;

  container.textContent = '';
  const root = document.createElement('div');
  root.className = 'ov-hub';
  root.innerHTML = `
    <section class="ov-hub__hero">
      <div class="ov-hub__hero-glow" aria-hidden="true"></div>
      <div class="ov-hub__hero-inner">
        <div class="ov-hub__hero-copy">
          <p class="ov-hub__welcome" id="ovHubWelcome">Welcome back, there</p>
          <h1 class="ov-hub__title">Sustainability Dashboard</h1>
          <p class="ov-hub__tagline">${BRAND_TAGLINE}</p>
          <p class="ov-hub__chip"><span aria-hidden="true">◎</span> Palm Oil Supply Chain Monitoring</p>
          <div class="ov-hub__actions">
            <button type="button" class="ov-hub__refresh ov-hub__glass-btn" id="ovHubRefresh">Refresh metrics</button>
          </div>
        </div>
        <div class="ov-hub__hero-stats" id="ovHubHeroStats"></div>
      </div>
    </section>

    <div class="ov-hub__section-head">
      <h2 class="ov-hub__section-title">Monitoring modules</h2>
      <span class="ov-hub__section-meta">8 active modules</span>
    </div>
    <div class="ov-hub__modules" id="ovHubModules"></div>

    <div class="ov-hub__bottom">
      <section class="ov-hub__panel">
        <h3 class="ov-hub__panel-title"><span aria-hidden="true">▥</span> Module completion</h3>
        <div class="ov-hub__progress" id="ovHubProgress"></div>
      </section>
      <section class="ov-hub__panel">
        <h3 class="ov-hub__panel-title"><span aria-hidden="true">↻</span> Recent activity</h3>
        <ul class="ov-hub__activity" id="ovHubActivity"></ul>
      </section>
    </div>
  `;
  container.appendChild(root);
  overviewHubRoot = root;
  bindQuarterCacheWatch_();

  root.querySelector('#ovHubRefresh').addEventListener('click', function() {
    refreshHub_(root, { force: true });
  });
  window.__overviewMetricsRefresh = function() {
    refreshHub_(root, { force: false });
  };
  refreshHub_(root, { force: false });
}
