import './overview-landing.css';
import { BRAND_TAGLINE } from './brand.js';

const BG_SVG = `<svg class="ov-ld__bg-svg" viewBox="0 0 300 420" fill="none" aria-hidden="true"><rect x="110" y="10" width="9" height="320" fill="#5a1515"/><rect x="88" y="76" width="52" height="7" fill="#5a1515"/><rect x="65" y="95" width="94" height="195" rx="2" fill="#5a1515"/><rect x="76" y="116" width="17" height="22" rx="1" fill="#e0d5d0"/><rect x="108" y="116" width="17" height="22" rx="1" fill="#e0d5d0"/><rect x="140" y="116" width="17" height="22" rx="1" fill="#e0d5d0"/><rect x="76" y="152" width="17" height="22" rx="1" fill="#e0d5d0"/><rect x="108" y="152" width="17" height="22" rx="1" fill="#e0d5d0"/><rect x="140" y="152" width="17" height="22" rx="1" fill="#e0d5d0"/><rect x="76" y="188" width="17" height="22" rx="1" fill="#e0d5d0"/><rect x="108" y="188" width="17" height="22" rx="1" fill="#e0d5d0"/><rect x="140" y="188" width="17" height="22" rx="1" fill="#e0d5d0"/><rect x="58" y="288" width="108" height="7" fill="#5a1515"/><rect x="180" y="155" width="7" height="140" fill="#5a1515"/><rect x="163" y="177" width="40" height="6" fill="#5a1515"/><rect x="155" y="182" width="56" height="113" rx="1" fill="#5a1515"/><rect x="164" y="198" width="13" height="15" rx="1" fill="#e0d5d0"/><rect x="186" y="198" width="13" height="15" rx="1" fill="#e0d5d0"/><rect x="164" y="224" width="13" height="15" rx="1" fill="#e0d5d0"/><rect x="186" y="224" width="13" height="15" rx="1" fill="#e0d5d0"/><rect x="148" y="293" width="68" height="7" fill="#5a1515"/><line x1="38" y1="300" x2="255" y2="300" stroke="#5a1515" stroke-width="3.5"/></svg>`;

const ICONS = {
  mill: `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" aria-hidden="true"><rect x="3" y="10" width="18" height="11" rx="1"/><path d="M7 10V7a2 2 0 014 0M13 10V7a2 2 0 014 0"/><line x1="12" y1="10" x2="12" y2="21"/></svg>`,
  company: `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" aria-hidden="true"><path d="M3 21h18"/><path d="M6 21V9l6-4 6 4v12"/><path d="M9 13h1"/><path d="M14 13h1"/></svg>`,
  ttp: `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" aria-hidden="true"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/></svg>`,
  grievance: `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" aria-hidden="true"><path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>`,
  bl: `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" aria-hidden="true"><rect x="1" y="3" width="15" height="13"/><polygon points="16 8 20 8 23 11 23 16 16 16 16 8"/><circle cx="5.5" cy="18.5" r="2.5"/><circle cx="18.5" cy="18.5" r="2.5"/></svg>`,
  facility: `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" aria-hidden="true"><path d="M12 20V10"/><path d="M18 20V4"/><path d="M6 20v-4"/></svg>`,
  eudr: `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" aria-hidden="true"><circle cx="12" cy="12" r="10"/><path d="M2 12h20"/><path d="M12 2a14 14 0 0 1 0 20 14 14 0 0 1 0-20"/></svg>`,
  contacts: `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" aria-hidden="true"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>`,
  questionnaire: `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" aria-hidden="true"><path d="M9 11l3 3L22 4"/><path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"/></svg>`,
  nbl: `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" aria-hidden="true"><circle cx="12" cy="12" r="10"/><line x1="4.93" y1="4.93" x2="19.07" y2="19.07"/></svg>`,
  dds: `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" aria-hidden="true"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/><line x1="10" y1="9" x2="8" y2="9"/></svg>`,
};

const CARDS = [
  {
    panel: 'mill-onboarding',
    title: 'Mill Onboarding',
    desc: 'Manage and register mill profiles across the supply chain.',
    icon: 'mill',
  },
  {
    panel: 'company-profile-list',
    title: 'Company Profile List',
    desc: 'Facility profiles — plant, site, address, capacity, coordinates, and certifications.',
    icon: 'company',
  },
  {
    panel: 'ttm-ttp',
    title: 'Traceability Data',
    desc: 'View and track trade monitoring data for each operational unit.',
    icon: 'ttp',
  },
  {
    panel: 'grievance',
    title: 'Grievance Mechanism',
    desc: 'Track, submit, and resolve grievances across supply chain partners.',
    icon: 'grievance',
  },
  {
    panel: 'bl-monitoring',
    title: 'Declaration Monitoring',
    desc: 'Shipping BLs, declarations, and SD monitoring with linked TTM / TTP data.',
    icon: 'bl',
  },
  {
    panel: 'no-buy-list',
    title: 'No Buy List',
    desc: 'NBL supplier registry and Unilever mill locations from your spreadsheet.',
    icon: 'nbl',
  },
  {
    panel: 'performa-facility',
    title: 'Facility Performance',
    desc: 'Facility KPIs, supply traceability, and risk metrics by CPO facility.',
    icon: 'facility',
  },
  {
    panel: 'eudr-potential',
    title: 'EUDR Potential',
    desc: 'EUDR readiness per mill — identity from Mill Onboarding; status on a separate sheet.',
    icon: 'eudr',
  },
  {
    panel: 'contact-list-supplier',
    title: 'Contact List Supplier',
    desc: 'Sustainability PIC contacts synced automatically from approved SDD screenings.',
    icon: 'contacts',
  },
  {
    panel: 'questionnaire-monitoring',
    title: 'Questionnaire Monitoring',
    desc: 'Monitor questionnaire progress per mill — status, milestones, and follow-up dates.',
    icon: 'questionnaire',
  },
  {
    panel: 'due-diligence-statement',
    title: 'Due Diligence Statement',
    desc: 'EUDR due diligence packages per shipment — create records, manage suppliers & plots, export PDF.',
    icon: 'dds',
  },
];

function formatWelcomeName(emailRaw) {
  const s = (emailRaw || '').trim();
  if (!s) return 'there';
  const local = s.split('@')[0] || s;
  return local
    .replace(/[._-]+/g, ' ')
    .split(' ')
    .filter(Boolean)
    .map((w) => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase())
    .join(' ');
}

/** Updates the overview hero line (call after login). */
export function updateOverviewWelcomeFromEmail(emailRaw) {
  const line = document.getElementById('ovLdWelcomeLine');
  if (!line) return;
  line.textContent = `Welcome, ${formatWelcomeName(emailRaw)}`;
}

function goPanel(panel) {
  if (typeof window.switchPanel === 'function') window.switchPanel(panel);
}

/** Builds overview DOM inside `#overview-root` (data-driven, no giant HTML partial). */
export function mountOverviewLanding(container) {
  if (!container) return;
  container.textContent = '';

  const root = document.createElement('div');
  root.className = 'ov-ld';

  const glow = document.createElement('div');
  glow.className = 'ov-ld__glow';
  root.appendChild(glow);

  const bgWrap = document.createElement('div');
  bgWrap.className = 'ov-ld__bg-buildings';
  bgWrap.setAttribute('aria-hidden', 'true');
  for (let i = 1; i <= 4; i += 1) {
    const tpl = document.createElement('template');
    tpl.innerHTML = BG_SVG.replace('ov-ld__bg-svg"', `ov-ld__bg-svg ov-ld__bg-svg--${i}"`).trim();
    const el = tpl.content.firstElementChild;
    if (el) bgWrap.appendChild(el);
  }
  root.appendChild(bgWrap);

  const main = document.createElement('div');
  main.className = 'ov-ld__main';

  const header = document.createElement('header');
  header.className = 'ov-ld__header';
  header.innerHTML = `
    <p class="ov-ld__welcome" id="ovLdWelcomeLine">Welcome, there</p>
    <h2 class="ov-ld__title">Sustainability Dashboard</h2>
    <p class="ov-ld__tagline">${BRAND_TAGLINE}</p>
    <p class="ov-ld__subtitle">Palm Oil Supply Chain Monitoring</p>
  `;
  main.appendChild(header);

  const grid = document.createElement('div');
  grid.className = 'ov-ld__cards';
  grid.setAttribute('role', 'navigation');
  grid.setAttribute('aria-label', 'Modules');

  CARDS.forEach((c) => {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'ov-ld__card';
    btn.dataset.panel = c.panel;
    btn.setAttribute('aria-label', c.title);
    const icon = document.createElement('div');
    icon.className = 'ov-ld__card-icon';
    icon.innerHTML = ICONS[c.icon] || '';
    const body = document.createElement('div');
    body.className = 'ov-ld__card-body';
    const t = document.createElement('div');
    t.className = 'ov-ld__card-title';
    t.textContent = c.title;
    const d = document.createElement('div');
    d.className = 'ov-ld__card-desc';
    d.textContent = c.desc;
    body.appendChild(t);
    body.appendChild(d);
    btn.appendChild(icon);
    btn.appendChild(body);
    grid.appendChild(btn);
  });

  grid.addEventListener('click', (e) => {
    const card = e.target.closest('.ov-ld__card');
    if (!card || !grid.contains(card)) return;
    const panel = card.dataset.panel;
    if (panel) goPanel(panel);
  });

  main.appendChild(grid);
  root.appendChild(main);
  container.appendChild(root);
}
