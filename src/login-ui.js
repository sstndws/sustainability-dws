import './login-shell.css';
import { BRAND_TAGLINE } from './brand.js';
import {
  allowLocalLogin,
  getHubLoginUrl,
  getHubPortalOrigin,
  isAuthGateEnabled,
} from './hub-sso.js';

function el(tag, className, props = {}) {
  const n = document.createElement(tag);
  if (className) n.className = className;
  const { textContent, htmlFor, children, ...rest } = props;
  if (textContent != null) n.textContent = textContent;
  if (htmlFor != null) n.htmlFor = htmlFor;
  if (Array.isArray(children)) children.forEach((c) => c && n.appendChild(c));
  Object.entries(rest).forEach(([k, v]) => {
    if (v == null) return;
    n.setAttribute(k, v);
  });
  return n;
}

function mountBrandCard_(card, headline, hint) {
  const brand = el('div', 'login-shell__brand');
  const logoWrap = el('div', 'login-shell__logo');
  logoWrap.innerHTML = `<svg viewBox="0 0 16 16" fill="none" aria-hidden="true"><path d="M8 2L14 6V10L8 14L2 10V6L8 2Z" fill="white" opacity="0.9"/><path d="M8 5L11 7V9L8 11L5 9V7L8 5Z" fill="#7a1515"/></svg>`;
  const titles = el('div', 'login-shell__titles');
  titles.appendChild(el('span', 'login-shell__name', { textContent: 'KPNCORP' }));
  titles.appendChild(el('span', 'login-shell__sub', { textContent: 'Downstream · Admin' }));
  titles.appendChild(el('span', 'login-shell__tagline', { textContent: BRAND_TAGLINE }));
  brand.appendChild(logoWrap);
  brand.appendChild(titles);

  card.appendChild(brand);
  card.appendChild(el('hr', 'login-shell__divider', { 'aria-hidden': 'true' }));
  card.appendChild(el('h1', 'login-shell__headline', { textContent: headline }));
  card.appendChild(el('p', 'login-shell__hint', { textContent: hint }));
}

/** Hub-only gate — full-bleed handoff screen (no password form). */
function mountHubRedirectGate_(loginRoot) {
  const hubLogin = getHubLoginUrl(
    typeof location !== 'undefined' ? location.origin + '/' : undefined
  );
  const hubHome = getHubPortalOrigin() + '/';

  loginRoot.classList.add('hub-gate-page');

  const stage = el('div', 'hub-gate');
  stage.setAttribute('role', 'status');
  stage.setAttribute('aria-live', 'polite');

  const backdrop = el('div', 'hub-gate__backdrop', { 'aria-hidden': 'true' });
  backdrop.appendChild(el('div', 'hub-gate__wash'));
  backdrop.appendChild(el('div', 'hub-gate__grain'));
  backdrop.appendChild(el('div', 'hub-gate__orb hub-gate__orb--a'));
  backdrop.appendChild(el('div', 'hub-gate__orb hub-gate__orb--b'));

  const content = el('div', 'hub-gate__content');

  const mark = el('div', 'hub-gate__mark');
  const logo = el('div', 'hub-gate__logo');
  logo.innerHTML = `<svg viewBox="0 0 16 16" fill="none" aria-hidden="true"><path d="M8 2L14 6V10L8 14L2 10V6L8 2Z" fill="white" opacity="0.95"/><path d="M8 5L11 7V9L8 11L5 9V7L8 5Z" fill="#7a1515"/></svg>`;
  const brandText = el('div', 'hub-gate__brand-text');
  brandText.appendChild(el('p', 'hub-gate__brand', { textContent: 'KPNCORP' }));
  brandText.appendChild(el('p', 'hub-gate__product', { textContent: 'Sustainability Dashboard' }));
  mark.appendChild(logo);
  mark.appendChild(brandText);

  const title = el('h1', 'hub-gate__title', { textContent: 'Continue through the Hub' });
  const lead = el('p', 'hub-gate__lead', {
    textContent:
      'This module opens only after you sign in at the Sustainability Hub Portal. You will be redirected automatically.',
  });

  const progress = el('div', 'hub-gate__progress', { 'aria-hidden': 'true' });
  progress.appendChild(el('span', 'hub-gate__progress-bar'));

  const status = el('p', 'hub-gate__status', {
    id: 'hubSsoStatus',
    textContent: 'Opening Hub Portal…',
  });

  const actions = el('div', 'hub-gate__actions');
  const primary = el('a', 'hub-gate__btn hub-gate__btn--primary', {
    href: hubLogin,
    id: 'btn-hub-redirect',
    textContent: 'Open Hub Portal',
  });
  const secondary = el('a', 'hub-gate__btn hub-gate__btn--ghost', {
    href: hubHome,
    textContent: 'Hub home',
  });
  actions.appendChild(primary);
  actions.appendChild(secondary);

  const footnote = el('p', 'hub-gate__footnote', { textContent: BRAND_TAGLINE });

  content.appendChild(mark);
  content.appendChild(title);
  content.appendChild(lead);
  content.appendChild(progress);
  content.appendChild(status);
  content.appendChild(actions);
  content.appendChild(footnote);

  stage.appendChild(backdrop);
  stage.appendChild(content);
  loginRoot.appendChild(stage);
}

function mountLocalLoginForm_(loginRoot) {
  loginRoot.classList.remove('hub-gate-page');
  const hubHome = getHubPortalOrigin() + '/';
  const backLink = el('a', 'login-shell__back', {
    href: hubHome,
    textContent: '← Back to Sustainability Hub Portal',
  });

  const glow = el('div', 'login-shell__glow');
  const shell = el('div', 'login-shell');
  const card = el('div', 'login-shell__card');

  mountBrandCard_(
    card,
    'Sign in',
    'Use your authorized account to access the sustainability dashboard.'
  );

  const fgEmail = el('div', 'form-group');
  fgEmail.appendChild(el('label', null, { htmlFor: 'loginEmail', textContent: 'Email address' }));
  fgEmail.appendChild(
    el('input', null, {
      type: 'text',
      id: 'loginEmail',
      placeholder: 'you@company.com',
      autocomplete: 'username',
      inputmode: 'email',
      spellcheck: 'false',
    })
  );

  const fgPass = el('div', 'form-group');
  fgPass.appendChild(el('label', null, { htmlFor: 'loginPass', textContent: 'Password' }));
  fgPass.appendChild(
    el('input', null, {
      type: 'password',
      id: 'loginPass',
      placeholder: '••••••••',
      autocomplete: 'current-password',
    })
  );

  const btn = el('button', 'login-submit', { type: 'button', id: 'btn-login-submit', textContent: 'Sign in' });
  const err = el('div', 'login-err', { id: 'loginErr', textContent: 'Invalid credentials. Please try again.' });

  card.appendChild(fgEmail);
  card.appendChild(fgPass);
  card.appendChild(btn);
  card.appendChild(err);

  shell.appendChild(card);
  loginRoot.appendChild(backLink);
  loginRoot.appendChild(glow);
  loginRoot.appendChild(shell);
}

/** Mounts sign-in / Hub gate into `#login`. */
export function mountLoginPage(loginRoot) {
  if (!loginRoot) return;
  loginRoot.textContent = '';
  loginRoot.classList.remove('hub-gate-page');

  if (isAuthGateEnabled() && !allowLocalLogin()) {
    mountHubRedirectGate_(loginRoot);
    return;
  }

  mountLocalLoginForm_(loginRoot);
}
