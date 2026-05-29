import './login-shell.css';
import { BRAND_TAGLINE } from './brand.js';

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

/** Mounts the sign-in form into `#login` (Vite-built DOM, minimal HTML partial). */
export function mountLoginPage(loginRoot) {
  if (!loginRoot) return;
  loginRoot.textContent = '';

  const glow = el('div', 'login-shell__glow');
  const shell = el('div', 'login-shell');
  const card = el('div', 'login-shell__card');

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
  card.appendChild(el('h1', 'login-shell__headline', { textContent: 'Sign in' }));
  card.appendChild(
    el('p', 'login-shell__hint', {
      textContent: 'Use your authorized account to access the sustainability dashboard.',
    })
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
  loginRoot.appendChild(glow);
  loginRoot.appendChild(shell);
}
