/**
 * Hub Portal SSO — login only at the hub; this app accepts a session bridge.
 *
 * Hub should open:
 *   https://sustainability-dws-psi.vercel.app/auth-bridge#access_token=...&refresh_token=...
 * or the same tokens as query params on /auth-bridge.
 */

const DEFAULT_HUB_ORIGIN = 'https://sustainability-hub-portal-eight.vercel.app';

export function isAuthGateEnabled() {
  return import.meta.env.VITE_AUTH_ENABLED === 'true';
}

/** Local email/password form (dev only). Default off when auth gate is on. */
export function allowLocalLogin() {
  return import.meta.env.VITE_ALLOW_LOCAL_LOGIN === 'true';
}

export function getHubPortalOrigin() {
  const fromEnv = String(import.meta.env.VITE_HUB_PORTAL_URL || '')
    .trim()
    .replace(/\/$/, '');
  return fromEnv || DEFAULT_HUB_ORIGIN;
}

export function getHubLoginUrl(redirectBack) {
  const origin = getHubPortalOrigin();
  let path = String(import.meta.env.VITE_HUB_LOGIN_PATH || '/login').trim() || '/login';
  let url;
  if (/^https?:\/\//i.test(path)) {
    url = new URL(path);
  } else {
    if (!path.startsWith('/')) path = '/' + path;
    url = new URL(origin + path);
  }
  // Hub LoginForm uses `next` (internal path). After login, user opens apps from Hub cards.
  url.searchParams.set('next', '/');
  if (redirectBack) url.searchParams.set('from', redirectBack);
  return url.toString();
}

export function redirectToHubLogin() {
  if (typeof window === 'undefined') return;
  const back =
    window.location.origin +
    (window.location.pathname.startsWith('/auth-bridge') ? '/' : window.location.pathname || '/');
  window.location.replace(getHubLoginUrl(back));
}

/**
 * Read access_token + refresh_token from hash or query (auth-bridge).
 * @returns {{ access_token: string, refresh_token: string } | null}
 */
export function extractBridgeTokens() {
  if (typeof window === 'undefined') return null;

  const fromParams = (params) => {
    const access_token = params.get('access_token');
    const refresh_token = params.get('refresh_token');
    if (access_token && refresh_token) return { access_token, refresh_token };
    return null;
  };

  const hash = String(window.location.hash || '').replace(/^#/, '');
  if (hash) {
    const fromHash = fromParams(new URLSearchParams(hash));
    if (fromHash) return fromHash;
  }

  const fromQuery = fromParams(new URLSearchParams(window.location.search || ''));
  if (fromQuery) return fromQuery;

  return null;
}

export function isAuthBridgePath() {
  if (typeof window === 'undefined') return false;
  const path = String(window.location.pathname || '').replace(/\/+$/, '') || '/';
  return path === '/auth-bridge';
}

/** Strip tokens from the address bar after setSession. */
export function clearBridgeParamsFromUrl() {
  if (typeof window === 'undefined' || !window.history || !window.history.replaceState) return;
  window.history.replaceState({}, document.title, '/');
}

/**
 * Apply hub-carried tokens into the local Supabase session.
 * @param {import('@supabase/supabase-js').SupabaseClient} supabase
 * @returns {Promise<{ ok: boolean, error?: string }>}
 */
export async function applyBridgeSession(supabase) {
  const tokens = extractBridgeTokens();
  if (!tokens) {
    if (isAuthBridgePath()) return { ok: false, error: 'missing_tokens' };
    return { ok: false, error: 'no_bridge' };
  }
  try {
    const { error } = await supabase.auth.setSession({
      access_token: tokens.access_token,
      refresh_token: tokens.refresh_token,
    });
    clearBridgeParamsFromUrl();
    if (error) return { ok: false, error: error.message || 'set_session_failed' };
    return { ok: true };
  } catch (e) {
    clearBridgeParamsFromUrl();
    return { ok: false, error: e && e.message ? e.message : String(e) };
  }
}
