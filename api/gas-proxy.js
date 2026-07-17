/**
 * Vercel serverless — menyembunyikan URL GAS dari browser.
 * Tanpa Supabase. Cukup set di Vercel: GAS_WEBAPP_URL (+ opsional GAS_API_SECRET).
 */
function env_(key) {
  return process.env[key] || '';
}

function parseBody_(req) {
  if (!req.body) return {};
  if (typeof req.body === 'string') {
    try { return JSON.parse(req.body); } catch (_) { return {}; }
  }
  return req.body;
}

function isSameSiteRequest_(req) {
  const host = String(req.headers.host || '').toLowerCase();
  const origin = String(req.headers.origin || '').toLowerCase();
  const referer = String(req.headers.referer || '').toLowerCase();
  if (!host) return false;
  if (origin && origin.indexOf(host) !== -1) return true;
  if (referer && referer.indexOf(host) !== -1) return true;
  return !origin && !referer;
}

export default async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store');

  if (req.method === 'OPTIONS') {
    res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    return res.status(204).end();
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  if (!isSameSiteRequest_(req)) {
    return res.status(403).json({ error: 'Forbidden' });
  }

  // Strip accidental quotes/whitespace from Vercel env paste.
  const gasUrl = String(env_('GAS_WEBAPP_URL') || '')
    .trim()
    .replace(/^["']|["']$/g, '')
    || 'https://script.google.com/macros/s/AKfycbyxbr4znky9OIANcAN5G7qzZbbQkUdx2q21UlwnOTTgoBJtT9cr7ORpz-U-aB41iIPF7g/exec';
  const gasSecret = env_('GAS_API_SECRET');

  if (!/^https:\/\/script\.google\.com\/macros\/s\/AKfycb[\w-]+\/exec\/?$/.test(gasUrl)) {
    console.error('[gas-proxy] invalid GAS_WEBAPP_URL', gasUrl.slice(0, 120));
    return res.status(502).json({
      error: 'Upstream data service unavailable',
      detail: 'Invalid GAS_WEBAPP_URL — must be https://script.google.com/macros/s/.../exec (no quotes)',
    });
  }

  const payload = parseBody_(req);
  const method = String(payload.method || 'GET').toUpperCase();
  const actor = 'vercel-proxy';

  try {
    if (method === 'GET') {
      const params = Object.assign({}, payload.params || {});
      if (gasSecret) {
        params.token = gasSecret;
        params._actor = actor;
      }
      const fullUrl = gasUrl.replace(/\/$/, '') + '?' + new URLSearchParams(params).toString();
      const gasRes = await fetch(fullUrl, {
        method: 'GET',
        redirect: 'follow',
        headers: { Accept: 'application/json' },
      });
      const text = await gasRes.text();
      res.status(gasRes.status);
      res.setHeader('Content-Type', 'application/json; charset=utf-8');
      return res.send(text);
    }

    if (method === 'POST') {
      const body = Object.assign({}, payload.body || {});
      if (gasSecret) {
        body.token = gasSecret;
        body._actor = actor;
      }
      const gasRes = await fetch(gasUrl, {
        method: 'POST',
        redirect: 'follow',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify(body),
      });
      const text = await gasRes.text();
      res.status(gasRes.status);
      res.setHeader('Content-Type', 'application/json; charset=utf-8');
      return res.send(text);
    }

    return res.status(400).json({ error: 'Invalid method' });
  } catch (err) {
    const msg = (err && err.message) ? String(err.message) : String(err);
    console.error('[gas-proxy] upstream error', msg, 'url=', gasUrl.slice(0, 80));
    return res.status(502).json({
      error: 'Upstream data service unavailable',
      detail: msg.slice(0, 200),
    });
  }
}
