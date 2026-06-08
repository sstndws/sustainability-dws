/**
 * Vercel serverless — sole bridge to Google Apps Script in secure mode.
 * Env (Project Settings, NOT VITE_*): GAS_WEBAPP_URL, GAS_API_SECRET,
 * VITE_SUPABASE_URL (or SUPABASE_URL), VITE_SUPABASE_ANON_KEY (or SUPABASE_ANON_KEY).
 */
import { createClient } from '@supabase/supabase-js';

function env_(key, alt) {
  return process.env[key] || process.env[alt] || '';
}

function parseBody_(req) {
  if (!req.body) return {};
  if (typeof req.body === 'string') {
    try { return JSON.parse(req.body); } catch (_) { return {}; }
  }
  return req.body;
}

export default async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store');

  if (req.method === 'OPTIONS') {
    res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Authorization, Content-Type');
    return res.status(204).end();
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const gasUrl = env_('GAS_WEBAPP_URL');
  const gasSecret = env_('GAS_API_SECRET');
  const supabaseUrl = env_('VITE_SUPABASE_URL', 'SUPABASE_URL');
  const supabaseAnon = env_('VITE_SUPABASE_ANON_KEY', 'SUPABASE_ANON_KEY');

  if (!gasUrl || !gasSecret || !supabaseUrl || !supabaseAnon) {
    console.error('[gas-proxy] Missing server env configuration');
    return res.status(500).json({ error: 'Server security not configured' });
  }

  const authHeader = String(req.headers.authorization || '');
  const jwt = authHeader.startsWith('Bearer ') ? authHeader.slice(7).trim() : '';
  if (!jwt) {
    return res.status(401).json({ error: 'Unauthorized — login required' });
  }

  const supabase = createClient(supabaseUrl.trim(), supabaseAnon.trim(), {
    auth: { persistSession: false, autoRefreshToken: false },
  });
  const { data: userData, error: userErr } = await supabase.auth.getUser(jwt);
  if (userErr || !userData || !userData.user) {
    return res.status(401).json({ error: 'Invalid or expired session' });
  }

  const actor = userData.user.email || userData.user.id || 'unknown';
  const payload = parseBody_(req);
  const method = String(payload.method || 'GET').toUpperCase();

  try {
    if (method === 'GET') {
      const params = Object.assign({}, payload.params || {}, {
        token: gasSecret,
        _actor: actor,
      });
      const fullUrl = gasUrl + '?' + new URLSearchParams(params).toString();
      const gasRes = await fetch(fullUrl, {
        method: 'GET',
        redirect: 'follow',
        headers: { 'Accept': 'application/json' },
      });
      const text = await gasRes.text();
      res.status(gasRes.status);
      res.setHeader('Content-Type', 'application/json; charset=utf-8');
      return res.send(text);
    }

    if (method === 'POST') {
      const body = Object.assign({}, payload.body || {}, {
        token: gasSecret,
        _actor: actor,
      });
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
    console.error('[gas-proxy] upstream error', err);
    return res.status(502).json({ error: 'Upstream data service unavailable' });
  }
}
