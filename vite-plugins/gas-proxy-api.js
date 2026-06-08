/**
 * Dev-only mirror of /api/gas-proxy (Vercel) when VITE_SECURE_GAS=true.
 */
import { loadEnv } from 'vite';
import { createClient } from '@supabase/supabase-js';

function readJsonBody_(req) {
  return new Promise(function(resolve, reject) {
    let raw = '';
    req.on('data', function(chunk) { raw += chunk; });
    req.on('end', function() {
      try { resolve(raw ? JSON.parse(raw) : {}); } catch (e) { reject(e); }
    });
    req.on('error', reject);
  });
}

export function gasProxyApiPlugin() {
  return {
    name: 'gas-proxy-api-dev',
    configureServer: function(server) {
      server.middlewares.use(async function(req, res, next) {
        if (!req.url || req.url.split('?')[0] !== '/api/gas-proxy') return next();

        const env = loadEnv(server.config.mode, process.cwd(), '');
        if (env.VITE_SECURE_GAS !== 'true') {
          res.statusCode = 503;
          res.setHeader('Content-Type', 'application/json');
          res.end(JSON.stringify({ error: 'VITE_SECURE_GAS not enabled' }));
          return;
        }

        const gasUrl = env.GAS_WEBAPP_URL;
        const gasSecret = env.GAS_API_SECRET;
        const supabaseUrl = env.VITE_SUPABASE_URL;
        const supabaseAnon = env.VITE_SUPABASE_ANON_KEY;

        if (!gasUrl || !gasSecret || !supabaseUrl || !supabaseAnon) {
          res.statusCode = 500;
          res.setHeader('Content-Type', 'application/json');
          res.end(JSON.stringify({ error: 'Missing GAS_WEBAPP_URL / GAS_API_SECRET in .env.local' }));
          return;
        }

        if (req.method === 'OPTIONS') {
          res.statusCode = 204;
          res.end();
          return;
        }

        if (req.method !== 'POST') {
          res.statusCode = 405;
          res.setHeader('Content-Type', 'application/json');
          res.end(JSON.stringify({ error: 'Method not allowed' }));
          return;
        }

        try {
          const auth = String(req.headers.authorization || '');
          const jwt = auth.startsWith('Bearer ') ? auth.slice(7).trim() : '';
          if (!jwt) {
            res.statusCode = 401;
            res.setHeader('Content-Type', 'application/json');
            res.end(JSON.stringify({ error: 'Unauthorized' }));
            return;
          }

          const supabase = createClient(supabaseUrl, supabaseAnon, {
            auth: { persistSession: false, autoRefreshToken: false },
          });
          const { data: userData, error: userErr } = await supabase.auth.getUser(jwt);
          if (userErr || !userData || !userData.user) {
            res.statusCode = 401;
            res.setHeader('Content-Type', 'application/json');
            res.end(JSON.stringify({ error: 'Invalid session' }));
            return;
          }

          const actor = userData.user.email || userData.user.id || 'unknown';
          const payload = await readJsonBody_(req);
          const method = String(payload.method || 'GET').toUpperCase();

          if (method === 'GET') {
            const params = Object.assign({}, payload.params || {}, { token: gasSecret, _actor: actor });
            const fullUrl = gasUrl + '?' + new URLSearchParams(params).toString();
            const gasRes = await fetch(fullUrl, { method: 'GET', redirect: 'follow' });
            const text = await gasRes.text();
            res.statusCode = gasRes.status;
            res.setHeader('Content-Type', 'application/json');
            res.end(text);
            return;
          }

          const body = Object.assign({}, payload.body || {}, { token: gasSecret, _actor: actor });
          const gasRes = await fetch(gasUrl, {
            method: 'POST',
            redirect: 'follow',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify(body),
          });
          const text = await gasRes.text();
          res.statusCode = gasRes.status;
          res.setHeader('Content-Type', 'application/json');
          res.end(text);
        } catch (err) {
          res.statusCode = 502;
          res.setHeader('Content-Type', 'application/json');
          res.end(JSON.stringify({ error: String(err && err.message ? err.message : err) }));
        }
      });
    },
  };
}
