/**
 * Dev mirror of /api/gas-proxy — tanpa Supabase.
 */
import { loadEnv } from 'vite';

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
        const gasUrl = env.GAS_WEBAPP_URL;
        const gasSecret = env.GAS_API_SECRET || '';

        if (!gasUrl) {
          res.statusCode = 500;
          res.setHeader('Content-Type', 'application/json');
          res.end(JSON.stringify({ error: 'Set GAS_WEBAPP_URL in .env.local' }));
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
          const payload = await readJsonBody_(req);
          const method = String(payload.method || 'GET').toUpperCase();

          if (method === 'GET') {
            const params = Object.assign({}, payload.params || {});
            if (gasSecret) params.token = gasSecret;
            const fullUrl = gasUrl + '?' + new URLSearchParams(params).toString();
            const gasRes = await fetch(fullUrl, { method: 'GET', redirect: 'follow' });
            const text = await gasRes.text();
            res.statusCode = gasRes.status;
            res.setHeader('Content-Type', 'application/json');
            res.end(text);
            return;
          }

          const body = Object.assign({}, payload.body || {});
          if (gasSecret) body.token = gasSecret;
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
