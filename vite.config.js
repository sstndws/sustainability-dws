import { defineConfig } from 'vite';
import htmlInclude from './vite-plugins/html-include.js';
import { gasProxyApiPlugin } from './vite-plugins/gas-proxy-api.js';

export default defineConfig({
  plugins: [htmlInclude(), gasProxyApiPlugin()],

  esbuild: {
    legalComments: 'none',
  },

  // ── Dev server — port khusus sustain-dashboard (hindari bentrok app lain di :5173) ──
  server: {
    // true = listen on all interfaces so both http://localhost:5340 and http://127.0.0.1:5340 work
    host: true,
    port: 5340,
    strictPort: false,
    open: true,
    // Pastikan entry HTML tidak disangka stale oleh browser saat dev
    headers: {
      'Cache-Control': 'no-store',
    },
    // Proxy Apps Script di dev — hindari CORS/adblock saat fetch dari 127.0.0.1
    proxy: {
      '/gas-api': {
        target: 'https://script.google.com',
        changeOrigin: true,
        secure: true,
        timeout: 120000,
        proxyTimeout: 120000,
        rewrite: function(path) {
          return path.replace(/^\/gas-api/, '');
        },
      },
    },
  },

  preview: {
    host: true,
    port: 4173,
    strictPort: true,
    open: true,
  },

  // ── Build ─────────────────────────────────────────────────
  build: {
    outDir: 'dist',
    assetsDir: 'assets',
    sourcemap: false,
    minify: 'esbuild',
    cssCodeSplit: true,
    // No chunk size limit — allow the large JS bundle as-is
    chunkSizeWarningLimit: Infinity,
    rollupOptions: {
      input: 'index.html',
      output: {
        manualChunks(id) {
          if (id.includes('node_modules')) return 'vendor';
        },
      },
    },
  },
});
