/**
 * Soft palm-plantation page background for Mill Executive PDF.
 * Bundled via Vite so export always finds the artwork (no /public fetch miss).
 */
import millExecBgAsset from './assets/mill-executive-bg.png?url';

let cachedDataUrl = null;
let loadPromise = null;

function drawProceduralPlantationBg_(ctx, w, h) {
  const g = ctx.createLinearGradient(0, 0, 0, h);
  g.addColorStop(0, '#F5F0E6');
  g.addColorStop(0.5, '#E6EDE3');
  g.addColorStop(1, '#D8E4D0');
  ctx.fillStyle = g;
  ctx.fillRect(0, 0, w, h);

  ctx.save();
  ctx.globalAlpha = 0.12;
  ctx.strokeStyle = '#3D5C45';
  ctx.lineWidth = 2;
  for (let y = h * 0.28; y < h * 1.05; y += 18) {
    ctx.beginPath();
    ctx.moveTo(-w * 0.1, y);
    ctx.lineTo(w * 1.1, y + h * 0.1);
    ctx.stroke();
  }
  ctx.restore();

  function palmSilhouette_(x, baseY, scale) {
    ctx.save();
    ctx.translate(x, baseY);
    ctx.scale(scale, scale);
    ctx.globalAlpha = 0.14;
    ctx.fillStyle = '#2F5238';
    ctx.fillRect(-4, -100, 8, 100);
    for (let i = 0; i < 7; i++) {
      const a = (-0.55 + i * 0.18) * Math.PI;
      ctx.beginPath();
      ctx.moveTo(0, -95);
      ctx.quadraticCurveTo(Math.cos(a) * 60, -95 + Math.sin(a) * 42, Math.cos(a) * 85, -72 + Math.sin(a) * 58);
      ctx.quadraticCurveTo(Math.cos(a) * 48, -88 + Math.sin(a) * 30, 0, -95);
      ctx.fill();
    }
    ctx.restore();
  }

  [
    [0.06, 0.95, 1], [0.14, 0.9, 1.2], [0.24, 0.96, 0.85], [0.68, 0.92, 1.1], [0.78, 0.95, 0.9], [0.9, 0.88, 1.15],
  ].forEach(function(p) {
    palmSilhouette_(w * p[0], h * p[1], p[2]);
  });

  applyReadableWash_(ctx, w, h, true);
}

/** Light frosted overlay — plantation still visible */
function applyReadableWash_(ctx, w, h, procedural) {
  const wash = ctx.createLinearGradient(0, 0, 0, h);
  if (procedural) {
    wash.addColorStop(0, 'rgba(255,255,255,0.28)');
    wash.addColorStop(0.55, 'rgba(255,255,255,0.22)');
    wash.addColorStop(1, 'rgba(255,255,255,0.26)');
  } else {
    wash.addColorStop(0, 'rgba(255,252,248,0.22)');
    wash.addColorStop(0.45, 'rgba(255,255,255,0.18)');
    wash.addColorStop(1, 'rgba(248,252,246,0.24)');
  }
  ctx.fillStyle = wash;
  ctx.fillRect(0, 0, w, h);

  const vignette = ctx.createRadialGradient(w / 2, h / 2, w * 0.12, w / 2, h / 2, w * 0.78);
  vignette.addColorStop(0, 'rgba(255,255,255,0)');
  vignette.addColorStop(1, 'rgba(255,255,255,0.12)');
  ctx.fillStyle = vignette;
  ctx.fillRect(0, 0, w, h);
}

function processPhotoToSoftBg_(ctx, w, h, img) {
  const iw = img.naturalWidth || img.width;
  const ih = img.naturalHeight || img.height;
  const scale = Math.max(w / iw, h / ih);
  const dw = iw * scale;
  const dh = ih * scale;
  const dx = (w - dw) / 2;
  const dy = (h - dh) / 2;
  ctx.drawImage(img, dx, dy, dw, dh);
  applyReadableWash_(ctx, w, h, false);
}

function loadImageFromUrl_(url) {
  return new Promise(function(resolve, reject) {
    fetch(url)
      .then(function(res) {
        if (!res.ok) throw new Error('HTTP ' + res.status);
        return res.blob();
      })
      .then(function(blob) {
        const objUrl = URL.createObjectURL(blob);
        const img = new Image();
        img.onload = function() {
          URL.revokeObjectURL(objUrl);
          resolve(img);
        };
        img.onerror = function() {
          URL.revokeObjectURL(objUrl);
          reject(new Error('decode failed'));
        };
        img.src = objUrl;
      })
      .catch(reject);
  });
}

/**
 * @returns {Promise<string>} JPEG data URL for full landscape page
 */
export async function getMillExecutiveBackgroundDataUrl_() {
  if (cachedDataUrl) return cachedDataUrl;
  if (loadPromise) return loadPromise;

  loadPromise = (async function() {
    // ~72–90 DPI A4 landscape — enough for soft wash, much cheaper than photo-native
    const w = 840;
    const h = 594;
    const canvas = document.createElement('canvas');
    canvas.width = w;
    canvas.height = h;
    const ctx = canvas.getContext('2d');
    if (!ctx) return '';

    try {
      const img = await loadImageFromUrl_(millExecBgAsset);
      processPhotoToSoftBg_(ctx, w, h, img);
    } catch (e) {
      console.warn('[Mill Executive] Background photo unavailable, using procedural:', e);
      drawProceduralPlantationBg_(ctx, w, h);
    }

    // Yield once before the expensive toDataURL encode.
    await new Promise(function(resolve) {
      requestAnimationFrame(function() { setTimeout(resolve, 0); });
    });
    cachedDataUrl = canvas.toDataURL('image/jpeg', 0.62);
    return cachedDataUrl;
  })();

  return loadPromise;
}
