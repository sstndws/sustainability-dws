/**
 * Embed DM Sans + Cormorant Garamond into jsPDF documents.
 * Fonts are fetched once from Google Fonts gstatic and cached in memory.
 */

const FONT_URLS = {
  dmRegular: 'https://fonts.gstatic.com/s/dmsans/v17/rP2tp2ywxg089UriI5-g4vlH9VoD8CmcqZG40F9JadbnoEwAopxhTg.ttf',
  dmBold: 'https://fonts.gstatic.com/s/dmsans/v17/rP2tp2ywxg089UriI5-g4vlH9VoD8CmcqZG40F9JadbnoEwARZthTg.ttf',
  cormorantSemi: 'https://fonts.gstatic.com/s/cormorantgaramond/v21/co3umX5slCNuHLi8bLeY9MK7whWMhyjypVO7abI26QOD_iE9GnM.ttf',
};

let _fontData = null;
let _loadPromise = null;
const _registeredDocs = new WeakSet();

function arrayBufferToBase64_(buffer) {
  const bytes = new Uint8Array(buffer);
  let binary = '';
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
  }
  return btoa(binary);
}

async function loadFontData_() {
  if (_fontData) return _fontData;
  if (!_loadPromise) {
    _loadPromise = Promise.all([
      fetch(FONT_URLS.dmRegular).then(function(r) {
        if (!r.ok) throw new Error('Failed to load DM Sans Regular');
        return r.arrayBuffer();
      }),
      fetch(FONT_URLS.dmBold).then(function(r) {
        if (!r.ok) throw new Error('Failed to load DM Sans Bold');
        return r.arrayBuffer();
      }),
      fetch(FONT_URLS.cormorantSemi).then(function(r) {
        if (!r.ok) throw new Error('Failed to load Cormorant Garamond');
        return r.arrayBuffer();
      }),
    ]).then(function(buffers) {
      _fontData = {
        dmRegular: arrayBufferToBase64_(buffers[0]),
        dmBold: arrayBufferToBase64_(buffers[1]),
        cormorantSemi: arrayBufferToBase64_(buffers[2]),
      };
      return _fontData;
    });
  }
  return _loadPromise;
}

function applyFontsToDoc_(doc, data) {
  doc.addFileToVFS('DMSans-Regular.ttf', data.dmRegular);
  doc.addFont('DMSans-Regular.ttf', 'DMSans', 'normal');
  doc.addFileToVFS('DMSans-Bold.ttf', data.dmBold);
  doc.addFont('DMSans-Bold.ttf', 'DMSans', 'bold');
  doc.addFileToVFS('Cormorant-SemiBold.ttf', data.cormorantSemi);
  doc.addFont('Cormorant-SemiBold.ttf', 'Cormorant', 'normal');
  _registeredDocs.add(doc);
}

/** Register fonts on a jsPDF instance (idempotent per doc, cached fetch). */
export async function registerPdfFonts_(doc) {
  if (_registeredDocs.has(doc)) return;
  const data = await loadFontData_();
  applyFontsToDoc_(doc, data);
}

/** @param {'sans'|'sans-bold'|'serif'} variant */
export function setPdfFont_(doc, variant) {
  if (variant === 'serif') doc.setFont('Cormorant', 'normal');
  else if (variant === 'sans-bold') doc.setFont('DMSans', 'bold');
  else doc.setFont('DMSans', 'normal');
}

export const PDF_FONT_SANS = "'DM Sans', system-ui, sans-serif";
