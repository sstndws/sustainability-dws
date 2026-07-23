/**
 * EUDR DDS PDF export — renders the filled DOCX template to PDF so layout
 * matches EUDR_DDS_Revised.pdf (same pipeline as the Word template).
 */
import { buildFilledDdsDocxBlob_ } from './dds-docx-fill.js';

const A4_W_PT = 595.28;
const A4_H_PT = 841.89;
const DEFAULT_PAGE_H_PX = 1123; // A4 @ 96dpi

function waitForLayout_(ms) {
  return new Promise(function(resolve) {
    requestAnimationFrame(function() {
      requestAnimationFrame(function() {
        setTimeout(resolve, ms || 80);
      });
    });
  });
}

function createRenderHost_() {
  const host = document.createElement('div');
  host.setAttribute('aria-hidden', 'true');
  host.style.cssText = [
    'position:fixed',
    'left:-12000px',
    'top:0',
    'width:794px',
    'background:#fff',
    'z-index:-1',
    'pointer-events:none',
    'overflow:visible',
  ].join(';');

  const styleContainer = document.createElement('div');
  const bodyContainer = document.createElement('div');
  host.appendChild(styleContainer);
  host.appendChild(bodyContainer);
  document.body.appendChild(host);

  return { host: host, styleContainer: styleContainer, bodyContainer: bodyContainer };
}

function injectPdfCaptureStyles_(styleContainer) {
  const style = document.createElement('style');
  style.textContent = [
    '.docx-wrapper{background:#fff!important;padding:0!important;align-items:flex-start!important;}',
    '.docx-wrapper>section.docx{box-shadow:none!important;margin:0!important;}',
  ].join('');
  styleContainer.appendChild(style);
}

function parsePageHeightPx_(pageEl) {
  const style = window.getComputedStyle(pageEl);
  const minH = parseFloat(style.minHeight);
  if (minH > 0) return minH;
  const h = parseFloat(style.height);
  if (h > 0) return h;
  return DEFAULT_PAGE_H_PX;
}

function relativeTop_(el, root) {
  const er = el.getBoundingClientRect();
  const rr = root.getBoundingClientRect();
  return er.top - rr.top + root.scrollTop;
}

function parseRgb_(color) {
  const m = String(color || '').match(/rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i);
  if (!m) return null;
  return { r: +m[1], g: +m[2], b: +m[3] };
}

function isSectionHeaderBg_(color) {
  const rgb = parseRgb_(color);
  if (!rgb) return false;
  return rgb.b > 120 && rgb.r < 80 && rgb.g > 70 && rgb.g < 140;
}

function isInfoBoxBg_(color) {
  const rgb = parseRgb_(color);
  if (!rgb) return false;
  return rgb.r > 240 && rgb.g > 230 && rgb.b < 220;
}

function collectTableRows_(rootEl) {
  return Array.from(rootEl.querySelectorAll('table tr')).map(function(row) {
    const top = relativeTop_(row, rootEl);
    return { top: top, height: row.offsetHeight, bottom: top + row.offsetHeight };
  }).sort(function(a, b) { return a.top - b.top; });
}

function collectBlockElements_(rootEl) {
  const blocks = [];
  rootEl.querySelectorAll('p').forEach(function(p) {
    const bg = window.getComputedStyle(p).backgroundColor;
    if (isSectionHeaderBg_(bg)) {
      blocks.push({
        top: relativeTop_(p, rootEl),
        height: p.offsetHeight,
        kind: 'header',
        minFollow: 120,
      });
    } else if (isInfoBoxBg_(bg)) {
      blocks.push({
        top: relativeTop_(p, rootEl),
        height: p.offsetHeight,
        kind: 'info',
        minFollow: 72,
      });
    }
  });
  rootEl.querySelectorAll('table').forEach(function(table) {
    const top = relativeTop_(table, rootEl);
    blocks.push({
      top: top,
      height: table.offsetHeight,
      kind: 'table',
      minFollow: 48,
    });
  });
  return blocks.sort(function(a, b) { return a.top - b.top; });
}

function adjustSliceEnd_(start, candidateEnd, rows, blocks) {
  let end = candidateEnd;
  const minSlice = 96;

  rows.forEach(function(row) {
    if (row.top >= end || row.bottom <= start + 8) return;
    if (row.top < end && row.bottom > end) {
      const onPage = end - row.top;
      const overflow = row.bottom - end;
      if (onPage >= row.height * 0.55) return;
      if (overflow <= 28) {
        end = row.bottom;
        return;
      }
      if (row.top > start + minSlice) end = Math.min(end, row.top);
    }
  });

  blocks.forEach(function(block) {
    const bottom = block.top + block.height;
    const minFollow = block.minFollow || 80;
    if (block.top >= start + 8 && block.top < end && bottom <= end + 2) {
      if (end - bottom < minFollow && block.top > start + minSlice) {
        end = Math.min(end, block.top);
      }
    }
  });

  if (end <= start + 40) return candidateEnd;
  return end;
}

function computePageSlices_(rootEl, pageHeight) {
  const totalHeight = rootEl.scrollHeight;
  if (totalHeight <= pageHeight + 1) {
    return [{ y: 0, h: totalHeight }];
  }

  const rows = collectTableRows_(rootEl);
  const blocks = collectBlockElements_(rootEl);
  const slices = [];
  let start = 0;

  while (start < totalHeight - 1) {
    let end = Math.min(start + pageHeight, totalHeight);
    if (end >= totalHeight) {
      slices.push({ y: start, h: totalHeight - start });
      break;
    }

    end = adjustSliceEnd_(start, end, rows, blocks);
    slices.push({ y: start, h: end - start });
    start = end;
  }

  return slices;
}

function collectPages_(bodyContainer) {
  const wrapper = bodyContainer.querySelector('.docx-wrapper');
  if (!wrapper) return [];

  const sections = wrapper.querySelectorAll('section.docx');
  if (sections.length) return Array.from(sections);

  return [wrapper];
}

async function captureSectionPages_(snap, pdf, pageEl, startPdfPageIndex) {
  const pageWidth = pageEl.offsetWidth || pageEl.clientWidth;
  const pageHeight = parsePageHeightPx_(pageEl);
  const slices = computePageSlices_(pageEl, pageHeight);
  let pdfPageIndex = startPdfPageIndex;

  for (let i = 0; i < slices.length; i++) {
    const slice = slices[i];
    const canvas = await snap(pageEl, {
      scale: 2,
      useCORS: true,
      logging: false,
      backgroundColor: '#ffffff',
      width: pageWidth,
      height: slice.h,
      windowWidth: pageWidth,
      windowHeight: slice.h,
      x: 0,
      y: slice.y,
    });

    const imgData = canvas.toDataURL('image/jpeg', 0.95);
    if (pdfPageIndex > 0) pdf.addPage();
    const imgHeightPt = (canvas.height * A4_W_PT) / canvas.width;
    pdf.addImage(imgData, 'JPEG', 0, 0, A4_W_PT, Math.min(imgHeightPt, A4_H_PT), undefined, 'FAST');
    pdfPageIndex += 1;
  }

  return pdfPageIndex;
}

function resolveHtml2Canvas_(mod) {
  if (typeof mod === 'function') return mod;
  if (mod && typeof mod.default === 'function') return mod.default;
  throw new Error('html2canvas module tidak valid');
}

/**
 * @param {{ blob: Blob, fileKey: string }} filled
 */
export async function exportDdsPdfFromBlob_(filled) {
  const [{ renderAsync }, html2canvasMod, { jsPDF }] = await Promise.all([
    import('docx-preview'),
    import('html2canvas'),
    import('jspdf'),
  ]);
  const snap = resolveHtml2Canvas_(html2canvasMod);

  const hostInfo = createRenderHost_();
  injectPdfCaptureStyles_(hostInfo.styleContainer);

  try {
    await renderAsync(filled.blob, hostInfo.bodyContainer, hostInfo.styleContainer, {
      className: 'docx',
      inWrapper: true,
      hideWrapperOnPrint: true,
      ignoreWidth: false,
      ignoreHeight: false,
      ignoreFonts: false,
      breakPages: true,
      ignoreLastRenderedPageBreak: false,
      renderHeaders: true,
      renderFooters: true,
      renderFootnotes: true,
      renderEndnotes: true,
      useBase64URL: true,
    });

    await waitForLayout_(180);

    const pages = collectPages_(hostInfo.bodyContainer);
    if (!pages.length) throw new Error('Gagal merender template DDS untuk PDF');

    const pdf = new jsPDF({ orientation: 'portrait', unit: 'pt', format: 'a4', compress: true });
    let pdfPageIndex = 0;

    for (let i = 0; i < pages.length; i++) {
      pdfPageIndex = await captureSectionPages_(snap, pdf, pages[i], pdfPageIndex);
    }

    const name = filled.exportBaseName || ('EUDR_DDS_' + filled.fileKey);
    pdf.save(name + '.pdf');
  } finally {
    hostInfo.host.remove();
  }
}

/**
 * @param {{ master: object, suppliers: object[], geolocation: object[], documents: object[] }} bundle
 */
export async function exportDdsPdf_(bundle) {
  const filled = await buildFilledDdsDocxBlob_(bundle);
  await exportDdsPdfFromBlob_(filled);
}
