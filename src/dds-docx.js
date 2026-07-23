/**
 * EUDR DDS DOCX export — fills the official template (same source as PDF).
 */
import { buildFilledDdsDocxBlob_ } from './dds-docx-fill.js';

function downloadBlob_(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.rel = 'noopener';
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(function() { URL.revokeObjectURL(url); }, 1000);
}

/**
 * @param {{ blob: Blob, fileKey: string }} filled
 */
export function downloadFilledDdsDocx_(filled) {
  const name = filled.exportBaseName || ('EUDR_DDS_' + filled.fileKey);
  downloadBlob_(filled.blob, name + '.docx');
}

/**
 * @param {{ master: object, suppliers: object[], geolocation: object[], documents: object[] }} bundle
 */
export async function exportDdsDocx_(bundle) {
  const filled = await buildFilledDdsDocxBlob_(bundle);
  downloadFilledDdsDocx_(filled);
}
