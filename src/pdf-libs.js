// jsPDF + AutoTable bundled at build time (no CDN / script-tag race).
import { jsPDF } from 'jspdf';
import 'jspdf-autotable';

let _verified = false;

function verifyAutoTable_() {
  if (_verified) return;
  const probe = new jsPDF({ unit: 'mm', format: 'a4' });
  if (typeof probe.autoTable !== 'function') {
    throw new Error('PDF AutoTable plugin failed to initialize.');
  }
  _verified = true;
}

/** Returns the jsPDF constructor with AutoTable already attached. */
export function getJsPDF() {
  verifyAutoTable_();
  return jsPDF;
}
