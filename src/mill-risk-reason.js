/**
 * Risk Reason — per-row explanation from Mill Onboarding sheet gaps.
 *
 * Risk factors (what is missing / failing):
 *   Coordinate     → No Coordinate
 *   Legality       → Legality Not Complete
 *   APL            → Non APL Area
 *   NDPE           → No NDPE Commitment
 *   Certification  → No Certification
 *
 * Pengurangan (risk reduction factors on the sheet):
 *   Deforestasi            → 1
 *   High Deforestation     → 2
 *   Have Grievance (Legality, Enviro)
 *
 * RESULT RISK LEVEL still drives HIGH/MEDIUM/LOW display; NBL can elevate to HIGH.
 */

const FIELD_ALIASES = {
  nbl: ['BUYER NO BUY LIST', 'Buyer No Buy List'],
  resultRisk: ['RESULT RISK LEVEL', 'Result Risk Level'],
  riskLevel: ['RISK LEVEL', 'Risk Level'],
  compliment: ['COMPLIMENT/NOT COMPLIMENT', 'Compliment/Not Compliment'],
  totalScore: ['TOTAL SCORE', 'Total Score'],
  coordinate: ['COORDINATES', 'Coordinates', 'COORDINATE', 'Coordinate'],
  millLoc: ['MILL LOC', 'MILL LOCATION', 'Mill Location', 'LOC'],
  ndpe: ['NDPE', 'NDPE COMMITMENT', 'NDPE Policy'],
  certification: ['CERTIFICATION', 'Certification'],
  totalCertification: ['TOTAL CERTIFICATION', 'Total Certification'],
  deforestationWidth: ['DEFORESTATION WIDTH', 'Deforestation Width'],
  riskReduction: ['RISK REDUCTION FACTOR', 'Risk Reduction Factor'],
  legalityDocs: ['HGU/HGB', 'IZIN LOKASI', 'IUP', 'IZIN LINGKUNGAN'],
  legalityGrievance: [
    'LEGALITY GRIEVANCE', 'LEGALITY GRIEVANCES', 'LEGALITY', 'Legality',
  ],
  environmentGrievance: [
    'ENVIRONMENT GRIEVANCE', 'ENVIRONMENT GRIEVANCES', 'ENVIRONMENT', 'Environment',
  ],
};

/** Ha threshold aligned with mill form Medium tier for deforestation width. */
const HIGH_DEFORESTATION_HA = 25;

function pickRowField_(row, keys) {
  if (!row || typeof row !== 'object') return '';
  const list = keys || [];
  for (let i = 0; i < list.length; i++) {
    const k = list[i];
    if (row[k] != null && String(row[k]).trim() !== '') return String(row[k]).trim();
  }
  return '';
}

function defaultIsYes_(val) {
  const s = String(val == null ? '' : val).trim().toLowerCase();
  return s === 'yes' || s === 'y' || s === '1' || s === 'true' || /\bnbl\b/.test(s) || s.includes('no buy');
}

function isExplicitNo_(val) {
  const s = String(val == null ? '' : val).trim().toLowerCase();
  return s === 'no' || s === 'n' || s === '0' || s === 'false' || s === 'tidak';
}

function isBlankCell_(val) {
  if (val == null) return true;
  const s = String(val).trim();
  return !s || s === '—' || s === '-' || /^no\s*data$/i.test(s) || /^n\/?a$/i.test(s);
}

function normalizeComplimentCode_(val) {
  const raw = String(val == null ? '' : val).trim();
  if (!raw) return '';
  const upper = raw.toUpperCase();
  if (upper === 'NC' || /NOT\s*COMPLIMENT/.test(upper)) return 'NC';
  if (upper === 'C' || /^COMPLIMENT$/.test(upper)) return 'C';
  return upper;
}

function parseTotalScore_(val) {
  if (val === '' || val == null) return null;
  const n = Number(String(val).replace(/,/g, '').trim());
  return Number.isFinite(n) ? n : null;
}

function normalizeRiskBand_(val) {
  const upper = String(val == null ? '' : val).trim().toUpperCase();
  if (!upper) return '';
  if (upper.includes('HIGH')) return 'HIGH';
  if (upper.includes('MEDIUM') || upper.includes('MED')) return 'MEDIUM';
  if (upper.includes('LOW')) return 'LOW';
  return upper;
}

function parseHaWidth_(raw) {
  if (isBlankCell_(raw)) return NaN;
  const n = parseFloat(String(raw).replace(/[^\d.,-]/g, '').replace(',', '.'));
  return Number.isFinite(n) ? n : NaN;
}

function hasValidCoordinate_(row) {
  const raw = pickRowField_(row, FIELD_ALIASES.coordinate);
  if (isBlankCell_(raw)) return false;
  if (/^(none|null|-+)$/i.test(raw)) return false;
  // Accept "lat, lng" or similar numeric pair.
  const nums = String(raw).match(/-?\d+(?:[.,]\d+)?/g);
  return !!(nums && nums.length >= 2);
}

function legalityNotComplete_(row) {
  const compliment = normalizeComplimentCode_(pickRowField_(row, FIELD_ALIASES.compliment));
  if (compliment === 'NC') return true;
  // Any legality doc explicitly No → incomplete.
  for (let i = 0; i < FIELD_ALIASES.legalityDocs.length; i++) {
    const k = FIELD_ALIASES.legalityDocs[i];
    const v = row && row[k] != null ? String(row[k]).trim() : '';
    if (isExplicitNo_(v)) return true;
  }
  return false;
}

function isNonApl_(row) {
  const loc = pickRowField_(row, FIELD_ALIASES.millLoc);
  if (isBlankCell_(loc)) return false;
  const u = loc.toUpperCase().replace(/\s+/g, ' ');
  if (u.includes('NON APL') || u.includes('NON-APL') || u === 'NONAPL') return true;
  if (u === 'APL' || u.includes('APL')) return false;
  return false;
}

function hasNoNdpe_(row) {
  const v = pickRowField_(row, FIELD_ALIASES.ndpe);
  if (isBlankCell_(v)) return true;
  return isExplicitNo_(v);
}

function hasNoCertification_(row) {
  const cert = pickRowField_(row, FIELD_ALIASES.certification);
  const total = pickRowField_(row, FIELD_ALIASES.totalCertification);
  if (!isBlankCell_(cert)) {
    const u = cert.toUpperCase();
    if (u === 'NO' || u === 'NONE' || u === 'N/A' || u === '-' || u === '—') return true;
    return false;
  }
  if (!isBlankCell_(total)) {
    const n = parseTotalScore_(total);
    if (n === 0) return true;
    if (n != null && n > 0) return false;
  }
  return true;
}

function grievanceYes_(row, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const v = row && row[aliases[i]] != null ? String(row[aliases[i]]).trim() : '';
    if (defaultIsYes_(v)) return true;
  }
  return false;
}

/**
 * Collect Risk / Pengurangan phrases from the sheet row (handwritten mapping).
 * @returns {string[]}
 */
export function millRiskReasonGaps_(row) {
  const gaps = [];
  if (!hasValidCoordinate_(row)) gaps.push('No Coordinate');
  if (legalityNotComplete_(row)) gaps.push('Legality Not Complete');
  if (isNonApl_(row)) gaps.push('Non APL Area');
  if (hasNoNdpe_(row)) gaps.push('No NDPE Commitment');
  if (hasNoCertification_(row)) gaps.push('No Certification');

  const rrf = parseTotalScore_(pickRowField_(row, FIELD_ALIASES.riskReduction));
  const defHa = parseHaWidth_(pickRowField_(row, FIELD_ALIASES.deforestationWidth));
  if (rrf === 2 || (!Number.isNaN(defHa) && defHa > HIGH_DEFORESTATION_HA)) {
    gaps.push('High Deforestation');
  } else if (rrf === 1 || (!Number.isNaN(defHa) && defHa > 0)) {
    gaps.push('Deforestation');
  }

  const hasLegGrv = grievanceYes_(row, FIELD_ALIASES.legalityGrievance);
  const hasEnvGrv = grievanceYes_(row, FIELD_ALIASES.environmentGrievance);
  if (hasLegGrv && hasEnvGrv) {
    gaps.push('Have Grievance (Legality, Enviro)');
  } else if (hasLegGrv) {
    gaps.push('Have Grievance (Legality)');
  } else if (hasEnvGrv) {
    gaps.push('Have Grievance (Enviro)');
  }

  return gaps;
}

export function millResolvedRiskLevelFromRow_(row) {
  const rr = pickRowField_(row, FIELD_ALIASES.resultRisk);
  if (rr) return rr;
  return pickRowField_(row, FIELD_ALIASES.riskLevel);
}

export function millRowIsHighRisk_(row) {
  return String(millResolvedRiskLevelFromRow_(row)).toLowerCase().includes('high');
}

function scorePhrase_(totalScore) {
  if (totalScore === null) return '';
  return ' (Total Score: ' + totalScore + ')';
}

function fallbackBandReason_(resolved, compliment, totalScore) {
  const scoreTxt = scorePhrase_(totalScore);
  if (resolved === 'HIGH') {
    if (compliment === 'NC') {
      return 'Legality: Not Compliment (NC) — Result Risk Level is HIGH';
    }
    if (compliment === 'C' && totalScore !== null && totalScore <= 2) {
      return 'Compliment (C) with low Total Score (≤ 2) — Result Risk Level is HIGH' + scoreTxt;
    }
    return 'Assessment outcome — Result Risk Level is HIGH';
  }
  if (resolved === 'MEDIUM') {
    if (compliment === 'C' && totalScore === 3) {
      return 'Compliment (C) with moderate Total Score (3) — Result Risk Level is MEDIUM';
    }
    return 'Assessment outcome — Result Risk Level is MEDIUM';
  }
  if (resolved === 'LOW') {
    if (compliment === 'C' && totalScore !== null && totalScore >= 4) {
      return 'Compliment (C) with strong Total Score (≥ 4) — Result Risk Level is LOW' + scoreTxt;
    }
    return 'Assessment outcome — Result Risk Level is LOW';
  }
  return resolved ? ('Result Risk Level: ' + resolved) : '';
}

/**
 * @param {object} row — mill sheet row
 * @param {{ millIsNblYes?: (val: unknown) => boolean }} [opts]
 * @returns {string} English explanation for Risk Reason column
 */
export function millRiskReason_(row, opts) {
  opts = opts || {};
  const isNblYes = typeof opts.millIsNblYes === 'function' ? opts.millIsNblYes : defaultIsYes_;

  const resolved = normalizeRiskBand_(millResolvedRiskLevelFromRow_(row));
  const nbl = pickRowField_(row, FIELD_ALIASES.nbl);
  const gaps = millRiskReasonGaps_(row);
  const parts = [];

  if (isNblYes(nbl)) {
    parts.push('On No Buy List — Result Risk Level elevated to HIGH');
  }

  if (gaps.length) {
    parts.push(gaps.join('; '));
  }

  if (parts.length) return parts.join(' · ');

  // No sheet gaps detected — fall back to band / score wording.
  if (!resolved) return '';
  const compliment = normalizeComplimentCode_(pickRowField_(row, FIELD_ALIASES.compliment));
  const totalScore = parseTotalScore_(pickRowField_(row, FIELD_ALIASES.totalScore));
  return fallbackBandReason_(resolved, compliment, totalScore);
}

/** @deprecated Use millRiskReason_ — kept for callers not yet migrated. */
export function millHighRiskReason_(row, opts) {
  return millRiskReason_(row, opts);
}
