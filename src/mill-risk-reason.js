/**
 * Risk-level explanation — mirrors Mill Onboarding sheet formulas (English copy).
 *
 * RESULT RISK LEVEL:
 *   IF BUYER NO BUY LIST = Yes → HIGH
 *   ELSE → RISK LEVEL
 *
 * RISK LEVEL:
 *   IF COMPLIMENT/NOT COMPLIMENT = NC → HIGH
 *   IF C and TOTAL SCORE >= 4 → LOW
 *   IF C and TOTAL SCORE <= 2 → HIGH
 *   IF C and TOTAL SCORE = 3 → MEDIUM
 */

const FIELD_ALIASES = {
  nbl: ['BUYER NO BUY LIST', 'Buyer No Buy List'],
  resultRisk: ['RESULT RISK LEVEL', 'Result Risk Level'],
  riskLevel: ['RISK LEVEL', 'Risk Level'],
  compliment: ['COMPLIMENT/NOT COMPLIMENT', 'Compliment/Not Compliment'],
  totalScore: ['TOTAL SCORE', 'Total Score'],
};

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

export function millResolvedRiskLevelFromRow_(row) {
  const rr = pickRowField_(row, FIELD_ALIASES.resultRisk);
  if (rr) return rr;
  return pickRowField_(row, FIELD_ALIASES.riskLevel);
}

export function millRowIsHighRisk_(row) {
  return String(millResolvedRiskLevelFromRow_(row)).toLowerCase().includes('high');
}

/**
 * @param {object} row — mill sheet row
 * @param {{ millIsNblYes?: (val: unknown) => boolean }} [opts]
 * @returns {string} English explanation for the resolved risk level
 */
export function millRiskReason_(row, opts) {
  opts = opts || {};
  const isNblYes = typeof opts.millIsNblYes === 'function' ? opts.millIsNblYes : defaultIsYes_;

  const resolved = normalizeRiskBand_(millResolvedRiskLevelFromRow_(row));
  if (!resolved) return '';

  const nbl = pickRowField_(row, FIELD_ALIASES.nbl);
  const compliment = normalizeComplimentCode_(pickRowField_(row, FIELD_ALIASES.compliment));
  const totalScore = parseTotalScore_(pickRowField_(row, FIELD_ALIASES.totalScore));

  if (isNblYes(nbl)) {
    return 'Listed on No Buy List — Result Risk Level set to HIGH';
  }

  if (compliment === 'NC') {
    return 'Legality status: Not Compliment (NC) — classified as HIGH risk';
  }

  if (compliment === 'C' && totalScore !== null) {
    if (totalScore <= 2) {
      return 'Compliment (C) with Total Score ≤ 2 — classified as HIGH risk (score: ' + totalScore + ')';
    }
    if (totalScore === 3) {
      return 'Compliment (C) with Total Score 3 — classified as MEDIUM risk';
    }
    if (totalScore >= 4) {
      return 'Compliment (C) with Total Score ≥ 4 — classified as LOW risk (score: ' + totalScore + ')';
    }
  }

  if (resolved === 'HIGH') {
    return 'Classified as HIGH risk based on assessment score and legality status';
  }
  if (resolved === 'MEDIUM') {
    return 'Classified as MEDIUM risk based on assessment score and legality status';
  }
  if (resolved === 'LOW') {
    return 'Classified as LOW risk based on assessment score and legality status';
  }

  return 'Risk level: ' + resolved;
}

/** @deprecated Use millRiskReason_ — kept for callers not yet migrated. */
export function millHighRiskReason_(row, opts) {
  return millRiskReason_(row, opts);
}
