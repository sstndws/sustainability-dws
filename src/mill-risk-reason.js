/**
 * Risk-level explanation — anchored to Result Risk Level shown in the UI (English copy).
 *
 * RESULT RISK LEVEL:
 *   IF BUYER NO BUY LIST = Yes → HIGH
 *   ELSE → RISK LEVEL
 *
 * RISK LEVEL (when Compliment = C):
 *   NC → HIGH · score ≤ 2 → HIGH · score 3 → MEDIUM · score ≥ 4 → LOW
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

function scorePhrase_(totalScore) {
  if (totalScore === null) return '';
  return ' (Total Score: ' + totalScore + ')';
}

/**
 * @param {object} row — mill sheet row
 * @param {{ millIsNblYes?: (val: unknown) => boolean }} [opts]
 * @returns {string} English explanation for the displayed Result Risk Level
 */
export function millRiskReason_(row, opts) {
  opts = opts || {};
  const isNblYes = typeof opts.millIsNblYes === 'function' ? opts.millIsNblYes : defaultIsYes_;

  const resolved = normalizeRiskBand_(millResolvedRiskLevelFromRow_(row));
  if (!resolved) return '';

  const nbl = pickRowField_(row, FIELD_ALIASES.nbl);
  const compliment = normalizeComplimentCode_(pickRowField_(row, FIELD_ALIASES.compliment));
  const totalScore = parseTotalScore_(pickRowField_(row, FIELD_ALIASES.totalScore));
  const scoreTxt = scorePhrase_(totalScore);

  if (resolved === 'HIGH') {
    if (isNblYes(nbl)) {
      return 'On No Buy List — Result Risk Level elevated to HIGH';
    }
    if (compliment === 'NC') {
      return 'Legality: Not Compliment (NC) — Result Risk Level is HIGH';
    }
    if (compliment === 'C' && totalScore !== null && totalScore <= 2) {
      return 'Compliment (C) with low Total Score (≤ 2) — Result Risk Level is HIGH' + scoreTxt;
    }
    if (compliment === 'C' && totalScore !== null) {
      return 'Compliment (C)' + scoreTxt + ' — Result Risk Level is HIGH';
    }
    return 'Assessment outcome — Result Risk Level is HIGH';
  }

  if (resolved === 'MEDIUM') {
    if (compliment === 'C' && totalScore === 3) {
      return 'Compliment (C) with moderate Total Score (3) — Result Risk Level is MEDIUM';
    }
    if (compliment === 'C' && totalScore !== null) {
      return 'Compliment (C)' + scoreTxt + ' — Result Risk Level is MEDIUM';
    }
    if (compliment === 'NC') {
      return 'Legality flagged — Result Risk Level is MEDIUM';
    }
    return 'Assessment outcome — Result Risk Level is MEDIUM';
  }

  if (resolved === 'LOW') {
    if (compliment === 'C' && totalScore !== null && totalScore >= 4) {
      return 'Compliment (C) with strong Total Score (≥ 4) — Result Risk Level is LOW' + scoreTxt;
    }
    if (compliment === 'C' && totalScore !== null) {
      return 'Compliment (C)' + scoreTxt + ' — Result Risk Level is LOW';
    }
    return 'Assessment outcome — Result Risk Level is LOW';
  }

  return 'Result Risk Level: ' + resolved;
}

/** @deprecated Use millRiskReason_ — kept for callers not yet migrated. */
export function millHighRiskReason_(row, opts) {
  return millRiskReason_(row, opts);
}
