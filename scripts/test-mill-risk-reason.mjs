import assert from 'node:assert/strict';
import {
  millRiskReason_,
  millHighRiskReason_,
  millRowIsHighRisk_,
} from '../src/mill-risk-reason.js';

function isNblYes(v) {
  return String(v || '').trim().toLowerCase() === 'yes';
}

const opts = { millIsNblYes: isNblYes };

assert.equal(
  millRiskReason_({
    'BUYER NO BUY LIST': 'Yes',
    'RESULT RISK LEVEL': 'HIGH',
  }, opts),
  'On No Buy List — Result Risk Level elevated to HIGH'
);

assert.equal(
  millRiskReason_({
    'BUYER NO BUY LIST': 'No',
    'COMPLIMENT/NOT COMPLIMENT': 'NC',
    'RESULT RISK LEVEL': 'HIGH',
  }, opts),
  'Legality: Not Compliment (NC) — Result Risk Level is HIGH'
);

assert.equal(
  millRiskReason_({
    'BUYER NO BUY LIST': 'No',
    'COMPLIMENT/NOT COMPLIMENT': 'C',
    'TOTAL SCORE': 1,
    'RESULT RISK LEVEL': 'HIGH',
  }, opts),
  'Compliment (C) with low Total Score (≤ 2) — Result Risk Level is HIGH (Total Score: 1)'
);

assert.equal(
  millRiskReason_({
    'BUYER NO BUY LIST': 'No',
    'COMPLIMENT/NOT COMPLIMENT': 'C',
    'TOTAL SCORE': 3,
    'RESULT RISK LEVEL': 'MEDIUM',
  }, opts),
  'Compliment (C) with moderate Total Score (3) — Result Risk Level is MEDIUM'
);

assert.equal(
  millRiskReason_({
    'BUYER NO BUY LIST': 'No',
    'COMPLIMENT/NOT COMPLIMENT': 'C',
    'TOTAL SCORE': 4,
    'RESULT RISK LEVEL': 'LOW',
  }, opts),
  'Compliment (C) with strong Total Score (≥ 4) — Result Risk Level is LOW (Total Score: 4)'
);

// Reason must follow displayed Result Risk Level, not score formula alone.
assert.equal(
  millRiskReason_({
    'BUYER NO BUY LIST': 'No',
    'COMPLIMENT/NOT COMPLIMENT': 'C',
    'TOTAL SCORE': 3,
    'RESULT RISK LEVEL': 'LOW',
  }, opts),
  'Compliment (C) (Total Score: 3) — Result Risk Level is LOW'
);

assert.equal(
  millRiskReason_({
    'BUYER NO BUY LIST': 'No',
    'COMPLIMENT/NOT COMPLIMENT': 'C',
    'TOTAL SCORE': 2,
    'RESULT RISK LEVEL': 'MEDIUM',
  }, opts),
  'Compliment (C) (Total Score: 2) — Result Risk Level is MEDIUM'
);

assert.equal(
  millHighRiskReason_({
    'BUYER NO BUY LIST': 'No',
    'COMPLIMENT/NOT COMPLIMENT': 'C',
    'TOTAL SCORE': 5,
    'RESULT RISK LEVEL': 'LOW',
  }, opts),
  'Compliment (C) with strong Total Score (≥ 4) — Result Risk Level is LOW (Total Score: 5)'
);

assert.equal(millRowIsHighRisk_({ 'RESULT RISK LEVEL': 'HIGH' }), true);
assert.equal(millRowIsHighRisk_({ 'RISK LEVEL': 'High' }), true);

console.log('test-mill-risk-reason: OK');
