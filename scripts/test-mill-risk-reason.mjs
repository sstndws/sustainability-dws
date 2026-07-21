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
  'Listed on No Buy List — Result Risk Level set to HIGH'
);

assert.equal(
  millRiskReason_({
    'BUYER NO BUY LIST': 'No',
    'COMPLIMENT/NOT COMPLIMENT': 'NC',
    'RESULT RISK LEVEL': 'HIGH',
    'RISK LEVEL': 'HIGH',
  }, opts),
  'Legality status: Not Compliment (NC) — classified as HIGH risk'
);

assert.equal(
  millRiskReason_({
    'BUYER NO BUY LIST': 'No',
    'COMPLIMENT/NOT COMPLIMENT': 'C',
    'TOTAL SCORE': 1,
    'RESULT RISK LEVEL': 'HIGH',
    'RISK LEVEL': 'HIGH',
  }, opts),
  'Compliment (C) with Total Score ≤ 2 — classified as HIGH risk (score: 1)'
);

assert.equal(
  millRiskReason_({
    'BUYER NO BUY LIST': 'No',
    'COMPLIMENT/NOT COMPLIMENT': 'C',
    'TOTAL SCORE': 3,
    'RESULT RISK LEVEL': 'MEDIUM',
    'RISK LEVEL': 'MEDIUM',
  }, opts),
  'Compliment (C) with Total Score 3 — classified as MEDIUM risk'
);

assert.equal(
  millRiskReason_({
    'BUYER NO BUY LIST': 'No',
    'COMPLIMENT/NOT COMPLIMENT': 'C',
    'TOTAL SCORE': 4,
    'RESULT RISK LEVEL': 'LOW',
    'RISK LEVEL': 'LOW',
  }, opts),
  'Compliment (C) with Total Score ≥ 4 — classified as LOW risk (score: 4)'
);

assert.equal(
  millHighRiskReason_({
    'BUYER NO BUY LIST': 'No',
    'COMPLIMENT/NOT COMPLIMENT': 'C',
    'TOTAL SCORE': 5,
    'RESULT RISK LEVEL': 'LOW',
  }, opts),
  'Compliment (C) with Total Score ≥ 4 — classified as LOW risk (score: 5)'
);

assert.equal(millRowIsHighRisk_({ 'RESULT RISK LEVEL': 'HIGH' }), true);
assert.equal(millRowIsHighRisk_({ 'RISK LEVEL': 'High' }), true);

console.log('test-mill-risk-reason: OK');
