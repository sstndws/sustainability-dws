import assert from 'node:assert/strict';
import {
  millHighRiskReason_,
  millRowIsHighRisk_,
} from '../src/mill-risk-reason.js';

function isNblYes(v) {
  return String(v || '').trim().toLowerCase() === 'yes';
}

assert.equal(
  millHighRiskReason_({
    'BUYER NO BUY LIST': 'Yes',
    'RESULT RISK LEVEL': 'HIGH',
  }, { millIsNblYes: isNblYes }),
  'No Buy List = Yes (Result Risk Level forced to HIGH)'
);

assert.equal(
  millHighRiskReason_({
    'BUYER NO BUY LIST': 'No',
    'COMPLIMENT/NOT COMPLIMENT': 'NC',
    'RESULT RISK LEVEL': 'HIGH',
    'RISK LEVEL': 'HIGH',
  }, { millIsNblYes: isNblYes }),
  'Legality: Not Compliment (NC) — Risk Level = HIGH'
);

assert.equal(
  millHighRiskReason_({
    'BUYER NO BUY LIST': 'No',
    'COMPLIMENT/NOT COMPLIMENT': 'C',
    'TOTAL SCORE': 1,
    'RESULT RISK LEVEL': 'HIGH',
    'RISK LEVEL': 'HIGH',
  }, { millIsNblYes: isNblYes }),
  'Compliment (C) with Total Score ≤ 2 — Risk Level = HIGH (score: 1)'
);

assert.equal(
  millHighRiskReason_({
    'BUYER NO BUY LIST': 'No',
    'COMPLIMENT/NOT COMPLIMENT': 'C',
    'TOTAL SCORE': 4,
    'RESULT RISK LEVEL': 'LOW',
  }, { millIsNblYes: isNblYes }),
  ''
);

assert.equal(millRowIsHighRisk_({ 'RESULT RISK LEVEL': 'HIGH' }), true);
assert.equal(millRowIsHighRisk_({ 'RISK LEVEL': 'High' }), true);

console.log('test-mill-risk-reason: OK');
