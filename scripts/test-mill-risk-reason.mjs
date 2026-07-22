import assert from 'node:assert/strict';
import {
  millRiskReason_,
  millRiskReasonGaps_,
  millHighRiskReason_,
  millRowIsHighRisk_,
} from '../src/mill-risk-reason.js';

function isNblYes(v) {
  return String(v || '').trim().toLowerCase() === 'yes';
}

const opts = { millIsNblYes: isNblYes };

const cleanLow = {
  'BUYER NO BUY LIST': 'No',
  'RESULT RISK LEVEL': 'LOW',
  'COMPLIMENT/NOT COMPLIMENT': 'C',
  'TOTAL SCORE': 4,
  COORDINATES: '0.11, 111.22',
  'MILL LOC': 'APL',
  NDPE: 'Yes',
  CERTIFICATION: 'ISPO',
  'DEFORESTATION WIDTH': 0,
  'LEGALITY GRIEVANCE': 'No',
  'ENVIRONMENT GRIEVANCE': 'No',
  'HGU/HGB': 'Yes',
  'IZIN LOKASI': 'Yes',
  IUP: 'Yes',
  'IZIN LINGKUNGAN': 'Yes',
};

assert.deepEqual(millRiskReasonGaps_(cleanLow), []);
assert.equal(
  millRiskReason_(cleanLow, opts),
  'Compliment (C) with strong Total Score (≥ 4) — Result Risk Level is LOW (Total Score: 4)'
);

assert.equal(
  millRiskReason_({
    ...cleanLow,
    'BUYER NO BUY LIST': 'Yes',
    'RESULT RISK LEVEL': 'HIGH',
  }, opts),
  'On No Buy List'
);

assert.equal(
  millRiskReason_({
    ...cleanLow,
    'COMPLIMENT/NOT COMPLIMENT': 'NC',
    'RESULT RISK LEVEL': 'HIGH',
  }, opts),
  'Legality Not Complete'
);

assert.deepEqual(
  millRiskReasonGaps_({
    COORDINATES: '',
    'MILL LOC': 'Non APL',
    NDPE: 'No',
    CERTIFICATION: '',
    'COMPLIMENT/NOT COMPLIMENT': 'C',
    'HGU/HGB': 'Yes',
    'IZIN LOKASI': 'Yes',
    IUP: 'Yes',
    'IZIN LINGKUNGAN': 'Yes',
  }),
  ['No Coordinate', 'Non APL Area', 'No NDPE Commitment', 'No Certification']
);

assert.equal(
  millRiskReason_({
    ...cleanLow,
    COORDINATES: '',
    NDPE: 'No',
    CERTIFICATION: '',
    'RESULT RISK LEVEL': 'HIGH',
  }, opts),
  'No Coordinate; No NDPE Commitment; No Certification'
);

assert.deepEqual(
  millRiskReasonGaps_({
    ...cleanLow,
    'DEFORESTATION WIDTH': 10,
    'RISK REDUCTION FACTOR': 1,
  }),
  ['Deforestation']
);

assert.deepEqual(
  millRiskReasonGaps_({
    ...cleanLow,
    'DEFORESTATION WIDTH': 40,
  }),
  ['High Deforestation']
);

assert.deepEqual(
  millRiskReasonGaps_({
    ...cleanLow,
    'LEGALITY GRIEVANCE': 'Yes',
    'ENVIRONMENT GRIEVANCE': 'Yes',
  }),
  ['Have Grievance (Legality, Enviro)']
);

assert.equal(
  millRiskReason_({
    ...cleanLow,
    'BUYER NO BUY LIST': 'Yes',
    'RESULT RISK LEVEL': 'HIGH',
    COORDINATES: '',
    NDPE: 'No',
  }, opts),
  'On No Buy List; No Coordinate; No NDPE Commitment'
);

assert.equal(millRowIsHighRisk_({ 'RESULT RISK LEVEL': 'HIGH' }), true);
assert.equal(millRowIsHighRisk_({ 'RISK LEVEL': 'High' }), true);
assert.equal(
  millHighRiskReason_(cleanLow, opts),
  millRiskReason_(cleanLow, opts)
);

console.log('test-mill-risk-reason: OK');
