'use strict';
// PHI minimization on caregiver-facing outputs (2026-08-12): the task sheet + chart image show the
// client as first name + last initial ("Darnelle D.") instead of the full name, so less identifier
// leaves the system when texted/emailed to a caregiver.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp } = require('./harness');

test('_caregiverClientLabel: prefers structured first/last -> "First L."', () => {
  const w = loadApp();
  assert.strictEqual(w._caregiverClientLabel({ firstName: 'Darnelle', lastName: 'Dubose' }, 'Darnelle Dubose'), 'Darnelle D.');
  assert.strictEqual(w._caregiverClientLabel({ firstName: 'Mary', lastName: 'oConnor' }, 'x'), 'Mary O.', 'last initial upper-cased');
});

test('_caregiverClientLabel: falls back to parsing the full name when fields are blank', () => {
  const w = loadApp();
  assert.strictEqual(w._caregiverClientLabel({}, 'Darnelle Dubose'), 'Darnelle D.');
  assert.strictEqual(w._caregiverClientLabel({}, 'Darnelle Marie Dubose'), 'Darnelle D.', 'uses the LAST token as surname');
});

test('_caregiverClientLabel: single name or empty is handled without a stray initial', () => {
  const w = loadApp();
  assert.strictEqual(w._caregiverClientLabel({ firstName: 'Cher' }, 'Cher'), 'Cher', 'no last name -> no initial');
  assert.strictEqual(w._caregiverClientLabel({}, 'Madonna'), 'Madonna');
  assert.strictEqual(w._caregiverClientLabel({}, ''), '', 'empty stays empty');
});
