'use strict';
// Every field the MSA-4676 map asks for must exist in the data dictionary. A missing key is silent:
// the lookup returns undefined, the stamper skips the field, and the CERTIFIED form goes out blank.
// That is what happened to "Start of Service" — nothing in the app or the compose window showed it,
// because the compose modal renders To/Subject/Body and never the PDF.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

test('the MSA-4676 stamps a Start of Service date', () => {
  const w = loadApp(); resetStorage(w);
  w.saveProfilesLS({ Jane: { clientName: 'Jane', firstName: 'Jane', lastName: 'Doe',
    medicaidId: '1234567', startDate: '2026-01-15', street: '1 Main', city: 'Detroit',
    state: 'MI', zip: '48201', county: 'Wayne', phone: '313-555-1000' } });
  const dict = w._buildFormDataDict('Jane');
  assert.ok(dict, 'the dictionary must build');
  assert.ok(dict.start_date, 'msa_start maps to start_date, which the dictionary never defined');
  assert.match(String(dict.start_date), /2026/, 'it must carry the client’s actual start date');
});

test('EVERY field the 4676 map asks for resolves in the dictionary', () => {
  const w = loadApp(); resetStorage(w);
  w.saveProfilesLS({ Jane: { clientName: 'Jane', firstName: 'Jane', lastName: 'Doe',
    startDate: '2026-01-15', medicaidId: '1', city: 'Detroit', state: 'MI', zip: '48201' } });
  const dict = w._buildFormDataDict('Jane') || {};
  const map = (w.STATE_FORM_INPUTS && w.STATE_FORM_INPUTS.msa4676) ||
              (w.STATE_FORMS && w.STATE_FORMS.msa4676 && w.STATE_FORMS.msa4676.inputs);
  if (!map) { assert.ok(dict.start_date, 'fallback: at least the regressed field is present'); return; }
  const missing = Object.keys(map).map((k) => map[k]).filter((key) => !(key in dict));
  assert.deepStrictEqual(missing, [],
    'these 4676 fields would stamp BLANK on a certified form: ' + missing.join(', '));
});
