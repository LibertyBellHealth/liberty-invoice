'use strict';
// The assistant can write values the rest of the app cannot represent, and it is the one surface
// where a model decision reaches PHI. These four gates came out of an adversarial sweep of it.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

test('a free-text date is refused instead of being stored and later blanked', () => {
  const w = loadApp(); resetStorage(w);
  w.saveProfilesLS({ Jane: { clientName: 'Jane' } });
  const bad = w._asstUpdateClient({ client_name: 'Jane', field: 'dob', value: 'May 12 1950' });
  assert.ok(bad && bad.error,
    'stored raw it looks right in the confirm dialog, then <input type="date"> reports "" and the next save blanks it');
});

test('an accepted date is normalised to the format the date input can hold', () => {
  const w = loadApp(); resetStorage(w);
  w.saveProfilesLS({ Jane: { clientName: 'Jane' } });
  let shown = '';
  w.showConfirm = (msg) => { shown = String(msg); };
  w._asstUpdateClient({ client_name: 'Jane', field: 'start_date', value: '05/12/2026' });
  assert.match(shown, /2026-05-12/, 'the value put in front of the owner must be the value that will be stored');
});

test('the update whitelist rejects inherited property names', () => {
  const w = loadApp(); resetStorage(w);
  w.saveProfilesLS({ Jane: { clientName: 'Jane' } });
  const r = w._asstUpdateClient({ client_name: 'Jane', field: 'constructor', value: 'x' });
  assert.ok(r && r.error, "'constructor' resolved through the prototype and passed the whitelist");
});

test('a 4676 with no recipient is refused, never defaulted', () => {
  const w = loadApp(); resetStorage(w);
  w.saveProfilesLS({ Jane: { clientName: 'Jane', caregiverId: 'cg1' } });
  w.saveCaregiversLS({ cg1: { id: 'cg1', name: 'Ann Lee', email: 'ann@example.com' } });
  return Promise.resolve(w._asstOpenEmail({ client_name: 'Jane', attach_form: 'MSA-4676',
    subject: 's', body: 'b' })).then((r) => {
    assert.ok(r && r.error && /goes to/.test(r.error),
      'the model was told an empty recipient routes to the caseworker while the code sent to the caregiver — ' +
      'so it reported the wrong recipient to the owner');
  });
});
