'use strict';
// Saving a caseworker read the (hidden) agency input and then BLANKED the value for any org that
// wasn't exactly 'MDHHS'. Every caseworker created before the `org` field existed has org === '',
// so editing ANY of them — even just correcting an email — destroyed the agency. The agency is the
// invoice "Bill To", so the first symptom was every one of that caseworker's invoices refusing to
// send with "Caseworker has no Agency set". Found by using the app, not by any test.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

const FIELDS = ['first','middle','last','title','agency','org','phone','fax','email',
                'street','city','state','zip','county','supervisor'];

function seedCwForm(w, cw, formValues) {
  resetStorage(w);
  if (!w.document.getElementById('cwi-first')) {
    w.document.body.insertAdjacentHTML('beforeend',
      FIELDS.map((f) => '<input id="cwi-' + f + '">').join('') +
      // written to AFTER the save; present only so the function runs to completion
      '<div id="cwDetailName"></div><div id="cwDetailMeta"></div><button id="cwSaveInfoBtn"></button>');
  }
  FIELDS.forEach((f) => { w.document.getElementById('cwi-' + f).value = ''; });
  Object.keys(formValues || {}).forEach((f) => { w.document.getElementById('cwi-' + f).value = formValues[f]; });
  w.saveCaseworkersLS([cw]);
  w.activeCwId = cw.id;
  w.showAlert = () => {}; w.showToast = () => {};
  w.saveCaseworkerAPI = () => Promise.resolve();
  w.addAuditEntry = () => {};        // posts to /audit; not what this test is about
  ['renderCaseworkerList','renderCwInfoPane','renderCaseworkers','updateStats']
    .forEach((fn) => { if (typeof w[fn] === 'function') w[fn] = () => {}; });
}

test('editing a legacy caseworker (org unset) keeps its agency', () => {
  const w = loadApp();
  seedCwForm(w, { id: 'cw1', name: 'Renee Feldman', agency: 'MDHHS - Wayne', org: '' },
             { first: 'Renee', last: 'Feldman', email: 'new@michigan.gov', org: '' });
  w.saveCwInfoPane();
  const saved = w.getCaseworkers().find((c) => c.id === 'cw1');
  assert.strictEqual(saved.agency, 'MDHHS - Wayne',
    'changing only the email destroyed the Bill To on every pre-org caseworker');
  assert.strictEqual(saved.email, 'new@michigan.gov', 'the edit itself must still apply');
});

test('an MDHHS caseworker can still edit its agency', () => {
  const w = loadApp();
  seedCwForm(w, { id: 'cw2', name: 'Marcus Ojeda', agency: 'MDHHS - Oakland', org: 'MDHHS' },
             { first: 'Marcus', last: 'Ojeda', org: 'MDHHS', agency: 'MDHHS - Macomb' });
  w.saveCwInfoPane();
  assert.strictEqual(w.getCaseworkers().find((c) => c.id === 'cw2').agency, 'MDHHS - Macomb',
    'the visible field must still be authoritative when org is MDHHS');
});

test('an MDHHS caseworker can deliberately clear its agency', () => {
  const w = loadApp();
  seedCwForm(w, { id: 'cw3', name: 'Priya Nair', agency: 'MDHHS - Macomb', org: 'MDHHS' },
             { first: 'Priya', last: 'Nair', org: 'MDHHS', agency: '' });
  w.saveCwInfoPane();
  assert.strictEqual(w.getCaseworkers().find((c) => c.id === 'cw3').agency, '',
    'clearing the visible field on purpose must still save');
});

test('a carrier caseworker keeps its stored agency rather than losing it silently', () => {
  const w = loadApp();
  seedCwForm(w, { id: 'cw4', name: 'Daniel Voss', agency: 'MDHHS - Wayne', org: 'Humana' },
             { first: 'Daniel', last: 'Voss', org: 'Humana', phone: '313-555-2000' });
  w.saveCwInfoPane();
  const saved = w.getCaseworkers().find((c) => c.id === 'cw4');
  assert.strictEqual(saved.agency, 'MDHHS - Wayne',
    'the input is hidden here, so a blanked agency could never be seen or typed back');
  assert.strictEqual(saved.phone, '313-555-2000');
});
