'use strict';
// The two "apply what you just reviewed" writes. Both run AFTER a review dialog, so both must write
// to the client the review was opened for — not whichever client is active by the time the owner
// clicks — and both must cope with that client having been deleted meanwhile.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

function app() {
  const w = loadApp();
  resetStorage(w);
  ['saveProfileSP', 'addAuditEntry', 'showToast', 'renderInfoPane', 'renderAuthPane',
   'saveCaseworkerAPI', 'logActivity', '_syncReassessTask', 'renderOverviewPane', 'updateStats',
   'renderSidebarClients', 'renderClientGrid'].forEach((f) => { w[f] = () => {}; });
  w.alerts = [];
  w.showAlert = (m) => { w.alerts.push(String(m)); };
  // The import also files the source PDF; stub the network so the test exercises the WRITE.
  w.fetch = () => Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({}) });
  w.saveProfilesLS({
    'Alice Adams': { clientName: 'Alice Adams', _dbId: 'A1' },
    'Bob Brown': { clientName: 'Bob Brown', _dbId: 'B2' },
  });
  w.activeProfileName = 'Alice Adams';
  return w;
}
const prof = (w, n) => w.getProfiles()[n] || {};

test('card details land on the client the review was opened for', () => {
  const w = app();
  w.activeProfileName = 'Bob Brown';                       // owner navigated away meanwhile
  w._applyCardFields({ ssn: '123-45-6789' }, 'Alice Adams');
  assert.strictEqual(prof(w, 'Alice Adams').ssn, '123-45-6789');
  assert.strictEqual(prof(w, 'Bob Brown').ssn, undefined, "Bob's record must be untouched");
});

test('applying to a client deleted mid-review is refused, not silently misapplied', () => {
  const w = app();
  w.saveProfilesLS({ 'Bob Brown': { clientName: 'Bob Brown', _dbId: 'B2' } });   // Alice removed
  w._applyCardFields({ ssn: '123-45-6789' }, 'Alice Adams');
  assert.match(w.alerts.join(' '), /no longer available/i);
  assert.strictEqual(prof(w, 'Bob Brown').ssn, undefined);
});

test('the authorization lands on the client the import was opened for', () => {
  const w = app();
  w.activeProfileName = 'Bob Brown';
  w._applyDhsImport({ name: 'x.pdf' },
    { hours: 62, minutes: 21, tasks: [], effectiveDate: '08/01/2026' },
    { target: 'Alice Adams' });
  assert.strictEqual(prof(w, 'Alice Adams').authorization.hours, 62);
  assert.strictEqual(prof(w, 'Bob Brown').authorization, undefined,
    "Bob must not receive Alice's authorization");
});

test('importing onto a client deleted mid-review is refused and says the form was NOT saved', () => {
  const w = app();
  w.saveProfilesLS({ 'Bob Brown': { clientName: 'Bob Brown', _dbId: 'B2' } });
  w._applyDhsImport({ name: 'x.pdf' }, { hours: 62, minutes: 21, tasks: [] }, { target: 'Alice Adams' });
  assert.match(w.alerts.join(' '), /NOT saved/i);
  assert.strictEqual(prof(w, 'Bob Brown').authorization, undefined);
});

test('linking a matched caseworker assigns that caseworker, not a new one', () => {
  const w = app();
  w.saveCaseworkersLS([{ id: 'cw_1', name: 'Addison Sawyer', email: 'SawyerA2@michigan.gov' }]);
  w._applyDhsImport({ name: 'x.pdf' }, { hours: 10, minutes: 0, tasks: [], aswName: 'A Sawyer' },
    { target: 'Alice Adams', match: { id: 'cw_1', name: 'Addison Sawyer' }, link: true });
  assert.strictEqual(prof(w, 'Alice Adams').caseworkerId, 'cw_1');
  assert.strictEqual(prof(w, 'Alice Adams').worker, 'Addison Sawyer');
  assert.strictEqual(w.getCaseworkers().length, 1, 'linking must not also create a duplicate');
});

test('the authorization records the hours and effective date it was imported with', () => {
  const w = app();
  w._applyDhsImport({ name: 'MSA-6064.pdf' },
    { hours: 73, minutes: 44, effectiveDate: '08/04/2026',
      tasks: [{ task: 'Bathing', perDay: '00:05', freq: '7 days per week', perMonth: '02:30' }] },
    { target: 'Alice Adams' });
  const a = prof(w, 'Alice Adams').authorization;
  assert.strictEqual(a.hours, 73);
  assert.strictEqual(a.minutes, 44);
  assert.strictEqual(a.effectiveDate, '08/04/2026');
  assert.strictEqual(a.tasks.length, 1);
  assert.strictEqual(a.sourceFile, 'MSA-6064.pdf', 'the record should say which file it came from');
});
