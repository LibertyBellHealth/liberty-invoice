'use strict';
// Deleting a roster record left dead references behind. The DB nulls its FK, but the LOCAL client
// record keeps the dead id — so the next client save posts it, the FK rejects it as a generic 500,
// the dirty flag pins the local copy, and the merge keeps it. That client can then NEVER be saved
// again, and a reload cannot fix it. Six of the seven delete paths each did their own thing.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

function quiet(w) {
  if (!w.Element.prototype.scrollIntoView) w.Element.prototype.scrollIntoView = function () {};
  w.showConfirm = (msg, onOk) => onOk();
  w.showAlert = () => {}; w.showToast = () => {}; w.addAuditEntry = () => {};
  w.logActivity = () => {}; w.aiTrack = () => {}; w.currentUserEmail = () => 'owner@example.com';
  w.fetch = () => Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({}) });
  ['saveProfileSP','saveCaregiverAPI','saveCaseworkerAPI','saveSupervisorAPI','deleteCaregiverAPI',
   'deleteCaseworkerAPI','deleteSupervisorAPI','renderCaregiverGrid','renderCaseworkerList','updateStats',
   'showCwGrid','hideCaseworkerForm','hideCgForm','backToSupList','renderSupervisorList',
   'refreshSupervisorDropdowns','renderCaregiverList','renderClientGrid','renderSidebarClients',
  ].forEach((f) => { w[f] = () => {}; });
}

test('deleting a caregiver clears it from every client that referenced it', () => {
  const w = loadApp(); resetStorage(w); quiet(w);
  w.document.body.insertAdjacentHTML('beforeend', '<div id="cgDetailView"></div><div id="cgGridView"></div>');
  w.saveCaregiversLS({ cg1: { id: 'cg1', name: 'Ann Lee' } });
  w.saveProfilesLS({ Jane: { clientName: 'Jane', caregiverId: 'cg1' },
                     Bob:  { clientName: 'Bob',  caregiverId: 'cg9' } });
  w._doDeleteCaregiver('cg1');
  assert.strictEqual(w.getProfiles().Jane.caregiverId, '',
    'a dead caregiver id makes the next save 500 and the merge then pins it — the client can never be saved again');
  assert.strictEqual(w.getProfiles().Bob.caregiverId, 'cg9', 'other clients are untouched');
});

test('deleting a caseworker clears BOTH the id and the name the billing run groups on', () => {
  const w = loadApp(); resetStorage(w); quiet(w);
  w.saveCaseworkersLS([{ id: 'cw1', name: 'Renee Feldman', email: 'rf@michigan.gov' }]);
  w.saveProfilesLS({ Jane: { clientName: 'Jane', caseworkerId: 'cw1', worker: 'Renee Feldman' } });
  w.activeCwId = 'cw1';
  w.deleteCaseworkerFromDetail();
  const jane = w.getProfiles().Jane;
  assert.strictEqual(jane.caseworkerId, '');
  assert.strictEqual(jane.worker, '',
    'worker is the NAME Send All groups on: left set, the client passes validation but resolves to no ' +
    'email and is skipped silently — the invoice looks ready and is never sent');
});

test('deleting a supervisor from the DETAIL page unassigns its caseworkers (the modal already did)', () => {
  const w = loadApp(); resetStorage(w); quiet(w);
  w.document.body.insertAdjacentHTML('beforeend', '<input id="supd-id" value="sup1">');
  w.saveSupervisorsLS({ sup1: { id: 'sup1', name: 'Pat Reed' } });
  w.saveCaseworkersLS([{ id: 'cw1', name: 'Renee Feldman', supervisor_id: 'sup1' },
                       { id: 'cw2', name: 'Marcus Ojeda', supervisor_id: 'sup2' }]);
  w.deleteSupervisorFromDetail();
  assert.strictEqual(w.getCaseworkers()[0].supervisor_id, '', 'orphaned supervisor_id fails the next caseworker save');
  assert.strictEqual(w.getCaseworkers()[1].supervisor_id, 'sup2');
});

test('a client pointing at a caseworker that no longer exists is blocked, not skipped', () => {
  const w = loadApp(); resetStorage(w); w.saveSigsLS([{ id: 1, data: 'x' }]);
  const prof = { clientName: 'Jane', medicaidId: '123', worker: 'Ghost Worker', caseworkerId: 'cw_gone' };
  const inv = { billingPeriod: '08/2026', status: 'draft', data: { svcHH: '10', svcMM: '00' } };
  const issues = w.validateInvoiceForSend('Jane', prof, inv, {});   // {} = not found
  assert.ok(issues.some((i) => /no longer exists/.test(i)),
    'without this the client is dropped from Send All with no error at all');
});
