'use strict';
// The roster delete dialogs read the whole collection to build their preview, then wrote it back
// after the user answered. Anything added or edited while the dialog sat open — by a background
// roster load, another tab, or another device — was erased by that stale snapshot. Same shape as
// the invoice save that could overwrite a different month: state read before a dialog, acted on
// after it. The single-supervisor delete already re-read correctly; these four did not.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

function app() {
  const w = loadApp();
  resetStorage(w);
  ['deleteCaregiverAPI','deleteCaseworkerAPI','deleteSupervisorAPI','saveCaregiverAPI',
   'saveCaseworkerAPI','saveSupervisorAPI','_detachDeletedRoster','logActivity','updateStats',
   'clearCgBulkSelect','clearSupBulkSelect','clearCwBulkSelect','showAlert','showToast',
   'refreshSupervisorDropdowns','renderCaseworkerList','renderSupervisorList','renderCaregiverList',
   'renderCaregiverGrid'].forEach((f) => { w[f] = () => {}; });
  // Answer every confirm immediately, but mutate the roster first — that is the race.
  w._pending = null;
  w.showConfirm = (msg, onOk) => { w._pending = onOk; };
  return w;
}

test('deleting caregivers in bulk keeps one added while the dialog was open', () => {
  const w = app();
  w.saveCaregiversLS({ a: { name: 'Alice' }, b: { name: 'Bob' } });
  w.cgBulkSelected = { a: true };
  w.bulkDeleteCaregivers();
  w.saveCaregiversLS(Object.assign(w.getCaregivers(), { c: { name: 'Carol' } }));  // arrives mid-dialog
  w._pending();
  const after = w.getCaregivers();
  assert.strictEqual(after.a, undefined, 'the selected caregiver should be gone');
  assert.ok(after.b, 'Bob must survive');
  assert.ok(after.c, 'Carol was added while the dialog was open and must not be erased');
});

test('deleting caseworkers in bulk keeps one added while the dialog was open', () => {
  const w = app();
  w.saveCaseworkersLS([{ id: 'cw_1', name: 'A Sawyer' }, { id: 'cw_2', name: 'R Feto' }]);
  w.cwBulkSelected = { cw_1: true };
  w.bulkDeleteCaseworkers();
  w.saveCaseworkersLS(w.getCaseworkers().concat([{ id: 'cw_3', name: 'T Coleman' }]));
  w._pending();
  const ids = w.getCaseworkers().map((c) => c.id);
  assert.ok(!ids.includes('cw_1'), 'the selected caseworker should be gone');
  assert.ok(ids.includes('cw_2'));
  assert.ok(ids.includes('cw_3'), 'the one added mid-dialog must not be erased');
});

test('deleting supervisors in bulk keeps one added while the dialog was open', () => {
  const w = app();
  w.saveSupervisorsLS({ s1: { name: 'Sup One' }, s2: { name: 'Sup Two' } });
  w.saveCaseworkersLS([{ id: 'cw_1', name: 'A Sawyer', supervisor_id: 's1' }]);
  w.supBulkSelected = { s1: true };
  w.bulkDeleteSupervisors();
  w.saveSupervisorsLS(Object.assign(w.getSupervisors(), { s3: { name: 'Sup Three' } }));
  w._pending();
  const after = w.getSupervisors();
  assert.strictEqual(after.s1, undefined);
  assert.ok(after.s2, 'Sup Two must survive');
  assert.ok(after.s3, 'the supervisor added mid-dialog must not be erased');
});

test('a bulk supervisor delete does not revert caseworker edits made mid-dialog', () => {
  const w = app();
  w.saveSupervisorsLS({ s1: { name: 'Sup One' } });
  w.saveCaseworkersLS([{ id: 'cw_1', name: 'A Sawyer', supervisor_id: 's1', phone: 'old' }]);
  w.supBulkSelected = { s1: true };
  w.bulkDeleteSupervisors();
  w.saveCaseworkersLS([{ id: 'cw_1', name: 'A Sawyer', supervisor_id: 's1', phone: 'NEW' }]);
  w._pending();
  const cw = w.getCaseworkers().find((c) => c.id === 'cw_1');
  assert.strictEqual(cw.phone, 'NEW', 'the concurrent edit must not be reverted');
  assert.strictEqual(cw.supervisor_id, '', 'and the supervisor is still unassigned');
});
