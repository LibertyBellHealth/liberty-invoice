'use strict';
// The mirror of save-roundtrip: a LOAD that silently drops or reverts what the user entered is the
// same harm as a bad save. All three of these were confirmed by an independent sweep.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

// ── The task dirty flag I added in batch C was never read ────────────────────────────────────
test('a failed task edit survives the next background load', () => {
  const w = loadApp(); resetStorage(w);
  const server = [{ id: '42', dbId: 42, text: 'OLD server text', done: false }];
  const local  = [{ id: '42', dbId: 42, text: 'EDIT that failed to save', done: true, _unsaved: true }];
  const merged = w._mergeByIdKeepUnsynced(server, local, '_synced');
  assert.strictEqual(merged.length, 1);
  assert.strictEqual(merged[0].text, 'EDIT that failed to save',
    'the flag was written for tasks but the merge had no _unsaved branch, so the edit was reverted silently');
});

test('a task with no failed-save marker still lets the server win', () => {
  const w = loadApp(); resetStorage(w);
  const merged = w._mergeByIdKeepUnsynced(
    [{ id: '42', dbId: 42, text: 'server' }], [{ id: '42', dbId: 42, text: 'stale local' }], '_synced');
  assert.strictEqual(merged[0].text, 'server', 'a clean local copy must not win — that would pin stale data');
});

test('a locally-added task with no server row is still kept', () => {
  const w = loadApp(); resetStorage(w);
  const merged = w._mergeByIdKeepUnsynced([], [{ id: 'td_new', text: 'added offline' }], '_synced');
  assert.strictEqual(merged.length, 1, 'the existing keep-unsynced behaviour must not regress');
});

// ── The day-grid pattern lived only in localStorage ──────────────────────────────────────────
test('a client’s default day-grid pattern survives a server-wins load', () => {
  const w = loadApp(); resetStorage(w);
  const pattern = { svc: [[true, false], [false, true]] };
  const server = { Jane: { clientName: 'Jane', _dbId: 7 } };            // server never carries tasks
  const local  = { Jane: { clientName: 'Jane', _dbId: 7, tasks: pattern } };
  const out = w._mergeProfilesLoad(server, local);
  assert.ok(out.Jane.tasks, 'tasks is absent from the load map AND from _clientSig, so it was dropped every refresh');
  assert.deepStrictEqual(JSON.parse(JSON.stringify(out.Jane.tasks)), pattern);
});

test('a server copy that DOES carry a pattern is not overwritten by the local one', () => {
  const w = loadApp(); resetStorage(w);
  const out = w._mergeProfilesLoad(
    { Jane: { clientName: 'Jane', _dbId: 7, tasks: { svc: [['server']] } } },
    { Jane: { clientName: 'Jane', _dbId: 7, tasks: { svc: [['local']] } } });
  assert.deepStrictEqual(JSON.parse(JSON.stringify(out.Jane.tasks)), { svc: [['server']] },
    'only fill where the server has none — a real edit must still win');
});

// ── Deleting a client left their SSN in memory, keyed by NAME ────────────────────────────────
test('deleting a client does not leave their SSN attached to the next client of the same name', () => {
  const w = loadApp(); resetStorage(w);
  w.showConfirm = (msg, onOk) => onOk();
  w.showAlert = () => {}; w.showToast = () => {};
  w.deleteProfileSP = () => Promise.resolve(); w.saveTaskAPI = () => Promise.resolve();
  w.addAuditEntry = () => {}; w.logActivity = () => {}; w.navHome = () => {}; w.aiTrack = () => {};
  w.fetch = () => Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({}) });
  w.saveProfilesLS({ 'John Smith': { clientName: 'John Smith', ssn: '111-22-3333' } });
  w.activeProfileName = 'John Smith';
  assert.strictEqual(w.getProfiles()['John Smith'].ssn, '111-22-3333', 'precondition: the SSN is in the overlay');
  w.deleteClient();
  // A DIFFERENT person who happens to share the name
  w.saveProfilesLS({ 'John Smith': { clientName: 'John Smith' } });
  assert.ok(!w.getProfiles()['John Smith'].ssn,
    'the deleted patient’s SSN was re-attached to the new record and encrypted into their row');
});

test('deleting a client detaches its tasks instead of orphaning them', () => {
  const w = loadApp(); resetStorage(w);
  w.showConfirm = (msg, onOk) => onOk();
  w.showAlert = () => {}; w.showToast = () => {};
  w.deleteProfileSP = () => Promise.resolve(); w.saveTaskAPI = () => Promise.resolve();
  w.addAuditEntry = () => {}; w.logActivity = () => {}; w.navHome = () => {}; w.aiTrack = () => {};
  w.fetch = () => Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({}) });
  w.saveProfilesLS({ 'John Smith': { clientName: 'John Smith' } });
  w.saveTodos([{ id: 'td_1', text: 'Reassessment', client: 'John Smith' }]);
  w.activeProfileName = 'John Smith';
  w.deleteClient();
  assert.strictEqual(w.getTodos()[0].client, '',
    'a name-keyed task pointing at a deleted client is invisible, and the next same-name client adopts it');
});
