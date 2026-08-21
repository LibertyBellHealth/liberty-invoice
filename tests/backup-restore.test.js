'use strict';
// Disaster-recovery round trip. Before this existed, a OneDrive backup could not be restored by the
// app that wrote it: importProfiles treated every top-level key as a client name, so a backup
// restored ZERO clients and instead created junk records called "_exportedAt", "clients", etc.
// This drives the REAL backup builder and the REAL import, with the device wiped in between.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

function seed(w) {
  w.saveProfilesLS({
    'Jane Doe': { clientName: 'Jane Doe', medicaidId: '1234567890', startDate: '2024-01-01',
                  invoices: [{ dbId: 1, billingPeriod: '07/2026', status: 'submitted',
                               data: { svcHH: '20', svcMM: '05' }, _synced: 'baseline' }] },
    'John Roe': { clientName: 'John Roe', medicaidId: '9876543210', invoices: [] },
  });
  w.saveCaregiversLS({ cg1: { name: 'Casey Giver', phone: '313-555-1212' } });
  w.saveCaseworkersLS([{ id: 'cw1', name: 'Case Worker', email: 'w@michigan.gov' }]);
}
function quiet(w) {
  w.showConfirm = (m, cb) => { if (typeof cb === 'function') cb(); };
  w.showAlert = () => {};
  w.saveProfileSP = () => Promise.resolve();
  w.saveCaregiverAPI = () => Promise.resolve();
  w.saveCaseworkerAPI = () => Promise.resolve();
}
function restore(w, json) {
  w.importProfiles({ target: { files: [new w.File([json], 'backup.json', { type: 'application/json' })], value: '' } });
  return new Promise(r => setTimeout(r, 200));
}

test('a OneDrive backup restores onto a wiped device', async () => {
  const w = loadApp(); resetStorage(w); quiet(w); seed(w);
  const backup = JSON.stringify(w._buildBackupPayload(false));
  resetStorage(w);                                    // the disaster
  assert.strictEqual(Object.keys(w.getProfiles()).length, 0, 'device is empty');

  await restore(w, backup);

  const p = w.getProfiles();
  assert.deepStrictEqual(Object.keys(p).sort(), ['Jane Doe', 'John Roe'],
    'both clients come back and NO metadata key becomes a client');
  assert.strictEqual(p['Jane Doe'].medicaidId, '1234567890', 'client fields survive');
  assert.strictEqual(p['Jane Doe'].invoices.length, 1, 'invoices survive');
  assert.strictEqual(p['Jane Doe'].invoices[0].billingPeriod, '07/2026');
  assert.strictEqual(p['Jane Doe'].invoices[0]._synced, undefined,
    'the sync baseline is cleared, so a restored invoice is actually re-sent to the DB');
  assert.strictEqual(Object.keys(w.getCaregivers()).length, 1, 'caregivers are restored');
  assert.strictEqual(w.getCaseworkers().length, 1, 'caseworkers are restored');
});

test('the flat JSON export still restores, and metadata keys never become clients', async () => {
  const w = loadApp(); resetStorage(w); quiet(w); seed(w);
  // The other supported shape: a plain client-name -> record map, with a stray metadata key.
  const flat = JSON.stringify({ _exportedAt: '2026-08-21T00:00:00Z',
                                'Jane Doe': { clientName: 'Jane Doe', medicaidId: '1234567890' } });
  resetStorage(w);
  await restore(w, flat);
  assert.deepStrictEqual(Object.keys(w.getProfiles()), ['Jane Doe'],
    'the flat export still works and "_exportedAt" is not treated as a client');
});

test('restoring does not overwrite roster records that already exist locally', async () => {
  const w = loadApp(); resetStorage(w); quiet(w); seed(w);
  const backup = JSON.stringify(w._buildBackupPayload(false));
  w.saveCaregiversLS({ cg1: { name: 'Casey Giver — EDITED SINCE BACKUP', phone: '313-555-9999' } });
  await restore(w, backup);
  assert.match(w.getCaregivers().cg1.name, /EDITED SINCE BACKUP/,
    'work done since the backup is never silently rolled back by a restore');
});

test('_parseBackupFile understands both shapes', () => {
  const w = loadApp();
  const wrapped = w._parseBackupFile(JSON.stringify({ _exportedAt: 'x', _appVersion: 'v13',
    clients: { A: {} }, caregivers: { c1: {} }, caseworkers: [{ id: 1 }], signatures: [] }));
  assert.deepStrictEqual(Object.keys(wrapped.clients), ['A']);
  assert.ok(wrapped.caregivers && wrapped.caseworkers, 'rosters are carried through');
  const flat = w._parseBackupFile(JSON.stringify({ _exportedAt: 'x', A: {}, B: {} }));
  assert.deepStrictEqual(Object.keys(flat.clients).sort(), ['A', 'B']);
  assert.strictEqual(flat.caregivers, null);
});
