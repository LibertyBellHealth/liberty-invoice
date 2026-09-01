'use strict';
// Two ways the disaster-recovery story was not true.
//
// 1. Caregiver SSNs are not in the roster and not in localStorage — the bulk load returns
//    ssn_last4 only, saveCaregiversLS strips the rest to the in-memory _cgSsnMem, and
//    getCaregivers() re-overlays just what that map holds: the caregivers someone happened to open
//    this session. _buildBackupPayload serialises getCaregivers(), so a file labelled FULL taken on
//    a fresh page load contained no caregiver SSNs at all.
// 2. importProfiles fired every server write fire-and-forget and then said "Import Complete".
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

// ── 1. a FULL backup must actually carry the SSNs ─────────────────────────────
test('_fetchSsnForBackup pulls SSNs the roster only knows the last 4 of', async () => {
  const w = loadApp(); resetStorage(w);
  // Exactly the real post-load shape: last4 present, full value absent.
  w.saveCaregiversLS({ cg1: { name: 'Casey Giver', ssnLast4: '6789' },
                       cg2: { name: 'Robin Aide',  ssnLast4: '4321' } });
  const asked = [];
  w.fetch = (url) => { asked.push(String(url));
    const id = String(url).match(/caregivers\/([^/]+)\/ssn/)[1];
    return Promise.resolve({ ok: true, json: () => Promise.resolve(
      { ssn: id === 'cg1' ? '123-45-6789' : '987-65-4321' }) }); };
  const res = await w._fetchSsnForBackup();
  assert.strictEqual(asked.length, 2, 'both caregivers were fetched');
  assert.strictEqual(res.ssns.cg1, '123-45-6789');
  assert.strictEqual(res.ssns.cg2, '987-65-4321');
  assert.strictEqual(res.failed.length, 0);
});

test('_fetchSsnForBackup reports what it could NOT read', async () => {
  const w = loadApp(); resetStorage(w);
  w.saveCaregiversLS({ cg1: { name: 'Casey Giver', ssnLast4: '6789' } });
  w.fetch = () => Promise.resolve({ ok: false, status: 500, json: () => Promise.resolve(null) });
  const res = await w._fetchSsnForBackup();
  assert.deepStrictEqual(JSON.parse(JSON.stringify(res.failed)), ['cg1'],
    'a backup you believe holds the SSNs but does not is worse than one you know does not');
});

test('a FULL backup serialises the fetched SSN, not the last 4', () => {
  const w = loadApp(); resetStorage(w);
  w.saveCaregiversLS({ cg1: { name: 'Casey Giver', ssnLast4: '6789' } });   // no full value in memory
  const payload = w._buildBackupPayload(true, {}, { cg1: '123-45-6789' });
  assert.strictEqual(payload.caregivers.cg1.ssn, '123-45-6789',
    'restoring this into an empty system is the whole point of a FULL backup');
});

test('a MASKED backup still carries no real SSN', () => {
  const w = loadApp(); resetStorage(w);
  w.saveCaregiversLS({ cg1: { name: 'Casey Giver', ssn: '123-45-6789' } });
  const payload = w._buildBackupPayload(false, {}, { cg1: '123-45-6789' });
  assert.strictEqual(payload.caregivers.cg1.ssn, '***-**-6789',
    'the weekly auto-backup is masked and goes to OneDrive — it must never carry a real SSN');
  assert.ok(!payload.caregivers.cg1.miloginPassword, 'nor the state portal credential');
});

// ── 2. restore must confirm persistence before claiming success ───────────────
function runImport(w, payload, saveImpl) {
  const shown = [];
  w.showConfirm = (msg, ok, opts) => { shown.push({ msg: String(msg), title: (opts || {}).title }); };
  w.showToast = () => {};
  w.showAlert = () => {};
  ['renderSidebarClients','renderClientGrid','updateStats','logActivity','renderNotesPane']
    .forEach((f) => { w[f] = () => {}; });
  w.saveProfileSP = saveImpl;
  const file = new w.File([JSON.stringify(payload)], 'backup.json', { type: 'application/json' });
  w.importProfiles({ target: { files: [file] } });
  return shown;
}

const BACKUP = { _exportedAt: '2026-08-01T00:00:00.000Z', _includesFullSSN: true,
  clients: { 'Jane Doe': { clientName: 'Jane Doe', invoices: [] },
             'John Roe': { clientName: 'John Roe', invoices: [] } } };

test('restore reports success only after every write is confirmed', async () => {
  const w = loadApp(); resetStorage(w);
  const shown = runImport(w, BACKUP, () => Promise.resolve({}));
  await new Promise(r => setTimeout(r, 60));
  assert.strictEqual(shown.length, 1, 'exactly one completion dialog');
  assert.strictEqual(shown[0].title, 'Import Complete');
  assert.match(shown[0].msg, /All 2 records confirmed saved to the database/);
});

test('THE BUG: a restore whose writes FAIL must not report success', async () => {
  const w = loadApp(); resetStorage(w);
  const shown = runImport(w, BACKUP, (name) =>
    name === 'John Roe' ? Promise.reject(new Error('HTTP 500')) : Promise.resolve({}));
  await new Promise(r => setTimeout(r, 60));
  assert.strictEqual(shown.length, 1);
  assert.strictEqual(shown[0].title, 'Import INCOMPLETE',
    'this said "Import Complete" while the record existed only in this browser');
  assert.match(shown[0].msg, /NOT FULLY RESTORED/);
  assert.match(shown[0].msg, /client John Roe/, 'and it names the record that did not land');
  assert.match(shown[0].msg, /1 record reached the database/);
});

test('the completion dialog waits — it does not fire before the writes settle', async () => {
  const w = loadApp(); resetStorage(w);
  let release; const gate = new Promise(r => { release = r; });
  const shown = runImport(w, BACKUP, () => gate);
  await new Promise(r => setTimeout(r, 40));
  assert.strictEqual(shown.length, 0, 'nothing may be claimed while the writes are still in flight');
  release({});
  await new Promise(r => setTimeout(r, 40));
  assert.strictEqual(shown.length, 1, 'and it reports once they settle');
});
