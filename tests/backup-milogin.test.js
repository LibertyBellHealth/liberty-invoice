'use strict';
// Owner's decision: MI Login credentials go in the FULL backup. They are State portal logins and
// resetting them is not always quick. Deliberately NOT in the weekly masked auto-backup — a
// credential in a routinely-generated file is a standing exposure.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

function seed(w) {
  resetStorage(w);
  w.saveCaregiversLS({
    cg1: { id: 'cg1', name: 'Ann Lee', miloginUsername: 'annlee', ssn: '123-45-6789' },
    cg2: { id: 'cg2', name: 'Bob Ray', miloginUsername: 'bobray' },
    cg3: { id: 'cg3', name: 'No Portal' },        // no MI Login account at all
  });
  w.saveProfilesLS({ Jane: { clientName: 'Jane', ssn: '987-65-4321' } });
}

test('a MASKED backup contains no credential and no real SSN', () => {
  const w = loadApp(); seed(w);
  const p = w._buildBackupPayload(false, { cg1: 'SECRET-1' });
  assert.ok(!p.caregivers.cg1.miloginPassword,
    'the weekly auto-backup is masked and must never carry a State portal credential');
  assert.match(String(p.caregivers.cg1.ssn || ''), /^\*\*\*-\*\*-/, 'SSN stays masked');
});

test('a FULL backup carries the credential so a wiped device can be rebuilt', () => {
  const w = loadApp(); seed(w);
  const p = w._buildBackupPayload(true, { cg1: 'SECRET-1', cg2: 'SECRET-2' });
  assert.strictEqual(p.caregivers.cg1.miloginPassword, 'SECRET-1');
  assert.strictEqual(p.caregivers.cg2.miloginPassword, 'SECRET-2');
  assert.ok(!p.caregivers.cg3.miloginPassword, 'a caregiver with no portal account gets nothing');
});

test('a credential that could not be read is simply absent, never blank', () => {
  const w = loadApp(); seed(w);
  const p = w._buildBackupPayload(true, { cg1: 'SECRET-1' });   // cg2 failed to fetch
  assert.strictEqual(p.caregivers.cg1.miloginPassword, 'SECRET-1');
  assert.ok(!('miloginPassword' in p.caregivers.cg2),
    'an empty string would be sent on restore and — but for the backend guard — could blank the stored one');
});

test('only caregivers with a portal username are fetched', async () => {
  const w = loadApp(); seed(w);
  const asked = [];
  w.fetch = (url) => {
    const m = String(url).match(/caregivers\/([^/]+)\/milogin/);
    if (m) asked.push(m[1]);
    return Promise.resolve({ ok: true, status: 200,
      json: () => Promise.resolve({ milogin_password: 'PW-' + (m ? m[1] : '') }) });
  };
  const r = await w._fetchMiloginForBackup();
  assert.strictEqual(asked.sort().join(','), 'cg1,cg2', 'cg3 has no MI Login account — do not ask for one');
  assert.strictEqual(r.creds.cg1, 'PW-cg1');
  assert.strictEqual(r.failed.length, 0);
});

test('a failed read is reported, not silently dropped', async () => {
  const w = loadApp(); seed(w);
  w.fetch = (url) => (/cg2/.test(String(url))
    ? Promise.resolve({ ok: false, status: 500, json: () => Promise.resolve({}) })
    : Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({ milogin_password: 'PW' }) }));
  const r = await w._fetchMiloginForBackup();
  assert.strictEqual(Array.from(r.failed).join(','), 'cg2',
    'a backup you believe holds the credentials but does not is worse than one you know does not');
  assert.strictEqual(r.creds.cg1, 'PW', 'the ones that DID read must still be included');
});
