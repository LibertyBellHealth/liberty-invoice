'use strict';
// Regression tests for the CODE_AUDIT_2 PHI fixes. Mutation testing showed SSN masking had ZERO
// coverage in either export path — removing the mask entirely kept the suite green. These pin the
// actual payloads that leave the device.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

const SSN = '123-45-6789';
const CG_SSN = '987-65-4321';

function seed(w) {
  resetStorage(w);
  w.saveProfilesLS({ 'Jane Doe': { clientName: 'Jane Doe', ssn: SSN, medicaidId: '1234567890',
                                   _clientSynced: w._clientSig({ clientName: 'Jane Doe', ssn: SSN }) } });
  w.saveCaregiversLS({ cg1: { name: 'Casey Giver', ssn: CG_SSN, miloginPassword: 'hunter2' } });
}

test('backup payload: masked mode contains NO full SSN — client or caregiver', () => {
  const w = loadApp(); seed(w);
  const json = JSON.stringify(w._buildBackupPayload(false));
  assert.ok(!json.includes(SSN), 'client SSN must not appear (it used to survive inside _clientSynced)');
  assert.ok(!json.includes(CG_SSN), 'caregiver SSN must not appear (caregivers were never masked at all)');
  assert.ok(!json.includes('hunter2'), 'MI Login password must never leave the device');
  assert.ok(json.includes('***-**-6789') && json.includes('***-**-4321'), 'both are masked to last 4');
});

test('backup payload: _clientSynced never leaves the device in EITHER mode', () => {
  const w = loadApp(); seed(w);
  for (const full of [false, true]) {
    const p = w._buildBackupPayload(full);
    assert.strictEqual(p.clients['Jane Doe']._clientSynced, undefined,
      'the dirty-tracking signature embeds the raw SSN — it must be stripped (full=' + full + ')');
  }
});

test('backup payload: the full-SSN mode is explicit, and records who exported it', () => {
  const w = loadApp(); seed(w);
  const p = w._buildBackupPayload(true);
  assert.strictEqual(p.clients['Jane Doe'].ssn, SSN, 'an explicit full export still works');
  assert.strictEqual(p._includesFullSSN, true);
  assert.notStrictEqual(p._exportedBy, 'unknown',
    'window.signedInEmail was never assigned, so every backup was attributed to "unknown"');
  assert.ok(p._exportedBy && p._exportedBy.length, 'an actor is always recorded');
});

test('JSON export: the masked download carries no full SSN and no _clientSynced', () => {
  const w = loadApp(); seed(w);
  let blob = '';
  w.Blob = function (parts) { blob = String(parts[0]); return { type: 'application/json' }; };
  w.URL.createObjectURL = () => 'blob:x';
  w.URL.revokeObjectURL = () => {};
  w.doExportProfiles(false);
  assert.ok(blob.length, 'export produced a payload');
  assert.ok(!blob.includes(SSN), 'no full SSN in a masked export');
  assert.ok(!blob.includes('_clientSynced'), 'no dirty-tracking signature in a masked export');
});

test('a bulk export is recorded in the activity log (HIPAA disclosure accounting)', () => {
  const w = loadApp(); seed(w);
  w.Blob = function (parts) { return { type: 'application/json' }; };
  w.URL.createObjectURL = () => 'blob:x';
  w.URL.revokeObjectURL = () => {};
  // The full-SSN path ends by opening the confirm modal, whose DOM doesn't exist in bare jsdom.
  // The audit write happens before that, so swallow the render error and assert on the record.
  try { w.doExportProfiles(true); } catch (e) { /* modal DOM only */ }
  const log = w.getActivity();
  assert.ok(log.some(e => /export/i.test(e.type) && /FULL-SSN/.test(e.text)),
    'a full-SSN roster export must leave a record — it previously left none at all');
});

test('pdf.js is loaded with eval disabled (CVE-2024-4367 mitigation)', () => {
  const fs = require('fs'), path = require('path');
  const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
  const calls = src.match(/getDocument\(\{[^}]*\}/g) || [];
  assert.ok(calls.length, 'found the pdf.js entry point');
  calls.forEach(c => assert.ok(/isEvalSupported\s*:\s*false/.test(c),
    'every getDocument call must disable eval while the vendored pdf.js is < 4.2.67: ' + c));
});
