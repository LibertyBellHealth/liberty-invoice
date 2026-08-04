'use strict';
// PHI / HIPAA guards — the highest-stakes logic in the app. A regression here is a PHI leak or a
// PHI-at-rest violation, not just a wrong value. We already shipped one bug in clearPHIFromStorage.
const { test, beforeEach } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

const w = loadApp();
beforeEach(() => resetStorage(w));

test('clearPHIFromStorage: wipes PHI keys but keeps the settings whitelist', () => {
  const ls = w.localStorage;
  // PHI / client-identifying
  ls.setItem('lhca_profiles', '{"Jane":{}}');
  ls.setItem('lhca_caregivers', '{}');
  ls.setItem('lhca_supervisors', '{}');          // was missed by the old fixed list
  ls.setItem('lhca_email_audit', '[]');
  ls.setItem('lhca_autogen_undo', '[]');
  ls.setItem('lhca_todos', '[]');
  // Non-PHI settings that MUST survive a wipe
  ls.setItem('lhca_state_rate', '27.00');
  ls.setItem('lhca_pdf_mode', 'dual');
  ls.setItem('lhca_last_onedrive_backup', '2026-08-01'); // the exact key whose wipe caused a real bug
  ls.setItem('lhca_cw_col_widths', '{}');
  ls.setItem('lhca_sup_page_size', '25');
  ls.setItem('theme', 'dark'); // non-lhca, must be untouched

  w.clearPHIFromStorage();

  for (const gone of ['lhca_profiles', 'lhca_caregivers', 'lhca_supervisors', 'lhca_email_audit', 'lhca_autogen_undo', 'lhca_todos']) {
    assert.strictEqual(ls.getItem(gone), null, `${gone} must be wiped`);
  }
  for (const kept of ['lhca_state_rate', 'lhca_pdf_mode', 'lhca_last_onedrive_backup', 'lhca_cw_col_widths', 'lhca_sup_page_size', 'theme']) {
    assert.notStrictEqual(ls.getItem(kept), null, `${kept} must survive the wipe`);
  }
});

test('clearPHIFromStorage: keeping lhca_last_onedrive_backup (regression guard)', () => {
  // Directly locks in the fix: the weekly-backup timestamp is a plain date, not PHI, and wiping it
  // made the "weekly" OneDrive backup re-fire every session.
  w.localStorage.setItem('lhca_last_onedrive_backup', '2026-08-01');
  w.localStorage.setItem('lhca_profiles', '{"x":{}}');
  w.clearPHIFromStorage();
  assert.strictEqual(w.localStorage.getItem('lhca_last_onedrive_backup'), '2026-08-01', 'backup timestamp survives');
  assert.strictEqual(w.localStorage.getItem('lhca_profiles'), null, 'PHI still wiped');
});

test('saveProfilesLS: SSN never reaches localStorage (PHI-at-rest invariant)', () => {
  w.saveProfilesLS({
    'Jane Doe': { clientName: 'Jane Doe', medicaidId: 'M123', ssn: '123-45-6789', _clientSynced: 'sig-embedding-ssn' },
  });
  const onDisk = JSON.parse(w.localStorage.getItem('lhca_profiles'))['Jane Doe'];
  assert.strictEqual(onDisk.ssn, undefined, 'ssn must NOT be written to disk');
  assert.strictEqual(onDisk._clientSynced, undefined, '_clientSynced embeds ssn and must NOT be on disk');
  assert.strictEqual(onDisk.medicaidId, 'M123', 'non-SSN fields still persist');
  // ...but the SSN is retained in the in-memory overlay so the app still has it this session
  assert.strictEqual(w._ssnMem['Jane Doe'], '123-45-6789', 'ssn kept in memory-only overlay');
});

test('formatSSN: formats to XXX-XX-XXXX and strips non-digits', () => {
  const mk = (v) => { const el = { value: v }; w.formatSSN(el); return el.value; };
  assert.strictEqual(mk('123456789'), '123-45-6789', 'full SSN formatted');
  assert.strictEqual(mk('12345'), '123-45', 'partial formats progressively');
  assert.strictEqual(mk('abc12-34x56789extra'), '123-45-6789', 'non-digits stripped, capped at 9 digits');
  assert.strictEqual(mk('12'), '12', 'too short: left as-is');
});
