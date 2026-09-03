'use strict';
// changeInvStatus re-read the profiles inside its confirmation callback but still used the array
// INDEX captured beforehand, and read activeProfileName at apply time. A background load replaces
// the invoices array (re-sorted, de-duplicated) and the owner can navigate away while the dialog
// sits open — so the status landed on a different invoice, or a different client's invoice.
// Eleventh instance of the read-before-a-dialog pattern; found while writing coverage for it.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

function app() {
  const w = loadApp();
  resetStorage(w);
  ['saveProfileSP', 'logActivity', 'updateStats', 'renderOverviewPane', 'addAuditEntry']
    .forEach((f) => { w[f] = () => {}; });
  w.pending = null;
  w.showConfirm = (msg, onOk, opts) => { w.pending = { ok: onOk, opts: opts || {} }; };
  w.activeProfileName = 'Jane Doe';
  w.saveProfilesLS({
    'Jane Doe': { clientName: 'Jane Doe', invoices: [
      { billingPeriod: '08/2026', status: 'draft', dbId: 'db_08', data: {} },
      { billingPeriod: '07/2026', status: 'submitted', dbId: 'db_07', data: {} }] },
    'Bob Roe': { clientName: 'Bob Roe', invoices: [
      { billingPeriod: '08/2026', status: 'draft', dbId: 'db_b08', data: {} }] },
  });
  return w;
}
const sel = (idx, value) => ({ dataset: { idx: String(idx) }, value: value, className: '' });
const statusOf = (w, client, period) =>
  (w.getProfiles()[client].invoices.find((i) => i.billingPeriod === period) || {}).status;

test('marking Paid asks first, and applies to the right invoice', () => {
  const w = app();
  w.changeInvStatus(sel(0, 'paid'));
  assert.ok(w.pending, 'marking Paid must be confirmed');
  w.pending.ok();
  assert.strictEqual(statusOf(w, 'Jane Doe', '08/2026'), 'paid');
  assert.strictEqual(statusOf(w, 'Jane Doe', '07/2026'), 'submitted', 'the other month is untouched');
});

test('an invoice arriving mid-dialog does not shift the change onto another month', () => {
  const w = app();
  w.changeInvStatus(sel(0, 'paid'));              // targeting 08/2026, currently index 0
  const store = w.getProfiles();
  store['Jane Doe'].invoices.unshift({ billingPeriod: '09/2026', status: 'draft', dbId: 'db_09', data: {} });
  w.saveProfilesLS(store);                        // a sync lands: every index moves
  w.pending.ok();
  assert.strictEqual(statusOf(w, 'Jane Doe', '08/2026'), 'paid', 'the targeted month must be the one marked');
  assert.strictEqual(statusOf(w, 'Jane Doe', '09/2026'), 'draft', 'the newcomer must not be marked Paid');
  assert.strictEqual(statusOf(w, 'Jane Doe', '07/2026'), 'submitted');
});

test('navigating to another client mid-dialog does not mark THEIR invoice', () => {
  const w = app();
  w.changeInvStatus(sel(0, 'paid'));
  w.activeProfileName = 'Bob Roe';                // owner moves on while the dialog is open
  w.pending.ok();
  assert.strictEqual(statusOf(w, 'Jane Doe', '08/2026'), 'paid');
  assert.strictEqual(statusOf(w, 'Bob Roe', '08/2026'), 'draft', "Bob's invoice must be untouched");
});

test('cancelling restores the dropdown and changes nothing', () => {
  const w = app();
  const s = sel(0, 'paid');
  w.changeInvStatus(s);
  w.pending.opts.onCancel();
  assert.strictEqual(s.value, 'draft', 'the dropdown must snap back');
  assert.strictEqual(statusOf(w, 'Jane Doe', '08/2026'), 'draft');
});

test('unlocking a Paid invoice is confirmed as the risky change it is', () => {
  const w = app();
  const store = w.getProfiles();
  store['Jane Doe'].invoices[0].status = 'paid';
  w.saveProfilesLS(store);
  w.changeInvStatus(sel(0, 'draft'));
  assert.ok(w.pending && w.pending.opts.danger, 'reverting Paid must be flagged dangerous');
  w.pending.ok();
  assert.strictEqual(statusOf(w, 'Jane Doe', '08/2026'), 'draft');
});

test('a no-op change does nothing at all', () => {
  const w = app();
  w.changeInvStatus(sel(0, 'draft'));
  assert.strictEqual(w.pending, null, 'draft -> draft must not prompt');
});
