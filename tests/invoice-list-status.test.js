'use strict';
// The attention row "N submitted invoices pending 30+ days — follow up on payment" opens the
// all-invoices list, but the status there was a static span: you arrived at the one place the
// workflow points you, and could not mark anything Paid. Reported from real use, 2026-09-03.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

function app() {
  const w = loadApp();
  resetStorage(w);
  ['saveProfileSP', 'logActivity', 'updateStats', 'renderOverviewPane', 'addAuditEntry', 'navDetail']
    .forEach((f) => { w[f] = () => {}; });
  w.pending = null;
  w.showConfirm = (msg, onOk, opts) => { w.pending = { msg: msg, ok: onOk, opts: opts || {} }; };
  ['allInvoicesModal', 'allInvModalList', 'allInvModalSubtitle', 'allInvModalTitle'].forEach((id) => {
    if (!w.document.getElementById(id)) {
      w.document.body.insertAdjacentHTML('beforeend', '<div id="' + id + '"></div>');
    }
  });
  w.saveProfilesLS({
    'Jane Doe': { clientName: 'Jane Doe', invoices: [
      { billingPeriod: '07/2026', status: 'submitted', dbId: 'db_07', savedAt: '7/1/2026', data: {} }] },
    'Bob Roe': { clientName: 'Bob Roe', invoices: [
      { billingPeriod: '07/2026', status: 'submitted', dbId: 'db_b07', savedAt: '7/2/2026', data: {} }] },
  });
  return w;
}
const statusOf = (w, c, p) =>
  (w.getProfiles()[c].invoices.find((i) => i.billingPeriod === p) || {}).status;
const selects = (w) => [...w.document.getElementById('allInvModalList').querySelectorAll('select')];

test('the list offers a real status control, not a label', () => {
  const w = app();
  w.openAllInvoicesModal('outstanding');
  const s = selects(w);
  assert.strictEqual(s.length, 2, 'every row needs its own control');
  assert.ok([...s[0].options].map((o) => o.value).includes('paid'));
});

test('marking Paid from the list asks first, then applies to that exact invoice', () => {
  const w = app();
  w.openAllInvoicesModal('outstanding');
  const s = selects(w).find((x) => x.getAttribute('data-client') === 'Jane Doe');
  s.value = 'paid';
  w.changeInvStatusFromList(s);
  assert.ok(w.pending, 'marking Paid must still be confirmed');
  w.pending.ok();
  assert.strictEqual(statusOf(w, 'Jane Doe', '07/2026'), 'paid');
  assert.strictEqual(statusOf(w, 'Bob Roe', '07/2026'), 'submitted', "the other client's row is untouched");
});

test('the control is keyed by client and period, not by row position', () => {
  const w = app();
  w.openAllInvoicesModal('outstanding');
  selects(w).forEach((s) => {
    assert.ok(s.getAttribute('data-client'), 'no client on the control');
    assert.ok(/^\d{2}\/\d{4}$/.test(s.getAttribute('data-period')), 'no billing period on the control');
  });
});

test('a paid invoice drops out of the outstanding list once marked', () => {
  const w = app();
  w.openAllInvoicesModal('outstanding');
  assert.strictEqual(selects(w).length, 2);
  const s = selects(w).find((x) => x.getAttribute('data-client') === 'Jane Doe');
  s.value = 'paid';
  w.changeInvStatusFromList(s);
  w.pending.ok();
  assert.strictEqual(selects(w).length, 1, 'the list should refresh and drop the paid invoice');
  assert.strictEqual(selects(w)[0].getAttribute('data-client'), 'Bob Roe');
});

test('cancelling changes nothing', () => {
  const w = app();
  w.openAllInvoicesModal('outstanding');
  const s = selects(w).find((x) => x.getAttribute('data-client') === 'Jane Doe');
  s.value = 'paid';
  w.changeInvStatusFromList(s);
  w.pending.opts.onCancel();
  assert.strictEqual(statusOf(w, 'Jane Doe', '07/2026'), 'submitted');
});

test('the detail-page dropdown still works and still guards Paid', () => {
  const w = app();
  w.activeProfileName = 'Jane Doe';
  const sel = { dataset: { idx: '0' }, value: 'paid', className: '' };
  w.changeInvStatus(sel);
  assert.ok(w.pending);
  w.pending.ok();
  assert.strictEqual(statusOf(w, 'Jane Doe', '07/2026'), 'paid');
});
