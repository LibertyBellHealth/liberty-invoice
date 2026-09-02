'use strict';
// findClientsEligibleForAutoGen is the only "does this client already have an invoice for this
// period" check, and it ran BEFORE the confirmation dialog. A background sync (or another device)
// could add that very invoice while the dialog sat open, and the generator then unshifted a second
// one unconditionally. The monthly preview picks a period's invoice with .find(), so it showed the
// fresh Draft and the same client+period went to the caseworker twice, with re-derived numbers.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

function app() {
  const w = loadApp();
  resetStorage(w);
  ['addAuditEntry','logActivity','aiTrack','saveProfileSP','updateStats','renderUndoBanner',
   'previewMonthlyInvoices','renderSidebarClients','renderClientGrid'].forEach((f) => { w[f] = () => {}; });
  w.isInvoiceAdmin = () => true;
  w.alerts = [];
  w.showAlert = (m) => { w.alerts.push(String(m)); };
  return w;
}
const AUTH = { hours: 20, minutes: 0, tasks: [
  { task: 'Bathing', perDay: '00:10', freq: '7 days per week', perMonth: '05:00' }] };
const client = (invoices) => ({ clientName: 'Jane Doe', clientStatus: 'active',
  startDate: '2026-01-01', authorization: AUTH, invoices: invoices || [] });
const count = (w, period) => (w.getProfiles()['Jane Doe'].invoices || [])
  .filter((i) => i.billingPeriod === period).length;

test('an invoice that appeared while the dialog was open is not duplicated', () => {
  const w = app();
  // Eligibility was decided when the client had none...
  const eligible = [{ name: 'Jane Doe' }];
  // ...but by the time the user confirms, one exists.
  w.saveProfilesLS({ 'Jane Doe': client([
    { billingPeriod: '09/2026', status: 'submitted', dbId: 'db_09', data: { svcHH: '20' } }]) });
  w._doAutoGenerateInvoices(eligible, '09/2026');
  assert.strictEqual(count(w, '09/2026'), 1, 'the same client+period must not be invoiced twice');
  assert.strictEqual(w.getProfiles()['Jane Doe'].invoices[0].status, 'submitted',
    'and the existing submitted invoice must survive');
});

test('the skip is reported, not silent', () => {
  const w = app();
  w.saveProfilesLS({ 'Jane Doe': client([
    { billingPeriod: '09/2026', status: 'draft', data: {} }]) });
  w._doAutoGenerateInvoices([{ name: 'Jane Doe' }], '09/2026');
  assert.ok(w.alerts.some((m) => /already/i.test(m) && /Jane Doe/.test(m)),
    'expected the owner to be told: ' + JSON.stringify(w.alerts));
});

test('a client who genuinely has no invoice still gets one', () => {
  const w = app();
  w.saveProfilesLS({ 'Jane Doe': client([]) });
  w._doAutoGenerateInvoices([{ name: 'Jane Doe' }], '09/2026');
  assert.strictEqual(count(w, '09/2026'), 1);
  assert.strictEqual(w.getProfiles()['Jane Doe'].invoices[0].status, 'draft');
});

test('one client being skipped does not stop the others generating', () => {
  const w = app();
  w.saveProfilesLS({
    'Jane Doe': client([{ billingPeriod: '09/2026', status: 'draft', data: {} }]),
    'Bob Roe': Object.assign(client([]), { clientName: 'Bob Roe' }),
  });
  w._doAutoGenerateInvoices([{ name: 'Jane Doe' }, { name: 'Bob Roe' }], '09/2026');
  assert.strictEqual(count(w, '09/2026'), 1, 'Jane keeps exactly one');
  assert.strictEqual((w.getProfiles()['Bob Roe'].invoices || []).length, 1, 'Bob still gets his');
});
