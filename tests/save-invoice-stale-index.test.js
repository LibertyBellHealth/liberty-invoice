'use strict';
// saveInvoiceToClient read the invoice's array INDEX and status, then awaited a human click on
// "Overwrite?". A background loadProfilesAPI replaces the whole invoices array — re-sorted by
// saved_at and de-duplicated — and can land while that dialog is open. The write then went to
// whatever now sat at the old index: a different month's invoice, overwritten with this month's
// figures and carrying that month's dbId, so the next sync pushed the wrong data onto that
// month's server row and reset it to Draft. The "already submitted" warning did not fire either,
// because the status was stale as well.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

const IDS = ['clientName','clientName2','medicaidId','worker','worker2','billingPeriod','billingPeriod2',
  'hourlyRate','billTo','svcHH','svcMM','cplxHH','cplxMM','p1HH','p1MM','grandHH','grandMM',
  'dateSubmitted','sigDate1','sigDate2'];

function app() {
  const w = loadApp();
  resetStorage(w);
  if (!w.document.getElementById('sigArea1')) {
    w.document.body.insertAdjacentHTML('beforeend',
      IDS.map((i) => '<input id="' + i + '">').join('') +
      '<input type="checkbox" id="showComplex"><div id="complexSection"></div>' +
      '<div id="sigArea1"></div><div id="sigArea2"></div>' +
      '<button id="saveInvoiceBtn"></button><div id="dupWarning"></div>');
  }
  w.resetSigArea(1); w.resetSigArea(2);
  ['addAuditEntry','logActivity','aiTrack','saveProfileSP','renderInvoiceHistory','updateStats']
    .forEach((f) => { w[f] = () => {}; });
  w.activeProfileName = 'Jane Doe';
  return w;
}
const inv = (period, status, dbId) => ({ billingPeriod: period, status, dbId,
  savedAt: '1/1/2026', data: { billingPeriod: period, svcHH: '10', svcMM: '00' } });
const periods = (w) => w.getProfiles()['Jane Doe'].invoices.map((i) => i.billingPeriod + ':' + i.status);

test('the write lands on the period it was for, even if the array moved meanwhile', () => {
  const w = app();
  w.saveProfilesLS({ 'Jane Doe': { clientName: 'Jane Doe', invoices: [
    inv('09/2026', 'submitted', 'db_09'), inv('08/2026', 'draft', 'db_08')] } });
  w.document.getElementById('billingPeriod').value = '08/2026';
  // A sync arrives while the dialog is open: a new invoice is unshifted, shifting every index by one.
  const store = w.getProfiles();
  store['Jane Doe'].invoices.unshift(inv('10/2026', 'draft', 'db_10'));
  w.saveProfilesLS(store);
  w._doSaveInvoiceToClient('08/2026');          // what the confirm callback now does
  const list = w.getProfiles()['Jane Doe'].invoices;
  const sept = list.find((i) => i.billingPeriod === '09/2026');
  assert.strictEqual(sept.status, 'submitted', 'September must not be un-submitted: ' + periods(w));
  assert.strictEqual(sept.dbId, 'db_09', 'nor may it take another row\'s database id');
  assert.strictEqual(list.filter((i) => i.billingPeriod === '08/2026').length, 1,
    'August must be updated in place, not duplicated: ' + periods(w));
});

test('the invoice that was targeted is the one updated', () => {
  const w = app();
  w.saveProfilesLS({ 'Jane Doe': { clientName: 'Jane Doe', invoices: [
    inv('09/2026', 'draft', 'db_09'), inv('08/2026', 'draft', 'db_08')] } });
  w.document.getElementById('billingPeriod').value = '08/2026';
  w.document.getElementById('svcHH').value = '62';
  w._doSaveInvoiceToClient('08/2026');
  const list = w.getProfiles()['Jane Doe'].invoices;
  assert.strictEqual(list.find((i) => i.billingPeriod === '08/2026').data.svcHH, '62');
  assert.strictEqual(list.find((i) => i.billingPeriod === '09/2026').data.svcHH, '10',
    'the untargeted month must be untouched');
});

test('a period that no longer exists is created rather than overwriting a neighbour', () => {
  const w = app();
  w.saveProfilesLS({ 'Jane Doe': { clientName: 'Jane Doe', invoices: [inv('09/2026', 'submitted', 'db_09')] } });
  w.document.getElementById('billingPeriod').value = '08/2026';
  w._doSaveInvoiceToClient('08/2026');
  const list = w.getProfiles()['Jane Doe'].invoices;
  assert.strictEqual(list.length, 2);
  assert.strictEqual(list.find((i) => i.billingPeriod === '09/2026').status, 'submitted');
});

test('a row marked Paid while the dialog was open is still refused', () => {
  const w = app();
  let alerted = '';
  w.showAlert = (m) => { alerted = String(m); };
  w.saveProfilesLS({ 'Jane Doe': { clientName: 'Jane Doe', invoices: [inv('08/2026', 'paid', 'db_08')] } });
  w.document.getElementById('billingPeriod').value = '08/2026';
  w.document.getElementById('svcHH').value = '99';
  w._doSaveInvoiceToClient('08/2026');
  assert.match(alerted, /Paid and cannot be overwritten/);
  assert.strictEqual(w.getProfiles()['Jane Doe'].invoices[0].data.svcHH, '10', 'a paid invoice must not change');
});

test('an existing invoice keeps its status through an overwrite', () => {
  const w = app();
  w.saveProfilesLS({ 'Jane Doe': { clientName: 'Jane Doe', invoices: [inv('08/2026', 'submitted', 'db_08')] } });
  w.document.getElementById('billingPeriod').value = '08/2026';
  w._doSaveInvoiceToClient('08/2026');
  assert.strictEqual(w.getProfiles()['Jane Doe'].invoices[0].status, 'submitted');
});
