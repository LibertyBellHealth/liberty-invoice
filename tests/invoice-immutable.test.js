'use strict';
// OWNER RULE (2026-08-23): "An old invoice should never be touched." An invoice is a certified
// record of what was billed to MDHHS. Both the edit path and the PDF/email path were re-deriving
// the hourly rate from today's Settings value, so a prior-year invoice displayed, re-saved and
// re-printed at the CURRENT rate. The state rate changes annually, so this rewrites history.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

function invoiceDom(w) {
  if (!w.document.getElementById('hourlyRate')) {
    w.document.body.insertAdjacentHTML('beforeend',
      ['clientName','clientName2','medicaidId','worker','worker2','billingPeriod','billingPeriod2',
       'hourlyRate','billTo','svcHH','svcMM','cplxHH','cplxMM','p1HH','p1MM','grandHH','grandMM',
       'dateSubmitted','sigDate1','sigDate2'].map((i) => '<input id="' + i + '">').join('') +
      '<input type="checkbox" id="showComplex"><div id="complexSection"></div>');
  }
  ['rebuild','applyStates','resetSigArea','renderNotesPane','updateInvTotals','toggleComplex']
    .forEach((f) => { if (typeof w[f] === 'function') w[f] = () => {}; });
}

test('opening an old invoice shows the rate it was BILLED at, not today’s', () => {
  const w = loadApp(); resetStorage(w); invoiceDom(w);
  w.localStorage.setItem('lhca_state_rate', '27.00');      // today's rate
  w.saveProfilesLS({ 'Jane Doe': { clientName: 'Jane Doe', invoices: [] } });
  w.activeProfileName = 'Jane Doe';
  w.applyFullInvoice({ clientName: 'Jane Doe', billingPeriod: '03/2025', hourlyRate: '24.50',
    svcHH: '20', svcMM: '00' });
  assert.strictEqual(w.document.getElementById('hourlyRate').value, '24.50',
    'a 2025 invoice must not be redisplayed at the 2026 rate — saving it then writes that rate over the original');
});

test('a NEW invoice still gets the current state rate', () => {
  const w = loadApp(); resetStorage(w); invoiceDom(w);
  w.localStorage.setItem('lhca_state_rate', '27.00');
  w.saveProfilesLS({ 'Jane Doe': { clientName: 'Jane Doe', invoices: [] } });
  w.activeProfileName = 'Jane Doe';
  w.applyFullInvoice({ clientName: 'Jane Doe', billingPeriod: '08/2026' });  // no stored rate
  assert.strictEqual(w.document.getElementById('hourlyRate').value, '27.00',
    'with no rate on the invoice the current Settings rate is correct');
});

test('re-printing or re-emailing an old invoice uses the rate it was billed at', async () => {
  const w = loadApp(); resetStorage(w); invoiceDom(w);
  w.localStorage.setItem('lhca_state_rate', '27.00');
  w.saveProfilesLS({ 'Jane Doe': { clientName: 'Jane Doe', medicaidId: '123', invoices: [] } });
  await w.loadInvoiceForCapture('Jane Doe',
    { billingPeriod: '03/2025', data: { hourlyRate: '24.50', svcHH: '20', svcMM: '00' } }, '03/2025');
  assert.strictEqual(w.document.getElementById('hourlyRate').value, '24.50',
    'the emailed/printed PDF of a past month certified today’s rate');
});
