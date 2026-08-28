'use strict';
// Two ways an invoice reached MDHHS with a number that was not true.
//
// 1. _startsInsidePeriod parsed the stored 'YYYY-MM-DD' start date with `new Date()`, which reads
//    it as UTC midnight, then queried it with LOCAL getters. Every date shifted back a day in
//    Michigan, so a start on the 2nd looked like the 1st — a FULL month — and auto-generate billed
//    the whole authorized month for a month that began on the 2nd.
// 2. applyFullInvoice re-derived "Bill To" from the caseworker's CURRENT agency, so opening an old
//    invoice rewrote who it had been billed to, and saving persisted the rewrite.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

// ── 1. calendar dates are components, not instants ────────────────────────────
test('_startsInsidePeriod: a start on the 2nd IS a partial month', () => {
  const w = loadApp(); resetStorage(w);
  assert.strictEqual(w._startsInsidePeriod({ startDate: '2026-08-02' }, '08/2026'), true,
    'the 2nd read back as the 1st under UTC parsing, so proration was never offered');
});

test('_startsInsidePeriod: a start on the 1st is a FULL month', () => {
  const w = loadApp(); resetStorage(w);
  assert.strictEqual(w._startsInsidePeriod({ startDate: '2026-08-01' }, '08/2026'), false);
});

test('_startsInsidePeriod: every day of the month is classified correctly', () => {
  const w = loadApp(); resetStorage(w);
  for (let d = 1; d <= 31; d++) {
    const sd = '2026-08-' + String(d).padStart(2, '0');
    assert.strictEqual(w._startsInsidePeriod({ startDate: sd }, '08/2026'), d > 1,
      sd + ' misclassified');
  }
});

test('_startsInsidePeriod: a start in a DIFFERENT month is not a partial month here', () => {
  const w = loadApp(); resetStorage(w);
  assert.strictEqual(w._startsInsidePeriod({ startDate: '2026-09-01' }, '08/2026'), false);
  assert.strictEqual(w._startsInsidePeriod({ startDate: '2026-07-15' }, '08/2026'), false);
});

test('auto-generate HOLDS BACK a client whose service starts on the 2nd', () => {
  const w = loadApp(); resetStorage(w);
  w.saveProfilesLS({ 'Jane Doe': { clientName: 'Jane Doe', clientStatus: 'active',
    startDate: '2026-08-02', invoices: [],
    authorization: { hours: '30', minutes: '0', effectiveDate: '08/02/2026', tasks: [] } } });
  const eligible = w.findClientsEligibleForAutoGen('08/2026');
  const jane = eligible.find(e => e.name === 'Jane Doe');
  assert.ok(jane, 'she is otherwise eligible');
  assert.strictEqual(jane.partialMonth, true,
    'without this she is auto-generated at the FULL authorized 30:00 for a month that began on the 2nd');
});

// ── 2. Bill To is a certified fact, not a lookup ──────────────────────────────
function invoiceDom(w) {
  if (!w.document.getElementById('billTo')) {
    w.document.body.insertAdjacentHTML('beforeend',
      ['clientName','clientName2','medicaidId','worker','worker2','billingPeriod','billingPeriod2',
       'hourlyRate','billTo','svcHH','svcMM','cplxHH','cplxMM','p1HH','p1MM','grandHH','grandMM',
       'dateSubmitted','sigDate1','sigDate2'].map((i) => '<input id="' + i + '">').join('') +
      '<input type="checkbox" id="showComplex"><div id="complexSection"></div>');
  }
  ['rebuild','applyStates','resetSigArea','renderNotesPane','updateInvTotals','toggleComplex']
    .forEach((f) => { if (typeof w[f] === 'function') w[f] = () => {}; });
}

test('opening an old invoice keeps the agency it was BILLED to', () => {
  const w = loadApp(); resetStorage(w); invoiceDom(w);
  // The caseworker has since moved to a different agency.
  w.saveCaseworkersLS([{ id: 'cw1', name: 'Pat Worker', agency: 'Wayne County DHHS TODAY' }]);
  w.saveProfilesLS({ 'Jane Doe': { clientName: 'Jane Doe', caseworkerId: 'cw1', worker: 'Pat Worker', invoices: [] } });
  w.activeProfileName = 'Jane Doe';
  w.applyFullInvoice({ clientName: 'Jane Doe', worker: 'Pat Worker', billingPeriod: '03/2025',
    billTo: 'Oakland County DHHS', hourlyRate: '24.50', svcHH: '20', svcMM: '00' });
  assert.strictEqual(w.document.getElementById('billTo').value, 'Oakland County DHHS',
    'a 2025 invoice must not be re-addressed to the agency the caseworker joined later — saving it then rewrites the record');
});

test('an invoice with no Bill To still falls back to the live caseworker', () => {
  const w = loadApp(); resetStorage(w); invoiceDom(w);
  w.saveCaseworkersLS([{ id: 'cw1', name: 'Pat Worker', agency: 'Wayne County DHHS' }]);
  w.saveProfilesLS({ 'Jane Doe': { clientName: 'Jane Doe', caseworkerId: 'cw1', worker: 'Pat Worker', invoices: [] } });
  w.activeProfileName = 'Jane Doe';
  w.applyFullInvoice({ clientName: 'Jane Doe', worker: 'Pat Worker', billingPeriod: '08/2026' });
  assert.strictEqual(w.document.getElementById('billTo').value, 'Wayne County DHHS',
    'a fresh invoice, or one saved before this field existed, still gets a Bill To');
});
