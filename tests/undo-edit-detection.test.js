'use strict';
// undoAutoGenBatch deletes auto-generated invoices, and refuses any the owner has since edited —
// decided by _quickInvoiceHash. That hash covered only service hours, complex hours and the day
// grid, so an invoice whose grand total, previous-page total, rate, dates or Bill To had been
// corrected hashed IDENTICALLY to an untouched one. Undo reported it pristine and deleted the
// corrections. Undo is irreversible, so anything editable has to count as an edit.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

const w = loadApp();
resetStorage(w);
const BASE = {
  clientName: 'Jane Doe', medicaidId: '1234567', worker: 'A Sawyer', billTo: 'MDHHS Macomb County',
  billingPeriod: '08/2026', hourlyRate: '27.00',
  svcHH: '20', svcMM: '00', cplxHH: '', cplxMM: '', p1HH: '', p1MM: '', grandHH: '20', grandMM: '00',
  dateSubmitted: '09/01/2026', sigDate1: '09/01/2026', sigDate2: '09/01/2026',
  hasComplex: false, sigId: '', tasks: { svc: [[true, false]], cplx: [] },
};
const changed = (over) => w._quickInvoiceHash(BASE) !== w._quickInvoiceHash(Object.assign({}, BASE, over));

test('an untouched invoice hashes the same', () => {
  assert.strictEqual(w._quickInvoiceHash(BASE), w._quickInvoiceHash(Object.assign({}, BASE)));
});

test('every billed figure the owner can correct counts as an edit', () => {
  const cases = {
    'service hours': { svcHH: '18' }, 'service minutes': { svcMM: '30' },
    'complex hours': { cplxHH: '2' }, 'complex minutes': { cplxMM: '15' },
    'previous-page hours': { p1HH: '2' }, 'previous-page minutes': { p1MM: '30' },
    'grand total hours': { grandHH: '18' }, 'grand total minutes': { grandMM: '30' },
    'hourly rate': { hourlyRate: '30.00' },
  };
  Object.keys(cases).forEach((k) => assert.ok(changed(cases[k]),
    k + ' was not detected — undo would delete that correction'));
});

test('dates, routing and the signature count too', () => {
  const cases = {
    'date submitted': { dateSubmitted: '09/15/2026' },
    'signature date 1': { sigDate1: '09/15/2026' },
    'signature date 2': { sigDate2: '09/15/2026' },
    'bill to': { billTo: 'MDHHS Wayne County' },
    'caseworker': { worker: 'R Feto' },
    'medicaid id': { medicaidId: '7654321' },
    'client name': { clientName: 'Bob Roe' },
    'complex care toggled': { hasComplex: true },
    'which signature certified it': { sigId: 'sig_B' },
  };
  Object.keys(cases).forEach((k) => assert.ok(changed(cases[k]), k + ' was not detected'));
});

test('a change to the day grid is still detected', () => {
  assert.ok(changed({ tasks: { svc: [[true, true]], cplx: [] } }));
  assert.ok(changed({ tasks: { svc: [[true, false]], cplx: [[true]] } }));
});

test('no data at all hashes to empty rather than throwing', () => {
  assert.strictEqual(w._quickInvoiceHash(null), '');
  assert.strictEqual(w._quickInvoiceHash(undefined), '');
  assert.doesNotThrow(() => w._quickInvoiceHash({}));
});
