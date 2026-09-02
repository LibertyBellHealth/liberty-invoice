'use strict';
// Reopening a stored invoice reset both signature areas and never put the stored one back. The PDF
// and email paths read the signature off the DOM, so a reopened CERTIFIED invoice printed and
// emailed with a blank line under "I certify that Liberty Home Care Assistance has provided all the
// services as checked above." captureFullInvoice() then read data-sig-id as '' and wrote that blank
// over the record, after which a reprint fell back to sigs[0] — re-certifying a sent invoice under
// a different person. That is the exact failure sigId was introduced to prevent.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

const IDS = ['clientName','clientName2','medicaidId','worker','worker2','billingPeriod','billingPeriod2',
  'hourlyRate','billTo','svcHH','svcMM','cplxHH','cplxMM','p1HH','p1MM','grandHH','grandMM',
  'dateSubmitted','sigDate1','sigDate2'];

function app(sigs) {
  const w = loadApp();
  resetStorage(w);           // signatures persist in localStorage between loadApp() calls
  // loadApp() reuses one jsdom document, so build the fields ONCE — appending them per test would
  // duplicate every id and getElementById would keep returning the first test's stale nodes.
  if (!w.document.getElementById('sigArea1')) {
    w.document.body.insertAdjacentHTML('beforeend',
      IDS.map((i) => '<input id="' + i + '">').join('') +
      '<input type="checkbox" id="showComplex"><div id="complexSection"></div>' +
      '<div id="sigArea1"></div><div id="sigArea2"></div>');
  }
  w.resetSigArea(1); w.resetSigArea(2);   // start every test from an unsigned form
  ['rebuild', 'applyStates', 'renderNotesPane'].forEach((f) => { if (typeof w[f] === 'function') w[f] = () => {}; });
  // Deep-copy: saveSigsLS mutates what it is handed, so a shared fixture leaks between tests.
  if (sigs) w.saveSigsLS(JSON.parse(JSON.stringify(sigs)));
  return w;
}
const SIGS = [{ id: 'sig_A', label: 'Owner A', data: 'data:image/png;base64,AAAA' },
               { id: 'sig_B', label: 'Owner B', data: 'data:image/png;base64,BBBB' }];
const invoice = (over) => Object.assign({ clientName: 'Jane Doe', billingPeriod: '08/2026',
  svcHH: '62', svcMM: '21', hourlyRate: '27.00', tasks: { svc: [], cplx: [] } }, over || {});

test('reopening a certified invoice replays the signature it was certified with', () => {
  const w = app(SIGS);
  w.applyFullInvoice(invoice({ sigId: 'sig_B' }));
  const el = w.document.getElementById('sigArea1');
  assert.strictEqual(el.tagName, 'IMG', 'the signature must be back on the form');
  assert.strictEqual(el.getAttribute('data-sig-id'), 'sig_B');
});

test('and re-capturing it does not blank the provenance', () => {
  const w = app(SIGS);
  w.applyFullInvoice(invoice({ sigId: 'sig_B' }));
  assert.strictEqual(w.captureFullInvoice().sigId, 'sig_B',
    'a blank here is written back over the certified record');
});

test('it never substitutes a different signature', () => {
  const w = app(SIGS);
  w.applyFullInvoice(invoice({ sigId: 'sig_GONE' }));   // certified with one since deleted
  const el = w.document.getElementById('sigArea1');
  assert.notStrictEqual(el.getAttribute('data-sig-id'), 'sig_A',
    'falling back to the first signature re-certifies under the wrong person');
  assert.strictEqual(w.captureFullInvoice().sigId, '');
});

test('an unsigned invoice stays unsigned', () => {
  const w = app(SIGS);
  w.applyFullInvoice(invoice({}));
  assert.strictEqual(w.document.getElementById('sigArea1').tagName, 'DIV');
  assert.strictEqual(w.captureFullInvoice().sigId, '');
});

test('a previous month\'s stamp still cannot carry over', () => {
  const w = app(SIGS);
  w.applyFullInvoice(invoice({ sigId: 'sig_B' }));
  assert.strictEqual(w.document.getElementById('sigArea1').getAttribute('data-sig-id'), 'sig_B');
  w.applyFullInvoice(invoice({ billingPeriod: '09/2026' }));   // next month, never signed
  assert.strictEqual(w.document.getElementById('sigArea1').tagName, 'DIV',
    'the prior invoice\'s signature must not persist');
});
