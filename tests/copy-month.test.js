'use strict';
// copyMonth rolls the invoice forward a month. It carries the task grid and the header across but
// must NOT carry the billed totals, the old rate, the old dates or the signature — each of those
// has been a real defect in this file before (a reprint rewriting history, a stale rate, a stamp
// certifying a month it was never placed on).
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

const IDS = ['clientName','clientName2','medicaidId','worker','worker2','billingPeriod','billingPeriod2',
  'hourlyRate','billTo','svcHH','svcMM','cplxHH','cplxMM','p1HH','p1MM','grandHH','grandMM',
  'dateSubmitted','sigDate1','sigDate2'];

function app(period, opts) {
  opts = opts || {};
  const w = loadApp();
  resetStorage(w);
  if (!w.document.getElementById('sigArea1')) {
    w.document.body.insertAdjacentHTML('beforeend',
      IDS.map((i) => '<input id="' + i + '">').join('') +
      '<input type="checkbox" id="showComplex"><div id="complexSection"></div>' +
      '<div id="sigArea1"></div><div id="sigArea2"></div>');
  }
  w.resetSigArea(1); w.resetSigArea(2);
  w.rebuiltWith = null;
  w.appliedStates = null;
  w.rebuild = (days) => { w.rebuiltWith = days; };
  w.applyStates = (s) => { w.appliedStates = s; };
  w.captureStates = () => opts.states || { svc: [[true]], cplx: [] };
  w.toggleComplex = () => {};
  w.showAlert = (m) => { w.alerted = String(m); };
  const set = (id, v) => { const e = w.document.getElementById(id); if (e) e.value = v; };
  Object.assign({}, opts.values || {});
  set('billingPeriod', period);
  set('clientName', 'Jane Doe'); set('medicaidId', '1234567'); set('worker', 'A Sawyer');
  set('billTo', 'MDHHS Macomb County'); set('hourlyRate', '25.00');
  set('svcHH', '62'); set('svcMM', '21'); set('grandHH', '62'); set('grandMM', '21');
  set('dateSubmitted', '01/01/2020'); set('sigDate1', '01/01/2020'); set('sigDate2', '01/01/2020');
  return w;
}
const val = (w, id) => w.document.getElementById(id).value;

test('the period advances one month', () => {
  const w = app('08/2026');
  w.copyMonth();
  assert.strictEqual(val(w, 'billingPeriod'), '09/2026');
  assert.strictEqual(val(w, 'billingPeriod2'), '09/2026');
});

test('December rolls into January of the next year', () => {
  const w = app('12/2026');
  w.copyMonth();
  assert.strictEqual(val(w, 'billingPeriod'), '01/2027');
});

test('billed totals are CLEARED, never carried into the new month', () => {
  const w = app('08/2026');
  w.copyMonth();
  ['svcHH', 'svcMM', 'cplxHH', 'cplxMM', 'p1HH', 'p1MM', 'grandHH', 'grandMM']
    .forEach((id) => assert.strictEqual(val(w, id), '', id + ' carried forward'));
});

test('the rate is re-stamped from Settings, not copied from the old invoice', () => {
  const w = app('08/2026');
  w.localStorage.setItem('lhca_state_rate', '28.50');
  w.copyMonth();
  assert.strictEqual(val(w, 'hourlyRate'), '28.50', 'a stale rate would bill last year\'s number');
});

test('dates are re-stamped to today', () => {
  const w = app('08/2026');
  w.copyMonth();
  ['dateSubmitted', 'sigDate1', 'sigDate2'].forEach((id) =>
    assert.strictEqual(val(w, id), w.today(), id + ' kept the old invoice\'s date'));
});

test('the signature is not carried onto the new month', () => {
  const w = app('08/2026');
  w.stampSignatureData(1, 'data:image/png;base64,AAA', 'sig_A');
  assert.strictEqual(w.document.getElementById('sigArea1').tagName, 'IMG');
  w.copyMonth();
  assert.strictEqual(w.document.getElementById('sigArea1').tagName, 'DIV',
    'a stamp must never certify a month it was not placed on');
});

test('the day grid is rebuilt for the NEW month\'s length', () => {
  const w = app('08/2026');    // 31 days -> September has 30
  w.copyMonth();
  assert.strictEqual(w.rebuiltWith, 30);
  const w2 = app('01/2026');   // 31 -> February 2026 has 28
  w2.copyMonth();
  assert.strictEqual(w2.rebuiltWith, 28);
});

test('client identity carries across unchanged', () => {
  const w = app('08/2026');
  w.copyMonth();
  assert.strictEqual(val(w, 'clientName'), 'Jane Doe');
  assert.strictEqual(val(w, 'medicaidId'), '1234567');
  assert.strictEqual(val(w, 'worker'), 'A Sawyer');
  assert.strictEqual(val(w, 'billTo'), 'MDHHS Macomb County');
});

test('a missing or malformed period is refused', () => {
  const w = app('');
  w.copyMonth();
  assert.match(String(w.alerted || ''), /billing period/i);
  assert.strictEqual(val(w, 'billingPeriod'), '', 'nothing should have been rolled forward');
});
