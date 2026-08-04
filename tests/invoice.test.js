'use strict';
// Invoice logic — the billing paths the audit flagged as uncovered. All pure functions from app.js.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp } = require('./harness');

test('daysIn: month lengths including leap years', () => {
  const w = loadApp();
  assert.strictEqual(w.daysIn('02', '2024'), 29, 'Feb leap year');
  assert.strictEqual(w.daysIn('02', '2025'), 28, 'Feb non-leap');
  assert.strictEqual(w.daysIn('04', '2026'), 30, 'April');
  assert.strictEqual(w.daysIn('12', '2026'), 31, 'December');
  assert.strictEqual(w.daysIn('13', '2026'), 31, 'invalid month -> 31 fallback');
  assert.strictEqual(w.daysIn('x', 'y'), 31, 'garbage -> 31 fallback, no crash');
});

test('_invoiceSig: detects edits so an invoice save actually fires', () => {
  const w = loadApp();
  const inv = { billingPeriod: '08/2026', status: 'draft', invoiceNote: '', data: { svcHH: '29', svcMM: '47' } };
  assert.strictEqual(w._invoiceSig(inv), w._invoiceSig({ ...inv }), 'identical -> same signature');
  assert.notStrictEqual(w._invoiceSig(inv), w._invoiceSig({ ...inv, status: 'submitted' }), 'status change detected');
  assert.notStrictEqual(w._invoiceSig(inv), w._invoiceSig({ ...inv, invoiceNote: 'paid' }), 'note change detected');
  assert.notStrictEqual(w._invoiceSig(inv), w._invoiceSig({ ...inv, data: { svcHH: '30', svcMM: '47' } }), 'hours change detected');
});

test('_quickInvoiceHash: stable for same data, changes on an hours edit', () => {
  const w = loadApp();
  const data = { svcHH: '29', svcMM: '47', cplxHH: '', cplxMM: '', tasks: { svc: [], cplx: [] } };
  assert.strictEqual(w._quickInvoiceHash(data), w._quickInvoiceHash({ ...data }), 'same -> same hash');
  assert.notStrictEqual(w._quickInvoiceHash(data), w._quickInvoiceHash({ ...data, svcHH: '30' }), 'hours edit -> different hash');
  assert.strictEqual(w._quickInvoiceHash(null), '', 'null -> empty, no crash');
});

// Build a [days][cols] boolean service grid.
function grid(days, cols, fill) {
  const g = [];
  for (let d = 0; d < days; d++) { const r = []; for (let c = 0; c < cols; c++) r.push(fill(d, c)); g.push(r); }
  return g;
}

test('generateNextMonthInvoiceData: carries service-day patterns forward correctly', () => {
  const w = loadApp();
  const COLS = 15, HOSP = 14;
  // Prior month = July 2026 (31 days). Column patterns exercise each carry-forward rule:
  const svc = grid(31, COLS, (d, c) => {
    if (c === 0) return true;      // checked every day -> "daily" -> stays daily
    if (c === 1) return false;     // never checked -> stays cleared
    if (c === 2) return d <= 4;    // 5 checks -> "shift" by +1 day
    if (c === HOSP) return true;   // Hospital (last col) -> ALWAYS cleared (by-exception)
    return false;
  });
  const prevInv = {
    billingPeriod: '07/2026', savedAt: 'x', status: 'submitted',
    data: { billingPeriod: '07/2026', svcHH: '29', svcMM: '47', tasks: { svc, cplx: [] } },
  };
  const r = w.generateNextMonthInvoiceData(prevInv, '08/2026');
  assert.ok(r, 'returns a new invoice');
  assert.strictEqual(r.billingPeriod, '08/2026', 'new billing period');
  assert.strictEqual(r.status, 'draft', 'auto-generated invoice is a Draft');
  assert.strictEqual(r.data.svcHH, '29', 'hours carried over from the prior invoice');
  const ns = r.data.tasks.svc;
  assert.strictEqual(ns.length, 31, 'August has 31 rows');
  assert.ok(ns.every(row => row[0] === true), 'daily column stays daily');
  assert.ok(ns.every(row => row[1] === false), 'empty column stays empty');
  assert.ok(ns.every(row => row[HOSP] === false), 'Hospital column always cleared');
  assert.strictEqual(ns[1][2], true, 'shift +1: new day 1 = prior day 0 (checked)');
  assert.strictEqual(ns[0][2], false, 'shift +1: new day 0 wraps to prior last day (unchecked)');
});

test('generateNextMonthInvoiceData: guards bad input instead of producing a garbage bill', () => {
  const w = loadApp();
  assert.strictEqual(w.generateNextMonthInvoiceData(null, '08/2026'), null, 'no prior invoice -> null');
  assert.strictEqual(w.generateNextMonthInvoiceData({ data: null }, '08/2026'), null, 'no data -> null');
  assert.strictEqual(w.generateNextMonthInvoiceData({ data: { billingPeriod: '07/2026' } }, 'bad'), null, 'bad period -> null');
});
