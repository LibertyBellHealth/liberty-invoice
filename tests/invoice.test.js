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

test('clientDueForInvoice: only flags a missing invoice when a start date is on/before the period (#10)', () => {
  const w = loadApp();
  assert.strictEqual(w.clientDueForInvoice({}, '07/2026'), false, 'no start date -> not due (this was the bug)');
  assert.strictEqual(w.clientDueForInvoice({ startDate: '2026-05-01' }, '07/2026'), true, 'started before the period -> due');
  assert.strictEqual(w.clientDueForInvoice({ startDate: '2026-09-01' }, '07/2026'), false, 'starts after the period -> not due');
  assert.strictEqual(w.clientDueForInvoice({ startDate: '2026-07-15' }, '07/2026'), true, 'starts within the period -> due');
});

test('clientInvoiceRate: client rate if set (from DHS or manual), else the Settings default (#1)', () => {
  const w = loadApp();
  w.localStorage.setItem('lhca_state_rate', '27.00');
  assert.strictEqual(w.clientInvoiceRate({ hourlyRate: '30.00' }), '30.00', 'client rate used');
  assert.strictEqual(w.clientInvoiceRate({ hourlyRate: '' }), '27.00', 'empty -> state default');
  assert.strictEqual(w.clientInvoiceRate({}), '27.00', 'no rate -> state default');
  assert.strictEqual(w.clientInvoiceRate({ hourlyRate: '0' }), '27.00', 'zero -> state default');
  assert.strictEqual(w.clientInvoiceRate({ hourlyRate: '$28.50' }), '28.50', 'strips formatting');
});

test('DHS rate is suggested as the client hourly rate, normalized to 2 decimals (#1)', () => {
  const w = loadApp();
  const ups = w._dhsSuggestedUpdates({ rate: 27 }, { hourlyRate: '' }, null);
  const rate = ups.find(u => u.field === 'hourlyRate');
  assert.ok(rate, 'rate is suggested when client has none');
  assert.strictEqual(rate.to, '27.00', 'normalized to 2 decimals');
  // stored "27.00" vs form 27 must NOT be flagged (same value, different format)
  const same = w._dhsSuggestedUpdates({ rate: 27 }, { hourlyRate: '27.00' }, null);
  assert.ok(!same.some(u => u.field === 'hourlyRate'), 'identical rate not flagged');
});

test('_dhsMapTaskToCol: every DHS-1210 task maps to the right service column (#1)', () => {
  const w = loadApp();
  const cols = w._dhsSvcColNames();
  const at = (name) => cols[w._dhsMapTaskToCol(name)];
  assert.strictEqual(at('Bathing'), 'Bathing', 'exact');
  assert.strictEqual(at('Transferring'), 'Transferring', 'exact');
  assert.strictEqual(at('Meal Preparation'), 'Meal Preparation', 'exact');
  assert.strictEqual(at('Shopping for Food/Meds'), 'Shopping', 'alias -> Shopping');
  assert.strictEqual(at('Travel For Shopping'), 'Travel Time for Shopping', 'alias -> Travel Time for Shopping');
  assert.strictEqual(w._dhsMapTaskToCol('Nonexistent Task'), -1, 'unknown -> -1 (surfaced, never wrong column)');
});

test('_dhsFreqToDays: frequency -> a reviewable day pattern', () => {
  const w = loadApp();
  assert.strictEqual(w._dhsFreqToDays('7 days per week', 31).length, 31, 'daily -> every day');
  assert.deepStrictEqual([...w._dhsFreqToDays('once per month', 31)], [0], 'monthly -> just the 1st');
  const wk = w._dhsFreqToDays('2 days per week', 14);
  assert.deepStrictEqual([...wk], [0, 1, 7, 8], '2x/week -> 2 days each week');
  assert.deepStrictEqual([...w._dhsFreqToDays('weird', 31)], [0], 'unknown -> 1st day only');
});

test('_dhsBuildFirstInvoice: builds a correct draft from the authorization, flags unmapped (#1)', () => {
  const w = loadApp();
  const res = {
    hours: 29, minutes: 47, rate: 27,
    tasks: [
      { task: 'Bathing', freq: '7 days per week' },
      { task: 'Shopping for Food/Meds', freq: 'Once per month' },
      { task: 'Mystery Service', freq: '1 day per week' }, // won't map
    ],
  };
  const built = w._dhsBuildFirstInvoice(res, { clientName: 'Jane', medicaidId: 'M1' }, '08/2026');
  assert.ok(built, 'returns a build');
  const cols = w._dhsSvcColNames();
  const svc = built.data.tasks.svc;
  assert.strictEqual(svc.length, 31, 'August has 31 day-rows');
  assert.ok(svc.every(r => r[cols.indexOf('Bathing')] === true), 'Bathing checked every day');
  assert.strictEqual(svc[0][cols.indexOf('Shopping')], true, 'Shopping checked on the 1st');
  assert.strictEqual(svc[5][cols.indexOf('Shopping')], false, 'Shopping not checked mid-month');
  assert.strictEqual(built.data.svcHH, '29', 'hours from authorization');
  assert.strictEqual(built.data.grandHH, '29', 'grand hours from authorization');
  assert.strictEqual(built.data.hourlyRate, '27.00', 'rate from the form');
  assert.strictEqual(built.data.billingPeriod, '08/2026');
  assert.deepStrictEqual([...built.unmapped], ['Mystery Service'], 'unmapped task surfaced, not dropped');
});

test('_dhsBuildFirstInvoice: bad period -> null (no garbage invoice)', () => {
  const w = loadApp();
  assert.strictEqual(w._dhsBuildFirstInvoice({ tasks: [] }, {}, 'bad'), null);
});

test('generateNextMonthInvoiceData: guards bad input instead of producing a garbage bill', () => {
  const w = loadApp();
  assert.strictEqual(w.generateNextMonthInvoiceData(null, '08/2026'), null, 'no prior invoice -> null');
  assert.strictEqual(w.generateNextMonthInvoiceData({ data: null }, '08/2026'), null, 'no data -> null');
  assert.strictEqual(w.generateNextMonthInvoiceData({ data: { billingPeriod: '07/2026' } }, 'bad'), null, 'bad period -> null');
});
