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

test('generateNextMonthInvoiceData: a 2x/month task keeps 2 checks across a 30->31 month (no phantom 3rd)', () => {
  const w = loadApp();
  const COLS = 15, LAUNDRY = 3;
  // June (30 days): Laundry on day 15 and day 30 (indices 14 and 29 — 29 is June's LAST day).
  const svc = grid(30, COLS, (d, c) => c === LAUNDRY && (d === 14 || d === 29));
  const prevInv = { billingPeriod: '06/2026', savedAt: 'x', status: 'submitted',
    data: { billingPeriod: '06/2026', svcHH: '3', svcMM: '13', tasks: { svc, cplx: [] } } };
  const r = w.generateNextMonthInvoiceData(prevInv, '07/2026');
  const ns = r.data.tasks.svc;
  assert.strictEqual(ns.length, 31, 'July has 31 rows');
  const checkedDays = ns.map((row, i) => row[LAUNDRY] ? i + 1 : null).filter(Boolean);
  // June 15 & 30 shift forward one day -> July 16 & 31 (exactly two, NOT the old buggy 1/16/31).
  // .join() keeps the compare realm-safe (jsdom array vs Node literal would fail deepStrictEqual).
  assert.strictEqual(checkedDays.join(','), '16,31', 'exactly two laundry days, shifted +1 — no phantom day 1');
});

test('generateNextMonthInvoiceData: a check on the last day of a longer month wraps once (no dupe)', () => {
  const w = loadApp();
  const COLS = 15, COL = 4;
  // A 31-day month with the task on its LAST day (index 30); next month is shorter (30 days).
  const svc = grid(31, COLS, (d, c) => c === COL && d === 30);
  const prevInv = { billingPeriod: '07/2026', savedAt: 'x', status: 'submitted',
    data: { billingPeriod: '07/2026', tasks: { svc, cplx: [] } } };
  const r = w.generateNextMonthInvoiceData(prevInv, '08/2026'); // Aug has 31, but test the shorter path:
  const r2 = w.generateNextMonthInvoiceData(prevInv, '09/2026'); // Sep has 30
  const count = r2.data.tasks.svc.reduce((n, row) => n + (row[COL] ? 1 : 0), 0);
  assert.strictEqual(count, 1, 'the single check is preserved exactly once (wrapped), never duplicated');
});

test('clientDueForInvoice: only flags a missing invoice when a start date is on/before the period (#10)', () => {
  const w = loadApp();
  assert.strictEqual(w.clientDueForInvoice({}, '07/2026'), false, 'no start date -> not due (this was the bug)');
  assert.strictEqual(w.clientDueForInvoice({ startDate: '2026-05-01' }, '07/2026'), true, 'started before the period -> due');
  assert.strictEqual(w.clientDueForInvoice({ startDate: '2026-09-01' }, '07/2026'), false, 'starts after the period -> not due');
  assert.strictEqual(w.clientDueForInvoice({ startDate: '2026-07-15' }, '07/2026'), true, 'starts within the period -> due');
});

test('clientInvoiceRate: ALWAYS the state rate — the client hourly field never bills', () => {
  const w = loadApp();
  w.localStorage.setItem('lhca_state_rate', '27.00');
  // Whatever is in the client's own rate field, the invoice bills the flat state rate.
  assert.strictEqual(w.clientInvoiceRate({ hourlyRate: '30.00' }), '27.00', 'client field ignored');
  assert.strictEqual(w.clientInvoiceRate({ hourlyRate: '13.50' }), '27.00', 'caregiver-pay value ignored');
  assert.strictEqual(w.clientInvoiceRate({ hourlyRate: '' }), '27.00', 'empty -> state rate');
  assert.strictEqual(w.clientInvoiceRate({}), '27.00', 'no field -> state rate');
  // Honors a changed Settings rate (the once-a-year state change).
  w.localStorage.setItem('lhca_state_rate', '28.00');
  assert.strictEqual(w.clientInvoiceRate({ hourlyRate: '30.00' }), '28.00', 'follows the Settings state rate');
});

test('DHS import no longer pushes the provider rate into the client hourly field (#1)', () => {
  const w = loadApp();
  // Even with a client rate that differs from the DHS provider rate, no hourlyRate change is suggested.
  const ups = w._dhsSuggestedUpdates({ rate: 27, medicaidId: '111' }, { hourlyRate: '', medicaidId: '000' }, null);
  assert.ok(!ups.some(u => u.field === 'hourlyRate'), 'client hourly rate is never suggested from the form');
  // other genuine suggestions still work (e.g. Medicaid ID)
  assert.ok(ups.some(u => u.field === 'medicaidId'), 'unrelated suggestions still fire');
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
  // Word-form per-month counts (MDHHS forms spell these out) — must place the right COUNT.
  assert.strictEqual(w._dhsFreqToDays('Twice per month', 31).length, 2, '"Twice per month" -> 2 days');
  assert.deepStrictEqual([...w._dhsFreqToDays('Twice per month', 30)], [0, 15], '2/month spread evenly');
  assert.strictEqual(w._dhsFreqToDays('three times per month', 31).length, 3, '"three times per month" -> 3 days');
  assert.strictEqual(w._dhsFreqToDays('2 times per month', 31).length, 2, 'numeric "times per month" too');
});

test('_dhsBuildFirstInvoice: builds a correct draft from the authorization, flags unmapped (#1)', () => {
  const w = loadApp();
  w.localStorage.setItem('lhca_state_rate', '27.00');  // invoice rate always comes from Settings state rate
  const res = {
    hours: 29, minutes: 47, rate: 99,   // form's printed rate is IGNORED — the invoice bills the state rate
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
  assert.ok(svc.every(r => r[cols.indexOf('Bathing')] === true), 'Bathing (7 days/week) checked EVERY day');
  // Shopping is "Once per month" -> exactly ONE checked day (which day now varies by billing period).
  const shopCount = svc.reduce((n, r) => n + (r[cols.indexOf('Shopping')] ? 1 : 0), 0);
  assert.strictEqual(shopCount, 1, 'Once per month -> exactly one Shopping day');
  assert.strictEqual(built.data.svcHH, '29', 'hours from authorization');
  assert.strictEqual(built.data.grandHH, '29', 'grand hours from authorization');
  assert.strictEqual(built.data.hourlyRate, '27.00', 'invoice bills the state rate, not the form rate (99)');
  assert.strictEqual(built.data.billingPeriod, '08/2026');
  assert.deepStrictEqual([...built.unmapped], ['Mystery Service'], 'unmapped task surfaced, not dropped');
});

test('_dhsTaskDays: daily task (7 days/week) -> every day of the month', () => {
  const w = loadApp();
  assert.strictEqual(w._dhsTaskDays({ freq: '7 days per week' }, 31, 5).length, 31, 'Jul -> 31');
  assert.strictEqual(w._dhsTaskDays({ freq: '7 days per week' }, 30, 5).length, 30, 'Jun -> 30');
});

test('_dhsTaskDays: count = floor(Time/Month ÷ Time/Day), never over-documents', () => {
  const w = loadApp();
  // Mobility: 2:00/month at 0:14/day -> floor(120/14)=8 (8*14=112=1:52, under 2:00 — the safe side).
  assert.strictEqual(w._dhsTaskDays({ freq: '2 days per week', perDay: '00:14', perMonth: '02:00' }, 31, 3).length, 8);
  // Laundry: 2:00/month at 1:00/day -> exactly 2.
  assert.strictEqual(w._dhsTaskDays({ freq: 'Twice per month', perDay: '01:00', perMonth: '02:00' }, 31, 3).length, 2);
});

test('_dhsTaskDays: same count each month but DIFFERENT days (varied, not identical)', () => {
  const w = loadApp();
  const task = { freq: '2 days per week', perDay: '00:14', perMonth: '02:00' };
  const jul = w._dhsTaskDays(task, 31, w._dhsPeriodSeed('07/2026')).join(',');
  const aug = w._dhsTaskDays(task, 31, w._dhsPeriodSeed('08/2026')).join(',');
  assert.strictEqual(jul.split(',').length, aug.split(',').length, 'same number of days each month');
  assert.notStrictEqual(jul, aug, 'but the specific days differ month to month');
});

test('_dhsSpreadDays / _dhsHmToMin / _dhsPeriodSeed: building blocks', () => {
  const w = loadApp();
  assert.strictEqual(w._dhsHmToMin('02:00'), 120);
  assert.strictEqual(w._dhsHmToMin('00:14'), 14);
  assert.strictEqual(w._dhsSpreadDays(2, 30, 0).join(','), '0,15', '2 spread across 30 with no offset');
  assert.strictEqual(w._dhsSpreadDays(2, 30, 5).join(','), '5,20', 'offset by seed');
  assert.strictEqual(w._dhsSpreadDays(40, 31, 0).length, 31, 'count>=days -> every day, no overflow');
  assert.notStrictEqual(w._dhsPeriodSeed('07/2026'), w._dhsPeriodSeed('08/2026'), 'consecutive months seed differs');
});

test('_dhsBuildFirstInvoice: only authorized tasks are checked; the rest stay empty', () => {
  const w = loadApp();
  const cols = w._dhsSvcColNames();
  const res = { hours: 2, minutes: 0, tasks: [
    { task: 'Bathing', freq: '7 days per week', perDay: '00:16', perMonth: '08:02' },
    { task: 'Laundry', freq: 'Twice per month', perDay: '01:00', perMonth: '02:00' },
  ]};
  const svc = w._dhsBuildFirstInvoice(res, { clientName: 'X' }, '07/2026').data.tasks.svc;
  const count = (name) => svc.reduce((n, r) => n + (r[cols.indexOf(name)] ? 1 : 0), 0);
  assert.strictEqual(count('Bathing'), 31, 'daily -> every day');
  assert.strictEqual(count('Laundry'), 2, 'twice per month -> 2');
  // Unauthorized columns must be entirely empty.
  ['Dressing', 'Eating', 'Toileting', 'Transferring', 'Medication', 'Hospital/Nursing Facility Stay']
    .forEach(n => assert.strictEqual(count(n), 0, n + ' not authorized -> empty'));
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
