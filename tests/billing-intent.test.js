'use strict';
// Two owner decisions (2026-08-21), both previously left to silent defaults:
//   1. a mid-month start is billed full-month or prorated DEPENDING ON THE CASE, so the app asks;
//   2. an "In Progress" client (stored 'inactive') is never invoiced, on every surface.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

test('_proratedFirstMonth: bills only the days actually served', () => {
  const w = loadApp();
  // 30-day month, service starts on the 16th → 15 of 30 days → half of 20:00
  const half = w._proratedFirstMonth({ hours: 20, minutes: 0 }, '09/16/2026');
  assert.strictEqual(half.daysInMonth, 30);
  assert.strictEqual(half.daysServed, 15);
  assert.strictEqual(half.hours, 10);
  assert.strictEqual(half.minutes, 0);

  // 31-day month, starts on the 21st → 11 of 31 days of 29:47 (1787 min) → 634 min = 10:34
  const p = w._proratedFirstMonth({ hours: 29, minutes: 47 }, '07/21/2026');
  assert.strictEqual(p.daysServed, 11);
  assert.strictEqual(p.hours * 60 + p.minutes, Math.round(1787 * 11 / 31));
});

test('_proratedFirstMonth: a 1st-of-month start is not prorated', () => {
  const w = loadApp();
  assert.strictEqual(w._proratedFirstMonth({ hours: 20, minutes: 0 }, '09/01/2026'), null,
    'nothing to prorate — the app must not ask a pointless question');
  assert.strictEqual(w._proratedFirstMonth({ hours: 20, minutes: 0 }, ''), null, 'no date → no prompt');
});

test('an In Progress client is excluded from every invoicing surface', () => {
  const w = loadApp(); resetStorage(w);
  const base = { startDate: '2020-01-01', invoices: [] };
  assert.strictEqual(w.isInvoiceableStatus({ clientStatus: 'active' }), true);
  assert.strictEqual(w.isInvoiceableStatus({}), true, 'no status set means active');
  ['inactive', 'lost', 'terminated'].forEach(st => {
    assert.strictEqual(w.isInvoiceableStatus({ clientStatus: st }), false, st + ' is not invoiceable');
    assert.strictEqual(w.clientDueForInvoice(Object.assign({ clientStatus: st }, base), '07/2026'), false,
      st + ' must not be flagged as due for an invoice');
  });
  assert.strictEqual(w.clientDueForInvoice(Object.assign({ clientStatus: 'active' }, base), '07/2026'), true,
    'an active client with a start date IS due');
});

test('auto-gen eligibility agrees with the missing-invoice predicate', () => {
  const w = loadApp(); resetStorage(w);
  w.saveProfilesLS({
    'Active One':  { clientName: 'Active One', clientStatus: 'active', startDate: '2020-01-01',
                     authorization: { hours: 20, minutes: 0, tasks: [] }, invoices: [] },
    'In Progress': { clientName: 'In Progress', clientStatus: 'inactive', startDate: '2020-01-01',
                     authorization: { hours: 20, minutes: 0, tasks: [] }, invoices: [] },
  });
  const eligible = w.findClientsEligibleForAutoGen('07/2026').map(e => e.name);
  assert.strictEqual(JSON.stringify(eligible), JSON.stringify(['Active One']),
    'the two surfaces used separate inline status tests and had drifted apart');
});
