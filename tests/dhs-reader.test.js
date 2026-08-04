'use strict';
// Smoke tests for the DHS-1210 reader (parseDHS1210). Feeds pre-grouped page lines (the same
// shape _dhsPageLines produces from pdf.js) so no PDF/pdf.js is needed — pure parsing logic.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp } = require('./harness');

// A realistic single-page extraction of the MDHHS-6064-P task table + approval header.
// The 8 task amounts sum to $804.47 and the per-month times sum to 29:47 — matching the form,
// so the parser's self-checks should reconcile.
const REAL_PAGE = [[
  'are approved for Home Help Services effective 08/01/2026. Services have been approved for 29 Hours',
  'and 47 Minutes per month.',
  'County 82-WAYNE',
  'Adult Services Worker (ASW) Name',
  'R Feto 313-505-1660',
  'FetoR@michigan.gov',
  'Provider Name Provider Pay Rate',
  'Liberty Home Care Assistance $ 27.00',
  'Bathing 00:18 7 days per week 09:02 $243.81',
  'Transferring 00:08 7 days per week 04:01 $108.36',
  'Housework 00:12 7 days per week 06:01 $162.54',
  'Laundry 00:45 1 day per week 03:13 $87.08',
  'Medication 00:04 7 days per week 02:00 $54.18',
  'Meal Preparation 01:00 1 day per week 04:18 $116.10',
  'Shopping for Food/Meds 01:00 Once per month 01:00 $27.00',
  'Travel For Shopping 00:12 Once per month 00:12 $5.40',
  'Total per month 29:47 $ 804.47',
]];

test('parseDHS1210: extracts the headline fields', () => {
  const w = loadApp();
  const r = w.parseDHS1210(REAL_PAGE);
  assert.strictEqual(r.hours, 29, 'approved hours');
  assert.strictEqual(r.minutes, 47, 'approved minutes');
  assert.strictEqual(r.effectiveDate, '08/01/2026', 'effective date');
  assert.strictEqual(r.reassessDate, '02/01/2027', 'reassessment = effective + 6 months');
  assert.strictEqual(r.rate, 27, 'provider pay rate');
  assert.strictEqual(r.printedTotal, 804.47, 'printed monthly total');
  assert.strictEqual(r.aswEmail, 'FetoR@michigan.gov', 'ASW email');
  assert.strictEqual(r.aswName, 'R Feto', 'ASW name');
  assert.strictEqual(r.county, '82-WAYNE', 'county');
});

test('parseDHS1210: extracts all task rows with their columns', () => {
  const w = loadApp();
  const r = w.parseDHS1210(REAL_PAGE);
  assert.strictEqual(r.tasks.length, 8, 'eight tasks');
  const bathing = r.tasks[0];
  assert.strictEqual(bathing.task, 'Bathing');
  assert.strictEqual(bathing.perDay, '00:18', 'Time/Day');
  assert.strictEqual(bathing.freq, '7 days per week', 'Number of Days');
  assert.strictEqual(bathing.perMonth, '09:02', 'Time/Month');
  assert.strictEqual(bathing.amount, 243.81, 'Amount');
  // task name containing a slash must not break the row parse
  assert.ok(r.tasks.some(t => t.task === 'Shopping for Food/Meds'), 'slash in task name parses');
});

test('parseDHS1210: self-checks reconcile against the form totals', () => {
  const w = loadApp();
  const r = w.parseDHS1210(REAL_PAGE);
  assert.strictEqual(r.taskAmountSum, 804.47, 'task $ sum equals the printed total');
  assert.strictEqual(r.amountReconciles, true, 'amounts reconcile');
  assert.strictEqual(r.timeReconciles, true, 'task minutes reconcile with approved hours');
  assert.strictEqual(r.warnings.length, 0, 'no warnings on a clean form');
});

test('parseDHS1210: flags a mismatch instead of silently trusting a bad parse', () => {
  const w = loadApp();
  // Same tasks but a wrong printed total — amounts must NOT reconcile.
  const bad = [REAL_PAGE[0].map(l => l.startsWith('Total per month') ? 'Total per month 29:47 $ 999.99' : l)];
  const r = w.parseDHS1210(bad);
  assert.strictEqual(r.amountReconciles, false, 'a wrong total is caught, not accepted');
});

test('parseDHS1210: missing approved-hours line produces a warning, not a crash', () => {
  const w = loadApp();
  const r = w.parseDHS1210([['Bathing 00:18 7 days per week 09:02 $243.81']]);
  assert.ok(r.warnings.includes('approved hours'), 'warns when approved hours cannot be read');
  assert.ok(Array.isArray(r.tasks), 'still returns tasks array');
});
