// Azure Document Intelligence returns a table as one CELL per line in reading order, so a task row
// arrives as 4-5 consecutive lines. The single-line row pattern matched none of them, so EVERY
// scanned or photographed authorization imported with an empty task table — no caregiver task
// sheet, no invoice day grid. Lines below are the real shape, captured from the OCR endpoint on
// 2026-09-01 using a synthetic form (no client data).
const test = require('node:test');
const assert = require('node:assert');
const { loadApp } = require('./harness');
const w = loadApp();

const OCR_CELLS = [[
  'MDHHS-6064-P, PROVIDER TIME AND TASK MANAGEMENT',
  'SECTION 2 - ADULT SERVICES INFORMATION',
  'ASW Email Address', 'ExampleT9@michigan.gov', 'Date', '08/04/2026',
  'SECTION 3 - TASKS',
  'Provider Name', 'Liberty Home Care Assistance',
  'Provider Pay Rate', '$ 27.00',
  'Authorized Tasks', 'Time / Day', 'Number of Days', 'Time /Month', 'Amount',
  'Bathing', '00:05', '7 days per week', '02:30', '$67.72',
  'Laundry', '01:38', '1 day per week', '07:01', '$189.63',
  'Meal Preparation', '00:50', '7 days per week', '25:05', '$677.25',
  'Shopping for Food/Meds', '00:35', '2 days per week', '05:01', '$135.45',
  'Total per month', '39:41', '$ 1,070.05',
]];

test('task rows are rebuilt from cell-per-line OCR output', () => {
  const r = w.parseDHS1210(OCR_CELLS);
  assert.strictEqual(r.tasks.length, 4);
  const first = { ...r.tasks[0] };
  assert.strictEqual(first.task, 'Bathing');
  assert.strictEqual(first.perDay, '00:05');
  assert.strictEqual(first.freq, '7 days per week');
  assert.strictEqual(first.perMonth, '02:30');
  assert.strictEqual(first.amount, 67.72);
  // A task name containing a slash must survive.
  assert.strictEqual([...r.tasks].map(t => t.task).join(','),
    'Bathing,Laundry,Meal Preparation,Shopping for Food/Meds');
});

test('column headers and section titles are not mistaken for tasks', () => {
  const names = [...w.parseDHS1210(OCR_CELLS).tasks].map(t => t.task);
  ['Authorized Tasks', 'Time / Day', 'Number of Days', 'Provider Name', 'Total per month']
    .forEach(h => assert.ok(!names.includes(h), h + ' was parsed as a task'));
});

test('provider name and pay rate are read from their own cells', () => {
  const r = w.parseDHS1210(OCR_CELLS);
  assert.strictEqual(r.rate, 27);
  assert.strictEqual(r.providerName, 'Liberty Home Care Assistance');
});

test('the rebuilt rows are still reconciled against the form\'s printed totals', () => {
  const r = w.parseDHS1210(OCR_CELLS);
  assert.strictEqual(r.taskMinuteSum, 2377);      // 2:30 + 7:01 + 25:05 + 5:01
  assert.strictEqual(r.approvedTotalMin, 2381);   // 39:41 printed on the form
  // 4 minutes short across 4 rows — the form's own per-row rounding, so it reconciles.
  assert.strictEqual(r.timeReconciles, true);
  assert.strictEqual(r.amountReconciles, true);
  assert.ok(![...r.warnings].some(x => /task table/.test(x)));
});

test('a normal single-line table is untouched by the fallback (no duplicate rows)', () => {
  const r = w.parseDHS1210([[
    'MDHHS-6064-P, PROVIDER TIME AND TASK MANAGEMENT',
    'Bathing 00:05 7 days per week 02:30 $67.72',
    'Laundry 01:38 1 day per week 07:01 $189.63',
    'Total per month 09:31 $ 257.35',
  ]]);
  assert.strictEqual(r.tasks.length, 2);
});
