// The MDHHS-6064 rounds each task row to the whole minute, so its rows add up ~2 minutes below the
// total the same form prints. That tripped "Task times DO NOT match the approved hours" on every
// real 6064. The shortfall is the form's own arithmetic and never reaches an invoice — invoices
// bill the approved total. A shortfall too large for rounding, or ANY overrun, must still fail.
const test = require('node:test');
const assert = require('node:assert');
const { loadApp } = require('./harness');
const w = loadApp();

// n task rows of one minute each, against a stated monthly total.
function form(rows, totalHH, totalMM) {
  const lines = ['MDHHS-6064-P, PROVIDER TIME AND TASK MANAGEMENT'];
  for (let i = 0; i < rows; i++) lines.push('Task' + i + ' 00:01 7 days per week 00:01 $27.09');
  lines.push('Total per month ' + totalHH + ':' + totalMM + ' $ 100.00');
  return w.parseDHS1210([lines]);
}

test('10 rows summing 2 minutes under the printed total still reconciles (form rounding)', () => {
  const r = form(10, '00', '12');       // rows sum 10, form says 12
  assert.strictEqual(r.taskMinuteSum, 10);
  assert.strictEqual(r.approvedTotalMin, 12);
  assert.strictEqual(r.timeReconciles, true);
});

test('a shortfall bigger than the row count is still a misread', () => {
  const r = form(3, '00', '30');        // 3 rows can only explain 3 minutes, not 27
  assert.strictEqual(r.timeReconciles, false);
});

test('task rows that OVERRUN the authorization still fail — the recoupment direction', () => {
  const r = form(10, '00', '08');       // rows sum 10 against 8 approved
  assert.strictEqual(r.timeReconciles, false);
});

test('an exact match reconciles', () => {
  assert.strictEqual(form(10, '00', '10').timeReconciles, true);
});

test('prorating a partial first month bills the exact share, never rounded up', () => {
  // 41:16 = 2476 min over a 31-day month, served from the 12th = 20 days.
  // 2476 * 20/31 = 1597.42 -> 1597 exact-nearest, NOT 1598.
  const p = w._proratedFirstMonth({ hours: 41, minutes: 16 }, '03/12/2026');
  assert.strictEqual(p.daysServed, 20);
  assert.strictEqual(p.hours * 60 + p.minutes, 1597);
});
