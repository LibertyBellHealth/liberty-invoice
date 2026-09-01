// A scanned/photographed authorization goes through OCR, which returns the task table as loose
// cells, so the row pattern matches nothing. The import then showed the approved hours, no task
// list, and a red "Task times DO NOT match the approved hours" — blaming a mismatch for a total
// miss. Reported 2026-09-01 from a real MDHHS-6064 image import.
const test = require('node:test');
const assert = require('node:assert');
const { loadApp } = require('./harness');
const w = loadApp();

// What OCR gives back: the header fields survive, the table rows do not.
const OCR_NO_TABLE = [[
  'MDHHS-6064-P, PROVIDER TIME AND TASK MANAGEMENT',
  'SECTION 2 - ADULT SERVICES INFORMATION',
  'S Bright 313-618-3273',
  'BrightS@michigan.gov 08/04/2026',
  'Total per month 73:44',
  'Bathing', '00:05', '7 days per week', '02:30', '$67.72',   // cells, not a row
]];

test('an unread task table is reported as its own miss', () => {
  const r = w.parseDHS1210(OCR_NO_TABLE);
  assert.strictEqual(r.tasks.length, 0);
  assert.ok([...r.warnings].some(x => /task table/.test(x)),
    'expected a task-table warning, got: ' + JSON.stringify([...r.warnings]));
});

test('no task rows means no reconciliation verdict — not a failing one', () => {
  const r = w.parseDHS1210(OCR_NO_TABLE);
  assert.strictEqual(r.hours, 73);        // hours still read
  assert.strictEqual(r.timeReconciles, undefined,
    'an empty task list must not render a red "times do not match"');
});

test('a form whose rows DO parse still reconciles as before', () => {
  const r = w.parseDHS1210([[
    'MDHHS-6064-P, PROVIDER TIME AND TASK MANAGEMENT',
    'Bathing 00:05 7 days per week 02:30 $67.72',
    'Medication 00:02 7 days per week 01:00 $27.09',
    'Total per month 03:30 $ 94.81',
  ]]);
  assert.strictEqual(r.tasks.length, 2);
  assert.strictEqual(r.timeReconciles, true);
  assert.ok(![...r.warnings].some(x => /task table/.test(x)));
});
