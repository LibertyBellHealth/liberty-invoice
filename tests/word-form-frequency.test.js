// A real DHS-1210 row reading "Laundry 01:00 Twice per month 02:00 $54.00" was silently DROPPED:
// the row pattern accepted only "N days per week/month" and "Once per week/month", while
// _dhsFreqToDays downstream had understood "Twice per month" all along. The task vanished from the
// authorization, the caregiver task sheet and the invoice day grid. Reported 2026-09-01 — the
// owner saw both reconciliation chips fail by exactly the missing row (2:00 and $54.00).
const test = require('node:test');
const assert = require('node:assert');
const { loadApp } = require('./harness');
const w = loadApp();

const FORM = extra => [[
  'DHS-1210-A, SERVICES AND PAYMENT APPROVAL NOTICE',
  'Bathing 00:16 7 days per week 08:02 $216.72',
  ...extra,
  'Total per month 10:02 $ 270.72',
]];

test('"Twice per month" is read as a task row', () => {
  const r = w.parseDHS1210(FORM(['Laundry 01:00 Twice per month 02:00 $54.00']));
  assert.strictEqual(r.tasks.length, 2);
  const laundry = { ...r.tasks[1] };
  assert.strictEqual(laundry.task, 'Laundry');
  assert.strictEqual(laundry.freq, 'Twice per month');
  assert.strictEqual(laundry.perMonth, '02:00');
  assert.strictEqual(laundry.amount, 54);
});

test('the other spelled-out frequencies _dhsFreqToDays supports also parse', () => {
  ['Once per month', 'Three times per month', 'Four times per month', 'Twice per week',
   'Once per week', '2 times per month', '3 days per week', 'Daily', 'Monthly']
    .forEach(freq => {
      const r = w.parseDHS1210(FORM(['Laundry 01:00 ' + freq + ' 02:00 $54.00']));
      assert.strictEqual(r.tasks.length, 2, freq + ' was not parsed as a row');
      assert.strictEqual({ ...r.tasks[1] }.freq, freq);
    });
});

test('a frequency wording we do not know is REPORTED, not silently dropped', () => {
  const r = w.parseDHS1210(FORM(['Laundry 01:00 Every other Tuesday 02:00 $54.00']));
  assert.strictEqual(r.tasks.length, 1);
  assert.ok([...r.warnings].some(x => /unrecognised frequency/.test(x) && /Every other Tuesday/.test(x)),
    'expected the unknown wording to be named, got: ' + JSON.stringify([...r.warnings]));
});

test('a form with no odd wording carries no frequency warning', () => {
  const r = w.parseDHS1210(FORM(['Laundry 01:00 Twice per month 02:00 $54.00']));
  assert.ok(![...r.warnings].some(x => /unrecognised frequency/.test(x)));
});

test('the OCR cell pass accepts the same wordings', () => {
  const r = w.parseDHS1210([[
    'MDHHS-6064-P, PROVIDER TIME AND TASK MANAGEMENT',
    'Laundry', '01:00', 'Twice per month', '02:00', '$54.00',
    'Total per month', '02:00', '$ 54.00',
  ]]);
  assert.strictEqual(r.tasks.length, 1);
  assert.strictEqual({ ...r.tasks[0] }.freq, 'Twice per month');
});
