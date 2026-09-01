// The Number of Days column was only read when it said "per week" / "per month". A form written
// "Twice a week" or "once a month" failed the row pattern and the task was skipped — the same
// silent loss as the "Twice per month" row, just a different preposition. Both wordings appear on
// real paperwork. (The task is no longer lost silently either way: an unrecognised frequency is
// now reported by name.)
const test = require('node:test');
const assert = require('node:assert');
const { loadApp } = require('./harness');
const w = loadApp();

const form = freq => [[
  'MDHHS-6064-P, PROVIDER TIME AND TASK MANAGEMENT',
  'Laundry 01:00 ' + freq + ' 02:00 $54.00',
  'Total per month 02:00 $ 54.00',
]];

test('"a week" and "a month" wordings parse as task rows', () => {
  ['Twice a week', 'once a month', '2 days a week', '1 time a week', '3 times a month']
    .forEach(f => {
      const r = w.parseDHS1210(form(f));
      assert.strictEqual(r.tasks.length, 1, f + ' was not read as a row');
      assert.strictEqual({ ...r.tasks[0] }.freq, f);
    });
});

test('they resolve to the same schedule as their "per" equivalents', () => {
  const pairs = [['Twice a week', 'Twice per week'], ['once a month', 'Once per month'],
                 ['2 days a week', '2 days per week'], ['3 times a month', '3 times per month']];
  pairs.forEach(([a, b]) => {
    const A = w._dhsFreqSpec(a), B = w._dhsFreqSpec(b);
    assert.ok(A && B, a + ' / ' + b);
    assert.strictEqual(A.per, B.per, a + ' vs ' + b);
    assert.strictEqual(A.n, B.n, a + ' vs ' + b);
  });
});

test('"per" wordings are unaffected', () => {
  const r = w.parseDHS1210(form('2 days per week'));
  assert.strictEqual(r.tasks.length, 1);
  assert.strictEqual(w._dhsFreqSpec('2 days per week').n, 2);
});

test('a wording we still cannot read is reported, not dropped in silence', () => {
  const r = w.parseDHS1210(form('every other Tuesday'));
  assert.strictEqual(r.tasks.length, 0);
  assert.ok([...r.warnings].some(x => /unrecognised frequency/.test(x) && /every other Tuesday/.test(x)),
    JSON.stringify([...r.warnings]));
});
