// _dhsFreqToDays parsed the Number of Days column itself and only understood the numeric
// "N days per week" form, so "Once per week", "Twice per week" and "Weekly" fell through to a
// single day a month. Same shape as the bug that dropped a "Twice per month" row outright: two
// readers of one field, one of them quietly narrower. It now delegates to _dhsFreqSpec.
const test = require('node:test');
const assert = require('node:assert');
const { loadApp } = require('./harness');
const w = loadApp();
const n = f => w._dhsFreqToDays(f, 31).length;

test('spelled-out weekly frequencies match their numeric equivalents', () => {
  assert.strictEqual(n('Once per week'), n('1 day per week'));
  assert.strictEqual(n('Twice per week'), n('2 days per week'));
  assert.strictEqual(n('Three times per week'), n('3 days per week'));
  assert.strictEqual(n('Weekly'), n('1 day per week'));
});

test('the numeric forms are unchanged', () => {
  assert.strictEqual(n('7 days per week'), 31);
  assert.strictEqual(n('1 day per week'), 5);
  assert.strictEqual(n('2 days per week'), 10);
  assert.strictEqual(n('3 days per week'), 15);
});

test('monthly frequencies are unchanged', () => {
  assert.strictEqual(n('Once per month'), 1);
  assert.strictEqual(n('Twice per month'), 2);
  assert.strictEqual(n('Three times per month'), 3);
  assert.strictEqual(n('2 days per month'), 2);
  assert.strictEqual(n('Monthly'), 1);
});

test('an unknown wording still marks one day rather than none', () => {
  assert.deepStrictEqual([...w._dhsFreqToDays('every other Tuesday', 31)], [0]);
  assert.deepStrictEqual([...w._dhsFreqToDays('', 31)], [0]);
});

test('the two readers agree on every wording the row pattern accepts', () => {
  ['7 days per week', '1 day per week', '2 days per week', 'Once per week', 'Twice per week',
   'Once per month', 'Twice per month', 'Four times per month', 'Daily', 'Weekly', 'Monthly']
    .forEach(f => {
      assert.ok(w._dhsFreqSpec(f), f + ' unreadable by _dhsFreqSpec');
      assert.ok(n(f) >= 1, f + ' produced no days');
    });
});
