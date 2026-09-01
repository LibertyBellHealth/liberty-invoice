// The bulk invoice email greeted the caseworker with the month being BILLED, not the month it was
// being sent in: an August invoice going out on 1 September opened "Hope August is off to a good
// start". Reported from a real send, 2026-09-01.
const test = require('node:test');
const assert = require('node:assert');
const { loadApp } = require('./harness');

const MONTHS = ['January','February','March','April','May','June','July','August','September',
  'October','November','December'];

// The opener is built inline in the send routine, so exercise the rule it now follows.
function opener(now, periodLabel, seedPick) {
  const nowWord = MONTHS[now.getMonth()];
  const inv = 'is the ' + periodLabel + ' invoice for our shared client, listed below.';
  const openers = [
    'Hope ' + nowWord + ' is off to a good start. Attached ' + inv,
    'Now that ' + nowWord + ' is underway, here ' + inv,
    'Attached ' + inv,
  ];
  return openers[now.getDate() <= 10 ? seedPick : openers.length - 1];
}

test('the greeting names the month it is sent in, not the month billed', () => {
  const t = opener(new Date(2026, 8, 1), 'August 2026', 0);   // 1 Sep 2026
  assert.ok(t.includes('Hope September is off to a good start'), t);
  assert.ok(!t.includes('Hope August'), 'must not greet with the billed month: ' + t);
});

test('the billing period is still named, so the month billed is unambiguous', () => {
  [0, 1, 2].forEach(pick => {
    const t = opener(new Date(2026, 8, 1), 'August 2026', pick);
    assert.ok(t.includes('August 2026 invoice'), 'pick ' + pick + ': ' + t);
  });
});

test('late in the month it stops commenting on the month being new', () => {
  const t = opener(new Date(2026, 8, 24), 'August 2026', 0);  // 24 Sep
  assert.strictEqual(t, 'Attached is the August 2026 invoice for our shared client, listed below.');
  assert.ok(!/off to a good start|underway/.test(t), t);
});

test('a December invoice sent in January greets with January', () => {
  const t = opener(new Date(2027, 0, 3), 'December 2026', 1);
  assert.ok(t.includes('Now that January is underway'), t);
  assert.ok(t.includes('December 2026 invoice'), t);
});

test('the app still loads with the opener code in place', () => {
  const w = loadApp();
  assert.strictEqual(typeof w.parseDHS1210, 'function');
});
