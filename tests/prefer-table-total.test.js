// A real packet states the approved total twice and they disagree: the DHS-1210-A cover letter says
// "62 Hours and 20 Minutes", the MDHHS-6064 task table says "Total per month 62:21", and the 13
// task rows sum to 62:21 exactly. The reader took the letter, so every invoice for that client
// billed a minute LESS than authorized. Owner ruled 2026-09-01: the task table wins.
const test = require('node:test');
const assert = require('node:assert');
const { loadApp } = require('./harness');
const w = loadApp();

const packet = (letter, table) => [[
  'DHS-1210-A, SERVICES AND PAYMENT APPROVAL NOTICE',
  ...(letter ? ['are approved for Home Help Services effective 03/01/2026. Services have been approved for ' + letter + ' per month.'] : []),
  'MDHHS-6064-P, PROVIDER TIME AND TASK MANAGEMENT',
  'Bathing 00:16 7 days per week 08:02 $216.72',
  ...(table ? ['Total per month ' + table + ' $ 1,683.45'] : []),
]];

test('when the two totals disagree, the task table wins', () => {
  const r = w.parseDHS1210(packet('62 Hours and 20 Minutes', '62:21'));
  assert.strictEqual(r.hours, 62);
  assert.strictEqual(r.minutes, 21);
});

test('the disagreement is reported, naming both figures', () => {
  const r = w.parseDHS1210(packet('62 Hours and 20 Minutes', '62:21'));
  assert.strictEqual(r.totalsDisagree, true);
  const warn = [...r.warnings].find(x => /two different monthly totals/.test(x));
  assert.ok(warn, 'expected a disagreement warning, got: ' + JSON.stringify([...r.warnings]));
  assert.ok(warn.includes('62:21') && warn.includes('62:20'), 'both figures must be named: ' + warn);
});

test('when they agree, nothing is flagged', () => {
  const r = w.parseDHS1210(packet('29 Hours and 47 Minutes', '29:47'));
  assert.strictEqual(r.hours, 29);
  assert.strictEqual(r.minutes, 47);
  assert.strictEqual(r.totalsDisagree, undefined);
  assert.ok(![...r.warnings].some(x => /two different monthly totals/.test(x)));
});

test('a standalone 6064 (table only) still reads its total', () => {
  const r = w.parseDHS1210(packet(null, '73:44'));
  assert.strictEqual(r.hours, 73);
  assert.strictEqual(r.minutes, 44);
  assert.strictEqual(r.totalsDisagree, undefined);
});

test('a cover letter with no task table still falls back to the letter', () => {
  const r = w.parseDHS1210(packet('41 Hours and 16 Minutes', null));
  assert.strictEqual(r.hours, 41);
  assert.strictEqual(r.minutes, 16);
});

test('minutes under ten keep two digits in the warning', () => {
  const warn = [...w.parseDHS1210(packet('62 Hours and 5 Minutes', '62:06')).warnings]
    .find(x => /two different monthly totals/.test(x));
  assert.ok(warn.includes('62:06') && warn.includes('62:05'), warn);
});
