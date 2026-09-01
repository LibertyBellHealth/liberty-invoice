// Every task with the same day-count landed on identical days, because the spread was seeded by the
// billing period alone. One client had Laundry, Shopping and Travel all starting the same day;
// another had Laundry, Meal Prep, Shopping and Travel stacked on one date. Owner: "I need it to
// look more human done and not generated" (2026-09-01). Days are now anchored to weekdays, one
// pattern per task group — travel time still rides along with the task it serves.
const test = require('node:test');
const assert = require('node:assert');
const { loadApp } = require('./harness');
const w = loadApp();

const COLS = w._dhsSvcColNames();
const col = n => COLS.indexOf(n);
const daysFor = (grid, name) => grid.map((r, i) => r[col(name)] ? i + 1 : null).filter(Boolean);

const TASKS = [
  { task: 'Bathing', perDay: '00:16', freq: '7 days per week', perMonth: '08:02' },
  { task: 'Laundry', perDay: '00:24', freq: '2 days per week', perMonth: '03:26' },
  { task: 'Shopping for Food/Meds', perDay: '00:17', freq: '2 days per week', perMonth: '02:26' },
  { task: 'Travel For Shopping', perDay: '00:30', freq: '2 days per week', perMonth: '04:18' },
];
const build = (period, prof) => w._dhsBuildFirstInvoice(
  { hours: 62, minutes: 21, tasks: TASKS }, prof || { clientName: 'T' }, period).data.tasks.svc;

test('a task and its travel time fall on exactly the same days', () => {
  const g = build('08/2026');
  assert.deepStrictEqual(daysFor(g, 'Shopping'), daysFor(g, 'Travel Time for Shopping'));
});

test('two unrelated tasks with the same day-count do NOT share a day pattern', () => {
  const g = build('08/2026');
  const laundry = daysFor(g, 'Laundry'), shopping = daysFor(g, 'Shopping');
  assert.strictEqual(laundry.length, shopping.length, 'same count expected for this fixture');
  assert.notDeepStrictEqual(laundry, shopping, 'laundry must not sit on the shopping days');
});

test('each task still gets exactly its authorized number of days', () => {
  const g = build('08/2026');
  assert.strictEqual(daysFor(g, 'Laundry').length, 8);    // 03:26 / 00:24
  assert.strictEqual(daysFor(g, 'Shopping').length, 8);   // 02:26 / 00:17
  assert.strictEqual(daysFor(g, 'Bathing').length, 31);   // daily
});

test('a four-a-month task lands on one weekday', () => {
  // September 2026 starts on a Tuesday.
  const g = w._dhsBuildFirstInvoice({ hours: 4, minutes: 0, tasks: [
    { task: 'Laundry', perDay: '01:00', freq: '1 day per week', perMonth: '04:00' }] },
    { clientName: 'T' }, '09/2026').data.tasks.svc;
  const days = daysFor(g, 'Laundry');
  assert.strictEqual(days.length, 4);
  const weekdays = new Set(days.map(d => (2 + d - 1) % 7));  // 1 Sep 2026 = Tuesday (2)
  assert.strictEqual(weekdays.size, 1, 'expected one weekday, got days ' + days);
});

test('the days still never fall before the service start date', () => {
  const g = build('08/2026', { clientName: 'T', startDate: '2026-08-20' });
  const all = [].concat(...COLS.map((_, c) => g.map((r, i) => r[c] ? i + 1 : null).filter(Boolean)));
  assert.ok(Math.min(...all) >= 20, 'earliest marked day was ' + Math.min(...all));
});

test('the pattern shifts month to month rather than repeating exactly', () => {
  assert.notDeepStrictEqual(daysFor(build('08/2026'), 'Laundry'),
                            daysFor(build('09/2026'), 'Laundry'));
});
