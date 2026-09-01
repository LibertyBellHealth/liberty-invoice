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

test('the day count comes from the Number of Days column, not the times', () => {
  const g = build('08/2026');
  // "2 days per week" = every occurrence of two weekdays in a 31-day August = 9 visits.
  // Dividing the times (03:26 / 00:24) would have given 8 — MDHHS fills this grid by days.
  assert.strictEqual(daysFor(g, 'Laundry').length, 9);
  assert.strictEqual(daysFor(g, 'Shopping').length, 9);
  assert.strictEqual(daysFor(g, 'Bathing').length, 31);   // 7 days per week = every day
});

test('a "once per month" task gets exactly one day', () => {
  const g = w._dhsBuildFirstInvoice({ hours: 1, minutes: 0, tasks: [
    { task: 'Shopping for Food/Meds', perDay: '01:00', freq: 'Once per month', perMonth: '01:00' }] },
    { clientName: 'T' }, '08/2026').data.tasks.svc;
  assert.strictEqual(daysFor(g, 'Shopping').length, 1);
});

test('an unreadable frequency falls back to the authorized time, floored', () => {
  const g = w._dhsBuildFirstInvoice({ hours: 3, minutes: 26, tasks: [
    { task: 'Laundry', perDay: '00:24', freq: 'every other Tuesday', perMonth: '03:26' }] },
    { clientName: 'T' }, '08/2026').data.tasks.svc;
  assert.strictEqual(daysFor(g, 'Laundry').length, 8);   // floor(206/24)
});

test('a once-a-week task lands on one weekday, every occurrence of it', () => {
  // September 2026 starts on a Tuesday; a given weekday occurs 4 or 5 times in the month.
  const g = w._dhsBuildFirstInvoice({ hours: 4, minutes: 0, tasks: [
    { task: 'Laundry', perDay: '01:00', freq: '1 day per week', perMonth: '04:00' }] },
    { clientName: 'T' }, '09/2026').data.tasks.svc;
  const days = daysFor(g, 'Laundry');
  const weekdays = new Set(days.map(d => (2 + d - 1) % 7));  // 1 Sep 2026 = Tuesday (2)
  assert.strictEqual(weekdays.size, 1, 'expected one weekday, got days ' + days);
  assert.ok(days.length === 4 || days.length === 5, 'got ' + days.length + ' days: ' + days);
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
