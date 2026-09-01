// The caregiver works from the task sheet, so its per-task times are padded to the next half hour
// with at least 15 minutes of headroom — the agency should never deliver less time than MDHHS
// authorized (owner's rule, 2026-09-01). Invoices are NOT padded: they bill the exact approved
// total, because billing over the authorization is what triggers a recoupment.
const test = require('node:test');
const assert = require('node:assert');
const { loadApp } = require('./harness');
const w = loadApp();

test('target lands on the next half hour, at least 15 minutes over', () => {
  assert.strictEqual(w._taskSheetTargetMin(80 * 60 + 36), 81 * 60);      // 80:36 -> 81:00 (+24)
  assert.strictEqual(w._taskSheetTargetMin(80 * 60 + 5), 80 * 60 + 30);  // 80:05 -> 80:30 (+25)
  assert.strictEqual(w._taskSheetTargetMin(80 * 60 + 55), 81 * 60 + 30); // 80:55 -> 81:30 (+35)
  assert.strictEqual(w._taskSheetTargetMin(80 * 60), 80 * 60 + 30);      // exact hour still gets headroom
  assert.strictEqual(w._taskSheetTargetMin(0), 0);
});

test('padded rows add up to the target exactly', () => {
  const tasks = [
    { task: 'Bathing', perDay: '00:05', freq: '7 days per week', perMonth: '02:30' },
    { task: 'Meal Preparation', perDay: '00:50', freq: '7 days per week', perMonth: '25:05' },
    { task: 'Laundry', perDay: '01:38', freq: '1 day per week', perMonth: '07:01' },
  ];
  const authMin = 34 * 60 + 36;                       // 34:36 -> target 35:00
  const out = w._taskSheetPaddedTasks(tasks, authMin);
  const sum = [...out].reduce((a, t) => a + w._dhsHmToMin(t.perMonth), 0);
  assert.strictEqual(sum, 35 * 60);
  // Every row got no less than it was authorized.
  const before = tasks.map(t => w._dhsHmToMin(t.perMonth));
  [...out].forEach((t, i) => assert.ok(w._dhsHmToMin(t.perMonth) >= before[i]));
});

test('task name, time/day and frequency are never altered', () => {
  const tasks = [{ task: 'Bathing', perDay: '00:05', freq: '7 days per week', perMonth: '02:30' }];
  const out = w._taskSheetPaddedTasks(tasks, 3 * 60);
  assert.strictEqual(out[0].task, 'Bathing');
  assert.strictEqual(out[0].perDay, '00:05');
  assert.strictEqual(out[0].freq, '7 days per week');
});

test('the caller\'s authorization objects are not mutated', () => {
  const tasks = [{ task: 'Bathing', perDay: '00:05', freq: '7 days per week', perMonth: '02:30' }];
  w._taskSheetPaddedTasks(tasks, 3 * 60);
  assert.strictEqual(tasks[0].perMonth, '02:30');
});

test('no authorized total, or rows already at target: left alone', () => {
  const tasks = [{ task: 'Bathing', perDay: '00:05', freq: '7 days per week', perMonth: '02:30' }];
  assert.strictEqual(w._taskSheetPaddedTasks(tasks, 0)[0].perMonth, '02:30');
  assert.strictEqual(w._taskSheetPaddedTasks(tasks, 60)[0].perMonth, '02:30'); // rows already exceed
});

test('the INVOICE still bills the exact authorized total, unpadded', () => {
  const prof = { clientName: 'Test', startDate: '' };
  const built = w._dhsBuildFirstInvoice(
    { hours: 80, minutes: 36, tasks: [{ task: 'Bathing', perDay: '00:05', freq: '7 days per week', perMonth: '02:30' }] },
    prof, '08/2026');
  assert.strictEqual(built.data.svcHH, '80');
  assert.strictEqual(built.data.svcMM, '36');
  assert.strictEqual(built.data.grandHH, '80');
  assert.strictEqual(built.data.grandMM, '36');
});
