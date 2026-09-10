'use strict';
// Caregivers had no way to tell how long to clock in for. The task rows each carry their own
// time and frequency, so nothing on the sheet answered it: a 7-days-per-week task and a
// 1-day-per-week task never add up to the same visit twice.
//
// The fix states the month's approved time spread evenly, over 28 days and rounded UP. That is
// deliberately generous: a calendar-month divisor (30) would cost the agency less, because a
// caregiver working the 28-day figure every day of a 30-day month delivers about five hours that
// cannot be billed — MDHHS is invoiced the authorization exactly, so the overrun is payroll.
// The owner was shown that cost and chose it (2026-09-10: "Its okay if they go over but not
// under"). These tests hold that ruling in place: being UNDER is the failure mode.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

// 34:36 authorized -> padded to 35:00 -> ceil(2100/28) = 75 min = "1h 15m" per day.
// A 30-day divisor would give 70 min = "1h 10m" — under, which is what this file exists to catch.
const AUTH = { hours: 34, minutes: 36, effectiveDate: '08/01/2026', tasks: [
  { task: 'Bathing', perDay: '00:05', freq: '7 days per week', perMonth: '02:30' },
  { task: 'Meal Preparation', perDay: '00:50', freq: '7 days per week', perMonth: '25:05' },
  { task: 'Laundry', perDay: '01:38', freq: '1 day per week', perMonth: '07:01' },
] };
const AUTH_MIN = AUTH.hours * 60 + AUTH.minutes;

function app(auth) {
  const w = loadApp();
  resetStorage(w);
  w.saveProfilesLS({ 'Jane Doe': { clientName: 'Jane Doe', firstName: 'Jane', lastName: 'Doe',
    authorization: JSON.parse(JSON.stringify(auth === undefined ? AUTH : auth)) } });
  w.activeProfileName = 'Jane Doe';
  w.showAlert = () => {};
  w.captured = '';
  w._openPhiWindow = () => ({ focus() {}, document: { write(h) { w.captured += h; }, close() {} } });
  return w;
}

test('the sheet gives one per-day figure a caregiver can clock to', () => {
  const w = app();
  w.exportCaregiverTaskSheet();
  // Match the VISIBLE row, not just the figure: the same string also sits in the hidden
  // plain-text block, so a bare indexOf('1h 15m') passes even with the printed row deleted.
  assert.ok(/About per day:<\/span>\s*<span class="v">1h 15m<\/span>/.test(w.captured),
    'no per-day average printed on the sheet — the caregiver still has nothing to clock against');
});

test('the per-day figure is the padded month over 28, rounded up', () => {
  const w = app();
  w.exportCaregiverTaskSheet();
  assert.ok(w.captured.indexOf('1h 10m') === -1,
    'that is 2100/30 — a calendar-month divisor leaves the caregiver under time');
  assert.strictEqual(w._taskSheetPerDayMin(AUTH_MIN), 75);
});

test('the figure is never under, on any length of month', () => {
  const w = app();
  const perDay = w._taskSheetPerDayMin(AUTH_MIN);
  // Four weeks of this must already meet the padded authorization — that is the point of /28.
  assert.ok(perDay * 28 >= w._taskSheetTargetMin(AUTH_MIN),
    'four weeks at the stated daily rate must reach the padded month');
  // And it must never fall short of the plain authorized daily share in a long month.
  assert.ok(perDay > AUTH_MIN / 31,
    'the daily figure must not dip under the authorized share of even a 31-day month');
});

test('rounding goes up, never to nearest', () => {
  const w = app();
  // 20:00 authorized -> padded 20:30 = 1230 min; 1230/28 = 43.9. Math.round would give 43 and
  // leave the caregiver under; ceil gives 44.
  assert.strictEqual(w._taskSheetPerDayMin(20 * 60), 44);
});

test('the emailed and copied text carries it too', () => {
  const w = app();
  w.exportCaregiverTaskSheet();
  assert.ok(/About per day \(average\): 1h 15m/.test(w.captured),
    'the plain-text body is what gets emailed/texted; it must carry the same figure');
});

test('the texted image carries it too', async () => {
  const w = app();
  // Capture the detached markup the image is rasterised from. Falling back to the sheet's HTML
  // would make this pass no matter what the image actually contained.
  let markup = '';
  const realCreate = w.document.createElement.bind(w.document);
  w.document.createElement = (tag) => {
    const el = realCreate(tag);
    if (String(tag).toLowerCase() === 'div') {
      Object.defineProperty(el, 'innerHTML', {
        configurable: true,
        set(v) { markup += String(v); this.__h = v; },
        get() { return this.__h || ''; },
      });
    }
    return el;
  };
  w.html2canvas = () => Promise.reject(new Error('no canvas in jsdom'));
  try { await w.shareCaregiverTaskImage(); } catch (e) { /* rasterising needs a canvas */ }
  w.document.createElement = realCreate;
  assert.ok(markup.length > 0, 'the image markup was never built — this test would prove nothing');
  assert.ok(markup.indexOf('About / day') !== -1, 'the image is what reaches the phone');
  assert.ok(markup.indexOf('1h 15m') !== -1, 'the image must show the same per-day figure');
});

test('an authorization with no approved total shows no per-day line', () => {
  const w = app({ effectiveDate: '08/01/2026', tasks: AUTH.tasks });
  w.exportCaregiverTaskSheet();
  assert.ok(w.captured.indexOf('About per day') === -1,
    'with nothing authorized there is no average to state');
});

// ── Per-visit totals, grouped by how often each task runs ──────────────────────────────────────
// The flat average answers "roughly how much a day" but never "how long is TODAY": a 7-day task
// and a 1-day-per-week task never land on the same visit twice. Owner, 2026-09-10: "7 days a week
// should be this much time per day. 3 days this much time. 1 day this much time."
const MIXED = [
  { task: 'Bathing',      perDay: '00:05', freq: '7 days per week',  perMonth: '02:30' },
  { task: 'Meal Prep',    perDay: '00:50', freq: '7 days per week',  perMonth: '25:05' },
  { task: 'Shopping',     perDay: '00:45', freq: '2 days per week',  perMonth: '06:26' },
  { task: 'Travel',       perDay: '00:30', freq: '2 days per week',  perMonth: '04:18' },
  { task: 'Laundry',      perDay: '01:38', freq: '1 day per week',   perMonth: '07:01' },
  { task: 'Nail Care',    perDay: '00:15', freq: 'Twice per month',  perMonth: '00:30' },
];

test('tasks are grouped by frequency and each group totalled', () => {
  const g = loadApp()._taskSheetFreqGroups(MIXED);
  // JSON.stringify, not deepStrictEqual: jsdom hands back arrays from another realm, which fail
  // reference-equality on their prototype even when every value matches.
  assert.strictEqual(JSON.stringify(g.map((x) => [x.label, x.text])), JSON.stringify([
    ['Every day', '55m'],          // 5 + 50 = 55; sub-hour reads "55m", not "0h 55m"
    ['2 days a week', '1h 15m'],   // 45 + 30 = 75
    ['1 day a week', '1h 40m'],    // 98 -> rounded UP to 100
    ['2 days a month', '15m'],
  ]));
});

test('groups run most-frequent first, so the everyday total is read first', () => {
  const ranks = loadApp()._taskSheetFreqGroups(MIXED).map((x) => x.label);
  assert.strictEqual(ranks[0], 'Every day');
  assert.ok(ranks.indexOf('2 days a week') < ranks.indexOf('1 day a week'),
    'a more frequent group must come first');
  assert.strictEqual(ranks[ranks.length - 1], '2 days a month');
});

test('group totals round UP, never down', () => {
  const g = loadApp()._taskSheetFreqGroups(MIXED);
  // Laundry is 1h38m = 98 min. Rounding to nearest 5 would state 100 either way, but rounding
  // DOWN would say 95 — five minutes of authorized care the caregiver never books.
  const laundry = g.find((x) => x.label === '1 day a week');
  assert.strictEqual(laundry.min, 100, '98 min must round up to 100, never down to 95');
  MIXED.forEach((t) => {
    const owner = g.find((x) => x.tasks.indexOf(t.task) !== -1);
    assert.ok(owner.min >= owner.tasks.reduce((a, name) => {
      const src = MIXED.find((m) => m.task === name);
      return a + (parseInt(src.perDay.split(':')[0], 10) * 60 + parseInt(src.perDay.split(':')[1], 10));
    }, 0), 'a group total must never be under the sum of its tasks');
  });
});

test('each task lands in exactly one group, none silently dropped', () => {
  const g = loadApp()._taskSheetFreqGroups(MIXED);
  const listed = g.reduce((acc, x) => acc.concat(x.tasks), []);
  assert.strictEqual(listed.length, MIXED.length,
    'a task missing from every group is time the caregiver never sees');
  assert.deepStrictEqual([...listed].sort(), MIXED.map((t) => t.task).sort());
});

test('an unrecognised frequency still shows up, labelled as written', () => {
  const g = loadApp()._taskSheetFreqGroups(
    [{ task: 'Respite', perDay: '02:00', freq: 'Every other Thursday', perMonth: '04:00' }]);
  assert.strictEqual(g.length, 1, 'an unknown frequency must not vanish from the sheet');
  assert.strictEqual(g[0].label, 'Every other Thursday');
  assert.strictEqual(g[0].text, '2h 0m');
});

test('the printed sheet and the texted image both carry the per-visit block', async () => {
  const w = app();
  w.exportCaregiverTaskSheet();
  assert.ok(/Time per visit/.test(w.captured), 'the sheet must carry the per-visit totals');
  assert.ok(/Every day<\/span><span class="vt">/.test(w.captured), 'as a printed row, not just in the text block');

  let markup = '';
  const realCreate = w.document.createElement.bind(w.document);
  w.document.createElement = (tag) => {
    const el = realCreate(tag);
    if (String(tag).toLowerCase() === 'div') {
      Object.defineProperty(el, 'innerHTML', {
        configurable: true, set(v) { markup += String(v); this.__h = v; }, get() { return this.__h || ''; },
      });
    }
    return el;
  };
  w.html2canvas = () => Promise.reject(new Error('no canvas in jsdom'));
  try { await w.shareCaregiverTaskImage(); } catch (e) { /* rasterising needs a canvas */ }
  w.document.createElement = realCreate;
  assert.ok(markup.length > 0, 'the image markup was never built — this test would prove nothing');
  assert.ok(markup.indexOf('Time per visit') !== -1, 'the image is what reaches the phone');
});
