'use strict';
// Caregivers had no way to tell how long to clock in for. The task rows each carry their own time
// and frequency, so nothing on the sheet answered it: a 7-days-per-week task and a 1-day-per-week
// task never add up to the same visit twice.
//
// A single flat average (the whole month over N days) was tried and removed — the "average day"
// does not exist, so it was wrong on every actual day of the week, and it read as contradicting
// the everyday total sitting next to it. What replaced it groups the rows by how often they run
// and totals each group; the caregiver adds up whichever lines apply today. Owner, 2026-09-10:
// "7 days a week should be this much time per day. Per 3 days. Per 2 days. Per 1 day."
//
// Group totals round UP to the next 5 minutes: being under the authorized time is the failure
// this guards against ("Its okay if they go over but not under").
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
