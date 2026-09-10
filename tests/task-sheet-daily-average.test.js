'use strict';
// Caregivers had no way to tell how long to clock in for. The task rows each carry their own
// time and frequency, so nothing on the sheet answered it: a 7-days-per-week task and a
// 1-day-per-week task never add up to the same visit twice.
//
// The fix states the month's approved time spread evenly. The divisor is the part worth guarding:
// a four-week (28) divisor reads as the safer, more generous choice, but a caregiver working it
// every day of a 30-day month delivers hours the agency can never bill — MDHHS is invoiced the
// authorization exactly, so the overrun is payroll the agency absorbs. 30 keeps the padding a
// one-off cushion instead of multiplying it by thirty.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

// 34:36 authorized -> padded to 35:00 -> 2100/30 = 70 min = "1h 10m" per day.
// The 28-day divisor would give 75 min = "1h 15m", which is the mistake this file exists to catch.
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
  // plain-text block, so a bare indexOf('1h 10m') passes even with the printed row deleted.
  assert.ok(/About per day:<\/span>\s*<span class="v">1h 10m<\/span>/.test(w.captured),
    'no per-day average printed on the sheet — the caregiver still has nothing to clock against');
});

test('the per-day figure is the padded month over 30, not 28', () => {
  const w = app();
  w.exportCaregiverTaskSheet();
  assert.ok(w.captured.indexOf('1h 15m') === -1,
    'that is 2100/28 — a four-week divisor overstates every day of a calendar month');
  assert.strictEqual(w._taskSheetPerDayMin(AUTH_MIN), 70);
});

test('the daily cushion is not multiplied across the month', () => {
  const w = app();
  const perDay = w._taskSheetPerDayMin(AUTH_MIN);
  // Worked every day, the daily figure must land on the padded target exactly — not past it.
  assert.strictEqual(perDay * 30, w._taskSheetTargetMin(AUTH_MIN),
    'the per-day figure must reconstruct the padded month, not exceed it');
  // And the whole month stays inside the "slightly over" rule rather than hours over.
  const overrun = perDay * 30 - AUTH_MIN;
  assert.ok(overrun > 0 && overrun < 60,
    'monthly overrun should be minutes, not hours — got ' + overrun + ' min');
});

test('the emailed and copied text carries it too', () => {
  const w = app();
  w.exportCaregiverTaskSheet();
  assert.ok(/About per day \(average\): 1h 10m/.test(w.captured),
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
  assert.ok(markup.indexOf('1h 10m') !== -1, 'the image must show the same per-day figure');
});

test('an authorization with no approved total shows no per-day line', () => {
  const w = app({ effectiveDate: '08/01/2026', tasks: AUTH.tasks });
  w.exportCaregiverTaskSheet();
  assert.ok(w.captured.indexOf('About per day') === -1,
    'with nothing authorized there is no average to state');
});
