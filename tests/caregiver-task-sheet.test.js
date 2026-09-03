'use strict';
// Two defects introduced by the task-sheet padding change (#132):
//  1. exportCaregiverTaskSheet read _sheetTasks, which was declared in renderAuthPane — so the whole
//     export threw ReferenceError before opening anything and the button did nothing, silently.
//  2. The printed table and the texted task IMAGE were still built from the RAW authorization while
//     the Authorization tab showed padded times, so the caregiver worked a schedule that
//     under-delivers the authorization — exactly what the padding exists to prevent.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

const AUTH = { hours: 34, minutes: 36, effectiveDate: '08/01/2026', tasks: [
  { task: 'Bathing', perDay: '00:05', freq: '7 days per week', perMonth: '02:30' },
  { task: 'Meal Preparation', perDay: '00:50', freq: '7 days per week', perMonth: '25:05' },
  { task: 'Laundry', perDay: '01:38', freq: '1 day per week', perMonth: '07:01' },
] };

function app() {
  const w = loadApp();
  resetStorage(w);
  w.saveProfilesLS({ 'Jane Doe': { clientName: 'Jane Doe', firstName: 'Jane', lastName: 'Doe',
    authorization: JSON.parse(JSON.stringify(AUTH)) } });
  w.activeProfileName = 'Jane Doe';
  w.showAlert = () => {};
  w.captured = '';
  // Capture what would be written into the opened window.
  w._openPhiWindow = () => ({ focus() {}, document: { write(h) { w.captured += h; }, close() {} } });
  return w;
}
const hhmmIn = (html) => (html.match(/\b\d{1,3}:\d{2}\b/g) || []);
const target = (w) => w._taskSheetTargetMin(AUTH.hours * 60 + AUTH.minutes);

test('the task sheet opens instead of throwing', () => {
  const w = app();
  assert.doesNotThrow(() => w.exportCaregiverTaskSheet(),
    'a ReferenceError here means the button does nothing at all');
  assert.ok(w.captured.length > 0, 'it must actually produce a sheet');
});

test('the printed sheet shows the PADDED times, matching the Authorization tab', () => {
  const w = app();
  w.exportCaregiverTaskSheet();
  const padded = w._taskSheetPaddedTasks(AUTH.tasks, AUTH.hours * 60 + AUTH.minutes);
  padded.forEach((t) => assert.ok(w.captured.indexOf(t.perMonth) !== -1,
    'padded time ' + t.perMonth + ' missing from the sheet'));
  // The sum of what the caregiver is shown must reach the padded target.
  const sum = [...padded].reduce((a, t) => a + w._dhsHmToMin(t.perMonth), 0);
  assert.strictEqual(sum, target(w));
});

test('the authorized total is still stated as authorized, not padded', () => {
  const w = app();
  w.exportCaregiverTaskSheet();
  assert.ok(/34h 36m|34:36/.test(w.captured),
    'the sheet must still say what MDHHS authorized: ' + hhmmIn(w.captured).join(','));
});

test('the texted task image uses the same padded times as the sheet', async () => {
  const w = app();
  // The image is built into a detached div before rasterising. Capture that markup directly —
  // falling back to the sheet's HTML would make this test pass no matter what the image contained.
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
  // It bails early unless html2canvas is loaded; stub it so the markup is actually built.
  w.html2canvas = () => Promise.reject(new Error('no canvas in jsdom'));
  try { await w.shareCaregiverTaskImage(); } catch (e) { /* rasterising needs a canvas */ }
  w.document.createElement = realCreate;
  assert.ok(markup.length > 0, 'the image markup was never built — this test would prove nothing');
  const padded = w._taskSheetPaddedTasks(AUTH.tasks, AUTH.hours * 60 + AUTH.minutes);
  padded.forEach((t) => assert.ok(markup.indexOf(t.perMonth) !== -1,
    'the image must not show a shorter schedule than the sheet: missing ' + t.perMonth));
  // And the raw, unpadded time must be gone from it.
  assert.ok(markup.indexOf('25:05') === -1, 'raw unpadded time still in the image');
});
