'use strict';
// Authorization time/date helpers + the client dirty-signature (what decides whether a save
// actually fires — a regression here means edits silently don't persist).
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp } = require('./harness');

test('_authHM: formats approved time as HH:MM like the DHS-1210', () => {
  const w = loadApp();
  assert.strictEqual(w._authHM({ hours: 29, minutes: 47 }), '29:47');
  assert.strictEqual(w._authHM({ hours: 30, minutes: 0 }), '30:00', 'pads minutes');
  assert.strictEqual(w._authHM({ hours: 8, minutes: 5 }), '8:05');
  assert.strictEqual(w._authHM({ hours: null }), '', 'no hours -> empty');
  assert.strictEqual(w._authHM(null), '', 'no authorization -> empty');
});

test('_parseHM: parses HH:MM (or bare hours) back to {hours,minutes}', () => {
  const w = loadApp();
  assert.deepStrictEqual({ ...w._parseHM('29:47') }, { hours: 29, minutes: 47 });
  assert.deepStrictEqual({ ...w._parseHM('30') }, { hours: 30, minutes: 0 }, 'bare hours -> :00');
  assert.deepStrictEqual({ ...w._parseHM('') }, { hours: null, minutes: null }, 'blank clears');
});

test('_authHM <-> _parseHM round-trips', () => {
  const w = loadApp();
  for (const a of [{ hours: 29, minutes: 47 }, { hours: 40, minutes: 0 }, { hours: 1, minutes: 9 }]) {
    assert.deepStrictEqual({ ...w._parseHM(w._authHM(a)) }, a, `round-trip ${a.hours}:${a.minutes}`);
  }
});

test('_mdyToYmd: converts MM/DD/YYYY to YYYY-MM-DD (task due dates)', () => {
  const w = loadApp();
  assert.strictEqual(w._mdyToYmd('08/01/2026'), '2026-08-01');
  assert.strictEqual(w._mdyToYmd('12/25/2026'), '2026-12-25');
  assert.strictEqual(w._mdyToYmd(''), '', 'blank -> blank');
  assert.strictEqual(w._mdyToYmd('not a date'), '', 'garbage -> blank, no crash');
});

test('_nextReassessment: next 6-month date on/after today (catches up old forms)', () => {
  const w = loadApp();
  const ref = new Date(2026, 7, 4); // fixed reference: Aug 4, 2026
  // Recent form: effective + 6mo is already in the future -> unchanged, with year rollover
  assert.strictEqual(w._nextReassessment('08/01/2026', ref), '02/01/2027', 'Aug -> Feb next year');
  assert.strictEqual(w._nextReassessment('12/01/2026', ref), '06/01/2027', 'Dec rolls into next year');
  assert.strictEqual(w._nextReassessment('10/15/2026', ref), '04/15/2027', 'keeps the day');
  // OLD form (the real scenario): effective years ago -> steps 6mo to the next upcoming date
  assert.strictEqual(w._nextReassessment('09/19/2024', ref), '09/19/2026', 'old form catches up, not a past date');
  // Boundary: a reassessment landing exactly on today stays today (not skipped forward)
  assert.strictEqual(w._nextReassessment('02/04/2026', ref), '08/04/2026', 'due exactly today stays');
  assert.strictEqual(w._nextReassessment('garbage', ref), '', 'invalid -> empty, no crash');
});

test('_clientSig: detects an authorization change (so the save actually fires)', () => {
  const w = loadApp();
  const base = { firstName: 'Jane', lastName: 'Doe', authorization: { hours: 29, minutes: 47 } };
  assert.strictEqual(w._clientSig(base), w._clientSig({ ...base }), 'identical data -> identical signature');
  const changed = { ...base, authorization: { hours: 30, minutes: 0 } };
  assert.notStrictEqual(w._clientSig(base), w._clientSig(changed), 'changed authorization -> different signature');
  const cleared = { ...base, authorization: null };
  assert.notStrictEqual(w._clientSig(base), w._clientSig(cleared), 'cleared authorization -> different signature');
});
