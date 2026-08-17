'use strict';
// localStorage round-trips — the working cache the whole app reads/writes. If a save/load here
// drops a field or throws on bad data, edits silently vanish. Uses the real get/save functions.
const { test, beforeEach } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

const w = loadApp();
beforeEach(() => resetStorage(w));

test('profiles: save then load round-trips, authorization included', () => {
  const p = {
    'Jane Doe': {
      clientName: 'Jane Doe', firstName: 'Jane', lastName: 'Doe',
      authorization: { hours: 29, minutes: 47, effectiveDate: '08/01/2026',
        tasks: [{ task: 'Bathing', perDay: '00:18', freq: '7 days per week', perMonth: '09:02', amount: 243.81 }] },
    },
  };
  w.saveProfilesLS(p);
  const back = w.getProfiles();
  assert.strictEqual(back['Jane Doe'].firstName, 'Jane');
  assert.strictEqual(back['Jane Doe'].authorization.hours, 29, 'nested authorization survives');
  assert.strictEqual(back['Jane Doe'].authorization.tasks[0].amount, 243.81, 'task amount survives');
});

test('profiles: empty or corrupt storage returns {} instead of throwing', () => {
  assert.strictEqual(Object.keys(w.getProfiles()).length, 0, 'empty -> {}');
  // Corrupt JSON. getProfiles() has a tick-scoped read-cache; the first call above cached {},
  // so we must clear it to force a REAL parse of the corrupt string — otherwise this test would
  // pass even if the try/catch guard were deleted (it's what an independent audit caught).
  w.localStorage.setItem('lhca_profiles', 'not valid json{');
  w._profilesCache = null;
  assert.strictEqual(Object.keys(w.getProfiles()).length, 0, 'corrupt -> {}, no crash');
});

test('caseworkers: title round-trips through save/load', () => {
  const arr = [{ id: 'cw_1', name: 'R Feto', title: 'Mrs.', email: 'FetoR@michigan.gov' }];
  w.saveCaseworkersLS(arr);
  const back = w.getCaseworkers();
  assert.strictEqual(back.length, 1);
  assert.strictEqual(back[0].title, 'Mrs.', 'title survives');
  assert.strictEqual(back[0].email, 'FetoR@michigan.gov');
});

test('caseworkers: corrupt storage returns [] not a throw', () => {
  w.localStorage.setItem('lhca_caseworkers', '}{bad');
  assert.strictEqual(w.getCaseworkers().length, 0, 'corrupt -> [], no crash');
});

test('supervisors: keyed-by-id map with title round-trips', () => {
  const map = { sup_1: { id: 'sup_1', name: 'Rebecca Feto', title: 'Mrs.', phone: '313-505-1660', email: 'FetoR@michigan.gov' } };
  w.saveSupervisorsLS(map);
  const back = w.getSupervisors();
  assert.strictEqual(back.sup_1.name, 'Rebecca Feto');
  assert.strictEqual(back.sup_1.title, 'Mrs.', 'supervisor title survives');
});
