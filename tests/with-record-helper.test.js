'use strict';
// stillOn / whenStillOn are the one place that answers "is this still the record on screen?".
// Reading a current-record global after an async gap is what put one person's data on another
// (#159 documents, #167 invoice status, #176 authorization, #179 MI Login password). Having ONE
// definition also stops the per-site checks drifting apart, which is its own bug source here.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

function app() {
  const w = loadApp();
  resetStorage(w);
  if (!w.document.getElementById('cg-editing-id')) {
    w.document.body.insertAdjacentHTML('beforeend', '<input id="cg-editing-id">');
  }
  w.document.getElementById('cg-editing-id').value = '';
  w.activeProfileName = ''; w.activeCgId = ''; w.activeCwId = '';
  return w;
}

test('a client matches only while it is the one open', () => {
  const w = app();
  w.activeProfileName = 'Alice Adams';
  assert.strictEqual(w.stillOn('client', 'Alice Adams'), true);
  w.activeProfileName = 'Bob Brown';
  assert.strictEqual(w.stillOn('client', 'Alice Adams'), false);
});

test('a caregiver is matched by the form it was opened for, not just the global', () => {
  const w = app();
  w.document.getElementById('cg-editing-id').value = 'cg_A';
  w.activeCgId = 'cg_B';                    // the pane moved on; the FORM is still A's
  assert.strictEqual(w.stillOn('caregiver', 'cg_A'), true, 'the open form wins');
  assert.strictEqual(w.stillOn('caregiver', 'cg_B'), false);
});

test('with no form open it falls back to the active caregiver', () => {
  const w = app();
  w.activeCgId = 'cg_B';
  assert.strictEqual(w.stillOn('caregiver', 'cg_B'), true);
});

test('a caseworker matches on its own global', () => {
  const w = app();
  w.activeCwId = 'cw_1';
  assert.strictEqual(w.stillOn('caseworker', 'cw_1'), true);
  assert.strictEqual(w.stillOn('caseworker', 'cw_2'), false);
});

test('an unknown kind is false, never true — it must fail closed', () => {
  const w = app();
  w.activeProfileName = 'Alice Adams';
  assert.strictEqual(w.stillOn('nonsense', 'Alice Adams'), false,
    'defaulting to true would write to whatever is on screen');
});

test('ids compare as text, so 7 and "7" are the same record', () => {
  const w = app();
  w.activeCwId = 7;
  assert.strictEqual(w.stillOn('caseworker', '7'), true);
});

test('whenStillOn runs the write only for the right record, and says whether it ran', () => {
  const w = app();
  w.activeProfileName = 'Alice Adams';
  let ran = 0;
  assert.strictEqual(w.whenStillOn('client', 'Alice Adams', () => { ran++; }), true);
  assert.strictEqual(ran, 1);
  w.activeProfileName = 'Bob Brown';
  assert.strictEqual(w.whenStillOn('client', 'Alice Adams', () => { ran++; }), false,
    'the caller must be able to report a skip rather than assume success');
  assert.strictEqual(ran, 1, 'the write must not have happened');
});

test('the existing guards delegate to it rather than re-deriving the check', () => {
  const src = require('node:fs').readFileSync(require('node:path').join(__dirname, '..', 'app.js'), 'utf8');
  assert.match(src, /function _miloginStillCurrent\(id\)\{ return stillOn\('caregiver', id\); \}/,
    'a second copy of this check is how the two drifted apart before');
  assert.ok(/function _docListStillCurrent[\s\S]{0,400}stillOn\(kind, id\)/.test(src));
});
