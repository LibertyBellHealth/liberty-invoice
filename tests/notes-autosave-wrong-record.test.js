'use strict';
// The caregiver and caseworker note panes debounced for 600ms and THEN read activeCgId/activeCwId
// to decide whose record to write. Type a note on caregiver A, click caregiver B inside that
// window, and the timer resolved the *now*-current id and wrote A's text into B's record — both
// localStorage and the backend, overwriting B's real notes. `ta` was still A's textarea holding A's
// text, and each render created a fresh timer handle, so the pending save was never cancelled.
//
// The same panes also wrote localStorage only INSIDE the timer, so closing the tab mid-debounce
// lost the note outright. The client pane (renderNotesPane) already had both right; this is that
// pattern ported across.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

function app() {
  const w = loadApp();
  resetStorage(w);
  ['cgNotesContent', 'cwNotesContent'].forEach((id) => {
    if (!w.document.getElementById(id)) {
      w.document.body.insertAdjacentHTML('beforeend', '<div id="' + id + '"></div>');
    }
    w.document.getElementById(id).innerHTML = '';
  });
  // Drop any debounced save left registered by an earlier test in this file.
  Object.keys(w._pendingNoteSaves || {}).forEach((k) => { delete w._pendingNoteSaves[k]; });
  w.activeCgId = ''; w.activeCwId = ''; w.activeProfileName = '';
  w.saved = [];
  w.saveCaregiverAPI = (id, cg) => { w.saved.push({ id, notes: cg && cg.notes }); return Promise.resolve(); };
  w.saveCaseworkerAPI = (cw) => { w.saved.push({ id: cw && cw.id, notes: cw && cw.notes }); return Promise.resolve(); };
  return w;
}

const typeInto = (w, areaId, text) => {
  const ta = w.document.getElementById(areaId);
  ta.value = text;
  ta.dispatchEvent(new w.Event('input'));
  return ta;
};
const settle = () => new Promise((r) => setTimeout(r, 0));

function twoCaregivers(w) {
  w.saveCaregiversLS({
    cg_A: { id: 'cg_A', name: 'Alice Aide', notes: 'A original' },
    cg_B: { id: 'cg_B', name: 'Bob Buddy', notes: 'B original' },
  });
}
function twoCaseworkers(w) {
  w.saveCaseworkersLS([
    { id: 'cw_A', name: 'Carol Case', notes: 'A original' },
    { id: 'cw_B', name: 'Dave Desk', notes: 'B original' },
  ]);
}

test('a caregiver note typed for A is not written onto B when you switch inside the debounce', async () => {
  const w = app();
  twoCaregivers(w);
  w.activeCgId = 'cg_A';
  w.renderCgNotesPane();
  typeInto(w, 'cgNotesArea', 'A private note');

  w.activeCgId = 'cg_B';           // owner clicks caregiver B before the 600ms flush
  w._flushPendingNoteSaves();      // the debounced backend save fires
  await settle();

  const cgs = w.getCaregivers();
  assert.strictEqual(cgs.cg_B.notes, 'B original', "B's notes must be untouched");
  assert.strictEqual(cgs.cg_A.notes, 'A private note', "A's note must be saved to A");
  assert.deepStrictEqual(w.saved, [{ id: 'cg_A', notes: 'A private note' }],
    'the backend save must name caregiver A, not whoever is on screen now');
});

test('a caregiver note reaches localStorage synchronously, so a tab close mid-debounce cannot lose it', () => {
  const w = app();
  twoCaregivers(w);
  w.activeCgId = 'cg_A';
  w.renderCgNotesPane();
  typeInto(w, 'cgNotesArea', 'typed then F4');
  // No flush, no timer: the note is already durable.
  assert.strictEqual(w.getCaregivers().cg_A.notes, 'typed then F4');
});

test('the pending caregiver save is registered so pagehide can flush it', async () => {
  const w = app();
  twoCaregivers(w);
  w.activeCgId = 'cg_A';
  w.renderCgNotesPane();
  typeInto(w, 'cgNotesArea', 'flush me');
  assert.ok(Object.prototype.hasOwnProperty.call(w._pendingNoteSaves, 'caregiver:cg_A'),
    'the backend save must be registered under its own caregiver key');
  w._flushPendingNoteSaves();
  await settle();
  assert.deepStrictEqual(w.saved, [{ id: 'cg_A', notes: 'flush me' }]);
});

test('"Saved ✓" does not appear under the caregiver you switched to', async () => {
  const w = app();
  twoCaregivers(w);
  w.activeCgId = 'cg_A';
  w.renderCgNotesPane();
  typeInto(w, 'cgNotesArea', 'A note');
  w.activeCgId = 'cg_B';
  w.renderCgNotesPane();           // B's pane, with B's flash element
  w._flushPendingNoteSaves();
  await settle();
  const flash = w.document.getElementById('cgNotesSavedFlash');
  assert.strictEqual(flash.style.display, 'none',
    "A's save must not tick 'Saved ✓' on the caregiver now on screen");
});

test('a caseworker note typed for A is not written onto B when you switch inside the debounce', async () => {
  const w = app();
  twoCaseworkers(w);
  w.activeCwId = 'cw_A';
  w.renderCwNotesPane();
  typeInto(w, 'cwNotesArea', 'A private note');

  w.activeCwId = 'cw_B';
  w._flushPendingNoteSaves();
  await settle();

  const arr = w.getCaseworkers();
  const byId = (id) => arr.find((c) => c.id === id);
  assert.strictEqual(byId('cw_B').notes, 'B original', "B's notes must be untouched");
  assert.strictEqual(byId('cw_A').notes, 'A private note', "A's note must be saved to A");
  assert.deepStrictEqual(w.saved, [{ id: 'cw_A', notes: 'A private note' }],
    'the backend save must name caseworker A');
});

test('a caseworker note reaches localStorage synchronously', () => {
  const w = app();
  twoCaseworkers(w);
  w.activeCwId = 'cw_A';
  w.renderCwNotesPane();
  typeInto(w, 'cwNotesArea', 'typed then F4');
  assert.strictEqual(w.getCaseworkers().find((c) => c.id === 'cw_A').notes, 'typed then F4');
});

// doSave is passed to flashQuietSave as the RETRY handler, so it can run long after the first
// attempt failed — by which time the operator may be looking at someone else. Deciding the flash
// target once, when the debounce fired, meant a successful retry ticked "Saved ✓" under whoever
// was on screen at retry time. Same family as the write, in the report.
function failingSave(w) {
  const retries = [];
  w.saveCaregiverAPI = () => Promise.reject(new Error('network error'));
  w._showSaveStatus = (state, label, onRetry) => { if (onRetry) retries.push(onRetry); };
  return retries;
}

test('a retry that succeeds does not tick "Saved ✓" under a different caregiver', async () => {
  const w = app();
  twoCaregivers(w);
  w.activeCgId = 'cg_A';
  w.renderCgNotesPane();
  const retries = failingSave(w);
  typeInto(w, 'cgNotesArea', 'A note');
  w._flushPendingNoteSaves();
  await settle();
  assert.strictEqual(retries.length, 1, 'the failed save should offer a retry');

  w.activeCgId = 'cg_B';
  w.renderCgNotesPane();                     // B's pane, B's flash element
  w.saveCaregiverAPI = () => Promise.resolve();
  retries[0]();                              // operator clicks retry while looking at B
  await settle();

  assert.strictEqual(w.document.getElementById('cgNotesSavedFlash').style.display, 'none',
    "A's retry ticked Saved on the caregiver now on screen");
});

test('a retry that succeeds while still on the same caregiver does tick', async () => {
  const w = app();
  twoCaregivers(w);
  w.activeCgId = 'cg_A';
  w.renderCgNotesPane();
  const retries = failingSave(w);
  typeInto(w, 'cgNotesArea', 'A note');
  w._flushPendingNoteSaves();
  await settle();

  w.saveCaregiverAPI = () => Promise.resolve();
  retries[0]();
  await settle();

  assert.strictEqual(w.document.getElementById('cgNotesSavedFlash').style.display, 'inline',
    'the guard must not stop a legitimate retry from confirming');
});

// doSave is the retry handler, and it used to close over the record captured when the debounce
// fired. A retry clicked after further typing re-sent the OLD note over the newer one, and the next
// roster load pulled that back into localStorage — so the newer text was gone from both.
test('a stale retry does not revert a newer note', async () => {
  const w = app();
  twoCaregivers(w);
  w.activeCgId = 'cg_A';
  w.renderCgNotesPane();
  const sentNotes = [];
  const retries = [];
  w.saveCaregiverAPI = (id, cg) => { sentNotes.push(cg && cg.notes); return Promise.reject(new Error('network error')); };
  w._showSaveStatus = (state, label, onRetry) => { if (onRetry) retries.push(onRetry); };

  typeInto(w, 'cgNotesArea', 'hello');
  w._flushPendingNoteSaves();
  await settle();
  assert.strictEqual(retries.length, 1, 'the failed save should offer a retry');

  // The operator keeps typing; that save succeeds.
  w.saveCaregiverAPI = (id, cg) => { sentNotes.push(cg && cg.notes); return Promise.resolve(); };
  typeInto(w, 'cgNotesArea', 'hello world');
  w._flushPendingNoteSaves();
  await settle();

  retries[0]();                 // then clicks the failure banner still on screen
  await settle();

  assert.strictEqual(sentNotes[sentNotes.length - 1], 'hello world',
    'the retry re-sent the note as it was when the first attempt failed: ' + JSON.stringify(sentNotes));
  assert.strictEqual(w.getCaregivers().cg_A.notes, 'hello world');
});
