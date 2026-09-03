'use strict';
// _clearAuth removes a client's DHS-1210 authorization — the thing that makes them billable and
// gates Active status. The dialog says "this client", but the callback read activeProfileName, so
// navigating to someone else before confirming removed THEIR authorization instead. Twelfth
// instance of the read-before-a-dialog pattern, found by triaging the untested writers.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

const AUTH = { hours: 62, minutes: 21, tasks: [], effectiveDate: '03/01/2026' };
function app() {
  const w = loadApp();
  resetStorage(w);
  ['saveProfileSP', 'renderAuthPane', 'addAuditEntry'].forEach((f) => { w[f] = () => {}; });
  w.alerts = [];
  w.showAlert = (m) => { w.alerts.push(String(m)); };
  w.pending = null;
  w.showConfirm = (msg, onOk) => { w.pending = { msg: msg, ok: onOk }; };
  w.saveProfilesLS({
    'Alice Adams': { clientName: 'Alice Adams', authorization: JSON.parse(JSON.stringify(AUTH)) },
    'Bob Brown': { clientName: 'Bob Brown', authorization: JSON.parse(JSON.stringify(AUTH)) },
  });
  w.activeProfileName = 'Alice Adams';
  return w;
}
const auth = (w, n) => (w.getProfiles()[n] || {}).authorization;

test('it removes the authorization from the client it was opened for', () => {
  const w = app();
  w._clearAuth();
  w.pending.ok();
  assert.strictEqual(auth(w, 'Alice Adams'), null);
  assert.ok(auth(w, 'Bob Brown'), "Bob's authorization must be untouched");
});

test('navigating to another client before confirming does not clear THEIRS', () => {
  const w = app();
  w._clearAuth();
  w.activeProfileName = 'Bob Brown';      // owner moves on while the dialog is open
  w.pending.ok();
  assert.strictEqual(auth(w, 'Alice Adams'), null, 'the client asked about is the one cleared');
  assert.ok(auth(w, 'Bob Brown'), 'the client merely being viewed must keep their authorization');
});

test('the dialog names the client, so it cannot be ambiguous', () => {
  const w = app();
  w._clearAuth();
  assert.match(w.pending.msg, /Alice Adams/);
});

test('a client deleted mid-dialog is reported, not silently skipped', () => {
  const w = app();
  w._clearAuth();
  w.saveProfilesLS({ 'Bob Brown': { clientName: 'Bob Brown', authorization: JSON.parse(JSON.stringify(AUTH)) } });
  w.pending.ok();
  assert.match(w.alerts.join(' '), /no longer available/i);
  assert.ok(auth(w, 'Bob Brown'), 'and nobody else is cleared instead');
});

test('with no client open it does nothing at all', () => {
  const w = app();
  w.activeProfileName = '';
  w._clearAuth();
  assert.strictEqual(w.pending, null);
});
