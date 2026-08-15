'use strict';
// Client-profile load = merge, not blind-replace (persistence audit follow-up, 2026-08-05). This is
// the store the earlier merge fixes left uncovered: loadProfilesAPI overwrote LS wholesale, guarded
// only by the in-flight window. A client save that FAILED (data left only in LS, _savesInFlight back
// to 0) was then silently reverted by the next background revalidate. These pin down the merge that
// closes that path — server wins for clean synced clients, unsynced local work is preserved, and
// cross-device deletes still propagate.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp } = require('./harness');

// Build a client whose _clientSynced baseline MATCHES its current fields (i.e. "clean", last save
// confirmed). We derive the baseline from the app's own _clientSig so the test tracks real logic.
function cleanClient(w, over) {
  const c = Object.assign({ clientName: 'X', firstName: 'X', clientStatus: 'active', _dbId: 1, invoices: [] }, over || {});
  c._clientSynced = w._clientSig(c);
  return c;
}

test('_profileHasUnsyncedChanges: clean client = false; an edited field = true', () => {
  const w = loadApp();
  const c = cleanClient(w, { firstName: 'Jane', _dbId: 7 });
  assert.strictEqual(w._profileHasUnsyncedChanges(c), false, 'baseline matches -> nothing pending');
  c.firstName = 'Janet';                                    // edit after the baseline
  assert.strictEqual(w._profileHasUnsyncedChanges(c), true, 'a changed field is detected as unsynced');
});

test('fresh session (no _clientSynced baseline) is NOT dirty -> server SSN wins, not blank local', () => {
  const w = loadApp();
  // On the FIRST load of a session, local profiles have SSN + _clientSynced stripped (memory-only).
  const local = { firstName: 'Jane', _dbId: 7, invoices: [], ssn: '' };   // no _clientSynced, no ssn
  assert.strictEqual(w._profileHasUnsyncedChanges(local), false,
    'without a baseline a client must NOT be treated as unsynced (that blanked the SSN)');
  // The merge must therefore take the SERVER copy, which carries the decrypted SSN.
  const server = { Jane: { clientName: 'Jane', firstName: 'Jane', _dbId: 7, ssn: '123-45-6789',
    invoices: [], _clientSynced: 'baseline' } };
  const out = w._mergeProfilesLoad(server, { Jane: local });
  assert.strictEqual(out.Jane.ssn, '123-45-6789', 'fresh load keeps the server SSN, not the blank local');
});

test('_profileHasUnsyncedChanges: a never-synced invoice (no dbId) counts as pending', () => {
  const w = loadApp();
  const c = cleanClient(w, { _dbId: 7 });
  c.invoices = [{ billingPeriod: '08/2026', status: 'draft', data: {} }]; // no dbId
  assert.strictEqual(w._profileHasUnsyncedChanges(c), true, 'a new local invoice must be protected');
});

test('_mergeProfilesLoad: clean synced client -> SERVER wins (fresh data)', () => {
  const w = loadApp();
  const server = { Jane: cleanClient(w, { clientName: 'Jane', firstName: 'Jane', phone: 'NEW', _dbId: 7 }) };
  const local  = { Jane: cleanClient(w, { clientName: 'Jane', firstName: 'Jane', phone: 'OLD', _dbId: 7 }) };
  const out = w._mergeProfilesLoad(server, local);
  assert.strictEqual(out.Jane.phone, 'NEW', 'a clean client takes the fresh server copy');
});

test('_mergeProfilesLoad: a FAILED edit to an existing client is preserved, not reverted', () => {
  const w = loadApp();
  const server = { Jane: cleanClient(w, { clientName: 'Jane', firstName: 'Jane', phone: 'SERVER', _dbId: 7 }) };
  // local copy has an unsaved edit: phone changed but _clientSynced still reflects the old baseline
  const local  = cleanClient(w, { clientName: 'Jane', firstName: 'Jane', phone: 'OLD', _dbId: 7 });
  local.phone = 'EDIT-THAT-FAILED-TO-SAVE';
  const out = w._mergeProfilesLoad(server, { Jane: local });
  assert.strictEqual(out.Jane.phone, 'EDIT-THAT-FAILED-TO-SAVE', 'the pending edit survives the background load');
});

test('_mergeProfilesLoad: an unsynced NEW client (no _dbId) survives the load', () => {
  const w = loadApp();
  const server = {};                                        // server has never heard of this client
  const local  = { Bob: { clientName: 'Bob', firstName: 'Bob', clientStatus: 'active', invoices: [] } }; // no _dbId
  const out = w._mergeProfilesLoad(server, local);
  assert.ok(out.Bob, 'a just-added client that has not synced yet is not wiped');
});

test('_mergeProfilesLoad: a synced client gone from the server was deleted elsewhere -> dropped', () => {
  const w = loadApp();
  const server = {};
  const local  = { Zed: cleanClient(w, { clientName: 'Zed', firstName: 'Zed', _dbId: 9 }) }; // synced, clean
  const out = w._mergeProfilesLoad(server, local);
  assert.ok(!out.Zed, 'a previously-synced clean client absent from the server is dropped (delete propagates)');
});

test('_mergeProfilesLoad: a client deleted elsewhere but with a pending local edit is KEPT (safe side)', () => {
  const w = loadApp();
  const server = {};
  const local  = cleanClient(w, { clientName: 'Zed', firstName: 'Zed', _dbId: 9 });
  local.firstName = 'Zed-edited';                           // unsaved edit
  const out = w._mergeProfilesLoad(server, { Zed: local });
  assert.ok(out.Zed, 'never discard unsaved work, even if the row was deleted on another device');
});

test('_mergeProfilesLoad: empty/garbage never throws', () => {
  const w = loadApp();
  assert.strictEqual(Object.keys(w._mergeProfilesLoad({}, {})).length, 0, 'empty -> empty');
  assert.doesNotThrow(() => w._mergeProfilesLoad({ A: null }, { B: null }), 'null entries tolerated');
});
