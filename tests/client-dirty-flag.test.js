'use strict';
// A failed client-field edit used to be silently reverted after a reload. _clientSynced (the dirty
// baseline) is memory-only BY DESIGN because its signature embeds the SSN, so after a cold load the
// dirty comparison was a no-op and _mergeProfilesLoad let the server copy win. _unsaved is the
// PHI-free durable flag that fixes it — same pattern the rosters already use.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

test('_unsaved survives a cold load and protects the edit', () => {
  const w = loadApp(); resetStorage(w);
  w.saveProfilesLS({ 'Jane Doe': { clientName: 'Jane Doe', phone: 'EDITED, save failed', _unsaved: true } });
  // simulate a reload: memory-only state is gone, only what's on disk remains
  w.eval('_clientSyncedMem = Object.create(null); _profilesCache = null;');
  const local = w.getProfiles()['Jane Doe'];
  assert.strictEqual(local._clientSynced, undefined, 'the baseline is memory-only, as designed');
  assert.strictEqual(w._profileHasUnsyncedChanges(local), true,
    'the durable flag is what keeps a failed edit alive across a reload');
  const merged = w._mergeProfilesLoad({ 'Jane Doe': { clientName: 'Jane Doe', phone: 'server copy' } },
                                      w.getProfiles());
  assert.strictEqual(merged['Jane Doe'].phone, 'EDITED, save failed',
    'the background load must not revert an edit the server never accepted');
});

test('a clean client still lets the server win', () => {
  const w = loadApp(); resetStorage(w);
  w.saveProfilesLS({ 'Jane Doe': { clientName: 'Jane Doe', phone: 'local' } });
  const merged = w._mergeProfilesLoad({ 'Jane Doe': { clientName: 'Jane Doe', phone: 'server copy' } },
                                      w.getProfiles());
  assert.strictEqual(merged['Jane Doe'].phone, 'server copy',
    'without a pending edit the server remains authoritative');
});

test('_markClientUnsaved sets and clears the flag', () => {
  const w = loadApp(); resetStorage(w);
  w.saveProfilesLS({ 'Jane Doe': { clientName: 'Jane Doe' } });
  w._markClientUnsaved('Jane Doe', true);
  assert.strictEqual(w.getProfiles()['Jane Doe']._unsaved, true);
  w._markClientUnsaved('Jane Doe', false);
  assert.strictEqual(w.getProfiles()['Jane Doe']._unsaved, undefined, 'a confirmed save clears it');
  assert.doesNotThrow(() => w._markClientUnsaved('No Such Client', true));
});

test('a 409 does NOT set the flag (that was the roster deadlock)', () => {
  const w = loadApp(); resetStorage(w);
  w.saveProfilesLS({ 'Jane Doe': { clientName: 'Jane Doe' } });
  const conflict = new Error('changed by someone else'); conflict.isConflict = true;
  // what saveProfileSP's rejection handler now does:
  w._markClientUnsaved('Jane Doe', !w._isConflict(conflict));
  assert.strictEqual(w.getProfiles()['Jane Doe']._unsaved, undefined,
    'on a conflict the SERVER is newer — pinning the local copy would deadlock every future save');
  w._markClientUnsaved('Jane Doe', !w._isConflict(new Error('network')));
  assert.strictEqual(w.getProfiles()['Jane Doe']._unsaved, true, 'a real failure still protects the edit');
});

test('the durable flag never carries PHI onto disk', () => {
  const w = loadApp(); resetStorage(w);
  w.saveProfilesLS({ 'Jane Doe': { clientName: 'Jane Doe', ssn: '123-45-6789',
                                   _clientSynced: w._clientSig({ ssn: '123-45-6789' }), _unsaved: true } });
  const onDisk = w.localStorage.getItem('lhca_profiles');
  assert.ok(!/123-45-6789/.test(onDisk), 'no SSN on disk — the reason _clientSynced is memory-only');
  assert.ok(!/_clientSynced/.test(onDisk), 'and its signature is stripped too');
  assert.ok(/_unsaved/.test(onDisk), 'while the PHI-free flag does persist, which is the point');
});
