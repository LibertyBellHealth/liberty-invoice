'use strict';
// Roster load = merge, not blind replace (persistence audit 2026-08-05). On a cold device a worker
// can add a caregiver/caseworker while the initial load is still in flight; a blind replace would
// wipe that unsynced row. The merge preserves an unsynced local addition (no _rowVersion) while
// still letting a cross-device DELETE propagate (a row that once had a _rowVersion but is gone from
// the server is dropped). Server data wins for rows the server knows about.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp } = require('./harness');

test('_mergeRosterMap: server wins shared ids; unsynced local addition survives', () => {
  const w = loadApp();
  const server = { a: { id: 'a', name: 'Alice (server)', _rowVersion: 'v2' } };
  const local = {
    a: { id: 'a', name: 'Alice (stale local)', _rowVersion: 'v1' },
    b: { id: 'b', name: 'Bob (just added, unsynced)' },      // no _rowVersion
  };
  const out = w._mergeRosterMap(server, local);
  assert.strictEqual(out.a.name, 'Alice (server)', 'server row wins for a shared id');
  assert.ok(out.b, 'the unsynced local addition is preserved');
  assert.strictEqual(out.b.name, 'Bob (just added, unsynced)');
});

test('_mergeRosterMap: a synced local row absent from the server was deleted elsewhere -> dropped', () => {
  const w = loadApp();
  const server = { a: { id: 'a', name: 'Alice', _rowVersion: 'v2' } };
  const local = {
    a: { id: 'a', name: 'Alice', _rowVersion: 'v2' },
    z: { id: 'z', name: 'Zed (deleted on another device)', _rowVersion: 'v5' }, // had a version
  };
  const out = w._mergeRosterMap(server, local);
  assert.ok(!out.z, 'a previously-synced row gone from the server is dropped (delete propagates)');
});

test('_mergeRosterMap: empty server response does not wipe unsynced local rows', () => {
  const w = loadApp();
  const out = w._mergeRosterMap({}, { b: { id: 'b', name: 'Bob (unsynced)' } });
  assert.ok(out.b, 'transient empty server must not erase a pending local addition');
});

test('_mergeRosterArr: caseworker array merges by id with the same rules', () => {
  const w = loadApp();
  const server = [{ id: 'cw1', name: 'CW One', _rowVersion: 'v2' }];
  const local = [
    { id: 'cw1', name: 'CW One (stale)', _rowVersion: 'v1' },
    { id: 'cw2', name: 'CW Two (unsynced)' },                 // preserved
    { id: 'cw3', name: 'CW Three (deleted elsewhere)', _rowVersion: 'v9' }, // dropped
  ];
  const out = w._mergeRosterArr(server, local);
  const byId = {}; out.forEach(c => { byId[c.id] = c; });
  assert.strictEqual(byId.cw1.name, 'CW One', 'server wins the shared id');
  assert.ok(byId.cw2, 'unsynced local caseworker preserved');
  assert.ok(!byId.cw3, 'previously-synced caseworker gone from server is dropped');
});

test('_mergeRosterArr: empty/garbage inputs never throw', () => {
  const w = loadApp();
  assert.deepStrictEqual([...w._mergeRosterArr(null, null)], []);
  assert.strictEqual(w._mergeRosterArr([{ id: 'x' }], null).length, 1, 'null local is fine');
});
