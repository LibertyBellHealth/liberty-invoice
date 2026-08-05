'use strict';
// Tasks & signatures load = merge, not blind replace (persistence audit 2026-08-05). Same cold-
// device race as the rosters: a task/signature created while the initial load is in flight would be
// wiped by a blind replace. _mergeByIdKeepUnsynced preserves unsynced local additions; for tasks it
// still drops rows deleted on another device (they carry a dbId).
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp } = require('./harness');

test('_mergeByIdKeepUnsynced: unsynced local task (no dbId) survives; server wins shared ids', () => {
  const w = loadApp();
  const server = [{ id: '1', dbId: 1, text: 'Server task' }];
  const local = [
    { id: '1', dbId: 1, text: 'Stale local copy' },
    { id: 'td_local', text: 'Just added, not yet saved' },   // no dbId
  ];
  const out = w._mergeByIdKeepUnsynced(server, local, 'dbId');
  const byId = {}; out.forEach(t => { byId[t.id] = t; });
  assert.strictEqual(byId['1'].text, 'Server task', 'server wins the shared id');
  assert.ok(byId['td_local'], 'the unsynced local task is preserved');
});

test('_mergeByIdKeepUnsynced: a synced task (had dbId) gone from server is dropped (delete propagates)', () => {
  const w = loadApp();
  const server = [{ id: '1', dbId: 1, text: 'Keep' }];
  const local = [
    { id: '1', dbId: 1, text: 'Keep' },
    { id: '9', dbId: 9, text: 'Deleted on another device' },
  ];
  const out = w._mergeByIdKeepUnsynced(server, local, 'dbId');
  assert.ok(!out.some(t => t.id === '9'), 'previously-synced task absent from server is dropped');
});

test('_mergeByIdKeepUnsynced: no syncedProp (signatures) keeps ALL local-only rows', () => {
  const w = loadApp();
  const server = [{ id: 's1', label: 'Server sig' }];
  const local = [
    { id: 's1', label: 'stale' },
    { id: 's2', label: 'Local sig not yet synced' },
  ];
  const out = w._mergeByIdKeepUnsynced(server, local);          // no syncedProp
  const byId = {}; out.forEach(s => { byId[s.id] = s; });
  assert.strictEqual(byId.s1.label, 'Server sig', 'server wins shared id');
  assert.ok(byId.s2, 'local-only signature always preserved (no sync marker to judge a delete)');
});

test('_mergeByIdKeepUnsynced: empty server does not wipe local; garbage never throws', () => {
  const w = loadApp();
  assert.strictEqual(w._mergeByIdKeepUnsynced([], [{ id: 'x' }], 'dbId').length, 1, 'empty server keeps local');
  assert.deepStrictEqual([...w._mergeByIdKeepUnsynced(null, null, 'dbId')], [], 'nulls -> empty, no throw');
});
