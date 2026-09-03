'use strict';
// _mergeRosterMap / _mergeRosterArr decide what survives every background roster load. Four rules
// are load-bearing: the server is authoritative for rows it returns, a FAILED local save wins over
// it, a never-synced local addition is kept, and a previously-synced row the server no longer
// returns was deleted elsewhere and is dropped. Nothing tested any of them.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

const w = loadApp();
resetStorage(w);
const synced = (over) => Object.assign({ name: 'x', _rowVersion: '00000001' }, over || {});
const unsynced = (over) => Object.assign({ name: 'x' }, over || {});

test('the server wins for rows it returns', () => {
  const out = w._mergeRosterMap({ a: synced({ name: 'Server Alice' }) },
                                { a: synced({ name: 'Stale Alice' }) });
  assert.strictEqual(out.a.name, 'Server Alice');
});

test('a FAILED local save wins over the server copy', () => {
  const out = w._mergeRosterMap({ a: synced({ name: 'Server Alice' }) },
                                { a: synced({ name: 'My edit', _unsaved: true }) });
  assert.strictEqual(out.a.name, 'My edit', 'a failed save must not be silently reverted');
});

test('a never-synced local addition survives the load', () => {
  const out = w._mergeRosterMap({ a: synced() }, { b: unsynced({ name: 'Just added' }) });
  assert.ok(out.b, 'a caregiver added while offline must not vanish');
  assert.strictEqual(out.b.name, 'Just added');
});

test('a synced row the server no longer returns is dropped — deleted elsewhere', () => {
  const out = w._mergeRosterMap({}, { a: synced({ name: 'Deleted on another device' }) });
  assert.strictEqual(out.a, undefined);
});

test('the array form (caseworkers) follows the same four rules', () => {
  const server = [{ id: 1, name: 'Server A', _rowVersion: 'v1' }];
  const local = [
    { id: 1, name: 'Stale A', _rowVersion: 'v1' },
    { id: 2, name: 'Failed edit', _rowVersion: 'v1', _unsaved: true },
    { id: 3, name: 'Just added' },
    { id: 4, name: 'Deleted elsewhere', _rowVersion: 'v1' },
  ];
  const out = w._mergeRosterArr(server, local);
  const byId = {}; [...out].forEach((x) => { byId[x.id] = x; });
  assert.strictEqual(byId[1].name, 'Server A', 'server wins');
  // id 2 failed to save AND is absent from the server, so the save never landed — keep it, or the
  // edit is lost with no trace. (My first assertion here had this backwards; the code is right.)
  assert.ok(byId[2], 'a failed save must survive even when the server does not return the row');
  assert.strictEqual(byId[2].name, 'Failed edit');
  assert.ok(byId[3], 'unsynced addition kept');
  assert.ok(!byId[4], 'synced row absent from the server is dropped');
});

test('an EMPTY server response drops every synced row', () => {
  // Documenting real behaviour, not endorsing it: the comment above the call site says the merge
  // "keeps local rows when the server response is transiently empty", but only UNSYNCED rows are
  // kept. A transient empty response therefore empties the visible roster until the next load.
  const out = w._mergeRosterMap({}, { a: synced(), b: unsynced({ name: 'kept' }) });
  assert.strictEqual(out.a, undefined, 'synced rows are dropped');
  assert.ok(out.b, 'only unsynced rows survive');
});

test('a null or malformed local row cannot crash the merge', () => {
  assert.doesNotThrow(() => w._mergeRosterMap({ a: synced() }, { b: null, c: undefined }));
  assert.doesNotThrow(() => w._mergeRosterArr([{ id: 1 }], [null, {}, { id: null }]));
});
