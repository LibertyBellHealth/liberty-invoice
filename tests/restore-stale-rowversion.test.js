'use strict';
// A backup carries the row_version each record had when it was taken. Restore replayed those as
// `expected_version`, so every record that STILL EXISTED on the server was rejected 409: the
// restore wrote nothing, reported "NOT FULLY RESTORED — check your connection", and could never
// succeed on retry. The stale token was persisted locally too, and _mergeProfilesLoad keeps the
// pending local copy, so the client stayed unsaveable across reloads. Restoring onto DELETED rows
// always worked (the 404 re-create path), which is why it hid.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp } = require('./harness');

test('a restored client carries no concurrency token', () => {
  const w = loadApp();
  const store = { 'Jane Doe': {
    clientName: 'Jane Doe', _dbId: 90, _rowVersion: '0000000000000011', _clientSynced: 'x',
    invoices: [{ billingPeriod: '08/2026', rowVersion: '00000000000065CE', _synced: 'y' }] } };
  w._clearSyncBaselines(store, ['Jane Doe']);
  const p = store['Jane Doe'];
  assert.strictEqual(p._rowVersion, undefined, 'the backup\'s client token must not be replayed');
  assert.strictEqual(p.invoices[0].rowVersion, undefined, 'nor the invoice\'s');
});

test('the sync baselines it already cleared are still cleared', () => {
  const w = loadApp();
  const store = { 'Jane Doe': { _clientSynced: 'x', invoices: [{ _synced: 'y' }] } };
  w._clearSyncBaselines(store, ['Jane Doe']);
  assert.strictEqual(store['Jane Doe']._clientSynced, undefined);
  assert.strictEqual(store['Jane Doe'].invoices[0]._synced, undefined);
});

test('the database id survives — it is what makes this an update, not a duplicate', () => {
  const w = loadApp();
  const store = { 'Jane Doe': { _dbId: 90, _rowVersion: 'stale', invoices: [] } };
  w._clearSyncBaselines(store, ['Jane Doe']);
  assert.strictEqual(store['Jane Doe']._dbId, 90);
});

test('with no token, the save omits expected_version entirely', () => {
  const w = loadApp();
  const sent = [];
  w.fetch = (url, opt) => { sent.push(JSON.parse(opt.body)); return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({}) }); };
  const store = { 'Jane Doe': { clientName: 'Jane Doe', _dbId: 90, _rowVersion: '0011', invoices: [] } };
  w._clearSyncBaselines(store, ['Jane Doe']);
  w.saveProfileSP('Jane Doe', store['Jane Doe']);
  const body = sent.find((b) => b && b.client_name === 'Jane Doe');
  if (body) assert.ok(!('expected_version' in body),
    'a restore must not claim to know the server version: ' + JSON.stringify(body.expected_version));
});

test('records not named in the restore are untouched', () => {
  const w = loadApp();
  const store = { A: { _rowVersion: 'keep', invoices: [] }, B: { _rowVersion: 'drop', invoices: [] } };
  w._clearSyncBaselines(store, ['B']);
  assert.strictEqual(store.A._rowVersion, 'keep');
  assert.strictEqual(store.B._rowVersion, undefined);
});
