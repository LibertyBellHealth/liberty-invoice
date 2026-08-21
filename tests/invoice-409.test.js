'use strict';
// The no-id invoice save returns 409 when the server already has that (client, period). Without
// adopting the returned id, the invoice stays dbId-less forever: _profileHasUnsyncedChanges keeps the
// local copy, the next save is another no-id POST, and the client can NEVER be saved again. That is
// the same deadlock shape the roster merge already had once — this pins it shut.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

function seedClientWithNewInvoice(w) {
  resetStorage(w);
  w.localStorage.setItem('lhca_id_map', JSON.stringify({ 'Jane Doe': 7 }));
  w.saveProfilesLS({ 'Jane Doe': { clientName: 'Jane Doe', _dbId: 7,
    invoices: [{ billingPeriod: '08/2026', savedAt: 'S1', status: 'draft', data: { svcHH: '20', svcMM: '00' } }] } });
}

test('a 409 on a new invoice adopts the server id instead of deadlocking', async () => {
  const w = loadApp(); seedClientWithNewInvoice(w);
  w.fetch = () => Promise.resolve({
    ok: false, status: 409,
    json: () => Promise.resolve({ error: 'already exists', id: 99, row_version: '00000000000004AB' })
  });
  await w.syncNewInvoices('Jane Doe', w.getProfiles()['Jane Doe']).catch(() => {});
  const inv = w.getProfiles()['Jane Doe'].invoices[0];
  assert.strictEqual(inv.dbId, 99,
    'without this the invoice never gets an id and every future save 409s again — permanently');
  assert.strictEqual(inv.rowVersion, '00000000000004AB',
    'the concurrency token is adopted too, so the next save is version-checked rather than blind');
});

test('the conflict is still reported — the first attempt must not silently overwrite', async () => {
  const w = loadApp(); seedClientWithNewInvoice(w);
  w.fetch = () => Promise.resolve({
    ok: false, status: 409,
    json: () => Promise.resolve({ id: 99, row_version: 'AB' })
  });
  let threw = null;
  await w.syncNewInvoices('Jane Doe', w.getProfiles()['Jane Doe']).catch(e => { threw = e; });
  assert.ok(threw, 'the save still fails loudly rather than quietly clobbering the other device');
  assert.strictEqual(threw.isConflict, true, 'and it is flagged as a conflict, not a generic error');
});

test('after adopting the id the invoice is no longer treated as never-synced', async () => {
  const w = loadApp(); seedClientWithNewInvoice(w);
  const before = w._profileHasUnsyncedChanges(w.getProfiles()['Jane Doe']);
  assert.strictEqual(before, true, 'a dbId-less invoice starts out pinning the local copy');
  w.fetch = () => Promise.resolve({ ok: false, status: 409, json: () => Promise.resolve({ id: 99, row_version: 'AB' }) });
  await w.syncNewInvoices('Jane Doe', w.getProfiles()['Jane Doe']).catch(() => {});
  const inv = w.getProfiles()['Jane Doe'].invoices[0];
  assert.ok(inv.dbId, 'it now has an id, so the next save takes the version-checked UPDATE path');
});
