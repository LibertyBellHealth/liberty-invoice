'use strict';
// A 409 on an invoice save means the server holds a row for this (client, period) that our local
// copy was NOT derived from.
//
// Two things must both be true, and they pull against each other:
//   • the local invoice must adopt the server's ID, or it stays dbId-less forever —
//     _profileHasUnsyncedChanges pins the local copy, every save is another no-id POST, and the
//     client can never be saved again;
//   • it must NOT adopt the server's row_version. `expected_version` means "the version whose
//     content I edited". Pinning the server's CURRENT version to content that never came from it
//     makes the backend's `AND row_version = @expectedVersion` guard compare the server against
//     itself — so the guard passes and the stale copy lands on top of the other device's invoice.
//     And it lands automatically: a 409 leaves _synced unset, so the invoice is still dirty and the
//     next save of that client for ANY reason re-sends it. No retry click required.
//
// The earlier version of this file asserted the version-adoption as correct behaviour.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage, jsonEqual } = require('./harness');

const SERVER_VER = '00000000000004AB';

function seedClientWithNewInvoice(w, data) {
  resetStorage(w);
  w.localStorage.setItem('lhca_id_map', JSON.stringify({ 'Jane Doe': 7 }));
  w.saveProfilesLS({ 'Jane Doe': { clientName: 'Jane Doe', _dbId: 7,
    invoices: [{ billingPeriod: '08/2026', savedAt: 'S1', status: 'draft',
                 data: data || { svcHH: '20', svcMM: '00' } }] } });
}

// POST /invoices -> 409 (+id/version); GET /invoices?clientId=7 -> what the server really holds.
function conflictingServer(w, serverInvoiceData) {
  w.fetch = (url, opt) => {
    if (!opt || opt.method !== 'POST') {
      return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(
        [{ id: 99, billing_period: '08/2026', status: 'draft', invoice_note: '',
           invoice_data: JSON.stringify(serverInvoiceData), row_version: SERVER_VER }]) });
    }
    return Promise.resolve({ ok: false, status: 409,
      json: () => Promise.resolve({ error: 'already exists', id: 99, row_version: SERVER_VER }) });
  };
}

test('a 409 adopts the server id — the invoice must not deadlock', async () => {
  const w = loadApp(); seedClientWithNewInvoice(w);
  conflictingServer(w, { svcHH: '11', svcMM: '00' });        // server holds something DIFFERENT
  await w.syncNewInvoices('Jane Doe', w.getProfiles()['Jane Doe']).catch(() => {});
  const inv = w.getProfiles()['Jane Doe'].invoices[0];
  assert.strictEqual(inv.dbId, 99, 'without an id every future save is another no-id POST — permanently');
});

test('a REAL conflict does not adopt the server version', async () => {
  const w = loadApp(); seedClientWithNewInvoice(w);
  conflictingServer(w, { svcHH: '11', svcMM: '00' });
  await w.syncNewInvoices('Jane Doe', w.getProfiles()['Jane Doe']).catch(() => {});
  const inv = w.getProfiles()['Jane Doe'].invoices[0];
  assert.strictEqual(inv._conflict, true, 'the invoice is flagged so it stops being sent');
  assert.ok(!inv.rowVersion, 'holding the server token here is what let the stale copy overwrite it');
  assert.strictEqual(inv._conflictVersion, SERVER_VER, 'kept only so a deliberate overwrite is still possible');
});

test('THE BUG: the next automatic save must not overwrite the server row', async () => {
  const w = loadApp(); seedClientWithNewInvoice(w, { svcHH: '99', svcMM: '00' });
  conflictingServer(w, { svcHH: '11', svcMM: '00' });
  await w.syncNewInvoices('Jane Doe', w.getProfiles()['Jane Doe']).catch(() => {});

  // Second save. Nothing changed on the server; the user did nothing. This is the save that used
  // to send the stale local invoice under the server's own version token.
  const posts = [];
  w.fetch = (url, opt) => {
    if (opt && opt.method === 'POST') { posts.push(JSON.parse(opt.body)); }
    return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({ id: 99, row_version: 'FF' }) });
  };
  await w.syncNewInvoices('Jane Doe', w.getProfiles()['Jane Doe']).catch(() => {});
  assert.deepStrictEqual(posts, [],
    'a conflicted invoice must never be written automatically — this POST destroyed the other device’s invoice');
});

test('a 409 whose server content is IDENTICAL is a lost response, not a conflict', async () => {
  const w = loadApp(); seedClientWithNewInvoice(w, { svcHH: '20', svcMM: '00' });
  conflictingServer(w, { svcHH: '20', svcMM: '00' });        // same content — our save did land
  await w.syncNewInvoices('Jane Doe', w.getProfiles()['Jane Doe']).catch(() => {});
  const inv = w.getProfiles()['Jane Doe'].invoices[0];
  assert.ok(!inv._conflict, 'nothing is in conflict — the response was simply lost');
  assert.strictEqual(inv.rowVersion, SERVER_VER, 'so the token IS adopted and normal saving resumes');
  assert.ok(inv._synced, 'and it is marked clean rather than re-sent forever');
});

test('an unreadable server fails CLOSED — treated as a conflict, never overwritten', async () => {
  const w = loadApp(); seedClientWithNewInvoice(w);
  w.fetch = (url, opt) => (!opt || opt.method !== 'POST')
    ? Promise.resolve({ ok: false, status: 500, json: () => Promise.resolve(null) })
    : Promise.resolve({ ok: false, status: 409, json: () => Promise.resolve({ id: 99, row_version: SERVER_VER }) });
  await w.syncNewInvoices('Jane Doe', w.getProfiles()['Jane Doe']).catch(() => {});
  const inv = w.getProfiles()['Jane Doe'].invoices[0];
  assert.strictEqual(inv._conflict, true, 'if we cannot see what the server holds we must not write over it');
});

test('the conflict is still reported loudly', async () => {
  const w = loadApp(); seedClientWithNewInvoice(w);
  conflictingServer(w, { svcHH: '11', svcMM: '00' });
  let threw = null;
  await w.syncNewInvoices('Jane Doe', w.getProfiles()['Jane Doe']).catch(e => { threw = e; });
  assert.ok(threw, 'the save fails rather than quietly clobbering the other device');
  assert.strictEqual(threw.isConflict, true, 'flagged as a conflict, not a generic error');
  assert.ok(jsonEqual(threw.conflictPeriods, ['08/2026']), 'and it names the period');
});

test('resolveInvoiceConflict("mine") is the only path to a deliberate overwrite', async () => {
  const w = loadApp(); seedClientWithNewInvoice(w, { svcHH: '99', svcMM: '00' });
  conflictingServer(w, { svcHH: '11', svcMM: '00' });
  await w.syncNewInvoices('Jane Doe', w.getProfiles()['Jane Doe']).catch(() => {});

  w.saveProfileSP = () => Promise.resolve();
  assert.strictEqual(w.resolveInvoiceConflict('Jane Doe', '08/2026', 'mine'), true);
  const inv = w.getProfiles()['Jane Doe'].invoices[0];
  assert.ok(!inv._conflict, 'cleared, so it sends again');
  assert.strictEqual(inv.rowVersion, SERVER_VER, 'now carrying the server version — by explicit choice');

  const posts = [];
  w.fetch = (url, opt) => { if (opt && opt.method === 'POST') posts.push(JSON.parse(opt.body));
    return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({ id: 99, row_version: 'FF' }) }); };
  await w.syncNewInvoices('Jane Doe', w.getProfiles()['Jane Doe']).catch(() => {});
  assert.strictEqual(posts.length, 1, 'and it is written');
  assert.match(String(posts[0].invoice_data), /"svcHH":"99"/, 'with the local content the user chose to keep');
});

test('resolveInvoiceConflict("server") drops the local copy', async () => {
  const w = loadApp(); seedClientWithNewInvoice(w, { svcHH: '99', svcMM: '00' });
  conflictingServer(w, { svcHH: '11', svcMM: '00' });
  await w.syncNewInvoices('Jane Doe', w.getProfiles()['Jane Doe']).catch(() => {});
  w.saveProfileSP = () => Promise.resolve();
  assert.strictEqual(w.resolveInvoiceConflict('Jane Doe', '08/2026', 'server'), true);
  assert.strictEqual(w.getProfiles()['Jane Doe'].invoices.length, 0,
    'gone locally, so the next background load brings the server copy back');
});
