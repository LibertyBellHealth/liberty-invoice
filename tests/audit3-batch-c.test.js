'use strict';
// Batch C: permanent loss on restore, and a retry that created a second billable client.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

// ── A lost CREATE response used to duplicate the client ──────────────────────────────────────
// The 404 re-create was guarded only by "the server said 404". If the create landed but its
// response was lost -- the very flakiness that caused the retry -- clicking Retry made a SECOND
// client row, and syncNewInvoices then wrote a full set of invoices under it.
function seed404(w) {
  resetStorage(w);
  w.saveProfilesLS({ 'Jane Doe': { clientName: 'Jane Doe', _dbId: 11, invoices: [] } });
  w.localStorage.setItem('lhca_id_map', JSON.stringify({ 'Jane Doe': 11 }));
}

test('a 404 re-create adopts an existing row instead of duplicating the client', async () => {
  const w = loadApp(); seed404(w);
  const calls = [];
  w.fetch = (url, opt) => {
    const method = (opt && opt.method) || 'GET';
    calls.push(method + ' ' + String(url).replace(/^.*\/api/, ''));
    if (method === 'POST') return Promise.resolve({ ok: false, status: 404, json: () => Promise.resolve({}) });
    // the lookup: the earlier create DID land, under a new id
    return Promise.resolve({ ok: true, status: 200,
      json: () => Promise.resolve([{ id: 77, client_name: 'Jane Doe', row_version_hex: 'BEEF' }]) });
  };
  await Promise.resolve(w.saveProfileSP('Jane Doe', w.getProfiles()['Jane Doe'], true)).catch(() => {});
  await new Promise((r) => setTimeout(r, 10));
  const posts = calls.filter((c) => c.startsWith('POST'));
  assert.strictEqual(posts.length, 1, 'only the original save should POST — the re-create must adopt: ' + calls.join(' | '));
  assert.strictEqual(JSON.parse(w.localStorage.getItem('lhca_id_map'))['Jane Doe'], 77,
    'the client must now point at the row that actually exists');
});

test('a 404 re-create still CREATES when the client genuinely is not there', async () => {
  const w = loadApp(); seed404(w);
  let creates = 0;
  w.fetch = (url, opt) => {
    const method = (opt && opt.method) || 'GET';
    if (method === 'POST') {
      const b = JSON.parse(opt.body || '{}');
      if (b.id) return Promise.resolve({ ok: false, status: 404, json: () => Promise.resolve({}) });
      creates++;
      return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({ id: 88, row_version: 'AA' }) });
    }
    return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve([]) });  // lookup finds nothing
  };
  await Promise.resolve(w.saveProfileSP('Jane Doe', w.getProfiles()['Jane Doe'], true)).catch(() => {});
  await new Promise((r) => setTimeout(r, 10));
  assert.strictEqual(creates, 1, 'a genuinely deleted row must still be re-created — that is what the retry is for');
});

test('two concurrent saves cannot both re-create the client', async () => {
  const w = loadApp(); seed404(w);
  let creates = 0;
  w.fetch = (url, opt) => {
    const method = (opt && opt.method) || 'GET';
    if (method === 'POST') {
      const b = JSON.parse(opt.body || '{}');
      if (b.id) return Promise.resolve({ ok: false, status: 404, json: () => Promise.resolve({}) });
      creates++;
      return new Promise((res) => setTimeout(() => res(
        { ok: true, status: 200, json: () => Promise.resolve({ id: 88, row_version: 'AA' }) }), 5));
    }
    return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve([]) });
  };
  const p = w.getProfiles()['Jane Doe'];
  await Promise.all([
    Promise.resolve(w.saveProfileSP('Jane Doe', p, true)).catch(() => {}),
    Promise.resolve(w.saveProfileSP('Jane Doe', p, true)).catch(() => {}),
  ]);
  await new Promise((r) => setTimeout(r, 20));
  assert.strictEqual(creates, 1, 'a deliberate save racing a note flush made two identical billable clients');
});

// ── Supervisors were in no backup at all ─────────────────────────────────────────────────────
test('the backup payload carries supervisors', () => {
  const w = loadApp(); resetStorage(w);
  w.saveSupervisorsLS({ sup1: { id: 'sup1', name: 'Pat Reed' } });
  const payload = w._buildBackupPayload(false);
  assert.ok(payload.supervisors, 'a wiped device restoring this backup lost every supervisor, permanently');
  assert.strictEqual(payload.supervisors.sup1.name, 'Pat Reed');
});

test('the importer surfaces supervisors from a backup file', () => {
  const w = loadApp();
  const parsed = w._parseBackupFile(JSON.stringify({
    _exportedAt: 'x', clients: {}, supervisors: { sup1: { id: 'sup1', name: 'Pat Reed' } } }));
  assert.ok(parsed.supervisors, 'the parser dropped supervisors on the floor');
  assert.strictEqual(parsed.supervisors.sup1.name, 'Pat Reed');
});

// ── Tasks were the one entity with no durable dirty flag ─────────────────────────────────────
test('a failed task save marks the task unsaved; a successful one clears it', async () => {
  const w = loadApp(); resetStorage(w);
  w.saveTodos([{ id: 'td_1', text: 'Call the caseworker', dbId: 5, done: false }]);
  w.fetch = () => Promise.resolve({ ok: false, status: 500, json: () => Promise.resolve({}) });
  await Promise.resolve(w.saveTaskAPI(w.getTodos()[0])).catch(() => {});
  await new Promise((r) => setTimeout(r, 10));
  assert.ok(w.getTodos()[0]._unsaved,
    'without the flag a failed edit to an already-synced task is reverted by the next load');
  w.fetch = () => Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({ id: 5 }) });
  await Promise.resolve(w.saveTaskAPI(w.getTodos()[0])).catch(() => {});
  await new Promise((r) => setTimeout(r, 10));
  assert.ok(!w.getTodos()[0]._unsaved, 'a successful save must clear it');
});

// ── Restore actually reaching the database ───────────────────────────────────────────────────
// Signatures were written to localStorage only, unlike every other roster, so the dialog reported
// them restored while they never reached the DB -- and a signature is what certifies an invoice.
// Supervisors had no restore branch at all.
function runImport(w, payload) {
  const posted = [];
  w.fetch = (url, opt) => {
    posted.push({ url: String(url).replace(/^.*\/api/, ''), method: (opt && opt.method) || 'GET',
                  body: opt && opt.body ? JSON.parse(opt.body) : null });
    return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({ id: 1, row_version: 'AA' }) });
  };
  w.showConfirm = () => {}; w.showAlert = () => {};
  w.renderSidebarClients = () => {}; w.renderClientGrid = () => {}; w.updateStats = () => {};
  w.importProfiles({ target: { files: [new w.File([JSON.stringify(payload)], 'backup.json',
    { type: 'application/json' })], value: '' } });
  return posted;
}

test('a restore pushes supervisors and signatures to the server, not just localStorage', async () => {
  const w = loadApp(); resetStorage(w);
  const posted = runImport(w, {
    _exportedAt: '2026-08-20T00:00:00Z',
    clients: { 'Jane Doe': { clientName: 'Jane Doe' } },
    supervisors: { sup1: { id: 'sup1', name: 'Pat Reed' } },
    signatures: [{ id: 'sig_1', label: 'Owner', data: 'data:image/png;base64,AAA' }],
  });
  await new Promise((r) => setTimeout(r, 30));
  assert.ok(w.getSupervisors().sup1, 'the supervisor must be restored locally');
  assert.ok(posted.some((p) => /supervisor/i.test(p.url)),
    'the supervisor never reached the DB: ' + posted.map((p) => p.url).join(' | '));
  assert.ok(posted.some((p) => /\/signatures/.test(p.url)),
    'the dialog said the signature was restored while it stayed on this device only');
  const sig = posted.find((p) => /\/signatures/.test(p.url));
  assert.strictEqual(sig.body.data_url, 'data:image/png;base64,AAA',
    'the API field is data_url; sending the local `data` shape would store an empty signature');
});

// ── The server's explanation was thrown away ─────────────────────────────────────────────────
// The backend rejects an over-long invoice note with a message naming the length, and now returns
// 503 "nothing was saved" when Key Vault is down. Both were discarded: the owner saw "HTTP 400".
test('a rejected save shows the server’s reason, not just the status code', async () => {
  const w = loadApp(); resetStorage(w);
  w.saveProfilesLS({ 'Jane Doe': { clientName: 'Jane Doe', _dbId: 7, invoices: [
    { billingPeriod: '08/2026', savedAt: 'S1', status: 'draft', data: {} }] } });
  w.localStorage.setItem('lhca_id_map', JSON.stringify({ 'Jane Doe': 7 }));
  w.fetch = () => Promise.resolve({ ok: false, status: 400,
    json: () => Promise.resolve({ error: 'invoice_note is 1240 characters; the limit is 1000' }) });
  const err = await w.syncNewInvoices('Jane Doe', w.getProfiles()['Jane Doe']).then(() => null, (e) => e);
  assert.ok(err, 'the save must reject');
  assert.match(err.message, /1240 characters/,
    'the owner could not tell an over-long note from a server outage — both said "HTTP 400"');
});

test('a rejected client save shows the server’s reason too', async () => {
  const w = loadApp(); resetStorage(w);
  w.saveProfilesLS({ 'Jane Doe': { clientName: 'Jane Doe', _dbId: 7, invoices: [] } });
  w.fetch = () => Promise.resolve({ ok: false, status: 503,
    json: () => Promise.resolve({ error: 'Could not encrypt the SSN (key service unavailable). Nothing was saved — try again.' }) });
  const err = await Promise.resolve(w.saveProfileSP('Jane Doe', w.getProfiles()['Jane Doe'], true))
    .then(() => null, (e) => e);
  assert.ok(err && /Nothing was saved/.test(err.message),
    '"HTTP 503" gave no hint that the SSN never reached the database');
});

test('a server with no JSON body still gives the status', async () => {
  const w = loadApp(); resetStorage(w);
  w.saveProfilesLS({ 'Jane Doe': { clientName: 'Jane Doe', _dbId: 7, invoices: [] } });
  w.fetch = () => Promise.resolve({ ok: false, status: 502, json: () => Promise.reject(new Error('not json')) });
  const err = await Promise.resolve(w.saveProfileSP('Jane Doe', w.getProfiles()['Jane Doe'], true))
    .then(() => null, (e) => e);
  assert.match(err.message, /502/, 'the fallback must not swallow the failure');
});
