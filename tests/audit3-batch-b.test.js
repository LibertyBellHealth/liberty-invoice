'use strict';
// Batch B of the adversarial diff review: the findings that cost money or silently stop syncing.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

const CW = { id: 1, name: 'Worker One', email: 'one@mi.gov', agency: 'Wayne DHHS' };
function invWith(data) { return { billingPeriod: '08/2026', status: 'draft', data: data }; }
function profWith(auth) {
  return { clientName: 'Jane Doe', medicaidId: '1234567', worker: 'Worker One',
           caseworkerId: 1, authorization: auth };
}

// ── Over-billing MDHHS ───────────────────────────────────────────────────────────────────────
// The cross-check compared SERVICE time to the authorization, but the number certified on the form
// is the billing-period GRAND total (service + complex care). Nothing else in the app compares
// these two numbers, and over-billing is what triggers a recoupment.
test('a complex-care invoice over the authorization is caught', () => {
  const w = loadApp(); resetStorage(w);
  w.saveSigsLS([{ id: 1, data: 'x' }]);
  const issues = w.validateInvoiceForSend('Jane Doe', profWith({ hours: '20', minutes: '0' }),
    invWith({ svcHH: '15', svcMM: '00', cplxHH: '10', cplxMM: '00', grandHH: '25', grandMM: '00' }), CW);
  assert.ok(issues.some((i) => /exceeds the authorized/.test(i)),
    'billed 25:00 against a 20:00 authorization — a 5-hour over-bill that used to go to MDHHS unflagged');
});

test('a service-only invoice over the authorization is still caught', () => {
  const w = loadApp(); resetStorage(w);
  w.saveSigsLS([{ id: 1, data: 'x' }]);
  const issues = w.validateInvoiceForSend('Jane Doe', profWith({ hours: '20', minutes: '0' }),
    invWith({ svcHH: '25', svcMM: '00' }), CW);
  assert.ok(issues.some((i) => /exceeds the authorized/.test(i)), 'the grand-blank fallback must still check svc');
});

test('an invoice within the authorization is clean', () => {
  const w = loadApp(); resetStorage(w);
  w.saveSigsLS([{ id: 1, data: 'x' }]);
  const issues = w.validateInvoiceForSend('Jane Doe', profWith({ hours: '20', minutes: '0' }),
    invWith({ svcHH: '15', svcMM: '00', cplxHH: '05', cplxMM: '00', grandHH: '20', grandMM: '00' }), CW);
  assert.ok(!issues.some((i) => /exceeds the authorized/.test(i)), 'exactly at the authorization is allowed');
});

test('the previous-page total is sanity-checked like every other time field', () => {
  const w = loadApp(); resetStorage(w);
  w.saveSigsLS([{ id: 1, data: 'x' }]);
  const issues = w.validateInvoiceForSend('Jane Doe', profWith(null),
    invWith({ svcHH: '20', svcMM: '00', p1HH: 'abc', p1MM: '999' }), CW);
  assert.ok(issues.some((i) => /Previous-page total/.test(i)),
    'p1HH/p1MM print on the certified form but were the one pair never validated');
});

// ── Minutes printed as decimals ──────────────────────────────────────────────────────────────
// The HTML/raster renderer copied the minute inputs verbatim, so a typed "5" printed as "20.5"
// (= 20h50m) instead of "20.05". captureInvoicePDF falls back to this path automatically.
test('the raster/print renderer zero-pads the minute fields', () => {
  const w = loadApp(); resetStorage(w);
  w.document.body.insertAdjacentHTML('beforeend',
    '<div id="page1"><input id="svcHH" value="20"><input id="svcMM" value="5">' +
    '<input id="grandHH" value="20"><input id="grandMM" value="5">' +
    '<input id="clientName" value="Jane Doe"></div>' +
    '<div id="complexSection"><table id="cplxTable"></table></div>' +
    '<input type="checkbox" id="showComplex">');
  w.document.getElementById('svcMM').value = '5';
  w.document.getElementById('grandMM').value = '5';
  const html = w.buildInvoiceHTML();
  assert.match(html, /id="svcMM"[^>]*value="05"/,
    'unpadded, "20" + "." + "5" reads as 20h50m on a form certified to MDHHS');
  assert.match(html, /id="grandMM"[^>]*value="05"/);
  assert.match(html, /id="svcHH"[^>]*value="20"/, 'hours are not padded');
  assert.match(html, /id="clientName"[^>]*value="Jane Doe"/, 'non-minute fields are untouched');
});

// ── The restore that stopped every other client syncing ──────────────────────────────────────
test('clearing sync baselines touches ONLY the imported clients', () => {
  const w = loadApp(); resetStorage(w);
  const store = {
    Imported: { clientName: 'Imported', _clientSynced: 'sigA', invoices: [{ _synced: 'i1' }] },
    Untouched: { clientName: 'Untouched', _clientSynced: 'sigB', invoices: [{ _synced: 'i2' }] },
  };
  w._clearSyncBaselines(store, ['Imported']);
  assert.strictEqual(store.Imported._clientSynced, undefined, 'the restored client re-syncs');
  assert.strictEqual(store.Imported.invoices[0]._synced, undefined);
  assert.strictEqual(store.Untouched._clientSynced, 'sigB',
    'wiping this pinned every other client to its LOCAL copy forever — another device stopped arriving');
  assert.strictEqual(store.Untouched.invoices[0]._synced, 'i2');
});

// ── A real failure must outrank a concurrent-edit conflict ───────────────────────────────────
test('when one invoice 409s and another genuinely fails, the hard failure wins', async () => {
  const w = loadApp(); resetStorage(w);
  w.saveProfilesLS({ 'Jane Doe': { clientName: 'Jane Doe', _dbId: 7, invoices: [
    { billingPeriod: '07/2026', savedAt: 'S1', status: 'draft', data: {} },
    { billingPeriod: '08/2026', savedAt: 'S2', status: 'draft', data: {} }] } });
  w.localStorage.setItem('lhca_id_map', JSON.stringify({ 'Jane Doe': 7 }));
  let n = 0;
  w.fetch = () => {
    n++;
    if (n === 1) return Promise.resolve({ ok: false, status: 409,
      json: () => Promise.resolve({ error: 'conflict', id: 99, row_version: 'AB' }) });
    return Promise.resolve({ ok: false, status: 500, json: () => Promise.resolve({}) });
  };
  const err = await w.syncNewInvoices('Jane Doe', w.getProfiles()['Jane Doe']).then(() => null, (e) => e);
  assert.ok(err, 'the save must reject');
  assert.ok(!err.isConflict,
    'a conflict-flagged error clears the dirty flag, leaving the genuinely-failed invoice unprotected');
});

test('a 409 does not clear a dirty flag an earlier failure set', async () => {
  const w = loadApp(); resetStorage(w);
  w.saveProfilesLS({ 'Jane Doe': { clientName: 'Jane Doe', _dbId: 7, invoices: [] } });
  w._markClientUnsaved('Jane Doe', true);                       // an earlier genuine failure
  assert.ok(w.getProfiles()['Jane Doe']._unsaved, 'precondition: the edit is protected');
  w.fetch = () => Promise.resolve({ ok: false, status: 409,
    json: () => Promise.resolve({ error: 'changed by someone else' }) });
  await Promise.resolve(w.saveProfileSP('Jane Doe', w.getProfiles()['Jane Doe'], true)).catch(() => {});
  await new Promise((r) => setTimeout(r, 5));
  assert.ok(w.getProfiles()['Jane Doe']._unsaved,
    'clearing on a 409 discarded the earlier edit at the next background load');
});
