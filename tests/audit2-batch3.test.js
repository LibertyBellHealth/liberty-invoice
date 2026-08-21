'use strict';
// Regression tests for the third audit-fix batch. Written after the independent verification pass,
// which found that several earlier "silent failure" claims were wrong because trackSave already
// surfaced them — so these pin only behaviour that was genuinely broken.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

test('validateInvoiceForSend: rejects out-of-range and non-numeric billed time', () => {
  const w = loadApp(); resetStorage(w);
  const prof = { medicaidId: '123', worker: 'W', authorization: { hours: 20, minutes: 0 } };
  const cw = { agency: 'MDHHS', email: 'w@michigan.gov' };
  const mk = (d) => w.validateInvoiceForSend('Jane', prof, { status: 'draft', data: d }, cw);
  assert.ok(mk({ svcHH: '20', svcMM: '75' }).some(i => /minutes must be 0-59/.test(i)), '75 minutes rejected');
  assert.ok(mk({ svcHH: 'abc', svcMM: '00' }).some(i => /hours is not a plain number/.test(i)), 'non-numeric hours rejected');
  assert.ok(mk({ svcHH: '20', svcMM: 'xx' }).some(i => /minutes is not a plain number/.test(i)), 'non-numeric minutes rejected');
  assert.ok(mk({ cplxHH: '4', cplxMM: '99' }).some(i => /Complex care time minutes/.test(i)), 'the complex-care pair is checked too');
  assert.ok(mk({ svcHH: '999', svcMM: '00' }).some(i => /more than a full month/.test(i)), 'absurd hours rejected');
  assert.strictEqual(mk({ svcHH: '20', svcMM: '00' }).filter(i => /minutes|plain number|full month/.test(i)).length, 0,
    'a valid 20:00 raises no time issue');
});

test('validateInvoiceForSend: flags billing MORE than the authorization allows', () => {
  const w = loadApp(); resetStorage(w);
  const prof = { medicaidId: '123', worker: 'W', authorization: { hours: 20, minutes: 0 } };
  const cw = { agency: 'MDHHS', email: 'w@michigan.gov' };
  const over = w.validateInvoiceForSend('Jane', prof, { status: 'draft', data: { svcHH: '25', svcMM: '00' } }, cw);
  assert.ok(over.some(i => /exceeds the authorized/.test(i)), 'over-billing the authorization is flagged');
  const ok = w.validateInvoiceForSend('Jane', prof, { status: 'draft', data: { svcHH: '19', svcMM: '30' } }, cw);
  assert.ok(!ok.some(i => /exceeds/.test(i)), 'billing under the authorization is fine');
});

test('_asstUpdateClient: cannot set Active without an authorization or a start date', () => {
  const w = loadApp(); resetStorage(w);
  w.saveProfilesLS({ 'No Auth': { clientName: 'No Auth' } });
  const r = w._asstUpdateClient({ client_name: 'No Auth', field: 'status', value: 'active' });
  assert.match(String(r.error || ''), /DHS-1210/, 'the AI path enforces the same gate as the client pane');
  const r2 = w._asstUpdateClient({ client_name: 'No Auth', field: 'status', value: 'banana' });
  assert.match(String(r2.error || ''), /Status must be one of/, 'free-text status is rejected');
});

test('import: _clearSyncBaselines forces a restored backup to actually re-sync', () => {
  const w = loadApp(); resetStorage(w);
  const restored = { 'Jane Doe': { clientName: 'Jane Doe', _clientSynced: 'stale-client-sig',
                                   invoices: [{ dbId: 5, billingPeriod: '07/2026', _synced: 'stale-sig' },
                                              { dbId: 6, billingPeriod: '08/2026', _synced: 'stale-sig-2' }] } };
  w._clearSyncBaselines(restored);
  const p = restored['Jane Doe'];
  assert.strictEqual(p._clientSynced, undefined, 'client baseline cleared');
  assert.strictEqual(p.invoices[0]._synced, undefined, 'invoice baseline cleared — otherwise syncNewInvoices skips it');
  assert.strictEqual(p.invoices[1]._synced, undefined, 'every invoice, not just the first');
  assert.strictEqual(p.invoices[0].dbId, 5, 'the db id is KEPT so the row updates rather than duplicating');
  assert.doesNotThrow(() => w._clearSyncBaselines({ Bad: null }), 'malformed entries do not throw');
});

test('applyStates: marks never survive onto a day the month does not have', () => {
  const w = loadApp(); resetStorage(w);
  // The harness DOM already holds the grid scaffolding (svcBody/cplxBody + the All rows); wiping it
  // here would break rebuild(), so reuse it.
  w.rebuild(31);
  const svc = Array.from({ length: 31 }, () => Array.from({ length: w.SVC }, () => true));  // every day checked
  w.applyStates({ svc, cplx: [] });
  w.rebuild(30);                                                        // now a 30-day month
  w.applyStates({ svc, cplx: [] });
  const rows = w.document.getElementById('svcBody').querySelectorAll('tr');
  const day31 = rows[30];
  if (day31) {
    const marked = day31.querySelectorAll('td.mc.on').length;
    assert.strictEqual(marked, 0, 'day 31 must carry no marks in a 30-day month — it prints on the certified form');
  }
});

test('_pushSettings returns a promise so a failed sync can be reported', () => {
  const w = loadApp(); resetStorage(w);
  w._apiToken = 'tok';
  w.fetch = () => Promise.resolve({ ok: false, status: 500 });
  const p = w._pushSettings({ state_rate: '27.00' });
  assert.ok(p && typeof p.then === 'function', 'callers can observe the outcome (it used to return undefined)');
  return p.then(() => assert.fail('a 500 must reject'), (e) => assert.match(String(e.message), /500/));
});
