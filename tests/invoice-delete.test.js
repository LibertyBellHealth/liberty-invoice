'use strict';
// deleteInvoiceAPI is server-first on purpose: the local row is only removed once the server
// confirms, so a failed delete cannot leave an invoice gone locally but alive on the server, where
// the next sync resurrects it. Nothing tested that, and a 404 — the row already deleted elsewhere —
// was treated as a failure, leaving the invoice stuck locally behind a retry that could never work.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

function app(response) {
  const w = loadApp();
  resetStorage(w);
  w.aiTrack = () => {};
  w.failures = [];
  w._showSaveStatus = (state, label) => { w.failures.push(state + '|' + label); };
  w.requests = [];
  w.fetch = (url, opt) => {
    w.requests.push({ url: String(url), method: (opt || {}).method });
    return typeof response === 'function' ? response() : Promise.resolve(response);
  };
  return w;
}
const settle = () => new Promise((r) => setTimeout(r, 10));

test('a local-only invoice (no database id) is removed without calling the server', async () => {
  const w = app({ ok: true, status: 200 });
  let removed = false;
  w.deleteInvoiceAPI(null, 'Jane Doe', '08/2026', () => { removed = true; });
  assert.strictEqual(removed, true);
  assert.strictEqual(w.requests.length, 0, 'there is nothing on the server to delete');
});

test('the local row is removed only after the server confirms', async () => {
  const w = app({ ok: true, status: 200 });
  let removed = false;
  w.deleteInvoiceAPI(90, 'Jane Doe', '08/2026', () => { removed = true; });
  assert.strictEqual(removed, false, 'must not remove locally before the server answers');
  await settle();
  assert.strictEqual(removed, true);
  assert.strictEqual(w.requests[0].method, 'DELETE');
  assert.match(w.requests[0].url, /\/invoices\/90$/);
});

test('a server error leaves the invoice in place and surfaces the failure', async () => {
  const w = app({ ok: false, status: 500 });
  let removed = false;
  w.deleteInvoiceAPI(90, 'Jane Doe', '08/2026', () => { removed = true; });
  await settle();
  assert.strictEqual(removed, false,
    'removing locally after a failed delete lets the next sync resurrect it');
  assert.ok(w.failures.some((f) => /^failed\|Delete invoice/.test(f)), 'the owner must be told');
});

test('a 404 counts as deleted — the row is already gone', async () => {
  const w = app({ ok: false, status: 404 });
  let removed = false;
  w.deleteInvoiceAPI(90, 'Jane Doe', '08/2026', () => { removed = true; });
  await settle();
  assert.strictEqual(removed, true, 'otherwise the invoice is stuck locally forever');
  assert.strictEqual(w.failures.length, 0, 'and nothing should be reported as broken');
});

test('a network failure is surfaced, not swallowed', async () => {
  const w = app(() => Promise.reject(new Error('offline')));
  let removed = false;
  w.deleteInvoiceAPI(90, 'Jane Doe', '08/2026', () => { removed = true; });
  await settle();
  assert.strictEqual(removed, false);
  assert.ok(w.failures.some((f) => /^failed\|Delete invoice/.test(f)));
});
