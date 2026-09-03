'use strict';
// Every other roster delete in this app surfaces a failure with a retry (D10: a "deleted"
// caregiver must not reappear on the next sync). The signature delete was missed: it ignored the
// response entirely, while loadSignaturesAPI refills the list from the server — so a refused
// delete brought the signature back with no explanation, after a dialog that said
// "This cannot be undone."
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

const tick = () => new Promise((r) => setTimeout(r, 0));

function withSig(w) {
  resetStorage(w);
  w.localStorage.setItem('lhca_signatures', JSON.stringify([{ id: 'sig_1', label: 'Tommy', data_url: 'data:,' }]));
}

test('a refused signature delete is surfaced, not swallowed', async () => {
  const w = loadApp();
  withSig(w);
  const surfaced = [];
  const origFetch = w.fetch, origShow = w._showSaveStatus, origRender = w.renderSigSettings;
  w.fetch = () => Promise.resolve({ ok: false, status: 500, json: () => Promise.resolve({ error: 'nope' }) });
  w._showSaveStatus = (state, msg, onRetry) => surfaced.push({ state, msg, hasRetry: typeof onRetry === 'function' });
  w.renderSigSettings = () => {};
  try {
    w.doDeleteSig('sig_1');
    await tick(); await tick();
  } finally {
    w.fetch = origFetch; w._showSaveStatus = origShow; w.renderSigSettings = origRender;
  }

  assert.strictEqual(surfaced.length, 1, 'a 500 on the delete produced no visible failure');
  assert.strictEqual(surfaced[0].state, 'failed');
  assert.match(surfaced[0].msg, /signature/i);
  assert.ok(surfaced[0].hasRetry, 'the failure offered no retry, unlike every other roster delete');
});

test('a successful signature delete stays silent and removes it locally', async () => {
  const w = loadApp();
  withSig(w);
  const surfaced = [];
  const origFetch = w.fetch, origShow = w._showSaveStatus, origRender = w.renderSigSettings;
  w.fetch = () => Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({}) });
  w._showSaveStatus = (state) => surfaced.push(state);
  w.renderSigSettings = () => {};
  try {
    w.doDeleteSig('sig_1');
    await tick(); await tick();
  } finally {
    w.fetch = origFetch; w._showSaveStatus = origShow; w.renderSigSettings = origRender;
  }

  assert.strictEqual(JSON.stringify(surfaced), '[]', 'a successful delete reported a failure');
  assert.strictEqual(JSON.parse(w.localStorage.getItem('lhca_signatures')).length, 0);
});
