// markInvoiceSubmitted flipped the invoice to Submitted and pushed it to the server, then threw the
// result away. It runs immediately after the invoice has been EMAILED, so a failed write leaves the
// invoice Submitted here and Draft on the server — another device (or this one after reloading from
// the server) offers it as unsent, and it goes to MDHHS twice. The email cannot be recalled, so the
// failure has to be visible.
const test = require('node:test');
const assert = require('node:assert');
const { loadApp } = require('./harness');

function setup() {
  const w = loadApp();
  w.saveProfilesLS({ 'Jane Doe': { clientName: 'Jane Doe', invoices: [
    { billingPeriod: '08/2026', status: 'draft', data: { svcHH: '20', svcMM: '00' } }] } });
  return w;
}

test('a successful write marks the invoice submitted', () => {
  const w = setup();
  w.saveProfileSP = () => Promise.resolve({ ok: true });
  w.markInvoiceSubmitted('Jane Doe', '08/2026');
  assert.strictEqual(w.getProfiles()['Jane Doe'].invoices[0].status, 'submitted');
});

test('a FAILED write is surfaced, naming the client and period', async () => {
  const w = setup();
  const shown = [];
  w._showSaveStatus = (state, label) => shown.push(state + '|' + label);
  w.saveProfileSP = () => Promise.reject(new Error('503'));
  w.markInvoiceSubmitted('Jane Doe', '08/2026');
  await new Promise(r => setTimeout(r, 10));
  assert.strictEqual(shown.length, 1, 'expected one surfaced failure, got ' + JSON.stringify(shown));
  assert.ok(/^failed\|/.test(shown[0]), shown[0]);
  assert.ok(shown[0].includes('Jane Doe') && shown[0].includes('08/2026'), shown[0]);
  // It must say the mail already went out — that is why this cannot be ignored.
  assert.ok(/EMAILED/.test(shown[0]), shown[0]);
});

test('a synchronous (non-promise) return does not blow up', () => {
  const w = setup();
  w.saveProfileSP = () => undefined;
  assert.doesNotThrow(() => w.markInvoiceSubmitted('Jane Doe', '08/2026'));
  assert.strictEqual(w.getProfiles()['Jane Doe'].invoices[0].status, 'submitted');
});

test('an already-paid invoice is left alone', () => {
  const w = loadApp();
  w.saveProfilesLS({ 'Jane Doe': { clientName: 'Jane Doe', invoices: [
    { billingPeriod: '08/2026', status: 'paid', data: {} }] } });
  let called = false;
  w.saveProfileSP = () => { called = true; return Promise.resolve({ ok: true }); };
  w.markInvoiceSubmitted('Jane Doe', '08/2026');
  assert.strictEqual(w.getProfiles()['Jane Doe'].invoices[0].status, 'paid');
  assert.strictEqual(called, false);
});
