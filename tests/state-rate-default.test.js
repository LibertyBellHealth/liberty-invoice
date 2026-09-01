// stateRate() falls back to a hardcoded 27.00 when nothing is saved, and the Settings box is filled
// by that same function — so a never-saved rate and a saved 27.00 looked identical. The day MDHHS
// changes the rate, an unsynced device would keep certifying 27.00 with nothing to show it was a
// guess. The fallback stays (invoices must not break); it is now visible.
const test = require('node:test');
const assert = require('node:assert');
const { loadApp } = require('./harness');

test('with nothing saved, the rate is the default AND is reported as such', () => {
  const w = loadApp();
  w.localStorage.removeItem('lhca_state_rate');
  assert.strictEqual(w.stateRate(), '27.00');
  assert.strictEqual(w.stateRateIsDefault(), true);
});

test('once a rate is saved, it is used and no longer flagged', () => {
  const w = loadApp();
  w.localStorage.setItem('lhca_state_rate', '28.50');
  assert.strictEqual(w.stateRate(), '28.50');
  assert.strictEqual(w.stateRateIsDefault(), false);
});

test('a saved rate that equals the default is still NOT flagged', () => {
  const w = loadApp();
  w.localStorage.setItem('lhca_state_rate', '27.00');
  assert.strictEqual(w.stateRate(), '27.00');
  assert.strictEqual(w.stateRateIsDefault(), false, 'a deliberate 27.00 must read as confirmed');
});

test('whitespace only counts as unsaved', () => {
  const w = loadApp();
  w.localStorage.setItem('lhca_state_rate', '   ');
  assert.strictEqual(w.stateRateIsDefault(), true);
  assert.strictEqual(w.stateRate(), '27.00');
});

test('invoices still bill the flat state rate whatever the client record says', () => {
  const w = loadApp();
  w.localStorage.setItem('lhca_state_rate', '28.50');
  assert.strictEqual(w.clientInvoiceRate({ hourlyRate: '15.00' }), '28.50');
});
