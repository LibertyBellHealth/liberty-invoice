'use strict';
// "1 submitted invoice pending 30+ days — follow up on payment" did nothing when clicked. Its
// handler was openAllInvoicesModal("outstanding"), interpolated into a DOUBLE-quoted onclick
// attribute — so the attribute ended at the first inner quote and the browser was left with
// onclick="openAllInvoicesModal(", a syntax error. Reported from real use, 2026-09-03.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

function panel(w) {
  if (!w.document.getElementById('attentionPanel')) {
    w.document.body.insertAdjacentHTML('beforeend', '<div id="attentionPanel"></div>');
  }
  return w.document.getElementById('attentionPanel');
}
function rowsWithHandlers(w) {
  return [...panel(w).querySelectorAll('.attn-item')]
    .map((el) => el.getAttribute('onclick')).filter(Boolean);
}

test('every attention row\'s click handler is valid JavaScript', () => {
  const w = loadApp();
  resetStorage(w);
  const old = new Date(Date.now() - 45 * 86400000).toLocaleString();
  w.saveProfilesLS({
    'Jane Doe': { clientName: 'Jane Doe', clientStatus: 'active', startDate: '2026-01-01',
      invoices: [{ billingPeriod: '07/2026', status: 'submitted', savedAt: old, data: {} }] },
  });
  w.saveTodos([{ id: 't1', text: 'overdue', due: '2020-01-01', done: false }]);
  if (typeof w.renderAttentionPanel === 'function') w.renderAttentionPanel();
  const handlers = rowsWithHandlers(w);
  assert.ok(handlers.length > 0, 'no attention rows rendered — the test would prove nothing');
  handlers.forEach((h) => {
    assert.doesNotThrow(() => new Function(h), 'unclickable row: onclick="' + h + '"');
    assert.ok(/\)\s*$/.test(h.trim()), 'handler looks truncated: ' + h);
  });
});

test('the 30+ day row opens the outstanding invoices list', () => {
  const w = loadApp();
  resetStorage(w);
  const old = new Date(Date.now() - 45 * 86400000).toLocaleString();
  w.saveProfilesLS({ 'Jane Doe': { clientName: 'Jane Doe', clientStatus: 'active',
    startDate: '2026-01-01',
    invoices: [{ billingPeriod: '07/2026', status: 'submitted', savedAt: old, data: {} }] } });
  if (typeof w.renderAttentionPanel === 'function') w.renderAttentionPanel();
  const row = [...panel(w).querySelectorAll('.attn-item')]
    .find((el) => /pending 30\+ days/.test(el.textContent));
  assert.ok(row, 'the 30+ day row did not render');
  const h = row.getAttribute('onclick');
  assert.match(h, /openAllInvoicesModal/);
  assert.doesNotThrow(() => new Function(h), 'onclick="' + h + '"');
  // And it actually calls through with the filter.
  let called = null;
  w.openAllInvoicesModal = (f) => { called = f; };
  new w.Function(h).call(w);
  assert.strictEqual(called, 'outstanding');
});

test('a handler containing double quotes is escaped, not truncated', () => {
  const w = loadApp();
  const d = w.document.createElement('div');
  d.innerHTML = '<div onclick="' + String('f("x")').replace(/"/g, '&quot;') + '">x</div>';
  assert.strictEqual(d.firstChild.getAttribute('onclick'), 'f("x")');
});
