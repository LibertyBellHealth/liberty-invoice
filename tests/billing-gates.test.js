'use strict';
// The BULK surfaces enforced "only Active clients are invoiced"; the three per-client entry points
// checked only for a carrier client. So an invoice for a Terminated client could be created and
// emailed to MDHHS one at a time.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

function quiet(w) {
  const alerts = [];
  w.showAlert = (m) => alerts.push(String(m));
  w.showPage = () => {}; w.showToast = () => {}; w.renderInvHistory = () => {};
  w.applyFullInvoice = () => {}; w.clearInvoiceForm = () => {}; w.rebuild = () => {};
  w.loadProfileIntoForm = () => {}; w.updateStats = () => {};
  w.fetch = () => Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({}) });
  return alerts;
}

for (const [status, label] of [['inactive', 'In Progress'], ['lost', 'Lost'], ['terminated', 'Terminated']]) {
  test(`navInvoice refuses a ${label} client`, () => {
    const w = loadApp(); resetStorage(w); const alerts = quiet(w);
    w.saveProfilesLS({ Jane: { clientName: 'Jane', clientStatus: status } });
    w.activeProfileName = 'Jane';
    w.navInvoice();
    assert.match(alerts.join(' '), /not invoiced|not an active client/i,
      `a ${label} client must not reach the invoice form — the bulk surfaces already refuse them`);
  });
}

test('navInvoice still opens for an Active client', () => {
  const w = loadApp(); resetStorage(w); const alerts = quiet(w);
  w.saveProfilesLS({ Jane: { clientName: 'Jane', clientStatus: 'active' } });
  w.activeProfileName = 'Jane';
  // The gate is what's under test. An Active client passes it and then walks into invoice-page DOM
  // this harness doesn't build — so a throw HERE means the gate let it through, which is the point.
  let threw = null;
  try { w.navInvoice(); } catch (e) { threw = e; }
  assert.ok(!alerts.some((a) => /not invoiced|not an active client/i.test(a)),
    'normal billing must not be blocked: ' + alerts.join(' | '));
  assert.ok(threw === null || /disabled|null/.test(String(threw.message)),
    'expected only missing-DOM fallout past the gate, got: ' + (threw && threw.message));
});

test('Email Worker refuses a terminated client', async () => {
  const w = loadApp(); resetStorage(w); const alerts = quiet(w);
  ['clientName','billingPeriod','activeAgentEmail','worker'].forEach((id) => {
    if (!w.document.getElementById(id)) w.document.body.insertAdjacentHTML('beforeend', '<input id="' + id + '">');
  });
  w.document.getElementById('clientName').value = 'Jane';
  w.document.getElementById('billingPeriod').value = '08/2026';
  w.saveProfilesLS({ Jane: { clientName: 'Jane', clientStatus: 'terminated' } });
  await w.sendEmail().catch(() => {});
  assert.match(alerts.join(' '), /not an Active client/i,
    'billing a terminated client for a post-termination month is a straight over-bill');
});

// ── Silence used to read as "checked and fine" ───────────────────────────────────────────────
test('an invoice with no DHS-1210 at all is flagged', () => {
  const w = loadApp(); resetStorage(w); w.saveSigsLS([{ id: 1, data: 'x' }]);
  const cw = { id: 1, name: 'W', email: 'w@mi.gov', agency: 'MDHHS - Wayne', org: 'MDHHS' };
  const issues = w.validateInvoiceForSend('Jane',
    { clientName: 'Jane', medicaidId: '1', caseworkerId: 1, worker: 'W' },
    { billingPeriod: '08/2026', status: 'draft', data: { svcHH: '99', svcMM: '00' } }, cw);
  assert.ok(issues.some((i) => /No DHS-1210 on file/.test(i)),
    '99 hours with nothing authorising it passed with no warning at all');
});

test('an authorization whose HOURS did not parse is NOT blocked', () => {
  const w = loadApp(); resetStorage(w); w.saveSigsLS([{ id: 1, data: 'x' }]);
  const cw = { id: 1, name: 'W', email: 'w@mi.gov', agency: 'MDHHS - Wayne', org: 'MDHHS' };
  const issues = w.validateInvoiceForSend('Jane',
    // A real DHS-1210 on file, but OCR missed the hours.
    { clientName: 'Jane', medicaidId: '1', caseworkerId: 1, worker: 'W',
      authorization: { effectiveDate: '01/01/2026', tasks: [{ task: 'Bathing' }] } },
    { billingPeriod: '08/2026', status: 'draft', data: { svcHH: '20', svcMM: '00' } }, cw);
  assert.strictEqual(issues.length, 0,
    'every issue blocks the batch send — blocking on a failed OCR read would stop legitimate ' +
    'billing for a client who IS authorised: ' + issues.join(' | '));
});
