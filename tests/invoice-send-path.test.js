'use strict';
// The invoice send path — money going to MDHHS under the owner's certification.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

// ── Auto-generate silently billed a full month for a mid-month start ─────────────────────────
// The manual path stops and ASKS full-month-or-prorate, because which is right depends on the case.
// Auto-generate billed the whole authorization: a 21 July start on 29:47 billed 29:47 where
// prorating gives 10:34 — over 19 hours, about $519, certified to MDHHS.
test('a client whose service starts mid-period is held back from auto-generate', () => {
  const w = loadApp(); resetStorage(w);
  const auth = { hours: '29', minutes: '47', effectiveDate: '07/21/2026', rate: '27.00', tasks: [] };
  w.saveProfilesLS({
    'Mid Month': { clientName: 'Mid Month', clientStatus: 'active', startDate: '2026-07-21',
                   authorization: auth, invoices: [] },
    'Full Month': { clientName: 'Full Month', clientStatus: 'active', startDate: '2025-01-01',
                    authorization: auth, invoices: [] },
  });
  const elig = w.findClientsEligibleForAutoGen('07/2026');
  const mid = elig.find((e) => e.name === 'Mid Month');
  const full = elig.find((e) => e.name === 'Full Month');
  assert.ok(mid && mid.partialMonth === true,
    'a partial first month must not be auto-billed in either direction — the owner chooses');
  assert.ok(full && !full.partialMonth, 'a client active all month is still auto-generated normally');
});

test('a start on the 1st is a full month, not a partial one', () => {
  const w = loadApp(); resetStorage(w);
  assert.strictEqual(w._startsInsidePeriod({ startDate: '2026-07-01' }, '07/2026'), false);
  assert.strictEqual(w._startsInsidePeriod({ startDate: '2026-07-21' }, '07/2026'), true);
  assert.strictEqual(w._startsInsidePeriod({ startDate: '2026-06-21' }, '07/2026'), false,
    'a start in a PREVIOUS month is a full month here');
  assert.strictEqual(w._startsInsidePeriod({ startDate: '' }, '07/2026'), false);
  assert.strictEqual(w._startsInsidePeriod({ startDate: 'not a date' }, '07/2026'), false);
});

// ── The batch could email a blank, signed, certified invoice ─────────────────────────────────
test('Send All skips a client with no invoice instead of emailing a blank certified form', async () => {
  const w = loadApp(); resetStorage(w);
  if (!w.document.getElementById('page-invoice')) {
    w.document.body.insertAdjacentHTML('beforeend', '<div id="page-invoice" class="page"></div>');
  }
  // Two clients ready per the preview, but one's invoice is gone by send time.
  w.saveProfilesLS({
    'Has Invoice': { clientName: 'Has Invoice',
      invoices: [{ billingPeriod: '08/2026', status: 'draft', data: { svcHH: '20', svcMM: '00' } }] },
    'No Invoice': { clientName: 'No Invoice', invoices: [] },
  });
  const captured = [];
  w.loadInvoiceForCapture = async (name) => { captured.push(name); };
  w.captureInvoicePDF = async () => 'BASE64';
  w.markInvoiceSubmitted = () => {}; w.closeMonthlyInvModal = () => {};
  w.updateStats = () => {}; w.showToast = () => {};
  const alerts = [];
  w.showAlert = (m) => alerts.push(String(m));
  w.sendMailWithPDF = async () => ({ ok: true });
  w.fetch = () => Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({}) });
  w.addAuditEntry = () => {}; w.logActivity = () => {};
  await w._doMonthlyEmailSendInner('cw@mi.gov', 'Worker One', '08/2026',
    [{ name: 'Has Invoice' }, { name: 'No Invoice' }], 0, [], []);
  assert.ok(!captured.includes('No Invoice'),
    'substituting {} produced a signed, certified, COMPLETELY BLANK MSA-1904');
  assert.ok(captured.includes('Has Invoice'), 'the good client must still be sent');
  assert.match(alerts.join(' '), /No Invoice/,
    'a silent skip would leave the owner believing that client was billed');
});

// ── sendEmail validated one invoice and sent a different one ─────────────────────────────────
// Validation ran against the SAVED invoice; the PDF is built from the live form 40 lines later. So
// an edit made after the last save was never examined — type 800 hours, hit Email Worker, and it
// certified 800 hours to MDHHS with no warning.
test('Email Worker validates what is on the form, not the last saved version', async () => {
  const w = loadApp(); resetStorage(w);
  ['clientName','billingPeriod','activeAgentEmail','worker'].forEach((id) => {
    if (!w.document.getElementById(id)) {
      w.document.body.insertAdjacentHTML('beforeend', '<input id="' + id + '">');
    }
  });
  w.document.getElementById('clientName').value = 'Jane Doe';
  w.document.getElementById('billingPeriod').value = '08/2026';
  w.document.getElementById('activeAgentEmail').value = 'cw@mi.gov';
  w.document.getElementById('worker').value = 'Worker One';
  // SAVED invoice is clean and within the authorization...
  w.saveProfilesLS({ 'Jane Doe': { clientName: 'Jane Doe', medicaidId: '1', worker: 'Worker One',
    caseworkerId: 1, authorization: { hours: '20', minutes: '0' },
    invoices: [{ billingPeriod: '08/2026', status: 'draft',
                 data: { svcHH: '20', svcMM: '00', grandHH: '20', grandMM: '00' } }] } });
  w.saveCaseworkersLS([{ id: 1, name: 'Worker One', email: 'cw@mi.gov', agency: 'MDHHS - Wayne', org: 'MDHHS' }]);
  w.saveSigsLS([{ id: 1, data: 'x' }]);
  // ...but the OPEN FORM now says 800 hours.
  w.captureFullInvoice = () => ({ svcHH: '800', svcMM: '00', grandHH: '800', grandMM: '00' });
  let shown = '';
  w.showConfirm = (msg, ok, opts) => { shown = String(msg); if (opts && opts.onCancel) opts.onCancel(); };
  w.showAlert = () => {};
  w.fetch = () => Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({}) });
  await w.sendEmail().catch(() => {});
  assert.match(shown, /exceeds the authorized/,
    'validating the SAVED invoice checked a different document than the one that reaches MDHHS');
});
