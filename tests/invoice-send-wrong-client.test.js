'use strict';
// sendEmail read activeProfileName a dozen times AFTER the "Invoice Has Issues" dialog. That is a
// real gap — hashchange reassigns the global on Back/Forward with no click on the page — so the
// invoice snapshot could be filed under one client while another client's PDF was emailed.
//
// The record is captured before the dialog now. The name typed on the form is NOT the record's
// identity: it is what prints on the PDF, and it is editable, so an invoice reopened for a renamed
// client, or one given a middle initial for MDHHS, must still send.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

function app() {
  const w = loadApp();
  resetStorage(w);
  if (!w.document.getElementById('clientName')) {
    w.document.body.insertAdjacentHTML('beforeend',
      '<input id="clientName"><input id="billingPeriod2x"><input id="activeAgentEmail">' +
      '<input id="worker"><button id="sendEmailInvBtn"></button>');
  }
  w.alerts = [];
  w.showAlert = (m) => w.alerts.push(String(m));
  w.showToast = () => {};
  w.validateInvoiceForSend = () => [];          // no issues, so no dialog — the check is what's under test
  w.captureFullInvoice = () => ({ svcHH: '20' });
  w.pdfs = 0;
  w.captureInvoicePDF = async () => { w.pdfs++; return 'BASE64'; };
  w.spToken = 'token';
  // saveProfileSP posts the snapshot, and sendEmail's outer try/catch turns any throw into a
  // generic "Error generating PDF" — which looks exactly like the guard having fired. Stub it.
  w.fetch = () => Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({}) });
  w.sendMailGraph = async () => ({ ok: true });
  w.saveProfilesLS({
    Alice: { clientName: 'Alice', clientStatus: 'active', invoices: [] },
    Bob:   { clientName: 'Bob',   clientStatus: 'active', invoices: [] },
  });
  w.saveCaseworkersLS([{ id: 'cw1', name: 'Casey', email: 'casey@example.com' }]);
  w.document.getElementById('clientName').value = 'Alice';
  w.document.getElementById('activeAgentEmail').value = 'casey@example.com';
  w.document.getElementById('worker').value = 'Casey';
  return w;
}
const bp = (w, v) => {
  let el = w.document.getElementById('billingPeriod');
  if (!el) { w.document.body.insertAdjacentHTML('beforeend', '<input id="billingPeriod">'); el = w.document.getElementById('billingPeriod'); }
  el.value = v;
};

test('the open client changing across the dialog stops the send', async () => {
  const w = app();
  bp(w, '08/2026');
  w.activeProfileName = 'Alice';
  let pressOk;
  w.validateInvoiceForSend = () => ['Something to confirm'];
  w.showConfirm = (msg, ok) => { pressOk = ok; };

  const sending = w.sendEmail();
  w.activeProfileName = 'Bob';    // Back button reaches another client while the dialog is up
  pressOk();
  await sending;

  assert.strictEqual(w.pdfs, 0, 'no PDF may be built once the open client has changed');
  assert.match(w.alerts.join(' '), /open client changed/i, w.alerts.join(' | '));
  assert.strictEqual(w.getProfiles().Bob.invoices.length, 0, "Alice's invoice was filed under Bob");
});

test('the send proceeds normally when nothing changed', async () => {
  const w = app();
  bp(w, '08/2026');
  w.activeProfileName = 'Alice';

  await w.sendEmail();

  assert.strictEqual(w.pdfs, 1, 'the guard must not block a normal send');
  assert.strictEqual(w.getProfiles().Alice.invoices.length, 1);
});

// The name on the form is editable and the stored snapshot keeps whatever it was written with, so
// requiring it to match the profile key made every historic invoice of a renamed client unsendable.
test('an invoice whose form name differs from the record still sends', async () => {
  const w = app();
  bp(w, '08/2026');
  w.activeProfileName = 'Alice';
  w.document.getElementById('clientName').value = 'Alice M. Adams';   // middle initial for MDHHS

  await w.sendEmail();

  assert.strictEqual(w.pdfs, 1, 'a differing form name is not a wrong client: ' + w.alerts.join(' | '));
  assert.strictEqual(w.getProfiles().Alice.invoices.length, 1,
    'the snapshot belongs to the open record regardless of what is typed on the form');
});
