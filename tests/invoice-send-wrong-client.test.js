'use strict';
// sendEmail had TWO independent sources of truth for who the invoice is for. `cn` comes from the
// invoice form, which is what captureInvoicePDF builds the document from — so it is the name the
// caseworker reads on the PDF. activeProfileName is a global, read a dozen times AFTER the
// "Invoice Has Issues" dialog, which is a real gap: the hash router reassigns the global on
// Back/Forward with no click on the page, and the overlay does not block it.
//
// So the invoice snapshot was persisted under one client while the first client's PDF was emailed.
// When two sources disagree there is no right one to pick, and this path sends a document out of
// the building — so it stops.
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

test('an invoice is not emailed when the open client is not the one on the form', async () => {
  const w = app();
  bp(w, '08/2026');
  w.activeProfileName = 'Bob';        // reached another client while the send was being set up

  await w.sendEmail();

  assert.strictEqual(w.pdfs, 0, 'no PDF may be built when the two sources disagree');
  assert.match(w.alerts.join(' '), /different client is open/i,
    'expected the wrong-client refusal, not some other failure: ' + w.alerts.join(' | '));
  assert.match(w.alerts.join(' '), /Alice/, 'the refusal must name the invoice’s client');
  assert.ok(!w.getProfiles().Bob.invoices.length,
    "Alice's invoice snapshot was filed under Bob");
});

test('the send proceeds normally when they agree', async () => {
  const w = app();
  bp(w, '08/2026');
  w.activeProfileName = 'Alice';

  await w.sendEmail();

  assert.strictEqual(w.pdfs, 1, 'the guard must not block a normal send');
  assert.strictEqual(w.getProfiles().Alice.invoices.length, 1,
    'the snapshot should be filed under the client on the invoice');
});

// The refusal above only fires when the form HAS a client name. With a blank one the old code still
// filed a snapshot under whoever happened to be open — an invoice nobody named, saved onto a real
// client's record. Using the form as the source of truth means a nameless invoice files nowhere.
test('an invoice with no client name on the form is not filed under the open client', async () => {
  const w = app();
  bp(w, '08/2026');
  w.document.getElementById('clientName').value = '';
  w.activeProfileName = 'Bob';

  await w.sendEmail();

  assert.strictEqual(w.getProfiles().Bob.invoices.length, 0,
    'a snapshot from a form naming nobody was written onto Bob');
});
