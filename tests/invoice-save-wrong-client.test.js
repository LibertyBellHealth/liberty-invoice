'use strict';
// saveInvoiceToClient reads activeProfileName, opens an "Overwrite Invoice?" dialog, and
// _doSaveInvoiceToClient then re-reads the global at write time. The dialog is a real gap — the
// hash router reassigns the global on Back/Forward with no click on the page — so the invoice on
// screen for one client was written onto another, along with its audit row.
//
// It also left the two buttons on the invoice page disagreeing: Email captures the client before
// the dialog, Save trusted the global.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

function app() {
  const w = loadApp();
  resetStorage(w);
  ['billingPeriod', 'saveInvoiceBtn', 'dupWarning'].forEach((id) => {
    if (!w.document.getElementById(id)) {
      w.document.body.insertAdjacentHTML('beforeend',
        id === 'billingPeriod' ? '<input id="billingPeriod">' : '<div id="' + id + '"></div>');
    }
  });
  w.document.getElementById('billingPeriod').value = '08/2026';
  w.audits = [];
  w.addAuditEntry = (name, action) => w.audits.push(name + ' | ' + action);
  w.captureFullInvoice = () => ({ svcHH: '20', marker: 'FORM' });
  w.saveProfileSP = () => Promise.resolve();
  w.logActivity = () => {}; w.updateStats = () => {}; w.aiTrack = () => {};
  w.showAlert = () => {};
  w.saveProfilesLS({
    Alice: { clientName: 'Alice', invoices: [{ billingPeriod: '08/2026', status: 'draft', data: { marker: 'OLD-A' } }] },
    Bob:   { clientName: 'Bob',   invoices: [{ billingPeriod: '08/2026', status: 'draft', data: { marker: 'OLD-B' } }] },
  });
  return w;
}
const invOf = (w, who) => w.getProfiles()[who].invoices.find((i) => i.billingPeriod === '08/2026');

test('an invoice saved for one client is not written onto whoever you reach mid-dialog', () => {
  const w = app();
  w.activeProfileName = 'Alice';
  let pressOk;
  w.showConfirm = (msg, ok) => { pressOk = ok; };

  w.saveInvoiceToClient();          // overwrite dialog opens for Alice
  w.activeProfileName = 'Bob';      // Back button reaches Bob while it is up
  pressOk();

  assert.strictEqual(invOf(w, 'Bob').data.marker, 'OLD-B',
    "Alice's on-screen invoice was written onto Bob");
  assert.strictEqual(invOf(w, 'Alice').data.marker, 'FORM',
    "the invoice belongs to the client it was opened for");
  assert.ok(w.audits.join(' ').includes('Alice'), 'the audit row must name Alice: ' + w.audits.join(' | '));
  assert.ok(!w.audits.join(' ').includes('Bob'), 'and must not name Bob');
});

test('a normal overwrite still works', () => {
  const w = app();
  w.activeProfileName = 'Alice';
  let pressOk;
  w.showConfirm = (msg, ok) => { pressOk = ok; };

  w.saveInvoiceToClient();
  pressOk();

  assert.strictEqual(invOf(w, 'Alice').data.marker, 'FORM');
  assert.strictEqual(invOf(w, 'Bob').data.marker, 'OLD-B');
});
