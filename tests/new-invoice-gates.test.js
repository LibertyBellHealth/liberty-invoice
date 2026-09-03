'use strict';
// confirmNewInvoice is the per-client entry point to creating an invoice. Every bulk surface refuses
// to invoice a carrier client or a non-active one; these single-client entry points once checked
// only for carrier, so an invoice for a terminated client could be created and emailed one at a
// time. These pin the gates.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

function app(prof) {
  const w = loadApp();
  resetStorage(w);
  ['showPage', 'bc', 'navHome', 'navDetail', 'rebuild', 'applyStates', 'toggleComplex',
   'renderNotesPane', 'resetSigArea', 'updateStats'].forEach((f) => { w[f] = () => {}; });
  w.alerts = [];
  w.showAlert = (m) => { w.alerts.push(String(m)); };
  // An allowed client proceeds into the invoice form, so those fields must exist.
  const IDS = ['clientName','clientName2','medicaidId','worker','worker2','billingPeriod',
    'billingPeriod2','hourlyRate','billTo','svcHH','svcMM','cplxHH','cplxMM','p1HH','p1MM',
    'grandHH','grandMM','dateSubmitted','sigDate1','sigDate2'];
  if (!w.document.getElementById('invClientTag')) {
    w.document.body.insertAdjacentHTML('beforeend',
      ['newInvChoiceModal', 'invClientTag', 'saveInvoiceBtn', 'topbarActions', 'dupWarning']
        .map((id) => '<div id="' + id + '"></div>').join('') +
      IDS.map((i) => '<input id="' + i + '">').join('') +
      '<input type="checkbox" id="showComplex"><div id="complexSection"></div>' +
      '<div id="sigArea1"></div><div id="sigArea2"></div>');
  }
  w.saveProfilesLS({ 'Jane Doe': Object.assign({ clientName: 'Jane Doe' }, prof) });
  w.activeProfileName = 'Jane Doe';
  return w;
}
const refused = (w) => w.alerts.length > 0;

test('a managed-care client is refused — billing goes through the carrier', () => {
  const w = app({ clientStatus: 'active', program: 'carrier', carrier: 'Priority Health' });
  w.confirmNewInvoice('blank');
  assert.ok(refused(w));
  assert.match(w.alerts[0], /managed-care|carrier/i);
});

test('an In Progress client is refused, and told so in the words the app displays', () => {
  const w = app({ clientStatus: 'inactive' });
  w.confirmNewInvoice('blank');
  assert.ok(refused(w));
  assert.match(w.alerts[0], /In Progress/, 'must not say "inactive" — that is not what the app shows');
});

test('terminated and lost clients are refused', () => {
  ['terminated', 'lost'].forEach((st) => {
    const w = app({ clientStatus: st });
    w.confirmNewInvoice('blank');
    assert.ok(refused(w), st + ' was allowed to be invoiced');
  });
});

test('an active CHAMPS client is allowed through', () => {
  const w = app({ clientStatus: 'active', program: 'champs' });
  w.confirmNewInvoice('blank');
  assert.strictEqual(w.alerts.length, 0, 'a normal client must not be blocked: ' + w.alerts.join('|'));
});

test('a client with no status set is treated as active', () => {
  const w = app({});
  w.confirmNewInvoice('blank');
  assert.strictEqual(w.alerts.length, 0);
});
