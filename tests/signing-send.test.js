'use strict';
// doSendForSignature creates a real signing request on the server — token, expiry, audit row — and
// only then emails the link. It validated the template and the date of birth but never checked that
// the caregiver HAS an email, so a caregiver without one produced a dangling request and an error
// that read as though the document had gone out. downloadSignedDoc is covered here too: it is the
// front end of the backend hash check fixed today, and had no test at all.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

function app(cg) {
  const w = loadApp();
  resetStorage(w);
  ['showToast', 'loadCgSigningRequests', 'closeSendSigModal', 'buildSigningEmail']
    .forEach((f) => { w[f] = () => ({ subject: 's', html: 'h' }); });
  w.alerts = [];
  w.showAlert = (m) => { w.alerts.push(String(m)); };
  w.calls = [];
  w.fetch = (url, opt) => { w.calls.push({ url: String(url), opt: opt });
    return Promise.resolve({ ok: true, status: 200,
      json: () => Promise.resolve({ signUrl: 'https://x/sign/1', expiresAt: '2026-09-17' }) }); };
  w.saveCaregiversLS({ cg1: Object.assign({ name: 'Sam Carer' }, cg || {}) });
  w.activeCgId = 'cg1';
  // loadApp reuses one jsdom document — build these ONCE, then reset them per test, or
  // getElementById keeps returning the first test's stale nodes.
  if (!w.document.getElementById('sendSigBtn')) {
    // The modal element itself, because doSendForSignature now takes the caregiver from the dialog
    // that names them rather than from the global. openSendForSignatureModal stamps it on creation.
    w.document.body.insertAdjacentHTML('beforeend',
      '<div id="sendSigModal"><button id="sendSigBtn"></button><div id="sendSigError"></div>' +
      '<select id="sendSigTemplate"><option value="7" selected>MSA-4676</option></select>' +
      '<input id="sendSigDob"></div>');
  }
  w.document.getElementById('sendSigModal').dataset.cgId = 'cg1';
  w.document.getElementById('sendSigDob').value = '1950-02-01';
  w.document.getElementById('sendSigError').textContent = '';
  w.document.getElementById('sendSigTemplate').value = '7';
  return w;
}
const err = (w) => w.document.getElementById('sendSigError').textContent;

test('a caregiver with no email is refused BEFORE any request is created', async () => {
  const w = app({ email: '' });
  await w.doSendForSignature();
  assert.match(err(w), /No email address on file/i);
  assert.strictEqual(w.calls.length, 0, 'no signing request may be created that cannot be delivered');
});

test('a whitespace-only email counts as missing', async () => {
  const w = app({ email: '   ' });
  await w.doSendForSignature();
  assert.match(err(w), /No email address on file/i);
  assert.strictEqual(w.calls.length, 0);
});

test('the refusal names the caregiver so it is actionable', async () => {
  const w = app({ email: '' });
  await w.doSendForSignature();
  assert.match(err(w), /Sam Carer/);
});

test('a missing date of birth is still refused, and before any request', async () => {
  const w = app({ email: 'sam@example.com' });
  w.document.getElementById('sendSigDob').value = '';
  await w.doSendForSignature();
  assert.match(err(w), /date of birth/i);
  assert.strictEqual(w.calls.length, 0);
});

test('a caregiver with an email does reach the send endpoint, with the DOB', async () => {
  const w = app({ email: 'sam@example.com' });
  w.spToken = null;                       // stop before the Graph email step
  await w.doSendForSignature();
  const call = w.calls.find((c) => /\/signing\/send$/.test(c.url));
  assert.ok(call, 'the request should have been created: ' + JSON.stringify(w.calls.map((c) => c.url)));
  const body = JSON.parse(call.opt.body);
  assert.strictEqual(body.recipientEmail, 'sam@example.com');
  assert.strictEqual(body.recipientDob, '1950-02-01', 'identity verification depends on this');
});

test('a refused signed-document download is surfaced, not swallowed', async () => {
  const w = app({ email: 'sam@example.com' });
  w.fetch = () => Promise.resolve({ ok: false, status: 409,
    json: () => Promise.resolve({ error: 'This signed document no longer matches the hash recorded when it was signed.' }) });
  await w.downloadSignedDoc(7);
  assert.strictEqual(w.alerts.length, 1, 'the owner must be told why nothing opened');
  assert.match(w.alerts[0], /no longer matches the hash/i);
});

test('a good signed-document download opens the returned URL', async () => {
  const w = app({ email: 'sam@example.com' });
  w.opened = null;
  w._openPhiWindow = (u) => { w.opened = u; };
  w.fetch = () => Promise.resolve({ ok: true, status: 200,
    json: () => Promise.resolve({ url: 'https://blob/signed.pdf?sig=x' }) });
  await w.downloadSignedDoc(7);
  assert.strictEqual(w.opened, 'https://blob/signed.pdf?sig=x');
  assert.strictEqual(w.alerts.length, 0);
});


// The gap here is the operator's own CLICK: openSendForSignatureModal builds the dialog for one
// caregiver, doSendForSignature runs whenever Send is pressed, and activeCgId can change in between
// with no click on the page — the hash router reassigns it on Back/Forward and the overlay does not
// block it. Neither function alone can see that gap; the modal has to carry the identity.
test('a link is not emailed to whoever is open when the dialog names someone else', async () => {
  const w = app({ email: 'sam@example.com' });
  w.saveCaregiversLS({
    cg1: { name: 'Sam Carer', email: 'sam@example.com' },
    cg2: { name: 'Dana Other', email: 'dana@example.com' },
  });
  w.activeCgId = 'cg2';               // operator reached another caregiver while the dialog was up

  await w.doSendForSignature();

  assert.strictEqual(w.calls.length, 0, 'no signing request may be created for a drifted record');
  assert.match(err(w), /Sam Carer/, 'the refusal must name who the link was for: ' + err(w));
  assert.match(err(w), /different caregiver is open/i);
});

test('the send still goes through for the caregiver the dialog was opened for', async () => {
  const w = app({ email: 'sam@example.com' });
  await w.doSendForSignature();
  const send = w.calls.find((c) => /\/signing\/send$/.test(c.url));
  assert.ok(send, 'the happy path must still reach the send endpoint: ' + JSON.stringify(w.calls.map((c) => c.url)));
  assert.strictEqual(JSON.parse(send.opt.body).caregiverId, 'cg1',
    'the request must name the caregiver the dialog was built for');
});

// The tests above stand the modal up by hand, so none of them would notice if the app stopped
// stamping the caregiver onto it — and the stamp is the whole mechanism. Build it the real way.
test('opening the dialog stamps the caregiver it was built for onto the modal', async () => {
  const w = app({ email: 'sam@example.com' });
  w.showAlert = (m) => w.alerts.push(String(m));
  w.fetch = (url) => Promise.resolve({ ok: true, status: 200,
    json: () => Promise.resolve([{ id: 7, name: 'MSA-4676', is_active: true }]) });
  w.activeCgId = 'cg1';

  await w.openSendForSignatureModal();

  const mdl = w.document.getElementById('sendSigModal');
  assert.ok(mdl, 'the modal was not created: ' + w.alerts.join(' | '));
  assert.strictEqual(mdl.dataset.cgId, 'cg1',
    'without the stamp the Send handler has nothing to act on but the global');
  mdl.remove();   // leave the document as the hand-built scaffold found it
});

// openSendForSignatureModal awaits a templates fetch BEFORE it builds the modal, and no overlay is
// on screen during that fetch — the whole page is clickable. The body was built from a caregiver
// captured before the await while the stamp was re-read from the global after it, so the dialog
// named one person and its stamp named another. The Send guard then compared that stamp against
// the same drifted global, agreed with itself, and sent to the wrong caregiver.
test('a caregiver clicked while the templates load does not hijack the dialog', async () => {
  const w = app({ email: 'sam@example.com' });
  w.saveCaregiversLS({
    cg1: { name: 'Sam Carer', email: 'sam@example.com' },
    cg2: { name: 'Dana Other', email: 'dana@example.com' },
  });
  w.activeCgId = 'cg1';
  let release;
  w.fetch = () => new Promise((res) => { release = () => res({ ok: true, status: 200,
    json: () => Promise.resolve([{ id: 7, name: 'MSA-4676', is_active: true }]) }); });

  const opening = w.openSendForSignatureModal();
  w.activeCgId = 'cg2';          // operator clicks another caregiver while templates are loading
  release();
  await opening;

  const mdl = w.document.getElementById('sendSigModal');
  assert.ok(/Sam Carer/.test(mdl.textContent), 'the body is built from the pre-await caregiver');
  assert.strictEqual(mdl.dataset.cgId, 'cg1',
    'the stamp must name the same caregiver the dialog names, not whoever was clicked mid-fetch');
  mdl.remove();
});
