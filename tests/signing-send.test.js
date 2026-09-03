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
    w.document.body.insertAdjacentHTML('beforeend',
      '<button id="sendSigBtn"></button><div id="sendSigError"></div>' +
      '<select id="sendSigTemplate"><option value="7" selected>MSA-4676</option></select>' +
      '<input id="sendSigDob">');
  }
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
