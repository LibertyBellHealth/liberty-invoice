'use strict';
// The signing request actions: resend (emails a NEW link to the recipient on file), revoke (kills
// the current link immediately), and the audit view. All three touch PHI or invalidate access, and
// none had a test.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

function app() {
  const w = loadApp();
  resetStorage(w);
  ['loadCgSigningRequests', 'showToast', 'logActivity'].forEach((f) => { w[f] = () => {}; });
  w.alerts = [];
  w.showAlert = (m) => { w.alerts.push(String(m)); };
  w.pending = null;
  w.showConfirm = (msg, onOk) => { w.pending = { msg: msg, ok: onOk }; };
  w.mails = [];
  w.sendMailWithPDF = (to, subj, body) => { w.mails.push({ to: to, subj: subj, body: body });
    return Promise.resolve({ ok: true }); };
  w.buildSigningEmail = (o) => ({ subject: 'Reminder — please sign ' + (o.docName || ''), html: '<p>' + (o.signUrl || '') + '</p>' });
  w.spToken = 'tok';
  w.saveCaregiversLS({ cg1: { name: 'Sam Carer', email: 'stale@old.example' } });
  w.activeCgId = 'cg1';
  w.requests = [];
  return w;
}
const settle = () => new Promise((r) => setTimeout(r, 0));

test('resend emails the recipient RECORDED ON THE REQUEST, not the local caregiver record', async () => {
  const w = app();
  w.fetch = (url, opt) => { w.requests.push({ url: String(url), method: (opt || {}).method });
    return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({
      recipientName: 'Sam Carer', recipientEmail: 'sam@current.example',
      signUrl: 'https://x/sign/abc', expiresAt: '2026-09-17', templateName: 'MSA-4676' }) }); };
  await w.resendSigningRequest(7);
  assert.strictEqual(w.mails.length, 1);
  assert.strictEqual(w.mails[0].to, 'sam@current.example',
    'the request is the source of truth for who it was sent to');
});

test('resend without a Microsoft sign-in refuses before hitting the server', async () => {
  const w = app();
  w.spToken = null;
  w.fetch = () => { throw new Error('should not be called'); };
  await w.resendSigningRequest(7);
  assert.match(w.alerts.join(' '), /Sign in with Microsoft/i);
});

test('a server refusal on resend is surfaced with its reason', async () => {
  const w = app();
  w.fetch = () => Promise.resolve({ ok: false, status: 410,
    json: () => Promise.resolve({ error: 'This request was already completed.' }) });
  await w.resendSigningRequest(7);
  assert.match(w.alerts.join(' '), /already completed/i);
  assert.strictEqual(w.mails.length, 0, 'nothing may be emailed when the server refused');
});

test('a failed email still gives the owner the link to send by hand', async () => {
  const w = app();
  w.fetch = () => Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({
    recipientName: 'Sam Carer', recipientEmail: 'sam@current.example',
    signUrl: 'https://x/sign/abc', expiresAt: '2026-09-17' }) });
  w.sendMailWithPDF = () => Promise.resolve({ ok: false, status: 403, err: 'no mailbox' });
  await w.resendSigningRequest(7);
  assert.match(w.alerts.join(' '), /https:\/\/x\/sign\/abc/,
    'the link must be recoverable, or the request is stranded');
});

test('revoking asks first and only then calls the server', async () => {
  const w = app();
  w.fetch = (url, opt) => { w.requests.push({ url: String(url), method: (opt || {}).method });
    return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({}) }); };
  w.revokeSigningRequest(7);
  assert.ok(w.pending, 'revoking kills a live link — it must be confirmed');
  assert.match(w.pending.msg, /stop working immediately/i);
  assert.strictEqual(w.requests.length, 0, 'nothing before the owner confirms');
  w.pending.ok();
  await settle();
  assert.match(w.requests[0].url, /\/signing\/7\/revoke$/);
  assert.strictEqual(w.requests[0].method, 'POST');
});

test('a failed revoke is surfaced — the link is still live', async () => {
  const w = app();
  w.fetch = () => Promise.resolve({ ok: false, status: 500, json: () => Promise.resolve({}) });
  w.revokeSigningRequest(7);
  w.pending.ok();
  await settle();
  assert.ok(w.alerts.length > 0, 'silence here means the owner believes a live link is dead');
});
