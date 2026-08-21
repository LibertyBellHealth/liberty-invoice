'use strict';
// Findings from the adversarial review of the #97-#111 diff. Every one of these was a REGRESSION
// introduced by a recent "fix" and shipped green, because the suite had no coverage for the paths.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

// ── 1-2. Send All aborted after caseworker #1 ────────────────────────────────────────────────
// `failedWorkers` was never declared and _doMonthlyEmailSendInner never returned, so the very first
// caseworker took the `else failedWorkers.push(...)` branch -> ReferenceError -> the catch pushed
// again -> uncaught, out of the loop. Caseworkers 2..N were never emailed, with no error shown, and
// #1's invoices were already flipped to Submitted so the run looked half-done.
function seedTwoCaseworkers(w) {
  resetStorage(w);
  w.spToken = 'tok';
  w.saveCaseworkersLS([{ id: 1, name: 'Worker One', email: 'one@mi.gov' },
                       { id: 2, name: 'Worker Two', email: 'two@mi.gov' }]);
  w.saveProfilesLS({
    'Client A': { clientName: 'Client A', status: 'active', caseworkerId: 1,
      invoices: [{ billingPeriod: '08/2026', status: 'draft', data: {} }] },
    'Client B': { clientName: 'Client B', status: 'active', caseworkerId: 2,
      invoices: [{ billingPeriod: '08/2026', status: 'draft', data: {} }] }
  });
  w.validateInvoiceForSend = () => [];
  w.isInvoiceableStatus = () => true;
  w.isCarrierClient = () => false;
  w.clientWasActiveInPeriod = () => true;
  w.showToast = () => {};
  w.previewMonthlyInvoices = () => {};
  w.updateStats = () => {};
  w.setTimeout = (fn) => { fn(); return 0; };            // skip the 3s inter-send throttle
  // showConfirm's async callback is not awaited by the caller, so hold it for the test to await.
  w._pendingConfirm = null;
  w.showConfirm = (msg, onOk) => { w._pendingConfirm = onOk(); return w._pendingConfirm; };
}

test('Send All reaches EVERY caseworker, not just the first', async () => {
  const w = loadApp(); seedTwoCaseworkers(w);
  const sentTo = [];
  w._doMonthlyEmailSend = async (email, wname) => { sentTo.push(wname); return { ok: true, sent: 1 }; };
  await w.sendAllCaseworkerEmails('08/2026');
  await w._pendingConfirm;
  assert.deepStrictEqual(sentTo.sort(), ['Worker One', 'Worker Two'],
    'a ReferenceError on the first caseworker used to abort the whole billing-day run');
});

test('Send All reports the caseworkers that failed instead of throwing', async () => {
  const w = loadApp(); seedTwoCaseworkers(w);
  const alerts = [];
  w.showAlert = (m) => alerts.push(String(m));
  w._doMonthlyEmailSend = async (email, wname) =>
    (wname === 'Worker Two' ? { ok: false, status: 403 } : { ok: true, sent: 1 });
  await w.sendAllCaseworkerEmails('08/2026');
  await w._pendingConfirm;
  const failMsg = alerts.find((a) => /NOT sent/.test(a)) || '';
  assert.match(failMsg, /Worker Two/, 'the failed caseworker must be named so the owner can re-send');
  assert.doesNotMatch(failMsg, /Worker One/, 'the caseworker that succeeded must not be listed as failed');
});

test('the batch send counts a caseworker as sent only when the send reported ok', async () => {
  const w = loadApp(); seedTwoCaseworkers(w);
  const toasts = [];
  w.showToast = (m) => toasts.push(String(m));
  w.showAlert = () => {};
  w._doMonthlyEmailSend = async () => ({ ok: false, status: 500 });
  await w.sendAllCaseworkerEmails('08/2026');
  await w._pendingConfirm;
  assert.match(toasts.join(' '), /0 caseworkers/, 'two failures must not be counted as sends');
});

// ── 3. First invoice from a DHS-1210 threw and saved nothing ─────────────────────────────────
// _createFirstInvoice was extracted from its caller for the proration prompt and kept referencing
// the caller's `p`. ReferenceError before saveProfilesLS AND before saveProfileSP, so the invoice
// existed only on a transient object -- the user saw no invoice and no error.
test('creating the first invoice from an authorization actually persists it', () => {
  const w = loadApp(); resetStorage(w);
  const auth = { effectiveDate: '01/15/2026', hours: '20', minutes: '0', rate: '5.50', tasks: [] };
  w.saveProfilesLS({ 'Jane Doe': { clientName: 'Jane Doe', status: 'active', authorization: auth, invoices: [] } });
  w.activeProfileName = 'Jane Doe';
  w.showAlert = () => {}; w.showConfirm = () => {};
  w.fetch = () => Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({}) });
  w.switchTab = () => {};   // DOM-only, and it runs AFTER the save this test is about
  const _realSP = w.saveProfileSP; w.saveProfileSP = () => Promise.resolve();
  try {
    const store = w.getProfiles();
    w._createFirstInvoice(store, store['Jane Doe'], auth, '01/2026', null);
    const saved = w.getProfiles()['Jane Doe'].invoices || [];
    assert.strictEqual(saved.length, 1, 'the invoice must reach localStorage, not just the in-memory copy');
    assert.strictEqual(saved[0].billingPeriod, '01/2026');
  } finally { w.saveProfileSP = _realSP; }   // loadApp caches one window: never leak a stub
});

test('the first invoice is pushed to the server too', () => {
  const w = loadApp(); resetStorage(w);
  const auth = { effectiveDate: '01/15/2026', hours: '20', minutes: '0', rate: '5.50', tasks: [] };
  w.saveProfilesLS({ 'Jane Doe': { clientName: 'Jane Doe', status: 'active', authorization: auth, invoices: [] } });
  w.activeProfileName = 'Jane Doe';
  w.showAlert = () => {};
  w.fetch = () => Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({}) });
  w.switchTab = () => {};   // DOM-only, and it runs AFTER the save this test is about
  let pushed = null;
  const _realSP2 = w.saveProfileSP; w.saveProfileSP = (n) => { pushed = n; return Promise.resolve(); };
  try {
    const store = w.getProfiles();
    w._createFirstInvoice(store, store['Jane Doe'], auth, '01/2026', null);
    assert.strictEqual(pushed, 'Jane Doe', 'the throw happened one statement before the server push');
  } finally { w.saveProfileSP = _realSP2; }
});

// ── 4. Prorated first month went negative on an OCR'd impossible date ────────────────────────
test('an impossible effective date is not prorated', () => {
  const w = loadApp();
  const a = { hours: '20', minutes: '0' };
  assert.strictEqual(w._proratedFirstMonth(a, '02/31/2026'), null,
    'inMonth-dd+1 went to -2, and the dialog offered to "Prorate to -2:-26"');
  assert.strictEqual(w._proratedFirstMonth(a, '02/29/2026'), null,
    '2026 is not a leap year: this produced a certified invoice with 0:00 billed');
});

test('a real mid-month start still prorates correctly', () => {
  const w = loadApp();
  const r = w._proratedFirstMonth({ hours: '20', minutes: '0' }, '01/16/2026');
  assert.strictEqual(r.daysServed, 16); assert.strictEqual(r.daysInMonth, 31);
  assert.strictEqual(r.hours, 10); assert.strictEqual(r.minutes, 19);  // 1200 * 16/31 = 619.35 -> 619
});

// ── 5. Bulk "Mark Active" authorization guard was dead code ──────────────────────────────────
// Array.from({...}) is [] for a plain object, so the guard never fired and bulk-activating a CHAMPS
// client with no DHS-1210 still worked -- the exact hole its own comment claims to close.
test('bulk Mark Active refuses a CHAMPS client with no authorization', () => {
  const w = loadApp(); resetStorage(w);
  w.saveProfilesLS({ 'No Auth': { clientName: 'No Auth', status: 'inactive' } });
  const alerts = [];
  w.showAlert = (m) => alerts.push(String(m));
  w.bulkSelected['No Auth'] = true;
  w.bulkSetClientStatus('active');
  assert.match(alerts.join(' '), /No Auth/, 'the guard must name the blocked client');
  assert.notStrictEqual(w.getProfiles()['No Auth'].status, 'active',
    'an Active client with no authorization gets picked up by invoice generation');
});

// ── 6. Restoring a MASKED backup destroyed real caregiver SSNs ───────────────────────────────
test('a masked SSN is never written back to the server as an SSN', async () => {
  const w = loadApp(); resetStorage(w);
  let body = null;
  w.fetch = (url, opt) => { body = JSON.parse((opt && opt.body) || '{}');
    return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({ id: 'cg1' }) }); };
  await Promise.resolve(w.saveCaregiverAPI('cg1', { name: 'Ann Lee', ssn: '***-**-1234' }, true)).catch(() => {});
  assert.ok(!('ssn' in body) || !body.ssn,
    'the backend stores what it is sent: this overwrote the encrypted SSN with the mask, irrecoverably');
});

test('a real SSN is still sent', async () => {
  const w = loadApp(); resetStorage(w);
  let body = null;
  w.fetch = (url, opt) => { body = JSON.parse((opt && opt.body) || '{}');
    return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({ id: 'cg1' }) }); };
  await Promise.resolve(w.saveCaregiverAPI('cg1', { name: 'Ann Lee', ssn: '123-45-6789' }, true)).catch(() => {});
  assert.strictEqual(body.ssn, '123-45-6789', 'the mask guard must not block a genuine SSN');
});

// ── 7. saveProfileSP wrote a stale id-map snapshot ───────────────────────────────────────────
// A restore saves every client in one synchronous pass; each resolver wrote back the map as it was
// at CALL time, so the last one to settle reverted its siblings to their dead ids. syncNewInvoices
// reads the map only -- so invoices were then billed against deleted client rows.
test('concurrent client saves do not clobber each other in the id map', async () => {
  const w = loadApp(); resetStorage(w);
  w.saveProfilesLS({ Alice: { clientName: 'Alice' }, Bob: { clientName: 'Bob' } });
  const ids = { Alice: 1001, Bob: 1002 };
  w.fetch = (url, opt) => {
    const b = JSON.parse((opt && opt.body) || '{}');
    return Promise.resolve({ ok: true, status: 200,
      json: () => Promise.resolve({ id: ids[b.client_name], row_version: 'V1' }) });
  };
  const p = w.getProfiles();
  await Promise.all([w.saveProfileSP('Alice', p.Alice, true), w.saveProfileSP('Bob', p.Bob, true)]
    .map((x) => Promise.resolve(x).catch(() => {})));
  const map = JSON.parse(w.localStorage.getItem('lhca_id_map') || '{}');
  assert.strictEqual(map.Alice, 1001, 'Bob’s resolver used to write back a snapshot without Alice');
  assert.strictEqual(map.Bob, 1002);
});

// ── 8. The inner send's RETURN CONTRACT ──────────────────────────────────────────────────────
// The batch loop reads _res.ok to decide sent-vs-failed. The inner returned undefined on BOTH
// branches, so every caseworker looked failed. The tests above stub _doMonthlyEmailSend, so they
// cannot see this -- these drive the real inner.
function seedInner(w) {
  resetStorage(w);
  if (!w.document.getElementById('page-invoice')) {
    w.document.body.insertAdjacentHTML('beforeend', '<div id="page-invoice" class="page"></div>');
  }
  w.saveProfilesLS({ 'Client A': { clientName: 'Client A',
    invoices: [{ billingPeriod: '08/2026', status: 'draft', data: {} }] } });
  w.loadInvoiceForCapture = async () => {};
  w.captureInvoicePDF = async () => 'BASE64PDF';
  w.markInvoiceSubmitted = () => {}; w.closeMonthlyInvModal = () => {};
  w.updateStats = () => {}; w.showToast = () => {}; w.showAlert = () => {};
}

test('a successful send REPORTS success to the batch caller', async () => {
  const w = loadApp(); seedInner(w);
  w.sendMailWithPDF = async () => ({ ok: true });
  const res = await w._doMonthlyEmailSendInner('one@mi.gov', 'Worker One', '08/2026',
    [{ name: 'Client A' }], 0, [], []);
  assert.ok(res && res.ok === true, 'returning undefined made the batch treat every send as failed');
  assert.strictEqual(res.sent, 1);
});

test('a failed send REPORTS the failure and its status', async () => {
  const w = loadApp(); seedInner(w);
  w.sendMailWithPDF = async () => ({ ok: false, status: 403, err: 'forbidden' });
  const res = await w._doMonthlyEmailSendInner('one@mi.gov', 'Worker One', '08/2026',
    [{ name: 'Client A' }], 0, [], []);
  assert.ok(res && res.ok === false, 'the batch must be able to tell a failure from a success');
  assert.strictEqual(res.status, 403);
});

test('an authorization with no billable time is never prorated', () => {
  const w = loadApp();
  assert.strictEqual(w._proratedFirstMonth({ hours: '0', minutes: '0' }, '01/16/2026'), null,
    'prorating nothing offers a certified invoice billing 0:00');
});
