'use strict';
// SIX of these were introduced by my own fixes this week. Each is pinned here so the next "fix"
// cannot quietly undo it.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

const CW = { id: 1, name: 'Worker One', email: 'w@mi.gov', agency: 'MDHHS - Wayne', org: 'MDHHS' };
function prof(auth) { return { clientName: 'Jane', medicaidId: '1', worker: 'Worker One',
  caseworkerId: 1, program: '', authorization: auth }; }
function inv(data) { return { billingPeriod: '08/2026', status: 'draft', data: data }; }

// ── The over-bill check was reading a HIDDEN field ───────────────────────────────────────────
test('an over-bill typed into the VISIBLE Total Time is caught', () => {
  const w = loadApp(); resetStorage(w); w.saveSigsLS([{ id: 1, data: 'x' }]);
  // grandHH lives inside #complexSection (display:none) and keeps its stale generated value;
  // svcHH is the field the owner edits and the one page 1 of the certified form prints.
  const issues = w.validateInvoiceForSend('Jane', prof({ hours: '20', minutes: '0' }),
    inv({ svcHH: '40', svcMM: '00', grandHH: '20', grandMM: '00' }), CW);
  assert.ok(issues.some((i) => /exceeds the authorized/.test(i)),
    'checking grand alone let a 20-hour over-bill reach MDHHS with no warning');
});

test('a complex-care total that exceeds the authorization is still caught', () => {
  const w = loadApp(); resetStorage(w); w.saveSigsLS([{ id: 1, data: 'x' }]);
  const issues = w.validateInvoiceForSend('Jane', prof({ hours: '20', minutes: '0' }),
    inv({ svcHH: '15', svcMM: '00', cplxHH: '10', cplxMM: '00', grandHH: '25', grandMM: '00' }), CW);
  assert.ok(issues.some((i) => /exceeds the authorized/.test(i)), 'the complex-care case must not regress');
});

test('complex care with a blank billing-period total is flagged', () => {
  const w = loadApp(); resetStorage(w); w.saveSigsLS([{ id: 1, data: 'x' }]);
  const issues = w.validateInvoiceForSend('Jane', prof({ hours: '40', minutes: '0' }),
    inv({ svcHH: '15', svcMM: '00', cplxHH: '10', cplxMM: '00', grandHH: '', grandMM: '' }), CW);
  assert.ok(issues.some((i) => /Total Time for Billing Period" is blank/.test(i)),
    'the form certifies svc + cplx across two pages while grand is blank');
});

test('an invoice within the authorization is still clean', () => {
  const w = loadApp(); resetStorage(w); w.saveSigsLS([{ id: 1, data: 'x' }]);
  const issues = w.validateInvoiceForSend('Jane', prof({ hours: '20', minutes: '0' }),
    inv({ svcHH: '20', svcMM: '00', grandHH: '20', grandMM: '00' }), CW);
  assert.strictEqual(issues.length, 0, 'must not cry wolf: ' + issues.join(' | '));
});

// ── The task dirty flag, for the THIRD time ──────────────────────────────────────────────────
test('a failed task edit survives a load using the REAL syncedProp', () => {
  const w = loadApp(); resetStorage(w);
  // Production passes 'dbId', and a locally created task keeps id='td_...' until a full reload —
  // so it is NOT in the server list by id and falls through the syncedProp drop.
  const server = [{ id: '42', dbId: 42, text: 'OLD server text' }];
  const local = [{ id: 'td_1700000000_ab', dbId: 42, text: 'EDIT that failed to save', _unsaved: true }];
  const merged = w._mergeByIdKeepUnsynced(server, local, 'dbId');
  assert.ok(merged.some((t) => t && t._unsaved && /EDIT that failed/.test(t.text)),
    'my first two attempts at this flag both missed the branch that actually fires in-session');
});

test('a synced task deleted on another device still propagates', () => {
  const w = loadApp(); resetStorage(w);
  const merged = w._mergeByIdKeepUnsynced([], [{ id: 'td_1', dbId: 9, text: 'deleted elsewhere' }], 'dbId');
  assert.strictEqual(merged.length, 0, 'the delete-propagation behaviour must not regress');
});

// ── Copy-from-last was inheriting last month's rate ──────────────────────────────────────────
test('the assistant refuses an impossible date', () => {
  const w = loadApp(); resetStorage(w);
  w.saveProfilesLS({ Jane: { clientName: 'Jane' } });
  ['13/05/2026', '02/31/2026', '2026-99-99'].forEach((bad) => {
    const r = w._asstUpdateClient({ client_name: 'Jane', field: 'dob', value: bad });
    assert.ok(r && r.error, bad + ' parsed "successfully" but <input type="date"> still rejects it, ' +
      'so the next save blanked the field — the exact bug the guard exists to stop');
  });
});

test('the assistant still accepts a real date', () => {
  const w = loadApp(); resetStorage(w);
  w.saveProfilesLS({ Jane: { clientName: 'Jane' } });
  let shown = '';
  w.showConfirm = (m) => { shown = String(m); };
  w._asstUpdateClient({ client_name: 'Jane', field: 'dob', value: '02/29/2024' });   // real leap day
  assert.match(shown, /2024-02-29/, 'a valid leap day must not be rejected');
});
