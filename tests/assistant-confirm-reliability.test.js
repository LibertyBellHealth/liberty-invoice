'use strict';
// The Assistant never writes without a confirmation — but the confirmation itself was unreliable.
// showConfirm is ONE modal that clones away the previous handler on every call, and the chat loop
// runs a turn's tool calls back to back, so a second update_client replaced the first's dialog: the
// owner saw only the last while the model was told both had been put to them, and the visible text
// could swap between reading and clicking. The callback also wrote back a roster snapshot taken
// before the dialog. And the off-roster recipient warning went only to the model.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

function app() {
  const w = loadApp();
  resetStorage(w);
  ['saveProfileSP', 'addAuditEntry', 'showToast', 'renderInfoPane'].forEach((f) => { w[f] = () => {}; });
  w._asstConfirmPending = false;   // module-level, and loadApp reuses the realm
  w.confirms = [];
  w.showConfirm = (msg, onOk, opts) => { w.confirms.push({ msg: msg, ok: onOk, opts: opts || {} }); };
  w.alerts = [];
  w.showAlert = (m) => { w.alerts.push(String(m)); };
  w.saveProfilesLS({
    'Jane Doe': { clientName: 'Jane Doe', phone: '111', clientStatus: 'active' },
    'Bob Roe': { clientName: 'Bob Roe', phone: '222', clientStatus: 'active' },
  });
  return w;
}
const upd = (w, client, field, value) =>
  w._asstUpdateClient({ client_name: client, field: field, value: value });

test('a second change is refused while the first confirmation is still open', () => {
  const w = app();
  const first = upd(w, 'Jane Doe', 'phone', '555-1234');
  assert.ok(first && first.opened, 'the first should open a dialog');
  const second = upd(w, 'Bob Roe', 'phone', '999-9999');
  assert.ok(second && second.error, 'the second must be refused, not silently replace the first');
  assert.match(second.error, /one at a time|already open/i);
  assert.strictEqual(w.confirms.length, 1, 'only one dialog may exist');
});

test('answering the first releases the lock so the next can be asked', () => {
  const w = app();
  upd(w, 'Jane Doe', 'phone', '555-1234');
  w.confirms[0].ok();
  const second = upd(w, 'Bob Roe', 'phone', '999-9999');
  assert.ok(second && second.opened, 'after answering, the next change may be put to the user');
});

test('cancelling also releases the lock', () => {
  const w = app();
  upd(w, 'Jane Doe', 'phone', '555-1234');
  assert.strictEqual(typeof w.confirms[0].opts.onCancel, 'function', 'cancel must clear the lock');
  w.confirms[0].opts.onCancel();
  assert.ok(upd(w, 'Bob Roe', 'phone', '999-9999').opened);
});

test('the write does not revert work done while the dialog was open', () => {
  const w = app();
  upd(w, 'Jane Doe', 'phone', '555-1234');
  // A client is added, and another edited, while the confirmation sits open. Build a FRESH object:
  // mutating the one the tool already holds would not distinguish a stale write from a fresh one.
  w.saveProfilesLS(JSON.parse(JSON.stringify({
    'Jane Doe': { clientName: 'Jane Doe', phone: '111', clientStatus: 'active' },
    'Bob Roe': { clientName: 'Bob Roe', phone: 'CHANGED', clientStatus: 'active' },
    'Carol Ng': { clientName: 'Carol Ng', phone: '333' },
  })));
  w.confirms[0].ok();
  const after = w.getProfiles();
  assert.strictEqual(after['Jane Doe'].phone, '555-1234', 'the approved change must apply');
  assert.ok(after['Carol Ng'], 'a client added mid-dialog must not be erased');
  assert.strictEqual(after['Bob Roe'].phone, 'CHANGED', 'a concurrent edit must not be reverted');
});

test('a client deleted while the dialog was open is reported, not resurrected', () => {
  const w = app();
  upd(w, 'Jane Doe', 'phone', '555-1234');
  w.saveProfilesLS(JSON.parse(JSON.stringify({
    'Bob Roe': { clientName: 'Bob Roe', phone: '222', clientStatus: 'active' },
  })));
  w.confirms[0].ok();
  assert.strictEqual(w.getProfiles()['Jane Doe'], undefined, 'must not be recreated by the write');
  assert.ok(w.alerts.some((m) => /no longer available/i.test(m)), 'the owner must be told');
});
