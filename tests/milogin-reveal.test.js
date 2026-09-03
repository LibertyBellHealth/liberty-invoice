'use strict';
// The MI Login password is a stored credential fetched on demand. The field it lands in
// (cg-milogin-pass) is a STATIC element in index.html, reused for every caregiver and merely
// cleared between them — so a fetch started for caregiver A that resolves after the owner has
// opened caregiver B wrote A's password into B's visible field, and saveCaregiver reads that field
// straight into B's record. Same shape as the document-list bug, with a secret instead of a file.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

function app() {
  const w = loadApp();
  resetStorage(w);
  if (!w.document.getElementById('cg-milogin-pass')) {
    w.document.body.insertAdjacentHTML('beforeend',
      '<input id="cg-milogin-pass" type="password"><input id="cg-editing-id">' +
      '<button id="revealBtn">Show</button>');
  }
  const f = w.document.getElementById('cg-milogin-pass');
  f.value = ''; f.type = 'password';
  w.document.getElementById('cg-editing-id').value = '';
  w.toggleMask = () => {};
  w._revealSecret = () => {};
  w.activeCgId = '';
  w.resolvers = [];
  w.fetch = () => new Promise((res) => { w.resolvers.push(res); });
  return w;
}
const field = (w) => w.document.getElementById('cg-milogin-pass');
const openCaregiver = (w, id) => { w.document.getElementById('cg-editing-id').value = id; w.activeCgId = id; };
const respond = (w, pw) => {
  const res = w.resolvers.shift();
  res({ ok: true, json: () => Promise.resolve({ milogin_password: pw }) });
  return new Promise((r) => setTimeout(r, 0));
};

test("a password arriving after the owner switched caregivers is NOT written", async () => {
  const w = app();
  openCaregiver(w, 'cg_A');
  w._revealMiloginField(field(w), 'cg_A');       // reveal started for A
  openCaregiver(w, 'cg_B');                       // owner opens B while it is in flight
  field(w).value = '';                            // the form was cleared for B
  await respond(w, 'AlicePassword123');
  assert.strictEqual(field(w).value, '',
    "A's credential must not land in B's form — saveCaregiver reads this field");
});

test('the password IS written when the same caregiver is still open', async () => {
  const w = app();
  openCaregiver(w, 'cg_A');
  w._revealMiloginField(field(w), 'cg_A');
  await respond(w, 'AlicePassword123');
  assert.strictEqual(field(w).value, 'AlicePassword123');
});

test('the Show button path is guarded the same way', async () => {
  const w = app();
  openCaregiver(w, 'cg_A');
  w.revealMilogin('cg-milogin-pass', w.document.getElementById('revealBtn'), 'cg_A');
  openCaregiver(w, 'cg_B');
  field(w).value = '';
  await respond(w, 'AlicePassword123');
  assert.strictEqual(field(w).value, '', "A's credential must not appear under B");
});

test('the Show button still reveals for the caregiver on screen', async () => {
  const w = app();
  openCaregiver(w, 'cg_A');
  w.revealMilogin('cg-milogin-pass', w.document.getElementById('revealBtn'), 'cg_A');
  await respond(w, 'AlicePassword123');
  assert.strictEqual(field(w).value, 'AlicePassword123');
});

test('a value the owner has typed is never overwritten by a late response', async () => {
  const w = app();
  openCaregiver(w, 'cg_A');
  w._revealMiloginField(field(w), 'cg_A');
  field(w).value = 'TypedByHand';
  await respond(w, 'FromServer');
  assert.strictEqual(field(w).value, 'TypedByHand');
});

test('a new caregiver with no id fetches nothing', () => {
  const w = app();
  w._revealMiloginField(field(w), '');
  assert.strictEqual(w.resolvers.length, 0, 'there is no stored credential to fetch yet');
});
