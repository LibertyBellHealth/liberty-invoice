'use strict';
// A documents response was rendered into whatever pane existed when it arrived, with no check that
// it still belonged to the client on screen — and it overwrote the global _docEditCtx with ITS ids.
// A client with many documents lists far slower (the backend signs each blob URL serially), so
// opening client A, switching to B, then opening B's Documents let A's later response repaint B's
// pane. The extract button then wrote A's card details onto B, the email button sent A's document
// to B's caregiver, and delete removed A's file.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

function app() {
  const w = loadApp();
  resetStorage(w);
  w.saveProfilesLS({
    'Alice Adams': { clientName: 'Alice Adams', _dbId: 'A1' },
    'Bob Brown': { clientName: 'Bob Brown', _dbId: 'B2' },
  });
  if (!w.document.getElementById('hcDocList')) {
    w.document.body.insertAdjacentHTML('beforeend', '<div id="hcDocList"></div>');
  }
  w.document.getElementById('hcDocList').innerHTML = '';
  w._docEditCtx = null;
  return w;
}
const docsFor = (who) => [{ id: who + '-doc', name: who + ' authorization.pdf', category: 'authorization', url: 'https://x/' + who }];

// loadHcDocs does not return its promise, so let the fetch chain settle before asserting.
const settle = () => new Promise((r) => setTimeout(r, 0));

// Resolve a documents fetch for `forClient` while `activeProfileName` is someone else.
async function deliverLate(w, forClientDbId, whileViewing, docs) {
  w.activeProfileName = whileViewing;
  w.fetch = () => Promise.resolve({ ok: true, json: () => Promise.resolve(docs) });
  w.loadHcDocs(forClientDbId);
  await settle();
}

test('a late listing for another client is discarded, not painted', async () => {
  const w = app();
  w.activeProfileName = 'Bob Brown';
  w.renderHcDocList('B2', docsFor('Bob'));            // Bob's own list is on screen
  const before = w.document.getElementById('hcDocList').innerHTML;
  await deliverLate(w, 'A1', 'Bob Brown', docsFor('Alice'));   // Alice's arrives late
  const after = w.document.getElementById('hcDocList').innerHTML;
  assert.strictEqual(after, before, 'Alice\'s documents must not repaint Bob\'s pane');
  assert.ok(after.indexOf('Alice') === -1, 'Alice\'s filename must not appear under Bob');
});

test('the late listing does not hijack the document action context', async () => {
  const w = app();
  w.activeProfileName = 'Bob Brown';
  w.renderHcDocList('B2', docsFor('Bob'));
  await deliverLate(w, 'A1', 'Bob Brown', docsFor('Alice'));
  assert.strictEqual(w._docEditCtx.clientId, 'B2',
    'extract/email/delete would otherwise act on the other client');
});

test('the listing for the client actually on screen still renders', async () => {
  const w = app();
  await deliverLate(w, 'B2', 'Bob Brown', docsFor('Bob'));
  const html = w.document.getElementById('hcDocList').innerHTML;
  assert.ok(html.indexOf('Bob') !== -1, 'the current client\'s own documents must still list');
  assert.strictEqual(w._docEditCtx.clientId, 'B2');
});

test('a late ERROR is not painted into another client\'s pane either', async () => {
  const w = app();
  w.activeProfileName = 'Bob Brown';
  w.renderHcDocList('B2', docsFor('Bob'));
  const before = w.document.getElementById('hcDocList').innerHTML;
  w.fetch = () => Promise.resolve({ ok: false, status: 500 });
  w.loadHcDocs('A1');
  await settle();
  assert.strictEqual(w.document.getElementById('hcDocList').innerHTML, before,
    'an error for another client must not replace this client\'s list');
});
