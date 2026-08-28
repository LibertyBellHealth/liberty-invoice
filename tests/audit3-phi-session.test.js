'use strict';
// A PHI document opened in its own tab is outside everything the sign-out wipe can reach: it is a
// separate document, so blanking inputs and raising the login wall does nothing to it. _phiWindows
// existed for exactly this, but only the caregiver task sheet was ever registered — the document
// lightbox fallback (SSN cards, licences, DHS authorizations), the signed agreement and both
// invoice PDF previews all called window.open directly and stayed open and readable afterwards.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

function trackOpens(w) {
  const opened = [];
  w.open = (url) => { const win = { url: String(url), closed: false, close() { this.closed = true; } };
                      opened.push(win); return win; };
  return opened;
}

test('_openPhiWindow registers the window so the wipe can reach it', () => {
  const w = loadApp(); resetStorage(w);
  const opened = trackOpens(w);
  w._phiWindows = [];
  w._openPhiWindow('blob:fake-ssn-card');
  assert.strictEqual(opened.length, 1);
  assert.strictEqual(w._phiWindows.length, 1, 'unregistered windows survive sign-out');
});

test('the document lightbox fallback opens a TRACKED window', () => {
  const w = loadApp(); resetStorage(w);
  const opened = trackOpens(w);
  w._phiWindows = [];
  // A .docx authorization: not an image and not a PDF, so it takes the new-tab fallback.
  w._docEditCtx = { docs: [{ url: 'https://blob/sas/authorization.docx', displayName: 'DHS-1210.docx' }] };
  w.openDocPreview(0);
  assert.strictEqual(opened.length, 1, 'it opened a tab');
  assert.strictEqual(w._phiWindows.length, 1,
    'this tab shows a client document and was never closed on sign-out');
});

test('clearPHIFromStorage closes every tracked PHI window', () => {
  const w = loadApp(); resetStorage(w);
  const opened = trackOpens(w);
  w._phiWindows = [];
  w._openPhiWindow('blob:ssn-card');
  w._openPhiWindow('blob:signed-agreement');
  w.clearPHIFromStorage();
  assert.ok(opened.every(x => x.closed), 'a PHI tab left open is PHI left on the screen');
  assert.strictEqual(w._phiWindows.length, 0);
});

test('no PHI document is opened with a bare window.open any more', () => {
  const fs = require('fs'), path = require('path');
  const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
  // Allowed: the helper itself, and the OneDrive folder link (a Microsoft page, not our PHI).
  const bare = (src.match(/window\.open\([^)]*\)/g) || [])
    .filter(m => !/window\.open\(url, target \|\| '_blank'\)/.test(m))
    .filter(m => !/info\.webUrl/.test(m));
  assert.deepStrictEqual(bare, [],
    'route it through _openPhiWindow so the sign-out wipe can close it: ' + bare.join(' | '));
});

test('sign-out raises the lock wall itself, not just via the MSAL popup', () => {
  const w = loadApp(); resetStorage(w);
  w.document.body.insertAdjacentHTML('beforeend',
    '<div id="loginWall" style="display:none"></div><div id="loginWallMsg"></div>' +
    '<button id="loginWallBtn"></button><button id="authBtn"></button><span id="authStatus"></span>');
  w.aiTrack = () => {};
  // The user dismisses / blocks the logout popup — the old code left the page as it was.
  w.msalInstance = { logoutPopup: () => { throw new Error('popup blocked'); } };
  try { w.signOut(); } catch (e) { /* the popup failing must not undo the lock */ }
  assert.strictEqual(w.document.getElementById('loginWall').style.display, 'flex',
    'rendered PHI stayed on screen when the logout popup did not complete');
});
