'use strict';
// Reactive auth-expiry handling (persistence audit 2026-08-05). After a device sleeps/wakes the AAD
// token can be expired while the UI still looks signed-in; the next API call 401s silently. These
// assert the visible banner appears on a 401 and clears on a later success — so a worker isn't
// typing into saves that quietly fail.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp } = require('./harness');

test('_showAuthExpiredBanner: creates ONE visible "Sign in again" banner, idempotent', () => {
  const w = loadApp();
  w._hideAuthExpiredBanner();                                  // reset shared state
  w._showAuthExpiredBanner();
  w._showAuthExpiredBanner();                                  // second call must not duplicate
  const bars = w.document.querySelectorAll('#authExpiredBar');
  assert.strictEqual(bars.length, 1, 'exactly one banner element');
  const bar = bars[0];
  assert.notStrictEqual(bar.style.display, 'none', 'banner is visible');
  assert.ok(/sign in again/i.test(bar.innerHTML), 'offers a re-auth action');
  assert.ok(/may not have saved/i.test(bar.innerHTML), 'warns a change may not have saved');
});

test('_hideAuthExpiredBanner: hides it and lets it re-show later', () => {
  const w = loadApp();
  w._showAuthExpiredBanner();
  w._hideAuthExpiredBanner();
  assert.strictEqual(w.document.getElementById('authExpiredBar').style.display, 'none', 'hidden');
  w._showAuthExpiredBanner();                                  // must be able to re-appear
  assert.notStrictEqual(w.document.getElementById('authExpiredBar').style.display, 'none', 're-shown');
  w._hideAuthExpiredBanner();
});

test('_onApiAuthFail: surfaces the banner even with no MSAL instance, without throwing', () => {
  const w = loadApp();
  w._hideAuthExpiredBanner();
  assert.doesNotThrow(() => w._onApiAuthFail());
  assert.notStrictEqual(w.document.getElementById('authExpiredBar').style.display, 'none',
    'a 401 always makes the expiry visible');
  w._hideAuthExpiredBanner();
});
