'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp } = require('./harness');

test('harness: app.js loads and its functions are defined', () => {
  const w = loadApp();
  for (const fn of ['parseDHS1210', '_authHM', '_parseHM', '_mdyToYmd', '_clientSig', 'getProfiles', 'saveProfilesLS']) {
    assert.strictEqual(typeof w[fn], 'function', `${fn} should be defined`);
  }
});
