'use strict';
// Home Phone client field (added 2026-08-05). The subtle risk: if homePhone isn't part of the
// client dirty-tracking signature, a home-phone-only edit looks "unchanged" and saveProfileSP skips
// the POST — so the edit never reaches the DB. These assert the field is tracked and distinct.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp } = require('./harness');

test('_clientSig: a home-phone edit changes the signature (so the save fires)', () => {
  const w = loadApp();
  const base = { firstName: 'Jane', phone: '313-111-2222', homePhone: '' };
  const sig0 = w._clientSig(base);
  const sig1 = w._clientSig(Object.assign({}, base, { homePhone: '313-999-8888' }));
  assert.notStrictEqual(sig0, sig1, 'adding a home phone must register as a change');
});

test('_clientSig: home phone is independent of the primary phone', () => {
  const w = loadApp();
  const a = w._clientSig({ phone: '313-111-2222', homePhone: '248-000-1111' });
  const b = w._clientSig({ phone: '248-000-1111', homePhone: '313-111-2222' }); // swapped
  assert.notStrictEqual(a, b, 'home phone and primary phone are separate fields, not interchangeable');
});
