'use strict';
// Doc-grid rendering + gating (bug-batch 2026-08-16). These would have caught the live bugs:
//   - the ✕ delete button generated a broken inline onclick when a filename contained an apostrophe
//     (category "Driver's License" prefixed into the name) → "missing ) after argument list";
//   - the ✉ email button wrongly showed on ID cards and the 🔍 extract button wrongly hid, because
//     the category was stored as the LABEL ("Driver's License") but gated by the KEY ("Drivers_License").
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp } = require('./harness');

function grid(w, docs) {
  const list = w.document.createElement('div');
  w.renderDocGrid(list, docs, { clientType: 'homecare', clientId: 5, refresh: function () {} });
  return list.innerHTML;
}

test('delete button is index-based — no filename interpolated into onclick', () => {
  const w = loadApp();
  // A name with an apostrophe is exactly what used to break the inline handler.
  const html = grid(w, [{ name: "Driver's_License__lic_abc123def.jpg", displayName: 'license.jpg', category: "Driver's License", url: '#' }]);
  assert.ok(html.includes('deleteDocAt(0)'), 'delete should call deleteDocAt(index)');
  assert.ok(!/deleteHcDoc\(/.test(html), 'must not interpolate a filename into the onclick');
  // No raw apostrophe from the name leaks into an onclick attribute (that was the syntax error).
  assert.ok(!/onclick="[^"]*Driver's/.test(html), 'no unescaped filename apostrophe in any onclick');
});

test('ID cards hide email + show extract — for BOTH the label and the key form of the category', () => {
  const w = loadApp();
  ["Driver's License", 'Drivers_License'].forEach(function (cat) {
    const html = grid(w, [{ name: 'x.jpg', displayName: 'x.jpg', category: cat, url: '#' }]);
    assert.ok(html.includes('extractCardFields(0)'), 'extract (🔍) should show on a license (' + cat + ')');
    assert.ok(!html.includes('emailDocToCaregiver('), 'email (✉) must be hidden on a license (' + cat + ')');
  });
});

test('a non-ID document (Other) shows email and no extract', () => {
  const w = loadApp();
  const html = grid(w, [{ name: 'auth.pdf', displayName: 'auth.pdf', category: 'Other', url: '#' }]);
  assert.ok(html.includes('emailDocToCaregiver(0)'), 'Other should allow emailing to a caregiver');
  assert.ok(!html.includes('extractCardFields('), 'Other should not offer card extraction');
});

test('_catKey resolves label, key, and free-text consistently', () => {
  const w = loadApp();
  assert.strictEqual(w._catKey("Driver's License"), 'Drivers_License');
  assert.strictEqual(w._catKey('Drivers_License'), 'Drivers_License');
  assert.strictEqual(w._catKey('Medicare Card'), 'Medicare_Card');
  assert.strictEqual(w._catKey('Something Custom'), 'Something Custom');
});
