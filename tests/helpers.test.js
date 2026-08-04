'use strict';
// Small but load-bearing helpers: HTML escaping (XSS defense — the audits found real XSS), and
// the status label gotcha where 'inactive' is shown as "In Progress".
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp } = require('./harness');

test('esc: escapes every HTML-significant character (XSS defense)', () => {
  const w = loadApp();
  assert.strictEqual(w.esc('<script>alert(1)</script>'), '&lt;script&gt;alert(1)&lt;/script&gt;');
  assert.strictEqual(w.esc('a & b'), 'a &amp; b');
  assert.strictEqual(w.esc('say "hi"'), 'say &quot;hi&quot;');
  assert.strictEqual(w.esc("it's"), 'it&#39;s');
  assert.strictEqual(w.esc(null), '', 'null -> empty string, no crash');
  // & must be escaped first so it doesn't double-escape the entities produced by the others
  assert.strictEqual(w.esc('<'), '&lt;');
  assert.ok(!w.esc('<img src=x onerror=alert(1)>').includes('<'), 'no raw < survives');
});

test('clientStatusLabel: the inactive -> "In Progress" gotcha holds', () => {
  const w = loadApp();
  assert.strictEqual(w.clientStatusLabel('inactive'), 'In Progress', 'stored inactive shows as In Progress');
  assert.strictEqual(w.clientStatusLabel('active'), 'Active');
  assert.strictEqual(w.clientStatusLabel('lost'), 'Lost');
  assert.strictEqual(w.clientStatusLabel(''), 'Active', 'blank defaults to Active');
});
