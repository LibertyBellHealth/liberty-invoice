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

test('escJsAttr: neutralizes JS-string breakout + HTML (the other XSS sink)', () => {
  const w = loadApp();
  // Used inside onclick="fn('<HERE>')" — a raw single quote would break out of the JS string.
  const out = w.escJsAttr("x');alert(1)//");
  assert.ok(!out.includes("'"), 'no bare single-quote survives (cannot break out of the JS string)');
  // Also HTML-escaped, so it is safe inside the attribute too.
  assert.ok(w.escJsAttr('<b>').includes('&lt;') && !w.escJsAttr('<b>').includes('<'), 'angle brackets escaped');
  assert.ok(!w.escJsAttr('a\nb\r\nc').match(/[\r\n]/), 'newlines stripped (cannot break the attribute)');
  assert.strictEqual(w.escJsAttr(null), '', 'null -> empty, no crash');
});

test('clientStatusLabel: the inactive -> "In Progress" gotcha holds', () => {
  const w = loadApp();
  assert.strictEqual(w.clientStatusLabel('inactive'), 'In Progress', 'stored inactive shows as In Progress');
  assert.strictEqual(w.clientStatusLabel('active'), 'Active');
  assert.strictEqual(w.clientStatusLabel('lost'), 'Lost');
  assert.strictEqual(w.clientStatusLabel(''), 'Active', 'blank defaults to Active');
});
