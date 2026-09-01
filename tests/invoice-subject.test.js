// Subject lines are the one part of an email that routinely sits unencrypted in server logs,
// backups and phone notification previews. The bulk send used a bare "INVOICE" for that reason,
// but the single-invoice button sent "Invoice August 2026 – <client name>", putting a client's
// name in the subject. Owner ruled 2026-09-01: always just INVOICE.
const test = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');

test('every invoice email subject is exactly "INVOICE"', () => {
  const subjects = [...src.matchAll(/var\s+subj\s*=\s*([^;]+);/g)].map(m => m[1].trim());
  const literal = subjects.filter(s => s === "'INVOICE'");
  const readFromForm = subjects.filter(s => s.includes('getElementById'));
  assert.strictEqual(literal.length, 2, 'expected both send paths to use a bare INVOICE, got: ' +
    JSON.stringify(subjects));
  // Anything else must be a user-typed subject, not a composed one.
  assert.strictEqual(literal.length + readFromForm.length, subjects.length,
    'a composed invoice subject reappeared: ' + JSON.stringify(subjects));
});

test('no invoice subject interpolates the client name or billing period', () => {
  const subjects = [...src.matchAll(/var\s+subj\s*=\s*([^;]+);/g)].map(m => m[1]);
  subjects.forEach(s => {
    assert.ok(!/\bcn\b|clientName|bpLabel|periodLabel/.test(s),
      'client/period must stay out of the subject: ' + s);
  });
});

test('the client and period are still stated in the email body', () => {
  assert.ok(/Attached is the invoice for <b>'\+esc\(cn\)\+'<\/b> for <b>'\+esc\(bpLabel\)/.test(src),
    'the single-invoice body should still name the client and period');
});
