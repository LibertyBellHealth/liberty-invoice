// The import review modal used to promise "Nothing was sent anywhere — parsed in your browser" on
// EVERY import. A scanned authorization has no text layer, so it is uploaded to Document
// Intelligence to be read — the notice was telling the owner their client's PHI stayed on the
// device when it had not. Reported 2026-09-01 from a real scanned MDHHS-6064.
const test = require('node:test');
const assert = require('node:assert');
const { loadApp } = require('./harness');

function render(res) {
  const w = loadApp();
  w.activeProfileName = 'Test Client';
  w.showDhsReview({ name: 'MSA-6064.pdf' }, res);
  const box = w.document.getElementById('dhsReviewModal');
  assert.ok(box, 'review modal did not render');
  return box.textContent;
}

const BASE = { formType: 'MDHHS-6064', warnings: [], tasks: [], hours: 73, minutes: 44 };

test('a scanned import says the file was uploaded to be read', () => {
  const t = render({ ...BASE, viaOcr: true });
  assert.ok(/uploaded to your agency's Azure Document Intelligence/.test(t),
    'expected the upload disclosure, got: ' + t.slice(0, 300));
  assert.ok(!/Nothing was sent anywhere/.test(t),
    'a scanned import must NOT claim nothing was sent');
});

test('a text-layer PDF still says it was parsed locally', () => {
  const t = render({ ...BASE });
  assert.ok(/Nothing was sent anywhere — parsed in your browser/.test(t));
  assert.ok(!/Document Intelligence/.test(t));
});

test('either way the notice names the file it read', () => {
  assert.ok(render({ ...BASE, viaOcr: true }).includes('MSA-6064.pdf'));
  assert.ok(render({ ...BASE }).includes('MSA-6064.pdf'));
});
