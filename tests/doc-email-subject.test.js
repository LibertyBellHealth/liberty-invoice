// Document emails put the client's name straight in the subject ("MSA-4676 Home Help Services
// Agreement — <client>"), and the non-4676 branch used the document's FILENAME, which is routinely
// the client's name too ("Delanor Simpson 4676.pdf"). Owner ruled 2026-09-01 that invoice subjects
// carry no client name; the same reasoning applies here — these also go to caseworkers.
const test = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');
const { loadApp } = require('./harness');
const w = loadApp();

test('the 4676 subject names the form, not the client', () => {
  assert.strictEqual(w._docEmailSubject({ category: 'agreement' }, true, false),
    'MSA-4676 Home Help Services Agreement');
  assert.strictEqual(w._docEmailSubject({ category: 'agreement' }, true, true),
    'Signed MSA-4676 Home Help Services Agreement');
});

test('a client-named file does not leak through the subject', () => {
  const d = { displayName: 'Delanor Simpson 4676.pdf', name: 'Delanor Simpson 4676.pdf', category: '' };
  ['', 'Signed '].forEach((_, i) => {
    const s = w._docEmailSubject(d, false, !!i);
    assert.ok(!/Delanor|Simpson/i.test(s), 'client name reached the subject: ' + s);
  });
});

test('an untagged document gets a neutral label rather than a filename', () => {
  assert.strictEqual(w._docEmailSubject({ displayName: 'Jane Doe SSN card.pdf' }, false, false),
    'Document');
});

test('no document subject interpolates the client name any more', () => {
  const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
  const subjects = [...src.matchAll(/\bsubject\s*=\s*([^;]+);/g)].map(m => m[1].trim());
  subjects.forEach(s => assert.ok(!/clientName/.test(s),
    'a subject still carries the client name: ' + s));
});
