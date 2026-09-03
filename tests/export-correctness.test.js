'use strict';
// Three export defects: a formula-executing CSV, a status the app never shows, and filenames dated
// in UTC so an evening export lands under tomorrow.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

const w = loadApp();
resetStorage(w);

test('a cell that would execute as a formula is neutralised', () => {
  const rows = [];
  w._asstTriggerDownload = (blob) => { rows.push(blob); };
  let csv = '';
  w.Blob = function (parts) { csv = String(parts[0]); return { parts: parts }; };
  w._asstDownloadCsv(['Name', 'Note'], [
    ['=HYPERLINK("http://evil","click")', 'ok'],
    ['+1-555-0100', '-2+3'],
    ['@SUM(A1)', 'plain'],
  ], 'x.csv');
  ['=HYPERLINK', '+1-555', '-2+3', '@SUM'].forEach((danger) => {
    const idx = csv.indexOf(danger);
    assert.ok(idx > 0 && "'\"".indexOf(csv[idx - 1]) !== -1,
      danger + ' is not prefixed — it executes when the file is opened: ' + csv.split('\r\n')[1]);
  });
});

test('ordinary values are not mangled', () => {
  let csv = '';
  w.Blob = function (parts) { csv = String(parts[0]); return {}; };
  w._asstTriggerDownload = () => {};
  w._asstDownloadCsv(['Name'], [['Jane Doe'], ['O\'Brien'], ['27.00']], 'x.csv');
  assert.ok(csv.indexOf('Jane Doe') !== -1 && csv.indexOf("'Jane") === -1);
  assert.ok(csv.indexOf('27.00') !== -1 && csv.indexOf("'27.00") === -1);
});

test('the roster export shows the status the app displays', () => {
  assert.strictEqual(w.clientStatusLabel('inactive'), 'In Progress');
  const src = require('node:fs').readFileSync(require('node:path').join(__dirname, '..', 'app.js'), 'utf8');
  const raw = src.match(/status:\(p\.clientStatus\|\|'active'\)/g) || [];
  assert.strictEqual(raw.length, 0,
    'an export still writes the raw status, which reads as dropped rather than onboarding');
});

test('export filenames use the local calendar day, not UTC', () => {
  const d = new Date();
  const local = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  assert.strictEqual(w._localYmd(), local);
  // The bug in one line: after 8pm Michigan time these differ.
  const evening = new Date('2026-09-03T21:30:00-04:00');
  assert.strictEqual(w._ymdOf(evening), '2026-09-03');
  assert.strictEqual(evening.toISOString().slice(0, 10), '2026-09-04');
});

test('no export path still dates itself in UTC', () => {
  const src = require('node:fs').readFileSync(require('node:path').join(__dirname, '..', 'app.js'), 'utf8');
  assert.strictEqual((src.match(/toISOString\(\)\.slice\(0,\s*10\)/g) || []).length, 0);
});
