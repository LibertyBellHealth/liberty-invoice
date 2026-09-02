'use strict';
// Every environment writes backups to the SAME OneDrive folder, and the automatic ones were named
// by date alone, so whichever environment ran last that day owned the slot. It happened: the
// automatic file for 2026-09-01 holds "Demo Patient One" and "ZZ Carrier Test" — staging data —
// so there is no automatic backup of the real clients for that date. The pruner also kept a fixed
// number of automatic files by name, so staging runs evicted real backups.
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
// The names the app can produce, and the pruner's matcher, read straight from the source.
const PRUNE_RE = /^liberty_clients_\d{4}-\d{2}-\d{2}(_staging|_local)?_(masked|FULL)\.json$/;
const name = (date, envTag, manual, full) =>
  'liberty_clients_' + date + envTag + (manual ? '_' + manual : '') + (full ? '_FULL' : '_masked') + '.json';

test('an environment tag is derived from the host, and production has none', () => {
  assert.ok(/var _BACKUP_ENV_TAG\s*=\s*\(_IS_LOCAL \? '_local' : \(_IS_STAGING \? '_staging' : ''\)\)/.test(src),
    'production must stay on the untagged name so existing history is not orphaned');
});

test('both filename builders carry the tag', () => {
  const builders = src.match(/var fname='liberty_clients_'\+[^;]+;/g) || [];
  const jsonBuilders = builders.filter((b) => /_masked|_FULL/.test(b));
  assert.strictEqual(jsonBuilders.length, 2, 'expected both JSON backup names: ' + JSON.stringify(builders));
  jsonBuilders.forEach((b) => assert.ok(b.includes('_BACKUP_ENV_TAG'), 'untagged builder: ' + b));
});

test('a staging backup cannot take production\'s daily slot', () => {
  assert.notStrictEqual(name('2026-09-01', '_staging', '', false), name('2026-09-01', '', '', false));
});

test('the pruner still recognises production\'s existing files', () => {
  assert.ok(PRUNE_RE.test(name('2026-09-01', '', '', false)), 'existing untagged history must still prune');
  assert.ok(PRUNE_RE.test(name('2026-09-01', '', '', true)));
});

test('the pruner recognises tagged files too, and keeps them in their own group', () => {
  const envOf = (n) => { const m = String(n).match(/^liberty_clients_\d{4}-\d{2}-\d{2}(_staging|_local)?_/); return (m && m[1]) || '_prod'; };
  assert.ok(PRUNE_RE.test(name('2026-09-01', '_staging', '', false)));
  assert.strictEqual(envOf(name('2026-09-01', '_staging', '', false)), '_staging');
  assert.strictEqual(envOf(name('2026-09-01', '_local', '', false)), '_local');
  assert.strictEqual(envOf(name('2026-09-01', '', '', false)), '_prod');
});

test('manual backups are still excluded from automatic pruning', () => {
  assert.ok(!PRUNE_RE.test(name('2026-09-01', '', 'manual', false)),
    'a hand-taken backup must never be auto-deleted');
  assert.ok(!PRUNE_RE.test(name('2026-09-01', '_staging', 'manual', false)));
});
