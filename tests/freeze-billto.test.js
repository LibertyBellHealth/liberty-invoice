'use strict';
// An invoice records WHO WAS BILLED. The reopen and print paths were fixed to prefer the Bill To
// stored on the invoice, falling back to the live caseworker only for records saved before that
// field existed — but _dhsBuildFirstInvoice never stored one, so every generated invoice took the
// fallback forever. Reassign a client, or move a caseworker between organisations, and a SUBMITTED
// invoice reprints with a Bill To that MDHHS never saw, while the caseworker NAME frozen on it
// stays the old one — a reprint that contradicts itself.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

const AUTH = { hours: 20, minutes: 0, tasks: [
  { task: 'Bathing', perDay: '00:10', freq: '7 days per week', perMonth: '05:00' }] };

function app(caseworkers) {
  const w = loadApp();
  resetStorage(w);
  w.saveCaseworkersLS(caseworkers || []);
  return w;
}
const build = (w, prof) => w._dhsBuildFirstInvoice(
  { hours: AUTH.hours, minutes: AUTH.minutes, tasks: AUTH.tasks }, prof, '08/2026').data;

test('a generated invoice records the Bill To it was created with', () => {
  const w = app([{ id: 'cw_1', name: 'A Sawyer', agency: 'MDHHS Macomb County' }]);
  const data = build(w, { clientName: 'Jane Doe', caseworkerId: 'cw_1', worker: 'A Sawyer' });
  assert.strictEqual(data.billTo, 'MDHHS Macomb County');
});

test('it does not change when the caseworker later moves office', () => {
  const w = app([{ id: 'cw_1', name: 'A Sawyer', agency: 'MDHHS Macomb County' }]);
  const data = build(w, { clientName: 'Jane Doe', caseworkerId: 'cw_1', worker: 'A Sawyer' });
  w.saveCaseworkersLS([{ id: 'cw_1', name: 'A Sawyer', agency: 'MDHHS Wayne County' }]);
  assert.strictEqual(data.billTo, 'MDHHS Macomb County',
    'the invoice already sent must keep the office it was billed to');
});

test('county is used when the caseworker has no agency', () => {
  const w = app([{ id: 'cw_2', name: 'R Feto', county: 'Wayne' }]);
  const data = build(w, { clientName: 'Jane Doe', caseworkerId: 'cw_2', worker: 'R Feto' });
  assert.strictEqual(data.billTo, 'Wayne');
});

test('matching by name works when the client has no caseworker id', () => {
  const w = app([{ id: 'cw_3', name: 'T Coleman', agency: 'MDHHS Wayne County' }]);
  const data = build(w, { clientName: 'Jane Doe', worker: 'T Coleman' });
  assert.strictEqual(data.billTo, 'MDHHS Wayne County');
});

test('an unknown caseworker yields an empty Bill To, not a wrong one', () => {
  const w = app([{ id: 'cw_1', name: 'A Sawyer', agency: 'MDHHS Macomb County' }]);
  const data = build(w, { clientName: 'Jane Doe', worker: 'Nobody At All' });
  assert.strictEqual(data.billTo, '');
});

test('the caseworker name is still frozen alongside it', () => {
  const w = app([{ id: 'cw_1', name: 'A Sawyer', agency: 'MDHHS Macomb County' }]);
  const data = build(w, { clientName: 'Jane Doe', caseworkerId: 'cw_1', worker: 'A Sawyer' });
  assert.strictEqual(data.worker, 'A Sawyer', 'name and Bill To must agree on the same record');
});
