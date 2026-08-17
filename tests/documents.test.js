'use strict';
// Document-load robustness (persistence audit 2026-08-05). The guarantee: a FAILED document-list
// load must NEVER render "No documents yet" — that would make a worker think files vanished and
// re-upload duplicates. It must show a retry, and a 401 must say to sign in again.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp } = require('./harness');

test('_renderDocLoadError: a 401 tells the worker to sign in again — never "No documents yet"', () => {
  const w = loadApp();
  w.document.body.innerHTML = '<div id="hcDocList"><div>No documents yet.</div></div>';
  w._renderDocLoadError('hcDocList', "loadHcDocs('42')", new Error('HTTP 401'));
  const html = w.document.getElementById('hcDocList').innerHTML;
  assert.ok(!/No documents yet/.test(html), 'must NOT claim the files are gone');
  assert.ok(/sign in again/i.test(html), 'a 401 tells the worker to re-authenticate');
  assert.ok(/Retry/.test(html) && /loadHcDocs\('42'\)/.test(html), 'offers a working retry');
});

test('_renderDocLoadError: a connection error shows retry, not a false-empty pane', () => {
  const w = loadApp();
  w.document.body.innerHTML = '<div id="cwDocListAzure"></div>';
  w._renderDocLoadError('cwDocListAzure', 'renderCwDocsPane()', new Error('Failed to fetch'));
  const html = w.document.getElementById('cwDocListAzure').innerHTML;
  assert.ok(!/No documents yet/.test(html), 'no false empty');
  assert.ok(/connection issue/i.test(html), 'names the connection problem');
  assert.ok(/files are safe/i.test(html), 'reassures the files still exist on the server');
  assert.ok(/renderCwDocsPane\(\)/.test(html), 'retry re-runs the caseworker loader');
});

test('_renderDocLoadError: a missing container is a no-op, never a throw', () => {
  const w = loadApp();
  assert.doesNotThrow(() => w._renderDocLoadError('nope-does-not-exist', 'x()', new Error('HTTP 500')));
});
