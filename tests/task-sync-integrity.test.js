'use strict';
// Three ways tasks lost or duplicated rows against the server:
//  1. deleteTodo read dbId from the snapshot taken BEFORE its confirm dialog, so a task whose create
//     POST resolved while the dialog was open was deleted locally with no DELETE sent — and the next
//     background load brought it straight back.
//  2. A follow-up shipped its parent's LOCAL id as parent_id. On the next load every task's id
//     becomes its numeric DB id, so the child pointed at nothing and multi-step workflows flattened.
//  3. saveTaskAPI sends `id: dbId || undefined` and dbId only exists after the first POST resolves,
//     so two quick saves INSERTed twice.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

function app() {
  const w = loadApp();
  resetStorage(w);
  ['renderTodos', 'updateTaskBadge', 'logActivity', 'showToast'].forEach((f) => { w[f] = () => {}; });
  w.deleted = [];
  w.deleteTaskAPI = (dbId) => { w.deleted.push(dbId); };
  w.pending = null;
  w.showConfirm = (msg, onOk) => { w.pending = onOk; };
  return w;
}
const settle = () => new Promise((r) => setTimeout(r, 0));

test('a dbId assigned while the confirm was open is still deleted on the server', () => {
  const w = app();
  w.saveTodos([{ id: 'td_1', text: 'Typo task', done: false }]);   // no dbId yet
  w.deleteTodo('td_1');
  // The create POST resolves mid-dialog and writes the id back.
  w.saveTodos([{ id: 'td_1', text: 'Typo task', done: false, dbId: 101 }]);
  w.pending();
  assert.ok(w.deleted.indexOf(101) !== -1,
    'no DELETE was sent, so the row survives and the task returns on the next sync');
  assert.strictEqual(w.getTodos().length, 0, 'and it must still be gone locally');
});

test('follow-ups are deleted with their parent, using current ids', () => {
  const w = app();
  w.saveTodos([{ id: 'td_1', text: 'Parent', done: false },
               { id: 'td_2', text: 'Step 2', parentId: 'td_1', done: false }]);
  w.deleteTodo('td_1');
  w.saveTodos([{ id: 'td_1', text: 'Parent', done: false, dbId: 101 },
               { id: 'td_2', text: 'Step 2', parentId: 'td_1', done: false, dbId: 102 }]);
  w.pending();
  assert.ok(w.deleted.indexOf(101) !== -1 && w.deleted.indexOf(102) !== -1,
    'both rows must be deleted server-side: ' + JSON.stringify([...w.deleted]));
  assert.strictEqual(w.getTodos().length, 0);
});

test('a follow-up is SENT with its parent\'s database id, never the local one', async () => {
  const w = app();
  const bodies = [];
  w.fetch = (url, opt) => { bodies.push(JSON.parse(opt.body));
    return Promise.resolve({ ok: true, json: () => Promise.resolve({ id: 202 }) }); };
  w.saveTodos([{ id: 'td_1', text: 'Parent', dbId: 101, done: false },
               { id: 'td_2', text: 'Step 2', parentId: 'td_1', done: false }]);
  const child = w.getTodos().find((t) => t.id === 'td_2');
  await Promise.resolve(w.saveTaskAPI(child)).catch(() => {});
  await settle();
  // Inspect the REQUEST, not just the helper — the call site is where this went wrong.
  assert.strictEqual(bodies.length, 1);
  assert.strictEqual(bodies[0].parent_id, 101,
    'sending the local id leaves the child orphaned after the next reload');
});

test('an unsynced parent yields null rather than an id the server cannot resolve', () => {
  const w = app();
  w.saveTodos([{ id: 'td_1', text: 'Parent', done: false },
               { id: 'td_2', text: 'Step 2', parentId: 'td_1', done: false }]);
  assert.strictEqual(w._taskParentDbId(w.getTodos().find((t) => t.id === 'td_2')), null);
});

test('a parent id that is already a database id passes through', () => {
  const w = app();
  w.saveTodos([{ id: '101', text: 'Parent', dbId: 101, done: false }]);
  assert.strictEqual(w._taskParentDbId({ id: '102', parentId: '101' }), 101);
});

test('two saves before the first create returns produce ONE server row', async () => {
  const w = app();
  const bodies = []; let nextId = 100;
  w.fetch = (url, opt) => { bodies.push(JSON.parse(opt.body)); const id = ++nextId;
    return Promise.resolve({ ok: true, json: () => Promise.resolve({ id: id }) }); };
  const todo = { id: 'td_9', text: 'Call caseworker', done: false };
  w.saveTodos([todo]);
  const a = w.saveTaskAPI(todo);
  const b = w.saveTaskAPI(todo);          // second save before the first resolves
  await Promise.all([a, b].map((p) => Promise.resolve(p).catch(() => {})));
  await settle();
  // The count alone proves nothing — two INSERTs is also two calls. What matters is that the
  // SECOND request carries the id assigned by the first, i.e. it updates rather than inserting.
  const inserts = bodies.filter((b2) => b2.id === undefined);
  assert.strictEqual(inserts.length, 1,
    'two id-less POSTs are two server rows: ' + JSON.stringify(bodies.map((b2) => b2.id)));
  const saved = w.getTodos().find((t) => t.id === 'td_9');
  assert.ok(saved && saved.dbId, 'the task must end up with a database id');
});
