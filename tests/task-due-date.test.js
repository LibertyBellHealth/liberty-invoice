'use strict';
// A due date is a calendar DAY. The overdue checks did `new Date(t.due) < new Date()`, which parses
// 'YYYY-MM-DD' as UTC midnight — 8pm the PREVIOUS EVENING in Michigan. So a task due today read as
// overdue all day, and tomorrow's task turned red at 8pm tonight. Four places did this: the sidebar
// badge, the attention panel, and two red-styling paths.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

const w = loadApp();
resetStorage(w);
const ymd = (offsetDays) => {
  const d = new Date();
  d.setDate(d.getDate() + offsetDays);
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
};

test('a task due TODAY is not overdue', () => {
  assert.strictEqual(w._isOverdueYmd(ymd(0)), false,
    'the whole point of a due date is that today still counts');
});

test('a task due TOMORROW is not overdue — including after 8pm', () => {
  assert.strictEqual(w._isOverdueYmd(ymd(1)), false);
});

test('a task due YESTERDAY is overdue', () => {
  assert.strictEqual(w._isOverdueYmd(ymd(-1)), true);
});

test('the local calendar day is used, not UTC', () => {
  // The bug in one line: this is what the old comparison did.
  const todayStr = ymd(0);
  assert.ok(new Date(todayStr) < new Date(), 'sanity: the old comparison did call today overdue');
  assert.strictEqual(w._isOverdueYmd(todayStr), false, 'the new one must not');
});

test('a missing or malformed due date is never overdue', () => {
  ['', null, undefined, 'not a date', '2026-13-99x'].forEach((v) => {
    assert.strictEqual(w._isOverdueYmd(v), false, 'unexpected overdue for ' + JSON.stringify(v));
  });
});

test('a full timestamp still compares by its date part', () => {
  assert.strictEqual(w._isOverdueYmd(ymd(0) + 'T00:00:00.000Z'), false);
  assert.strictEqual(w._isOverdueYmd(ymd(-1) + 'T23:59:59.000Z'), true);
});

test('the badge counts only genuinely overdue tasks', () => {
  w.saveTodos([
    { id: 't1', text: 'due today', due: ymd(0), done: false },
    { id: 't2', text: 'due tomorrow', due: ymd(1), done: false },
    { id: 't3', text: 'due yesterday', due: ymd(-1), done: false },
    { id: 't4', text: 'overdue but done', due: ymd(-3), done: true },
  ]);
  const overdue = w.getTodos().filter((t) => !t.done && t.due && w._isOverdueYmd(t.due));
  // Arrays cross the jsdom realm boundary, so compare by value.
  assert.strictEqual(JSON.stringify([...overdue].map((t) => t.id)), JSON.stringify(['t3']));
});
