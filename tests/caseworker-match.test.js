// The import matched an ASW to the roster by EMAIL ONLY. When the email could not be read it fell
// through to "Add <name> as a caseworker" with the box checked, creating a SECOND record for
// someone already on file — a duplicate "A Sawyer" next to "Addison Sawyer", same email, same
// phone. Reported 2026-09-01. MDHHS prints initial + surname, so name matching must handle that.
const test = require('node:test');
const assert = require('node:assert');
const { loadApp } = require('./harness');
const w = loadApp();

const ROSTER = [
  { id: 'cw_1', name: 'Addison Sawyer', email: 'SawyerA2@michigan.gov', phone: '586.770.9560' },
  { id: 'cw_2', name: 'Rahma Feto', email: 'FetoR@michigan.gov', phone: '3135051660' },
  { id: 'cw_3', name: 'Tammy Coleman', email: 'ColemanT1@michigan.gov', phone: '313-804-6694' },
];
const match = res => w._dhsMatchCaseworker(res, ROSTER);

test('email match wins, case-insensitively', () => {
  assert.strictEqual(match({ aswEmail: 'sawyera2@MICHIGAN.gov' }).id, 'cw_1');
});

test('no email: the printed "A Sawyer" still finds Addison Sawyer', () => {
  assert.strictEqual(match({ aswName: 'A Sawyer' }).id, 'cw_1');
});

test('no email: phone matches despite different punctuation', () => {
  assert.strictEqual(match({ aswPhone: '586-770-9560' }).id, 'cw_1');
});

test('a genuinely new worker is NOT matched to anyone', () => {
  assert.strictEqual(match({ aswName: 'B Newcomer', aswEmail: 'NewcomerB@michigan.gov' }), null);
});

test('a different first initial on the same surname is not a match', () => {
  assert.strictEqual(match({ aswName: 'J Sawyer' }), null);
});

test('a shared first name with a different surname is not a match', () => {
  assert.strictEqual(match({ aswName: 'Addison Brown' }), null);
});

test('the full name matches a roster entry stored with only an initial', () => {
  const roster = [{ id: 'cw_9', name: 'A Sawyer', email: '', phone: '' }];
  assert.strictEqual(w._dhsMatchCaseworker({ aswName: 'Addison Sawyer' }, roster).id, 'cw_9');
});

test('nothing to match on returns null rather than a wrong guess', () => {
  assert.strictEqual(match({}), null);
  assert.strictEqual(match({ aswName: 'Cher' }), null);   // single word, no surname
});
