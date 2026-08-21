'use strict';
// Appendix B fixes. Each pins a behaviour that was genuinely broken — the verification pass refuted
// several sibling findings whose "silent failure" was already surfaced elsewhere, so those are absent.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

test('caregiver name is built identically by both save paths', () => {
  const w = loadApp();
  assert.strictEqual(w._cgDisplayName('Casey', 'M', 'Giver'), 'Casey M Giver');
  assert.strictEqual(w._cgDisplayName('Casey', '', 'Giver'), 'Casey Giver');
  assert.strictEqual(w._cgDisplayName('Casey', undefined, 'Giver'), 'Casey Giver',
    'the two paths disagreeing here silently renamed the record and orphaned its audit history');
});

test('parallel deletes do not resurrect each other in lhca_id_map', async () => {
  const w = loadApp(); resetStorage(w);
  w.localStorage.setItem('lhca_id_map', JSON.stringify({ A: 1, B: 2, C: 3 }));
  const resolvers = [];
  w.fetch = () => new Promise(res => resolvers.push(() => res({ ok: true, status: 200 })));
  // three deletes fired in one synchronous pass, as bulkDelete does
  w.deleteProfileSP('A'); w.deleteProfileSP('B'); w.deleteProfileSP('C');
  resolvers.forEach(fn => fn());                       // then all three responses land
  await new Promise(r => setTimeout(r, 30));
  const map = w.getIdMap();
  assert.deepStrictEqual(Object.keys(map), [],
    'each delete must re-read the map — closing over a snapshot let the last writer restore the others');
});

test('the disclosure CSV neutralises formula injection and flags truncation', () => {
  const w = loadApp(); resetStorage(w);
  let csv = '';
  w.Blob = function (parts) { csv = String(parts[0]); return {}; };
  w.URL.createObjectURL = () => 'blob:x'; w.URL.revokeObjectURL = () => {};
  w._doEmailAuditCSVDownload([{ timestamp: 't', sentBy: 'a@b.c', type: 'x',
                                recipient: '=cmd|calc!A1', caseworkerName: 'W', billingPeriod: '07/2026',
                                attachmentCount: 1, clientNames: ['Jane'], success: true }]);
  assert.ok(/'=cmd/.test(csv), 'a leading = is neutralised so Excel does not evaluate it');
  assert.ok(!/WARNING/.test(csv), 'a short export carries no truncation warning');

  const many = Array.from({ length: 1000 }, () => ({ timestamp: 't', recipient: 'r', clientNames: [], success: true }));
  w._doEmailAuditCSVDownload(many);
  assert.ok(/INCOMPLETE/.test(csv),
    'at the 1000-row backend cap the file must say it is incomplete — it is a disclosure register');
});

test('the idle wipe closes PHI-bearing popups it opened', () => {
  const w = loadApp(); resetStorage(w);
  let closed = false;
  w._phiWindows.push({ closed: false, close() { closed = true; } });
  w.clearPHIFromStorage();
  assert.strictEqual(closed, true,
    'a caregiver task sheet is a separate document — wiping storage never reached it');
  assert.strictEqual(w._phiWindows.length, 0, 'the list is reset');
});
