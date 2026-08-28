'use strict';
// Regression tests for the CODE_AUDIT_2 billing/data-integrity fixes. Several of these pin CALL SITES
// rather than pure functions — the audit found that the suite tested `_padMin` itself while the code
// that prints minutes went half-patched, so the renderer is tested here directly.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

// ── C6: every minutes field the PDF draws must be two digits ──
function fakePdf(sink) {
  return { setFont(){}, setFontSize(){}, getTextWidth(){ return 10; }, setDrawColor(){}, setLineWidth(){},
           line(){}, rect(){}, setFillColor(){}, addPage(){}, setTextColor(){}, splitTextToSize(t){ return [t]; },
           text(t){ sink.push(String(t)); } };
}

// The renderer runs to completion against the harness DOM. It is NOT wrapped in try/catch: an
// exception here means the invoice PDF does not render at all, which must fail the suite. The
// previous version swallowed every exception and then asserted that the (empty) output contained
// no bad minutes — so replacing the whole renderer with `throw` left all 249 tests green.
// Every assertion below is therefore paired: prove the field was DRAWN, then prove it was correct.
test('PDF renderer: draws the invoice, and no minutes field is a single digit', () => {
  const w = loadApp();
  const data = { clientName:'Jane Doe', billingPeriod:'08/2026', hourlyRate:'27.00',
                 svcHH:'20', svcMM:'5', cplxHH:'4', cplxMM:'5', p1HH:'20', p1MM:'5',
                 grandHH:'24', grandMM:'5', hasComplex:true, tasks:{ svc:[], cplx:[] } };
  const REQUIRED = {
    1: ['Jane Doe', '08/2026', '27.00', '20.05'],           // client, period, rate, Total Time
    2: ['Jane Doe', '08/2026', '4.05', '20.05', '24.05'],   // complex care, previous page, grand total
  };
  for (const page2 of [false, true]) {
    const pageNo = page2 ? 2 : 1;
    const out = [];
    w.drawInvoicePageVector(fakePdf(out), data, page2, 31);

    assert.ok(out.length > 20, 'page ' + pageNo + ' drew almost nothing (' + out.length + ' strings)');
    for (const expected of REQUIRED[pageNo]) {
      assert.ok(out.some(s => String(s).trim() === expected),
        'page ' + pageNo + ' never drew ' + JSON.stringify(expected) +
        ' — a certified form missing this field is not a passing render');
    }
    const singleDigitMinutes = out.filter(s => /^\d+\.\d$/.test(String(s).trim()));
    assert.deepStrictEqual(singleDigitMinutes, [],
      'page ' + pageNo + ' drew a 1-digit minute (reads as tens of minutes on a state invoice): ' +
      singleDigitMinutes.join(', '));
  }
});

test('_padMin is applied to all four minute fields in the renderer source', () => {
  // Belt-and-braces: the renderer is DOM-heavy, so also assert no raw `data.<x>MM||''` interpolation
  // survives next to a '.' concatenation — that exact shape is what shipped the bug twice.
  const fs = require('fs'), path = require('path');
  const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
  const raw = src.match(/\+\s*'\.'\s*\+\s*\(data\.\w+MM\s*\|\|\s*''\)/g) || [];
  assert.deepStrictEqual(raw, [], 'unpadded minutes interpolation found: ' + raw.join(' | '));
});

// ── C7: a 409 conflict must NOT pin the stale local row ──
test('roster conflict: a 409 clears _unsaved so the server copy can win', () => {
  const w = loadApp(); resetStorage(w);
  assert.strictEqual(w._isConflict({ isConflict: true }), true);
  assert.strictEqual(w._isConflict(new Error('HTTP 500')), false);

  w.saveCaregiversLS({ cg1: { name: 'Edited locally', _rowVersion: 'v1', _unsaved: true } });
  // what the save's rejection handler now does for a conflict:
  w._rosterMarkUnsaved('caregiver', 'cg1', !w._isConflict({ isConflict: true }));
  assert.strictEqual(w.getCaregivers().cg1._unsaved, undefined, 'a conflict must not keep the stale row');

  const server = { cg1: { name: 'Server copy', _rowVersion: 'v2' } };
  assert.strictEqual(w._mergeRosterMap(server, w.getCaregivers()).cg1.name, 'Server copy',
    'after a conflict the server row wins, so the next save carries a fresh _rowVersion');

  // a genuine failure still keeps the local edit
  w._rosterMarkUnsaved('caregiver', 'cg1', !w._isConflict(new Error('network')));
  assert.strictEqual(w.getCaregivers().cg1._unsaved, true, 'a real failure still protects the local edit');
});

// ── H4: the task's db id must reach localStorage ──
test('saveTaskAPI: the new dbId is written back to localStorage', async () => {
  const w = loadApp(); resetStorage(w);
  w.spToken = 'tok'; w._apiToken = 'tok';
  w.saveTodos([{ id: 'td_1', text: 'Call caseworker', done: false }]);
  w.fetch = () => Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({ id: 4242 }) });
  await w.saveTaskAPI(w.getTodos()[0]);
  assert.strictEqual(w.getTodos()[0].dbId, 4242,
    'without this the next sync sees the task as unsynced AND accepts the server copy — a duplicate');
});

// ── H7: the Excel export must read the field that actually holds status ──
test('exportClientsXLSX: Status comes from clientStatus, not the never-set p.status', () => {
  const w = loadApp(); resetStorage(w);
  const sheets = [];                       // capture EVERY sheet, not just the last one
  w.XLSX = { utils: { json_to_sheet: (r) => { sheets.push(r); return {}; }, book_new: () => ({}),
                      book_append_sheet: () => {}, sheet_add_aoa: () => {}, aoa_to_sheet: () => ({}) },
             writeFile: () => {} };
  w.saveProfilesLS({
    'Active Client': { clientName: 'Active Client', clientStatus: 'active' },
    'Ended Client':  { clientName: 'Ended Client',  clientStatus: 'terminated' },
  });
  try { w.exportClientsXLSX(); } catch (e) { /* later sheets may need more stubs */ }

  const clientSheet = sheets.find(rows => Array.isArray(rows) && rows.length &&
                                          Object.prototype.hasOwnProperty.call(rows[0], 'Client Name'));
  assert.ok(clientSheet, 'the client sheet was produced');
  const row = clientSheet.find(r => r['Client Name'] === 'Ended Client');
  assert.ok(row, 'the terminated client is in the export');
  assert.strictEqual(row['Status'], w.clientStatusLabel('terminated'),
    'Status must come from clientStatus — p.status is never set, so it exported everyone as "active"');
});
