'use strict';
// The MSA-1904 is a CERTIFIED record. Anything on it that is re-derived at print time rather than
// replayed from the invoice can silently rewrite a document that was already sent.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

function invoiceDom(w) {
  if (!w.document.getElementById('hourlyRate')) {
    w.document.body.insertAdjacentHTML('beforeend',
      ['clientName','clientName2','medicaidId','worker','worker2','billingPeriod','billingPeriod2',
       'hourlyRate','billTo','svcHH','svcMM','cplxHH','cplxMM','p1HH','p1MM','grandHH','grandMM',
       'dateSubmitted','sigDate1','sigDate2'].map((i) => '<input id="' + i + '">').join('') +
      '<input type="checkbox" id="showComplex"><div id="complexSection"></div>' +
      '<div id="sigArea1"></div><div id="sigArea2"></div>');
  }
  ['rebuild','applyStates','renderNotesPane'].forEach((f) => { if (typeof w[f] === 'function') w[f] = () => {}; });
}

test('a reprint replays the signature the invoice was certified with', async () => {
  const w = loadApp(); resetStorage(w); invoiceDom(w);
  w.saveSigsLS([{ id: 'sig_a', label: 'Owner', data: 'data:image/png;base64,AAA' },
                { id: 'sig_b', label: 'Second', data: 'data:image/png;base64,BBB' }]);
  w.saveProfilesLS({ Jane: { clientName: 'Jane', medicaidId: '1', invoices: [] } });
  await w.loadInvoiceForCapture('Jane',
    { billingPeriod: '03/2026', data: { sigId: 'sig_b', svcHH: '20', svcMM: '00' } }, '03/2026');
  const img = w.document.getElementById('sigArea1');
  assert.strictEqual(img.getAttribute('data-sig-id'), 'sig_b',
    'it always took sigs[0], so placing your SECOND signature emailed your FIRST — and deleting or ' +
    'reordering signatures re-certified every already-sent invoice under a different person');
});

test('an invoice with no recorded signature still falls back to the primary one', async () => {
  const w = loadApp(); resetStorage(w); invoiceDom(w);
  w.saveSigsLS([{ id: 'sig_a', label: 'Owner', data: 'data:image/png;base64,AAA' }]);
  w.saveProfilesLS({ Jane: { clientName: 'Jane', invoices: [] } });
  await w.loadInvoiceForCapture('Jane', { billingPeriod: '03/2026', data: { svcHH: '20' } }, '03/2026');
  assert.ok(w.document.getElementById('sigArea1').src, 'records saved before the id existed must still print signed');
});

test('a reprint carries the dates the invoice was signed on, not today', async () => {
  const w = loadApp(); resetStorage(w); invoiceDom(w);
  w.saveSigsLS([{ id: 'sig_a', data: 'x' }]);
  w.saveProfilesLS({ Jane: { clientName: 'Jane', invoices: [] } });
  await w.loadInvoiceForCapture('Jane', { billingPeriod: '03/2026',
    data: { svcHH: '20', dateSubmitted: '03/31/2026', sigDate1: '03/31/2026', sigDate2: '03/31/2026' } }, '03/2026');
  assert.strictEqual(w.document.getElementById('sigDate1').value, '03/31/2026',
    'reprinting a March invoice in August produced a certified form dated August');
  assert.strictEqual(w.document.getElementById('dateSubmitted').value, '03/31/2026');
});

test('Bill To prints what the invoice was sent with, even if the caseworker changed', async () => {
  const w = loadApp(); resetStorage(w); invoiceDom(w);
  w.saveSigsLS([{ id: 'sig_a', data: 'x' }]);
  // The caseworker on file now has NO agency — previously that printed Bill To blank on the PDF
  // while the screen still showed the saved value.
  w.saveCaseworkersLS([{ id: 1, name: 'Worker One', agency: '', county: '' }]);
  w.saveProfilesLS({ Jane: { clientName: 'Jane', caseworkerId: 1, worker: 'Worker One', invoices: [] } });
  await w.loadInvoiceForCapture('Jane',
    { billingPeriod: '03/2026', data: { svcHH: '20', billTo: 'MDHHS - Wayne', worker: 'Worker One' } }, '03/2026');
  assert.strictEqual(w.document.getElementById('billTo').value, 'MDHHS - Wayne',
    'the certified form must carry what it was sent with');
});

// ── The day grid certifies service that did not happen ───────────────────────────────────────
test('the generated day grid never marks days before service started', () => {
  const w = loadApp(); resetStorage(w);
  const prof = { clientName: 'Late Start', startDate: '2026-07-21' };
  const built = w._dhsBuildFirstInvoice(
    { hours: '29', minutes: '47', rate: '27.00', tasks: [{ task: 'Bathing', freq: 'Daily' }] },
    prof, '07/2026');
  assert.ok(built && built.data && built.data.tasks && built.data.tasks.svc, 'the grid must build');
  const grid = built.data.tasks.svc;
  const marked = grid.map((row, i) => (row.some(Boolean) ? i + 1 : null)).filter(Boolean);
  assert.ok(marked.length, 'some days must be marked');
  assert.strictEqual(Math.min.apply(null, marked), 21,
    'the grid is the "verification of services" — marks before the start date certify visits that ' +
    'never happened. Earliest marked day was ' + Math.min.apply(null, marked));
});

test('a client active all month is still marked from day 1', () => {
  const w = loadApp(); resetStorage(w);
  const built = w._dhsBuildFirstInvoice(
    { hours: '29', minutes: '47', rate: '27.00', tasks: [{ task: 'Bathing', freq: 'Daily' }] },
    { clientName: 'Full', startDate: '2025-01-01' }, '07/2026');
  const grid = built.data.tasks.svc;
  const marked = grid.map((row, i) => (row.some(Boolean) ? i + 1 : null)).filter(Boolean);
  assert.strictEqual(Math.min.apply(null, marked), 1, 'a full month must not be truncated');
});
