'use strict';
// The MSA-1904 PDF is the document MDHHS actually receives, and nothing tested it: an earlier audit
// found the whole suite still passed with the invoice renderer replaced by `throw`. These tests
// render a real PDF through captureInvoicePDFVector and read it back — the certified values from
// the text, and the day-grid marks from the drawing operations, since a checked cell is two
// diagonal strokes rather than a glyph.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp } = require('./harness');
const { jsPDF } = require('jspdf');

const IDS = ['clientName','clientName2','medicaidId','worker','worker2','billingPeriod','billingPeriod2',
  'hourlyRate','billTo','svcHH','svcMM','cplxHH','cplxMM','p1HH','p1MM','grandHH','grandMM',
  'dateSubmitted','sigDate1','sigDate2'];

// A day grid: `marks` is a list of [dayIndex, columnIndex] pairs to check.
function gridWith(marks, days) {
  const g = [];
  for (let d = 0; d < (days || 31); d++) g.push(new Array(15).fill(false));
  marks.forEach(([d, c]) => { g[d][c] = true; });
  return g;
}

async function render(opts) {
  opts = opts || {};
  const w = loadApp();
  w.jspdf = { jsPDF };
  w.document.body.insertAdjacentHTML('beforeend',
    IDS.map((i) => '<input id="' + i + '">').join('') +
    '<input type="checkbox" id="showComplex"><div id="sigArea1"></div><div id="sigArea2"></div>');
  const set = (id, v) => { const e = w.document.getElementById(id); if (e) e.value = v; };
  const vals = Object.assign({
    clientName: 'Jane Doe', medicaidId: '1234567', worker: 'A Sawyer', billingPeriod: '08/2026',
    hourlyRate: '27.00', svcHH: '62', svcMM: '21', grandHH: '62', grandMM: '21',
    billTo: 'MDHHS', dateSubmitted: '09/01/2026', sigDate1: '09/01/2026', sigDate2: '09/01/2026',
  }, opts.values || {});
  Object.keys(vals).forEach((k) => set(k, vals[k]));
  if (opts.hasComplex) w.document.getElementById('showComplex').checked = true;
  w.captureStates = () => ({ svc: opts.svc || gridWith([]), cplx: opts.cplx || [] });
  const raw = Buffer.from(await w.captureInvoicePDFVector(), 'base64').toString('latin1');
  return {
    raw,
    // Text drawn on the page, in order.
    text: (raw.match(/\((?:\\.|[^()])*\)\s*Tj/g) || []).map((s) => s.replace(/\s*Tj$/, '').slice(1, -1)),
    // Every stroked line segment. A checked cell contributes exactly two.
    lines: (raw.match(/\d+\.?\d*\s+\d+\.?\d*\s+l\b/g) || []).length,
    pages: (raw.match(/\/Type\s*\/Page[^s]/g) || []).length,
  };
}
const has = (t, s) => t.some((x) => x.indexOf(s) !== -1);

test('the certified header carries the client, Medicaid ID, period and rate', async () => {
  const { text } = await render();
  ['Jane Doe', '1234567', '08/2026', '27.00', 'MDHHS', 'A Sawyer'].forEach((v) =>
    assert.ok(has(text, v), 'missing from the PDF: ' + v));
  assert.ok(has(text, 'HOME HELP AGENCY INVOICE'));
  assert.ok(has(text, 'MSA-1904'), 'the form number identifies what was certified');
});

test('billed time prints as HH.MM with the minutes padded', async () => {
  // "20.5" on this form reads as 20 hours 50 minutes — 45 minutes of over-billed time.
  // Pass a SINGLE digit — the value the padding exists for. '05' would pass either way.
  const { text } = await render({ values: { svcHH: '20', svcMM: '5', grandHH: '20', grandMM: '5' } });
  assert.ok(has(text, '20.05'), 'expected 20.05, got: ' + JSON.stringify(text.slice(-8)));
  assert.ok(!has(text, '20.5'), 'a bare .5 reads as 50 minutes on this form');
});

test('the total shown is the total billed', async () => {
  // The label and the figure are drawn as separate text runs on the same baseline.
  const { text } = await render();
  assert.ok(has(text, 'Total Time for Services Above:'), 'total label missing');
  assert.ok(has(text, '62.21'), 'billed total missing: ' + JSON.stringify(text.slice(-8)));
});

test('every day of the month gets a row', async () => {
  const { text } = await render();
  for (let d = 1; d <= 31; d++) assert.ok(has(text, String(d)), 'no row for day ' + d);
});

test('each checked cell draws exactly two strokes, and none are drawn when nothing is checked', async () => {
  const none = await render({ svc: gridWith([]) });
  const one = await render({ svc: gridWith([[0, 0]]) });
  const four = await render({ svc: gridWith([[0, 0], [1, 0], [2, 8], [30, 14]]) });
  assert.strictEqual(one.lines - none.lines, 2, 'one checked cell should add two strokes');
  assert.strictEqual(four.lines - none.lines, 8, 'four checked cells should add eight');
});

test('the certification sentence names the agency', async () => {
  const { text } = await render();
  assert.ok(has(text, 'I certify that Liberty Home Care Assistance has provided all the services as checked above.'));
});

test('a complex-care invoice gets its second page; an ordinary one does not', async () => {
  const plain = await render();
  const complex = await render({ hasComplex: true, cplx: gridWith([[0, 0]]) });
  assert.strictEqual(plain.pages, 1);
  assert.strictEqual(complex.pages, 2, 'complex care belongs on page 2');
});

test('complex care checked but with an empty grid does NOT add a page', async () => {
  const { pages } = await render({ hasComplex: true, cplx: gridWith([]) });
  assert.strictEqual(pages, 1);
});
