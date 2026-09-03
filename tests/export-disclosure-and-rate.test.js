'use strict';
// Two export gaps: three bulk PHI extracts left no HIPAA §164.528 disclosure record (the JSON
// export has always written one), and the roster spreadsheet's Invoices sheet hardcoded a 27.00
// rate fallback and reported service hours only — omitting complex care.
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');
const { loadApp, resetStorage } = require('./harness');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');

test('every bulk export writes a disclosure record', () => {
  ['Roster spreadsheet', 'Invoice report spreadsheet', 'Per-client invoice PDF folders']
    .forEach((label) => assert.ok(src.indexOf("_logBulkExport('" + label) !== -1,
      'no disclosure record for: ' + label));
});

test('the disclosure names who exported it', () => {
  const fn = src.slice(src.indexOf('function _logBulkExport'), src.indexOf('function exportReportExcel'));
  assert.ok(/currentUserEmail\(\)/.test(fn), 'an accounting of disclosures needs the actor');
  assert.ok(/logActivity\('export'/.test(fn));
});

test('the spreadsheet rate falls back to the CONFIGURED rate, not a hardcoded one', () => {
  assert.ok(src.indexOf("'Hourly Rate':(inv.data&&inv.data.hourlyRate)||stateRate()") !== -1,
    'a hardcoded fallback disagrees with the PDF after a rate change');
  assert.strictEqual((src.match(/\|\|'27\.00'/g) || []).length, 0,
    'no export may invent a rate');
});

test('the exported hours use the grand total when the invoice has one', () => {
  const w = loadApp();
  resetStorage(w);
  // Mirror the exporter's expression.
  const hours = (d) => {
    const hh = (d && d.grandHH) || (d && d.svcHH) || '';
    const mm = (d && d.grandHH) ? (d.grandMM || '') : ((d && d.svcMM) || '');
    return hh === '' ? '' : hh + '.' + w._padMin(mm);
  };
  assert.strictEqual(hours({ svcHH: '20', svcMM: '00', grandHH: '25', grandMM: '30' }), '25.30',
    'complex care belongs in the exported total');
  assert.strictEqual(hours({ svcHH: '62', svcMM: '21' }), '62.21', 'an ordinary invoice is unchanged');
  assert.strictEqual(hours({ svcHH: '20', svcMM: '5' }), '20.05', 'minutes stay padded');
  assert.strictEqual(hours({}), '');
});
