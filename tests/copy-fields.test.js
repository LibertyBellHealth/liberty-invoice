'use strict';
// Hover-copy on detail-form fields (ported from Health CRM, 2026-08-16). Locks in the copy
// FORMATTING (dash-free IDs keeping letters for MBI/DL, MM/DD/YYYY dates) and the generic
// button attachment (all filled text/date fields, skip checkboxes/no-id, idempotent).
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp } = require('./harness');

function inp(w, id, val, type) { const e = w.document.createElement('input'); e.id = id; e.type = type || 'text'; e.value = val; return e; }

test('_copyFmt: IDs strip separators but KEEP letters; dates → MM/DD/YYYY; else as-is', () => {
  const w = loadApp();
  assert.strictEqual(w._copyFmt(inp(w, 'ei-ssn', '123-45-6789')), '123456789');
  assert.strictEqual(w._copyFmt(inp(w, 'ei-medicaid', '1234 56 7890')), '1234567890');
  assert.strictEqual(w._copyFmt(inp(w, 'ei-medicare', '1EG4-TE5-MK73')), '1EG4TE5MK73'); // MBI keeps letters
  assert.strictEqual(w._copyFmt(inp(w, 'cgi-dl', 'D123-4567')), 'D1234567');             // DL keeps letter
  assert.strictEqual(w._copyFmt(inp(w, 'ei-dob', '1958-01-06', 'date')), '01/06/1958');
  assert.strictEqual(w._copyFmt(inp(w, 'ei-phone', '313-555-0100')), '313-555-0100');    // plain → unchanged
  assert.strictEqual(w._copyFmt(inp(w, 'ei-first', '')), '');                            // empty → nothing
});

test('wireCopyableFields: adds a copy button per filled field, skips checkboxes/no-id, idempotent', () => {
  const w = loadApp();
  const root = w.document.createElement('div'); root.id = 'copyTestRoot';
  root.innerHTML =
    '<div class="info-field"><label>SSN</label><input id="t-ssn" value="123-45-6789"></div>' +
    '<div class="info-field"><label>DOB</label><input id="t-dob" type="date" value="1958-01-06"></div>' +
    '<div class="info-field"><label>Chk</label><input id="t-chk" type="checkbox"></div>' +
    '<div class="info-field"><label>NoId</label><input value="x"></div>';
  w.document.body.appendChild(root);
  w.wireCopyableFields(root);
  assert.strictEqual(root.querySelectorAll('.copy-hover-btn').length, 2, 'button only on the two real text/date inputs');
  assert.ok(root.querySelector('#t-ssn').closest('.info-field').classList.contains('field-copyable'));
  w.wireCopyableFields(root); // run again
  assert.strictEqual(root.querySelectorAll('.copy-hover-btn').length, 2, 'idempotent — no duplicate buttons');
});
