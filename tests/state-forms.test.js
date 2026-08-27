'use strict';
// The FIRST version of this file was vacuous: it looked for `STATE_FORM_INPUTS` / `STATE_FORMS`,
// neither of which exists (the real name is STATE_FORM_INPUT_MAPS), so it always took the fallback
// branch and asserted almost nothing. It would not have caught a single one of the defects below.
// This version reads the REAL PDF templates and asserts against their actual field names.
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const { loadApp, resetStorage } = require('./harness');

const FORMS_DIR = path.join(__dirname, '..', 'forms');
let PDFDocument = null;
try { PDFDocument = require(path.join(__dirname, '..', '..', 'crm-backend', 'node_modules', 'pdf-lib')).PDFDocument; }
catch (e) { /* pdf-lib unavailable — the template tests below self-skip */ }

function seed(w) {
  resetStorage(w);
  w.saveProfilesLS({ 'Jane Doe': { clientName: 'Jane Doe', firstName: 'Jane', lastName: 'Doe',
    medicaidId: '1234567', startDate: '2026-01-15', dob: '1950-04-02', street: '1 Main',
    city: 'Detroit', state: 'MI', zip: '48201', county: 'Wayne', phone: '313-555-1000',
    worker: 'Worker One', caseworkerId: 1 } });
  w.saveCaseworkersLS([{ id: 1, name: 'Worker One', phone: '313-555-2000', email: 'w@mi.gov' }]);
  return w._buildFormDataDict('Jane Doe') || {};
}

test('the MSA-4676 stamps a Start of Service date', () => {
  const w = loadApp(); const dict = seed(w);
  assert.ok(dict.start_date, 'msa_start maps to start_date, which the dictionary never defined');
  assert.match(String(dict.start_date), /2026/);
});

test('EVERY field the 4676 coordinate map declares resolves in the dictionary', () => {
  const w = loadApp(); const dict = seed(w);
  const map = w.STATE_FORM_INPUT_MAPS && w.STATE_FORM_INPUT_MAPS.msa4676;
  assert.ok(map, 'STATE_FORM_INPUT_MAPS.msa4676 must exist — the old test looked for the wrong name ' +
                 'and silently passed for weeks');
  const missing = Object.keys(map).map((k) => map[k]).filter((key) => !(key in dict));
  assert.deepStrictEqual(missing, [], 'these would stamp BLANK: ' + missing.join(', '));
});

// ── Against the REAL templates ───────────────────────────────────────────────────────────────
// The live defect was invisible to any JS-only test: the template was swapped for a fillable
// AcroForm in May and the code kept drawing at the OLD coordinates, so every real box went out
// empty. Only reading the PDF catches that.
for (const [file, formType] of [['MSA-4676.pdf', 'msa4676'], ['DHS-390.pdf', 'dhs390']]) {
  test(file + ': every text field the template declares gets a value (or is deliberately blank)', async (t) => {
    if (!PDFDocument) return t.skip('pdf-lib not available');
    const p = path.join(FORMS_DIR, file);
    if (!fs.existsSync(p)) return t.skip(file + ' not present');
    const w = loadApp(); const dict = seed(w);
    const doc = await PDFDocument.load(fs.readFileSync(p));
    const fields = doc.getForm().getFields();
    assert.ok(fields.length > 0, file + ' is a fillable AcroForm — the code must fill it, not draw over it');
    const text = fields.filter((f) => typeof f.setText === 'function').map((f) => f.getName());
    const unresolved = text.filter((n) => !w._matchAcroFormField(n, dict, formType));
    // Fields the CLIENT/CAREGIVER fills in by hand are legitimately blank.
    const byHand = (w.STATE_FORM_FIELD_MAPS && w.STATE_FORM_FIELD_MAPS[formType]) || {};
    // A signature line is deliberately blank — the person signs it.
    const unexpected = unresolved.filter((n) => !(n in byHand) && !/signature/i.test(n));
    assert.deepStrictEqual(unexpected, [],
      file + ' fields that resolve to nothing and are not marked manual-fill: ' + unexpected.join(' | '));
  });
}

test('MSA-4676: the code does not stamp at coordinates outside the real field boxes', async (t) => {
  if (!PDFDocument) return t.skip('pdf-lib not available');
  const p = path.join(FORMS_DIR, 'MSA-4676.pdf');
  if (!fs.existsSync(p)) return t.skip('template not present');
  const w = loadApp();
  const doc = await PDFDocument.load(fs.readFileSync(p));
  const isAcro = doc.getForm().getFields().length > 0;
  // With a fillable template the coordinate overlay must NOT be the path taken — it was written for
  // a template replaced in May 2026 and its positions match none of the current boxes.
  assert.ok(isAcro, 'template is fillable');
  const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
  assert.match(src, /_acroFields\.length \? \[\] : \(def\.fields\|\|\[\]\)/,
    'buildStateFormBytes must skip the stale coordinate overlay whenever the template has real fields');
});
