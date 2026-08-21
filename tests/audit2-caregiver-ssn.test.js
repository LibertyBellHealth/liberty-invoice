'use strict';
// The caregiver SSN moved off the bulk roster load and is now fetched when the field is focused.
// The risk that introduces is blanking: the field starts empty, so a save must not write that
// emptiness over a stored SSN. These tests pin that, and the masked-export fallback.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

// saveCgInfoPane reads a fixed set of cgi-* inputs; build them so the real function can run.
const FIELDS = ['first','last','middle','nickname','status','phone','email','dl','ssn','street','city',
                'state','zip','county','dob','gender','hire','emptype','pay','champs','maxhours',
                'certs','ec-name','ec-phone','milogin-user','milogin-pw','notes'];
function pane(w, values) {
  // saveCgInfoPane also repaints the header and the save button; supply them so the real function
  // can run to completion instead of throwing before the assertions.
  w.document.body.insertAdjacentHTML('beforeend',
    '<div id="cgInfoPaneTest">' + FIELDS.map(f => '<input id="cgi-' + f + '">').join('') +
    '<span id="cgDetailName"></span><span id="cgDetailMeta"></span>' +
    '<button id="cgSaveInfoBtn"></button></div>');
  Object.keys(values || {}).forEach(k => { const el = w.document.getElementById('cgi-' + k); if (el) el.value = values[k]; });
}

test('an untouched (never-focused) SSN field does not blank the stored SSN', () => {
  const w = loadApp(); resetStorage(w);
  w.saveCaregiversLS({ cg1: { name: 'Casey Giver', ssn: '123-45-6789', ssnLast4: '6789' } });
  w.activeCgId = 'cg1';
  let sent = null;
  w.saveCaregiverAPI = (id, cg) => { sent = cg; return Promise.resolve({}); };   // capture what we'd send
  pane(w, { first: 'Casey', last: 'Giver', ssn: '' });      // SSN box empty — the field was never focused
  w.saveCgInfoPane();
  assert.ok(sent, 'the save ran');
  assert.strictEqual(sent.ssn, '123-45-6789',
    'the record sent to the API must still carry the stored SSN — an empty box means "not fetched", not "clear it"');
  assert.strictEqual(w.getCaregivers().cg1.ssn, '123-45-6789', 'and the local copy is intact');
});

test('a typed SSN still saves', () => {
  const w = loadApp(); resetStorage(w);
  w.saveCaregiversLS({ cg1: { name: 'Casey Giver', ssn: '' } });
  w.activeCgId = 'cg1';
  w.saveCaregiverAPI = () => Promise.resolve({});
  pane(w, { first: 'Casey', last: 'Giver', ssn: '987-65-4321' });
  w.saveCgInfoPane();
  assert.strictEqual(w.getCaregivers().cg1.ssn, '987-65-4321', 'a real edit is written through');
});

test('_revealCgSsnField fetches on focus and caches the result', async () => {
  const w = loadApp(); resetStorage(w);
  w.saveCaregiversLS({ cg1: { name: 'Casey Giver', ssn: '', ssnLast4: '6789' } });
  let calls = 0;
  w.fetch = (url) => { calls++; assert.match(String(url), /\/caregivers\/cg1\/ssn$/);
                       return Promise.resolve({ ok: true, json: () => Promise.resolve({ ssn: '123-45-6789' }) }); };
  const inp = w.document.createElement('input');
  inp.type = 'password'; w.document.body.appendChild(inp);
  w._revealCgSsnField(inp, 'cg1');
  await new Promise(r => setTimeout(r, 10));
  assert.strictEqual(inp.value, '123-45-6789', 'the real SSN arrives on focus');
  assert.strictEqual(inp.type, 'text', 'and the field is revealed');
  assert.strictEqual(w.getCaregivers().cg1.ssn, '123-45-6789', 'cached so a second focus does not refetch');
  w._revealCgSsnField(inp, 'cg1');
  await new Promise(r => setTimeout(r, 10));
  assert.strictEqual(calls, 1, 'no refetch once the value is present');
});

test('the masked export still shows the last 4 when the full SSN was never fetched', () => {
  const w = loadApp(); resetStorage(w);
  let sheets = [];
  w.XLSX = { utils: { json_to_sheet: (r) => { sheets.push(r); return {}; }, book_new: () => ({}),
                      book_append_sheet: () => {}, sheet_add_aoa: () => {}, aoa_to_sheet: () => ({}) },
             writeFile: () => {} };
  w.saveCaregiversLS({ cg1: { name: 'Casey Giver', ssn: '', ssnLast4: '6789' } });
  w.saveProfilesLS({ 'Jane Doe': { clientName: 'Jane Doe' } });
  try { w.exportClientsXLSX(); } catch (e) { /* later sheets may need more stubs */ }
  // Both the client and caregiver sheets carry an 'SSN (masked)' column — pick the caregiver one by
  // its own row, not by the column name.
  let row = null;
  sheets.forEach(rows => { if (Array.isArray(rows)) rows.forEach(r => { if (r && r['Name'] === 'Casey Giver') row = r; }); });
  assert.ok(row, 'the caregiver row was exported');
  assert.strictEqual(row['SSN (masked)'], '***-**-6789',
    'the export never needed the full SSN — last 4 is enough and is all the roster now carries');
});
