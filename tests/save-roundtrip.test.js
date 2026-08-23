'use strict';
// ROUND-TRIP: populate a record with every field set, render the form exactly as the UI does, save
// it untouched, and assert NOTHING changed. Any field a save handler silently drops or blanks fails
// here. This is the mechanical detector for the class of bug that destroyed a caseworker's agency
// and that three audits and 175 tests missed -- both instances "looked deliberate" on inspection.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

function quiet(w) {
  w.showAlert = () => {}; w.showToast = () => {}; w.showConfirm = () => {};
  w.addAuditEntry = () => {}; w.logActivity = () => {};
  w.fetch = () => Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({}) });
  ['saveProfileSP','saveCaregiverAPI','saveCaseworkerAPI','saveSupervisorAPI','renderSidebarClients',
   'renderClientGrid','updateStats','renderCaseworkerList','renderCaregiverList','renderInvHistory',
   'switchTab','refreshSupervisorDropdowns','renderSupervisorList','renderCgList','navDetail',
  ].forEach((f) => { if (typeof w[f] === 'function' || w[f] === undefined) w[f] = () => {}; });
}
function host(w, id) {
  if (!w.document.getElementById(id)) {
    w.document.body.insertAdjacentHTML('beforeend', '<div id="' + id + '"></div>');
  }
  return w.document.getElementById(id);
}
// Compare stored-before vs stored-after, ignoring keys the handler is expected to rewrite.
// Prove the save actually RAN. Without this, a handler that returns early passes every
// "field preserved" assertion trivially — which is exactly what happened here.
function positiveControl(w, inputId, value, read) {
  const el = w.document.getElementById(inputId);
  assert.ok(el, 'positive control input #' + inputId + ' must exist');
  el.value = value;
  return () => assert.strictEqual(read(), value,
    'the save did not run at all — every "preserved" assertion below would pass vacuously');
}
function diffFields(before, after, ignore) {
  const lost = [];
  Object.keys(before).forEach((k) => {
    if (ignore.includes(k)) return;
    const b = JSON.stringify(before[k]), a = JSON.stringify(after[k]);
    if (b !== a) lost.push(k + ': ' + b + ' -> ' + a);
  });
  return lost;
}

test('CAREGIVER: saving an untouched form preserves every field', () => {
  const w = loadApp(); resetStorage(w); quiet(w);
  host(w, 'cgInfoContent');
  w.document.body.insertAdjacentHTML('beforeend',
    '<div id="cgDetailName"></div><div id="cgDetailMeta"></div><button id="cgSaveInfoBtn"></button>');
  w.renderCaregiverGrid = () => {};
  const cg = { id: 'cg1', name: 'Ann B Lee', firstName: 'Ann', middleName: 'B', lastName: 'Lee',
    nickname: 'Annie', status: 'active', emptype: 'W2', phone: '313-555-0100',
    email: 'ann@example.com', hireDate: '2025-03-01', payRate: '15.50', dob: '1980-04-02',
    driversLicense: 'L1234', champsId: 'CH99', gender: 'F', street: '1 Main', city: 'Detroit',
    state: 'MI', zip: '48201', county: 'Wayne', notes: 'Reliable', miloginUsername: 'annlee',
    maxHours: '40', certifications: 'CPR', ecName: 'Bob Lee', ecPhone: '313-555-0199' };
  w.saveCaregiversLS({ cg1: cg });
  w.activeCgId = 'cg1';
  const before = JSON.parse(JSON.stringify(w.getCaregivers().cg1));
  w.renderCgInfoPane();
  const checkCg = positiveControl(w, 'cgi-phone', '313-555-9999', () => w.getCaregivers().cg1.phone);
  w.saveCgInfoPane();
  checkCg();
  const after = w.getCaregivers().cg1;
  const lost = diffFields(before, after, ['ssn','miloginPassword','_rowVersion','_unsaved','phone']);
  assert.deepStrictEqual(lost, [], 'fields destroyed by an untouched save:\n  ' + lost.join('\n  '));
});

test('CASEWORKER: saving an untouched form preserves every field', () => {
  const w = loadApp(); resetStorage(w); quiet(w);
  host(w, 'cwInfoContent');
  w.document.body.insertAdjacentHTML('beforeend',
    '<div id="cwDetailName"></div><div id="cwDetailMeta"></div><button id="cwSaveInfoBtn"></button>');
  const cw = { id: 'cw1', name: 'Renee Feldman', first_name: 'Renee', last_name: 'Feldman',
    middle_name: '', nickname: '', title: 'Case Manager', agency: 'MDHHS - Wayne', org: '',
    phone: '313-555-8830', fax: '313-555-8831', email: 'rf@michigan.gov', street: '2 State',
    city: 'Detroit', state: 'MI', zip: '48226', county: 'Wayne', notes: 'Primary contact',
    supervisor_id: 'sup1' };
  w.saveCaseworkersLS([cw]);
  w.activeCwId = 'cw1';
  const before = JSON.parse(JSON.stringify(w.getCaseworkers()[0]));
  w.renderCwInfoPane();
  const checkCw = positiveControl(w, 'cwi-phone', '313-555-9999',
    () => w.getCaseworkers()[0].phone);
  w.saveCwInfoPane();
  checkCw();
  const after = w.getCaseworkers()[0];
  const lost = diffFields(before, after, ['_rowVersion','_unsaved','phone']);
  assert.deepStrictEqual(lost, [], 'fields destroyed by an untouched save:\n  ' + lost.join('\n  '));
});

test('SUPERVISOR: saving preserves the concurrency token', () => {
  const w = loadApp(); resetStorage(w); quiet(w);
  w.document.body.insertAdjacentHTML('beforeend',
    '<input id="supd-id"><input id="supd-name"><input id="supd-title">' +
    '<input id="supd-phone"><input id="supd-email"><div id="supDetailName"></div>');
  const sup = { id: 'sup1', name: 'Pat Reed', title: 'Supervisor', phone: '313-555-7000',
                email: 'pat@michigan.gov', _rowVersion: '00000000000012AB' };
  w.saveSupervisorsLS({ sup1: sup });
  ['id','name','title','phone','email'].forEach((f) => {
    w.document.getElementById('supd-' + f).value = sup[f === 'id' ? 'id' : f] || '';
  });
  w.document.getElementById('supd-phone').value = '313-555-9999';
  w.saveSupDetail();
  const after = w.getSupervisors().sup1;
  assert.strictEqual(after.phone, '313-555-9999', 'the save must actually have run');
  assert.strictEqual(after.name, 'Pat Reed');
  assert.strictEqual(after._rowVersion, '00000000000012AB',
    'saveSupDetail rebuilds the record from 5 fields, dropping the row_version — so the next save ' +
    'sends no expected_version and a concurrent edit is silently overwritten');
});

test('CLIENT: saving an untouched form preserves every field', () => {
  const w = loadApp(); resetStorage(w); quiet(w);
  host(w, 'infoGrid');
  w.document.body.insertAdjacentHTML('beforeend',
    '<div id="detailName"></div><div id="detailMeta"></div><button id="saveInfoBtn"></button>' +
    '<div id="breadcrumb"></div>');
  const prof = { clientName: 'Jane Doe', firstName: 'Jane', middleName: '', lastName: 'Doe',
    nickname: 'Janie', medicaidId: '1234567', medicare: '1EG4TE5MK73', program: 'carrier',
    carrier: 'Humana', memberId: 'HUM-99', clientStatus: 'active', startDate: '2025-01-01',
    dob: '1950-02-03', gender: 'Female', driversLicense: 'D9999', phone: '313-555-4000',
    homePhone: '313-555-4001', clientEmail: 'jane@example.com', street: '9 Elm', city: 'Detroit',
    state: 'MI', zip: '48202', county: 'Wayne', worker: 'Renee Feldman', caseworkerId: 'cw1',
    caregiverId: 'cg1', clientNotes: 'Prefers morning visits', liveIn: false, invoices: [] };
  w.saveProfilesLS({ 'Jane Doe': prof });
  w.activeProfileName = 'Jane Doe';
  const before = JSON.parse(JSON.stringify(w.getProfiles()['Jane Doe']));
  w.renderInfoPane();
  const checkCl = positiveControl(w, 'ei-medicaid', '9998887',
    () => (w.getProfiles()['Jane Doe'] || {}).medicaidId);
  w.saveClientInfo();
  checkCl();
  const after = w.getProfiles()['Jane Doe'];
  assert.ok(after, 'the client must still exist under the same name');
  const lost = diffFields(before, after, ['ssn','_rowVersion','_unsaved','_clientSynced','_dbId','medicaidId']);
  assert.deepStrictEqual(lost, [], 'fields destroyed by an untouched save:\n  ' + lost.join('\n  '));
});

test('CLIENT: a legacy record with no program keeps its carrier details', () => {
  const w = loadApp(); resetStorage(w); quiet(w);
  host(w, 'infoGrid');
  const prof = { clientName: 'Legacy Client', firstName: 'Legacy', lastName: 'Client',
    program: '', carrier: 'Humana', memberId: 'HUM-77', clientStatus: 'active',
    startDate: '2025-01-01', invoices: [] };   // required when Active, or the save bails and this test proves nothing
  w.saveProfilesLS({ 'Legacy Client': prof });
  w.activeProfileName = 'Legacy Client';
  w.renderInfoPane();
  const checkLg = positiveControl(w, 'ei-medicaid', '5554443',
    () => (w.getProfiles()['Legacy Client'] || {}).medicaidId);
  w.saveClientInfo();
  checkLg();
  const after = w.getProfiles()['Legacy Client'];
  assert.strictEqual(after.carrier, 'Humana',
    'program is unset on legacy records, and unset is not "carrier" — so an untouched save wiped it');
  assert.strictEqual(after.memberId, 'HUM-77');
});

test('CLIENT: switching the program AWAY from carrier still clears the carrier details', () => {
  const w = loadApp(); resetStorage(w); quiet(w);
  host(w, 'infoGrid');
  w.saveProfilesLS({ 'Switcher Client': { clientName: 'Switcher Client', firstName: 'Switcher',
    lastName: 'Client', program: 'carrier', carrier: 'Humana', memberId: 'HUM-55',
    clientStatus: 'active', startDate: '2025-01-01', invoices: [] } });
  w.activeProfileName = 'Switcher Client';
  w.renderInfoPane();
  w.document.getElementById('ei-program').value = 'champs';   // the user actually switches it
  w.saveClientInfo();
  const after = w.getProfiles()['Switcher Client'];
  assert.strictEqual(after.program, 'champs', 'the switch must apply');
  assert.strictEqual(after.carrier, '', 'a real switch away from carrier still clears the stale carrier');
  assert.strictEqual(after.memberId, '');
});

test('CAREGIVER: a blank employment type is not silently promoted to Full-Time', () => {
  const w = loadApp(); resetStorage(w); quiet(w);
  host(w, 'cgInfoContent');
  w.document.body.insertAdjacentHTML('beforeend',
    '<div id="cgDetailName"></div><div id="cgDetailMeta"></div><button id="cgSaveInfoBtn"></button>');
  w.renderCaregiverGrid = () => {};
  w.saveCaregiversLS({ cg2: { id: 'cg2', name: 'Blank Type', firstName: 'Blank', lastName: 'Type',
    status: 'active', emptype: '' } });
  w.activeCgId = 'cg2';
  w.renderCgInfoPane();
  const check = positiveControl(w, 'cgi-phone', '313-555-1212', () => w.getCaregivers().cg2.phone);
  w.saveCgInfoPane();
  check();
  assert.strictEqual(w.getCaregivers().cg2.emptype, '',
    'the select had no blank option, so an unset employment type rendered as Full-Time and SAVED ' +
    'as Full-Time — a silent change to a plausible wrong value, not an obvious blank');
});
