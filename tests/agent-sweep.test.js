'use strict';
// Findings from two agent sweeps of the "a save silently rewrites stored data" class, after the
// round-trip harness landed. Each was verified by reading the cited code before fixing.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

function quiet(w) {
  if (!w.Element.prototype.scrollIntoView) w.Element.prototype.scrollIntoView = function () {};
  w.showAlert = () => {}; w.showToast = () => {}; w.showConfirm = () => {};
  w.addAuditEntry = () => {}; w.logActivity = () => {};
  w.fetch = () => Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({}) });
  ['saveProfileSP','saveCaregiverAPI','saveCaseworkerAPI','saveSupervisorAPI','saveTaskAPI',
   'renderSidebarClients','renderClientGrid','updateStats','renderCaseworkerList','renderTodos',
   'updateTaskBadge','renderInvHistory','switchTab','refreshSupervisorDropdowns','renderCaregiverGrid',
  ].forEach((f) => { w[f] = () => {}; });
}

// ── The supervisor modal was the unfixed twin of saveSupDetail ───────────────────────────────
test('the supervisor EDIT MODAL merges instead of replacing the record', () => {
  const w = loadApp(); resetStorage(w); quiet(w);
  w.saveSupervisorsLS({ sup1: { id: 'sup1', name: 'Pat Reed', title: 'Mr.', phone: '313-555-7000',
                                email: 'pat@michigan.gov', _rowVersion: '00000000000012AB' } });
  w.document.body.insertAdjacentHTML('beforeend',
    '<div id="supModal"><input id="sup-name" value="Pat Reed"><input id="sup-phone" value="313-555-7000">' +
    '<input id="sup-email" value="pat@michigan.gov"><select id="sup-title"><option>Mr.</option></select></div>');
  w.document.getElementById('sup-title').value = 'Mr.';
  w._saveSupervisorFromModal('sup1');
  const after = w.getSupervisors().sup1;
  assert.strictEqual(after.name, 'Pat Reed', 'the save must have run');
  assert.strictEqual(after._rowVersion, '00000000000012AB',
    'rebuilding from 5 fields dropped the concurrency token, so the next save could clobber another device');
});

// ── Selects that could not hold what was stored ──────────────────────────────────────────────
test('a client carrier outside the built-in list survives a save', () => {
  const w = loadApp(); resetStorage(w); quiet(w);
  w.document.body.insertAdjacentHTML('beforeend',
    '<div id="infoGrid"></div><div id="detailName"></div><div id="detailMeta"></div>' +
    '<button id="saveInfoBtn"></button><div id="breadcrumb"></div>');
  w.saveProfilesLS({ 'Carrier Client': { clientName: 'Carrier Client', firstName: 'Carrier',
    lastName: 'Client', program: 'carrier', carrier: 'Blue Cross Complete', memberId: 'BC-1',
    gender: 'M', clientStatus: 'active', startDate: '2025-01-01', invoices: [] } });
  w.activeProfileName = 'Carrier Client';
  w.renderInfoPane();
  w.saveClientInfo();
  const after = w.getProfiles()['Carrier Client'];
  assert.strictEqual(after.carrier, 'Blue Cross Complete',
    'the AI assistant writes carrier as free text, so a plan outside the 8 built-ins is reachable');
  assert.strictEqual(after.gender, 'M', 'a short-form gender from an import must not be blanked');
});

// ── A renamed client used to orphan its tasks ────────────────────────────────────────────────
test('renaming a client moves its tasks with it', () => {
  const w = loadApp(); resetStorage(w); quiet(w);
  w.document.body.insertAdjacentHTML('beforeend',
    '<div id="infoGrid"></div><div id="detailName"></div><div id="detailMeta"></div>' +
    '<button id="saveInfoBtn"></button><div id="breadcrumb"></div>');
  w.saveProfilesLS({ 'Jane Doe': { clientName: 'Jane Doe', firstName: 'Jane', lastName: 'Doe',
    clientStatus: 'active', startDate: '2025-01-01', invoices: [] } });
  w.saveTodos([{ id: 'td_1', text: 'Reassessment due', client: 'Jane Doe', done: false }]);
  w.activeProfileName = 'Jane Doe';
  w.renderInfoPane();
  w.document.getElementById('ei-last').value = 'Doe-Smith';
  w.saveClientInfo();
  assert.ok(w.getProfiles()['Jane Doe-Smith'], 'the rename must have applied');
  assert.strictEqual(w.getTodos()[0].client, 'Jane Doe-Smith',
    'tasks are keyed by client NAME; left behind they vanish from the client’s Overview');
});

test('editing a task whose client is not in the current roster keeps the link', () => {
  const w = loadApp(); resetStorage(w); quiet(w);
  w.document.body.insertAdjacentHTML('beforeend',
    '<div id="taskEditModal"><div id="taskEditTitle"></div><div id="taskEditSubtitle"></div>' +
    '<input id="taskEditName"><input id="taskEditDue"><textarea id="taskEditNote"></textarea>' +
    '<select id="taskEditClient"></select><button id="taskEditSaveBtn"></button>' +
    '<button id="taskEditCancelBtn"></button></div>');
  w.saveProfilesLS({});                       // roster not loaded yet
  w.showTaskEditModal({ client: 'Ghost Client', name: 'Call the caseworker', due: '', note: '' },
                      function () {});
  const sel = w.document.getElementById('taskEditClient');
  assert.strictEqual(sel.value, 'Ghost Client',
    'the select offered only "— No client —", so saving a due-date change unlinked the task');
});

// NOT a finding — recorded so it isn't chased again: the agent sweep flagged saveCgInfoPane for
// writing an empty cgi-milogin-pass over the stored password. It cannot: saveCaregiversLS strips
// miloginPassword before persisting (HIPAA — app.js "never persist MI Login passwords"), and the
// backend only writes the column when a non-empty value is supplied (pwProvided). A one-line
// truthy guard was added anyway as defence-in-depth, but nothing was being lost.
