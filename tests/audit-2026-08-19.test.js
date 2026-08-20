'use strict';
// Regression tests for the 2026-08-19 audit fixes (see the audit doc's HIGH/MEDIUM list). Each test
// pins the BEHAVIOUR that was wrong, so a future refactor can't quietly reintroduce it.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

// ── MSA-4676 routing (owner-confirmed workflow, 2026-08-20): the CAREGIVER signs the generated
// form first; the signed copy is then sent from Documents to the caseworker. open_email only ever
// attaches a freshly generated (unsigned) form, so that is the caregiver leg. ──
test('open_email: a generated 4676 with no recipient goes to the caregiver (the signature step)', async () => {
  const w = loadApp(); resetStorage(w);
  w.localStorage.setItem('lhca_caregivers', JSON.stringify({ cg1: { name: 'Casey Giver', email: 'caregiver@example.com' } }));
  w.localStorage.setItem('lhca_caseworkers', JSON.stringify([{ id: 'cw1', name: 'Case Worker', email: 'worker@michigan.gov' }]));
  w.localStorage.setItem('lhca_profiles', JSON.stringify({ 'Jane Doe': { clientName: 'Jane Doe', caregiverId: 'cg1', caseworkerId: 'cw1' } }));
  w._profilesCache = null;
  // PDF generation can't run under jsdom, so the result is the form-build error — what this pins is
  // that resolution did NOT fail on the caseworker leg (which would error about a caseworker email).
  const r = await w._asstOpenEmail({ client_name: 'Jane Doe', attach_form: 'msa4676' });
  assert.ok(!/caseworker/i.test(JSON.stringify(r)), 'an unsigned 4676 is not routed to the caseworker');

  // With no caregiver email on file it errors ABOUT the caregiver — proof of which branch ran.
  w.localStorage.setItem('lhca_caregivers', JSON.stringify({ cg1: { name: 'Casey Giver', email: '' } }));
  const r2 = await w._asstOpenEmail({ client_name: 'Jane Doe', attach_form: 'msa4676' });
  assert.match(String(r2.error || ''), /caregiver/i);
});

test('open_email: an explicit caseworker recipient still routes there', async () => {
  const w = loadApp(); resetStorage(w);
  w.localStorage.setItem('lhca_caregivers', JSON.stringify({ cg1: { name: 'Casey Giver', email: 'caregiver@example.com' } }));
  w.localStorage.setItem('lhca_caseworkers', JSON.stringify([{ id: 'cw1', name: 'Case Worker', email: 'worker@michigan.gov' }]));
  w.localStorage.setItem('lhca_profiles', JSON.stringify({ 'Jane Doe': { clientName: 'Jane Doe', caregiverId: 'cg1', caseworkerId: 'cw1' } }));
  w._profilesCache = null;
  const r = await w._asstOpenEmail({ client_name: 'Jane Doe', recipient: 'caseworker', subject: 's', body: 'b' });
  assert.strictEqual(r.to, 'worker@michigan.gov');
});

test('open_email: a non-4676 email with no recipient still defaults to the caregiver', async () => {
  const w = loadApp(); resetStorage(w);
  w.localStorage.setItem('lhca_caregivers', JSON.stringify({ cg1: { name: 'Casey Giver', email: 'caregiver@example.com' } }));
  w.localStorage.setItem('lhca_profiles', JSON.stringify({ 'Jane Doe': { clientName: 'Jane Doe', caregiverId: 'cg1' } }));
  w._profilesCache = null;
  const r = await w._asstOpenEmail({ client_name: 'Jane Doe', subject: 'hi', body: 'hello' });
  assert.strictEqual(r.to, 'caregiver@example.com');
});

// ── H2: "unbilled" defaults to the PREVIOUS month (a month is billed on the 1st of the next) ──
test('_asstPrevPeriod: returns the previous month as MM/YYYY', () => {
  const w = loadApp();
  const now = new Date();
  const prev = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const expect = String(prev.getMonth() + 1).padStart(2, '0') + '/' + prev.getFullYear();
  assert.strictEqual(w._asstPrevPeriod(), expect);
  assert.strictEqual(w._asstBillingPeriodNorm('last month'), expect);
});

test('query_billing unbilled: defaults to the previous month, not the current one', () => {
  const w = loadApp(); resetStorage(w);
  const prev = w._asstPrevPeriod();
  const t = w.today().split('/'); const cur = t[0] + '/' + t[2];
  // Active client, started long ago, WITH the previous month's invoice already on file.
  w.localStorage.setItem('lhca_profiles', JSON.stringify({
    'Jane Doe': { clientName: 'Jane Doe', startDate: '2020-01-01', clientStatus: 'active',
                  invoices: [{ billingPeriod: prev, status: 'draft' }] },
  }));
  w._profilesCache = null;
  const r = w._asstQueryBilling({ what: 'unbilled' });
  assert.strictEqual(r.period, prev, 'defaults to the previous month');
  assert.notStrictEqual(r.period, cur);
  assert.strictEqual(r.count, 0, 'a client billed for the previous month is not "unbilled"');
});

// ── H3: an export covers the FULL matched set, not the 80-row display cap ──
test('query_roster: no_cap returns every matched row (exports must not silently truncate)', () => {
  const w = loadApp(); resetStorage(w);
  const profs = {};
  for (let i = 0; i < 120; i++) profs['Client ' + i] = { clientName: 'Client ' + i, clientStatus: 'active' };
  w.localStorage.setItem('lhca_profiles', JSON.stringify(profs));
  w._profilesCache = null;
  const capped = w._asstQueryRoster({ action: 'list', filters: [] });
  assert.strictEqual(capped.clients.length, 80, 'the model-facing result stays capped');
  const full = w._asstQueryRoster({ action: 'list', filters: [], no_cap: true });
  assert.strictEqual(full.clients.length, 120, 'the export path gets every row');
  assert.strictEqual(full.truncated, false);
});

// ── M1: minutes are two digits — "HH.MM" renders a bare 5 as fifty minutes ──
test('_padMin: single-digit minutes are zero-padded, everything else untouched', () => {
  const w = loadApp();
  assert.strictEqual(w._padMin(5), '05');
  assert.strictEqual(w._padMin('5'), '05');
  assert.strictEqual(w._padMin('05'), '05');
  assert.strictEqual(w._padMin('30'), '30');
  assert.strictEqual(w._padMin(''), '');
  assert.strictEqual(w._padMin(null), '');
});

test('_dhsBuildFirstInvoice: 20h 5m becomes "20" / "05", never "20.5"', () => {
  const w = loadApp();
  const built = w._dhsBuildFirstInvoice({ hours: 20, minutes: 5, tasks: [] }, { clientName: 'Jane Doe' }, '07/2026');
  assert.strictEqual(built.data.svcHH, '20');
  assert.strictEqual(built.data.svcMM, '05');
  assert.strictEqual(built.data.grandMM, '05');
});

// ── M2: an authorization with no parsed HOURS must not generate a blank-total invoice ──
test('hasBillableAuthorization: an effective date alone is not billable', () => {
  const w = loadApp();
  assert.strictEqual(w.hasAuthorization({ authorization: { effectiveDate: '07/01/2026' } }), true);
  assert.strictEqual(w.hasBillableAuthorization({ authorization: { effectiveDate: '07/01/2026' } }), false);
  assert.strictEqual(w.hasBillableAuthorization({ authorization: { hours: 0, minutes: 0 } }), true);
  assert.strictEqual(w.hasBillableAuthorization({ authorization: { hours: '', tasks: [{ task: 'x' }] } }), false);
  assert.strictEqual(w.hasBillableAuthorization({ authorization: { hours: 29, minutes: 47 } }), true);
});

test('findClientsEligibleForAutoGen: skips a client whose authorization has no hours', () => {
  const w = loadApp(); resetStorage(w);
  w.localStorage.setItem('lhca_profiles', JSON.stringify({
    'No Hours': { clientName: 'No Hours', startDate: '2020-01-01', clientStatus: 'active',
                  authorization: { effectiveDate: '01/01/2020', tasks: [] }, invoices: [] },
    'Has Hours': { clientName: 'Has Hours', startDate: '2020-01-01', clientStatus: 'active',
                   authorization: { hours: 29, minutes: 47, tasks: [] }, invoices: [] },
  }));
  w._profilesCache = null;
  const names = w.findClientsEligibleForAutoGen('07/2026').map(e => e.name);
  assert.strictEqual(JSON.stringify(names), JSON.stringify(['Has Hours']));   // realm-safe compare
});

// ── M3: an OCR'd value never pre-checks an overwrite of data already on file ──
test('_dhsSuggestedUpdates: filling a blank is pre-checked; replacing a value is not', () => {
  const w = loadApp();
  const res = { medicaidId: '1234567890', aswEmail: 'new@michigan.gov', aswPhone: '313-555-1212' };
  const fill = w._dhsSuggestedUpdates(res, { medicaidId: '' }, null)[0];
  assert.strictEqual(fill.field, 'medicaidId');
  assert.strictEqual(fill.overwrite, false, 'a blank field is a fill, so it pre-checks');
  const over = w._dhsSuggestedUpdates(res, { medicaidId: '9999999999' }, null)[0];
  assert.strictEqual(over.overwrite, true, 'replacing a stored Medicaid ID is flagged as an overwrite');
  const cw = w._dhsSuggestedUpdates(res, {}, { id: 'cw1', email: 'old@michigan.gov', phone: '313-000-0000' });
  const email = cw.find(u => u.field === 'email');
  assert.strictEqual(email.overwrite, true);
  assert.strictEqual(email.verify, true, 'the ASW email/phone are heuristic reads — always flagged');
  assert.strictEqual(email.shared, true, 'the caseworker record is shared across clients');
});

// ── M5: a roster row whose save FAILED survives the next background load ──
test('_mergeRosterMap: a failed-save local row (_unsaved) wins over the stale server copy', () => {
  const w = loadApp();
  const server = { a: { id: 'a', name: 'Alice (server, pre-edit)', _rowVersion: 'v2' } };
  const local = { a: { id: 'a', name: 'Alice EDITED', _rowVersion: 'v2', _unsaved: true } };
  assert.strictEqual(w._mergeRosterMap(server, local).a.name, 'Alice EDITED');
  // …and once the save succeeds (flag cleared), the server is authoritative again.
  const saved = { a: { id: 'a', name: 'Alice EDITED', _rowVersion: 'v2' } };
  assert.strictEqual(w._mergeRosterMap(server, saved).a.name, 'Alice (server, pre-edit)');
});

test('_mergeRosterArr: a failed-save caseworker row wins in place (no duplicate row)', () => {
  const w = loadApp();
  const server = [{ id: 'cw1', name: 'Server Copy', _rowVersion: 'v2' }];
  const local = [{ id: 'cw1', name: 'Edited, save failed', _rowVersion: 'v2', _unsaved: true }];
  const out = w._mergeRosterArr(server, local);
  assert.strictEqual(out.length, 1, 'replaced in place, not appended');
  assert.strictEqual(out[0].name, 'Edited, save failed');
});

// ── M6: the caregiver-email gate is enforced for SSN-bearing categories ──
test('document email gate: SSN-bearing and ID-card categories are blocked, forms are not', () => {
  const w = loadApp();
  ['SSN_Card', 'Drivers_License', 'Insurance_Card', 'Medicare_Card', 'Medicaid_Card', 'I9_W4', 'Background_Check']
    .forEach(c => assert.ok(w._DOC_EMAIL_BLOCKED[c], c + ' must never be emailable to a caregiver'));
  assert.ok(!w._DOC_EMAIL_BLOCKED['Authorization']);
  assert.ok(w._DOC_EMAIL_SAFE['Authorization'], 'authorizations send without an extra prompt');
  assert.ok(!w._DOC_EMAIL_SAFE['Other'], '"Other" is the default bucket — it must warn first');
});

// ── LOW: assistant filters — day-precision dates, and "In Progress" == stored 'inactive' ──
test('_asstMatchFilter: before/after respect the day of month', () => {
  const w = loadApp();
  const f = { start_date: '05/20/2026' };
  assert.strictEqual(w._asstMatchFilter(f, { field: 'start_date', op: 'after', value: '05/15/2026' }), true);
  assert.strictEqual(w._asstMatchFilter(f, { field: 'start_date', op: 'after', value: '05/25/2026' }), false);
  assert.strictEqual(w._asstMatchFilter(f, { field: 'start_date', op: 'before', value: '05/25/2026' }), true);
});

test('_asstMatchFilter: status "In Progress" matches the stored value "inactive"', () => {
  const w = loadApp();
  assert.strictEqual(w._asstMatchFilter({ status: 'inactive' }, { field: 'status', op: 'eq', value: 'In Progress' }), true);
  assert.strictEqual(w._asstMatchFilter({ status: 'active' }, { field: 'status', op: 'eq', value: 'in progress' }), false);
});
