'use strict';
// Carrier (managed-care) authorization: a weekly unit schedule turned into monthly authorized units.
// Figures come from a real DAAA / Priority Health T1019 notice whose August posting was 202 units / $1,506.92.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

// Sun..Sat: Thu + Sat 12 units, the other five days 10 = 74/week.
const DAAA = { procedure: 'T1019', rate: '7.46', unitMinutes: 15,
  startDate: '2026-08-13', endDate: '2027-02-12', week: [10, 10, 10, 10, 12, 10, 12] };

function quiet(w) {
  w.showAlert = (m) => { w._alerts.push(m); }; w._alerts = [];
  w.showToast = () => {}; w.addAuditEntry = () => {}; w.saveTaskAPI = () => {};
  w.updateTaskBadge = () => {}; w.aiTrack = () => {}; w.logActivity = () => {};
}

test('months: the real schedule reproduces the carrier portal figures', () => {
  const w = loadApp();
  const months = w._carrierAuthMonths(DAAA).map((m) => ({ ...m }));
  assert.deepStrictEqual([...months.map((m) => m.ym)],
    ['2026-08', '2026-09', '2026-10', '2026-11', '2026-12', '2027-01', '2027-02']);
  assert.strictEqual(months[0].from, '2026-08-13', 'first month starts on the start date, not the 1st');
  assert.strictEqual(months[0].units, 202, 'Aug 13-31 posted as 202 units');
  assert.strictEqual(months[0].amount, 1506.92, '202 x $7.46');
  assert.strictEqual(months[1].units, 316);
  assert.strictEqual(months[2].units, 330);
  assert.strictEqual(months[6].to, '2027-02-12', 'last month stops at the end date');
  assert.strictEqual(months[6].units, 126, 'Feb 1-12 2027, end date included');
});

test('months: start and end dates are both counted', () => {
  const w = loadApp();
  const oneThu = w._carrierAuthMonths({ ...DAAA, startDate: '2026-08-13', endDate: '2026-08-13' });
  assert.strictEqual(oneThu.length, 1);
  assert.strictEqual(oneThu[0].units, 12, 'a single Thursday');
});

test('months: every date counted exactly once across DST changes', () => {
  const w = loadApp();
  const ones = { startDate: '2026-01-01', endDate: '2026-12-31', week: [1, 1, 1, 1, 1, 1, 1] };
  const units = [...w._carrierAuthMonths(ones).map((m) => m.units)];
  assert.deepStrictEqual(units, [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]);
});

test('months: bad or missing dates give no rows instead of throwing', () => {
  const w = loadApp();
  assert.strictEqual(w._carrierAuthMonths(null).length, 0);
  assert.strictEqual(w._carrierAuthMonths({ ...DAAA, startDate: '' }).length, 0);
  assert.strictEqual(w._carrierAuthMonths({ ...DAAA, startDate: '2026-02-30' }).length, 0, 'impossible date');
  assert.strictEqual(w._carrierAuthMonths({ ...DAAA, endDate: '2026-08-01' }).length, 0, 'end before start');
  const open = w._carrierAuthMonths({ ...DAAA, endDate: '' });
  assert.ok(open.length >= 2 && open.length <= 36, 'no end date still shows a bounded range');
});

test('money and hours: rate per unit, 15-minute units', () => {
  const w = loadApp();
  assert.strictEqual(w._caWeekUnits(DAAA), 74);
  assert.strictEqual(w._caHoursLabel(74, 15), '18h 30m');
  assert.strictEqual(w._caHoursLabel(202, 15), '50h 30m');
  assert.strictEqual(w._caAmount(74, 7.46), 552.04);
  assert.strictEqual(w._caRate({ rate: '$7.46' }), 7.46);
  assert.strictEqual(w._caRate({ rate: '' }), null, 'no rate → no dollar figures, not $0');
  assert.strictEqual(w._caMoney(2461.8), '$2,461.80');
});

test('_clientSig: a schedule change makes the client save fire', () => {
  const w = loadApp();
  const base = { firstName: 'Jane', program: 'carrier', carrierAuth: DAAA };
  const posted = { ...base, carrierAuth: { ...DAAA, posted: { '2026-08': 202 } } };
  assert.notStrictEqual(w._clientSig(base), w._clientSig(posted), 'posted units must count as an edit');
  assert.notStrictEqual(w._clientSig(base), w._clientSig({ ...base, carrierAuth: null }), 'removal must count');
});

test('save sends carrier_authorization; load parses it back', async () => {
  const w = loadApp(); resetStorage(w); quiet(w);
  let body = null;
  w.fetch = (url, opts) => {
    if (opts && opts.method === 'POST' && /homecare-clients$/.test(url)) body = JSON.parse(opts.body);
    return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({ id: 7, row_version: '00000000000007D1' }) });
  };
  const prof = { clientName: 'Pat Carrier', firstName: 'Pat', lastName: 'Carrier', program: 'carrier',
    carrier: 'Priority Health', carrierAuth: DAAA, clientStatus: 'active', invoices: [] };
  w.saveProfilesLS({ 'Pat Carrier': prof });
  await w.saveProfileSP('Pat Carrier', prof, true);
  assert.ok(body, 'the client POST ran');
  assert.strictEqual(body.carrier_authorization.startDate, '2026-08-13');
  assert.deepStrictEqual([...body.carrier_authorization.week], DAAA.week);
  assert.strictEqual(w._parseAuth(JSON.stringify(DAAA)).endDate, '2027-02-12', 'SQL returns the column as a string');
});

test('view: renders the schedule and the monthly table for a carrier client', () => {
  const w = loadApp(); resetStorage(w); quiet(w);
  w.document.body.insertAdjacentHTML('beforeend', '<div id="authContent"></div>');
  w.saveProfilesLS({ 'Pat Carrier': { clientName: 'Pat Carrier', program: 'carrier', carrier: 'Priority Health',
    carrierAuth: { ...DAAA, posted: { '2026-08': 202, '2026-09': 300 } }, invoices: [] } });
  w.activeProfileName = 'Pat Carrier';
  w.renderAuthPane(false);
  const html = w.document.getElementById('authContent').innerHTML;
  assert.match(html, /Authorization \(Priority Health\)/);
  assert.match(html, /74 units · 18h 30m · \$552\.04/);
  assert.match(html, /\$7\.46\/unit · \$29\.84\/hr/);
  assert.match(html, /\$1,506\.92/);
  assert.match(w.document.getElementById('ca-diff-2026-08').textContent, /all posted/);
  assert.match(w.document.getElementById('ca-diff-2026-09').textContent, /16 left/);
  assert.doesNotMatch(html, /Import DHS-1210/, 'no DHS-1210 import for a carrier client');
});

test('view: a CHAMPS client still gets the DHS-1210 pane', () => {
  const w = loadApp(); resetStorage(w); quiet(w);
  w.document.body.insertAdjacentHTML('beforeend', '<div id="authContent"></div>');
  w.saveProfilesLS({ 'Cham Client': { clientName: 'Cham Client', program: 'champs', carrierAuth: DAAA, invoices: [] } });
  w.activeProfileName = 'Cham Client';
  w.renderAuthPane(false);
  assert.match(w.document.getElementById('authContent').innerHTML, /Import DHS-1210/);
});

test('edit → save stores the schedule and adds one renewal task 30 days before the end', () => {
  const w = loadApp(); resetStorage(w); quiet(w);
  let posts = 0; w.saveProfileSP = () => { posts++; return Promise.resolve(); };
  w.document.body.insertAdjacentHTML('beforeend', '<div id="authContent"></div>');
  w.saveProfilesLS({ 'Pat Carrier': { clientName: 'Pat Carrier', program: 'carrier', carrier: 'Priority Health',
    carrierAuth: { posted: { '2026-08': 202 }, startDate: '2026-08-13', week: [1, 0, 0, 0, 0, 0, 0] }, invoices: [] } });
  w.activeProfileName = 'Pat Carrier';
  w.renderAuthPane(true);
  const set = (id, v) => { w.document.getElementById(id).value = v; };
  set('ca-proc', 'T1019'); set('ca-rate', '7.46'); set('ca-unit-min', '15');
  set('ca-start', '2026-08-13'); set('ca-end', '2027-02-12');
  DAAA.week.forEach((u, i) => set('ca-day-' + i, String(u)));
  w.saveCarrierAuthPane();
  assert.deepStrictEqual(w._alerts, []);
  assert.strictEqual(posts, 1, 'saved to the server');
  const ca = w.getProfiles()['Pat Carrier'].carrierAuth;
  assert.deepStrictEqual([...ca.week], DAAA.week);
  assert.strictEqual(ca.unitMinutes, 15);
  assert.strictEqual(ca.posted['2026-08'], 202, 'editing the schedule keeps posted units');
  let tasks = w.getTodos().filter((t) => /^Carrier authorization ends /.test(t.text));
  assert.strictEqual(tasks.length, 1);
  assert.strictEqual(tasks[0].due, '2027-01-13');
  // Saving again (renewal moved out) updates the same task rather than adding another.
  w.renderAuthPane(true); set('ca-end', '2027-08-12'); w.saveCarrierAuthPane();
  tasks = w.getTodos().filter((t) => /^Carrier authorization ends /.test(t.text));
  assert.strictEqual(tasks.length, 1);
  assert.strictEqual(tasks[0].due, '2027-07-13');
});

test('edit → save rejects bad input without saving', () => {
  const w = loadApp(); resetStorage(w); quiet(w);
  let posts = 0; w.saveProfileSP = () => { posts++; return Promise.resolve(); };
  w.document.body.insertAdjacentHTML('beforeend', '<div id="authContent"></div>');
  w.saveProfilesLS({ 'Pat Carrier': { clientName: 'Pat Carrier', program: 'carrier', invoices: [] } });
  w.activeProfileName = 'Pat Carrier';
  const attempt = (fill) => {
    w.renderAuthPane(true);
    w.document.getElementById('ca-start').value = '2026-08-13';
    w.document.getElementById('ca-day-4').value = '12';
    fill((id, v) => { w.document.getElementById(id).value = v; });
    w._alerts = []; w.saveCarrierAuthPane(); return w._alerts.length;
  };
  assert.strictEqual(attempt((s) => s('ca-start', '')), 1, 'start date required');
  assert.strictEqual(attempt((s) => s('ca-end', '2026-08-01')), 1, 'end before start');
  assert.strictEqual(attempt((s) => s('ca-day-4', '')), 1, 'at least one day');
  assert.strictEqual(attempt((s) => s('ca-rate', 'abc')), 1, 'rate must be a number');
  assert.strictEqual(attempt((s) => s('ca-unit-min', '0')), 1, 'unit length > 0');
  assert.strictEqual(posts, 0);
  assert.strictEqual(w.getProfiles()['Pat Carrier'].carrierAuth, undefined);
});

test('posted units: saved per month for the right client, blank clears', () => {
  const w = loadApp(); resetStorage(w); quiet(w);
  const saved = []; w.saveProfileSP = (name) => { saved.push(name); return Promise.resolve(); };
  w.saveProfilesLS({ 'Pat Carrier': { clientName: 'Pat Carrier', program: 'carrier', carrierAuth: DAAA, invoices: [] },
    'Other Client': { clientName: 'Other Client', program: 'carrier', invoices: [] } });
  w.activeProfileName = 'Other Client';   // navigated away before the field lost focus
  w._carrierPostedChange('Pat Carrier', '2026-08', 202, '202');
  assert.strictEqual(w.getProfiles()['Pat Carrier'].carrierAuth.posted['2026-08'], 202);
  assert.deepStrictEqual(saved, ['Pat Carrier']);
  w._carrierPostedChange('Pat Carrier', '2026-08', 202, '2.5');
  assert.strictEqual(w._alerts.length, 1, 'fractional units rejected');
  assert.strictEqual(w.getProfiles()['Pat Carrier'].carrierAuth.posted['2026-08'], 202);
  w._carrierPostedChange('Pat Carrier', '2026-08', 202, '');
  assert.ok(!('2026-08' in w.getProfiles()['Pat Carrier'].carrierAuth.posted));
});

test('saving the client info tab keeps the carrier schedule', () => {
  const w = loadApp(); resetStorage(w); quiet(w);
  if (!w.Element.prototype.scrollIntoView) w.Element.prototype.scrollIntoView = function () {};
  w.showConfirm = () => {}; w.logActivity = () => {};
  w.fetch = () => Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({}) });
  ['saveProfileSP', 'renderSidebarClients', 'renderClientGrid', 'updateStats', 'switchTab', 'navDetail']
    .forEach((f) => { w[f] = () => {}; });
  w.document.body.insertAdjacentHTML('beforeend', '<div id="infoGrid"></div><div id="detailName"></div>' +
    '<div id="detailMeta"></div><button id="saveInfoBtn"></button><div id="breadcrumb"></div>');
  w.saveProfilesLS({ 'Pat Carrier': { clientName: 'Pat Carrier', firstName: 'Pat', lastName: 'Carrier',
    program: 'carrier', carrier: 'Priority Health', memberId: 'PH-1', clientStatus: 'active',
    startDate: '2026-08-13', carrierAuth: DAAA, invoices: [] } });
  w.activeProfileName = 'Pat Carrier';
  w.renderInfoPane();
  w.document.getElementById('ei-phone').value = '313-555-0000';
  w.saveClientInfo();
  const after = w.getProfiles()['Pat Carrier'];
  assert.strictEqual(after.phone, '313-555-0000', 'the save ran');
  assert.strictEqual(JSON.stringify(after.carrierAuth), JSON.stringify(DAAA));
});

test('remove: switching clients while the confirm is open removes it from the ORIGINAL client only', () => {
  const w = loadApp(); resetStorage(w); quiet(w);
  w.saveProfileSP = () => Promise.resolve();
  let confirmCb = null; w.showConfirm = (msg, cb) => { confirmCb = cb; };
  w.document.body.insertAdjacentHTML('beforeend', '<div id="authContent"></div>');
  w.saveProfilesLS({ 'Pat Carrier': { clientName: 'Pat Carrier', program: 'carrier', carrierAuth: DAAA, invoices: [] },
    'Other Client': { clientName: 'Other Client', program: 'carrier', carrierAuth: DAAA, invoices: [] } });
  w.activeProfileName = 'Pat Carrier';
  w._clearCarrierAuth();
  w.activeProfileName = 'Other Client';
  confirmCb();
  assert.strictEqual(w.getProfiles()['Pat Carrier'].carrierAuth, null);
  assert.ok(w.getProfiles()['Other Client'].carrierAuth, 'the client you switched to keeps theirs');
});

test('posted units typed back to back save one at a time, the last save carrying every value', async () => {
  const w = loadApp(); resetStorage(w); quiet(w);
  const calls = []; let release = [];
  w.saveProfileSP = (name, rec) => {
    calls.push(JSON.parse(JSON.stringify(rec.carrierAuth.posted)));
    return new Promise((res) => release.push(res));
  };
  w.saveProfilesLS({ 'Pat Carrier': { clientName: 'Pat Carrier', program: 'carrier', carrierAuth: DAAA, invoices: [] } });
  w._carrierPostedChange('Pat Carrier', '2026-08', 202, '202');
  w._carrierPostedChange('Pat Carrier', '2026-09', 316, '150');
  w._carrierPostedChange('Pat Carrier', '2026-10', 330, '20');
  assert.strictEqual(calls.length, 1, 'no second POST while the first is in flight — it would 409 on a stale row version');
  release.shift()();
  await new Promise((r) => setTimeout(r, 0));
  assert.strictEqual(calls.length, 2, 'the pending edits go out as ONE follow-up save');
  assert.deepStrictEqual(calls[1], { '2026-08': 202, '2026-09': 150, '2026-10': 20 });
  release.shift()();
  await new Promise((r) => setTimeout(r, 0));
  assert.strictEqual(calls.length, 2);
  w._carrierPostedChange('Pat Carrier', '2026-11', 316, '1');
  assert.strictEqual(calls.length, 3, 'the queue is idle again, so a later edit saves straight away');
});
