'use strict';
// Found by driving the real app, not by a test: a note typed and still inside its 600ms autosave
// window was replaced by the previous note if a roster reload landed in that window — and the old
// note was then saved back to the server, while the textarea still showed what had been typed.
//
// The reload's merge only keeps a local row flagged `_unsaved` ("local copy is newer, keep it"),
// which was set for a FAILED save but never for a PENDING one. And the real saveCaregiverAPI clears
// that flag on success unconditionally, so typing more while a save is in flight left the newer
// text exposed to the next reload. These tests use the real save functions and route fetch by URL,
// because the defect lives inside them.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

const tick = () => new Promise((r) => setTimeout(r, 0));
const settle = async () => { for (let i = 0; i < 6; i++) await tick(); };
const ok = (body) => Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(body) });

function app() {
  const w = loadApp();
  resetStorage(w);
  ['cgNotesContent', 'cwNotesContent'].forEach((id) => {
    if (!w.document.getElementById(id)) w.document.body.insertAdjacentHTML('beforeend', '<div id="' + id + '"></div>');
    w.document.getElementById(id).innerHTML = '';
  });
  Object.keys(w._pendingNoteSaves || {}).forEach((k) => { delete w._pendingNoteSaves[k]; });
  w._showSaveStatus = () => {};
  // aiTrack is defined in index.html, which the harness does not load. Without it saveProfileSP's
  // success path throws, falls into the FAILURE branch and sets _unsaved — so a test asserting the
  // flag survives would pass without the code under test ever running.
  w.aiTrack = () => {};
  w.renderCgGrid = () => {}; w.renderCwGrid = () => {}; w.renderCaregiverGrid = () => {};
  w.eval('_savesInFlight = 0;');
  return w;
}

// server holds `serverNotes`; every caregiver POST is recorded; a held POST can be released later.
function caregiverServer(w, serverNotes) {
  const s = { notes: serverNotes, posts: [], hold: false, release: null };
  w.fetch = (url, opt) => {
    const u = String(url), m = (opt && opt.method) || 'GET';
    if (/\/caregivers$/.test(u) && m === 'POST') {
      const b = JSON.parse(opt.body); s.posts.push(b.notes);
      const done = () => { s.notes = b.notes; return ok({ row_version: 'bbbbbbbbbbbbbbbb' }); };
      if (s.hold) return new Promise((res) => { s.release = () => res(done()); });
      return done();
    }
    if (/\/caregivers$/.test(u)) {
      return ok([{ id: 'cg_A', name: 'Alice Aide', status: 'active', notes: s.notes, row_version_hex: 'aaaaaaaaaaaaaaaa' }]);
    }
    return ok([]);
  };
  return s;
}
const typeCg = (w, text) => {
  const ta = w.document.getElementById('cgNotesArea');
  ta.value = text; ta.dispatchEvent(new w.Event('input'));
};

test('a roster reload inside the autosave window does not revert a typed caregiver note', async () => {
  const w = app();
  w.saveCaregiversLS({ cg_A: { id: 'cg_A', name: 'Alice Aide', notes: 'old', _rowVersion: 'aaaaaaaaaaaaaaaa' } });
  const srv = caregiverServer(w, 'old');
  w.activeCgId = 'cg_A';
  w.renderCgNotesPane();

  typeCg(w, 'new');
  w.loadCaregiversAPI();              // reload lands before the 600ms save fires
  await settle();
  assert.strictEqual(w.getCaregivers().cg_A.notes, 'new', 'the reload put the old note back');

  w._flushPendingNoteSaves();
  await settle();
  assert.strictEqual(srv.posts[srv.posts.length - 1], 'new',
    'the autosave sent the old note back to the server: ' + JSON.stringify(srv.posts));
});

test('typing more while a save is in flight is not exposed to the next reload', async () => {
  const w = app();
  w.saveCaregiversLS({ cg_A: { id: 'cg_A', name: 'Alice Aide', notes: '', _rowVersion: 'aaaaaaaaaaaaaaaa' } });
  const srv = caregiverServer(w, '');
  w.activeCgId = 'cg_A';
  w.renderCgNotesPane();

  srv.hold = true;
  typeCg(w, 'a');
  w._flushPendingNoteSaves();          // save of "a" goes out and is held
  await settle();
  typeCg(w, 'ab');                     // operator keeps typing while it is in flight
  srv.hold = false;
  srv.release();                       // "a" succeeds — the real save clears _unsaved here
  await settle();
  w.loadCaregiversAPI();               // a reload lands while "ab" is still pending
  await settle();

  assert.strictEqual(w.getCaregivers().cg_A.notes, 'ab',
    'the save of "a" cleared the flag while "ab" was unsaved, so the reload reverted it');
});

test('once the typed note is confirmed saved, a reload takes the server copy again', async () => {
  const w = app();
  w.saveCaregiversLS({ cg_A: { id: 'cg_A', name: 'Alice Aide', notes: 'old', _rowVersion: 'aaaaaaaaaaaaaaaa' } });
  const srv = caregiverServer(w, 'old');
  w.activeCgId = 'cg_A';
  w.renderCgNotesPane();

  typeCg(w, 'new');
  w._flushPendingNoteSaves();
  await settle();
  srv.notes = 'edited on another device';
  w.loadCaregiversAPI();
  await settle();

  assert.strictEqual(w.getCaregivers().cg_A.notes, 'edited on another device',
    'the flag must clear once the note is saved, or this device pins its copy over everyone else forever');
});

test('a roster reload inside the autosave window does not revert a typed caseworker note', async () => {
  const w = app();
  w.saveCaseworkersLS([{ id: 'cw_A', name: 'Carol Case', notes: 'old', _rowVersion: 'aaaaaaaaaaaaaaaa' }]);
  const posts = [];
  w.fetch = (url, opt) => {
    const u = String(url), m = (opt && opt.method) || 'GET';
    if (/\/caseworkers$/.test(u) && m === 'POST') { posts.push(JSON.parse(opt.body).notes); return ok({ row_version: 'bbbbbbbbbbbbbbbb' }); }
    if (/\/caseworkers$/.test(u)) return ok([{ id: 'cw_A', name: 'Carol Case', notes: 'old', row_version_hex: 'aaaaaaaaaaaaaaaa' }]);
    return ok([]);
  };
  w.activeCwId = 'cw_A';
  w.renderCwNotesPane();
  const ta = w.document.getElementById('cwNotesArea');
  ta.value = 'new'; ta.dispatchEvent(new w.Event('input'));

  w.loadCaseworkersAPI();
  await settle();
  assert.strictEqual(w.getCaseworkers().find((c) => c.id === 'cw_A').notes, 'new', 'the reload put the old note back');

  w._flushPendingNoteSaves();
  await settle();
  assert.strictEqual(posts[posts.length - 1], 'new', 'the old note was sent back: ' + JSON.stringify(posts));
});

// Client notes are mostly covered already: _clientSig includes clientNotes, so once a session has a
// baseline a typed note reads as unsynced and the merge keeps it. But that baseline is memory-only
// and absent until the first load of a session lands — and the merge deliberately lets the server
// win when there is no baseline. So a note typed during that first load was still exposed.
function clientPane(w) {
  if (!w.document.getElementById('clientNotesArea')) {
    w.document.body.insertAdjacentHTML('beforeend',
      '<textarea id="clientNotesArea"></textarea><div id="invNotesContent"></div><span id="notesSavedFlash"></span>');
  }
}

test('a client note typed before the first load has a baseline survives that load', async () => {
  const w = app();
  clientPane(w);
  w.saveProfilesLS({ Alice: { clientName: 'Alice', clientNotes: 'old', _dbId: 7, invoices: [] } });
  w.eval('if (typeof _clientSyncedMem === "object") { for (var k in _clientSyncedMem) delete _clientSyncedMem[k]; }');
  w.fetch = () => ok({});
  w.activeProfileName = 'Alice';
  w.renderNotesPane();
  const ta = w.document.getElementById('clientNotesArea');
  ta.value = 'new'; ta.dispatchEvent(new w.Event('input'));

  // The first load of the session lands, with the server still holding the old note.
  const server = { Alice: { clientName: 'Alice', clientNotes: 'old', _dbId: 7, invoices: [] } };
  const merged = w._mergeProfilesLoad(server, w.getProfiles());
  assert.strictEqual(merged.Alice.clientNotes, 'new', 'the first load put the old client note back');
});

test('caseworker: typing more while a save is in flight is not exposed to the next reload', async () => {
  const w = app();
  w.saveCaseworkersLS([{ id: 'cw_A', name: 'Carol Case', notes: '', _rowVersion: 'aaaaaaaaaaaaaaaa' }]);
  let serverNotes = '', release = null, hold = true;
  w.fetch = (url, opt) => {
    const u = String(url), m = (opt && opt.method) || 'GET';
    if (/\/caseworkers$/.test(u) && m === 'POST') {
      const b = JSON.parse(opt.body);
      const done = () => { serverNotes = b.notes; return ok({ row_version: 'bbbbbbbbbbbbbbbb' }); };
      if (hold) return new Promise((res) => { release = () => res(done()); });
      return done();
    }
    if (/\/caseworkers$/.test(u)) return ok([{ id: 'cw_A', name: 'Carol Case', notes: serverNotes, row_version_hex: 'aaaaaaaaaaaaaaaa' }]);
    return ok([]);
  };
  w.activeCwId = 'cw_A';
  w.renderCwNotesPane();
  const ta = () => w.document.getElementById('cwNotesArea');
  ta().value = 'a'; ta().dispatchEvent(new w.Event('input'));
  w._flushPendingNoteSaves();
  await settle();
  ta().value = 'ab'; ta().dispatchEvent(new w.Event('input'));
  hold = false; release();
  await settle();
  w.loadCaseworkersAPI();
  await settle();
  assert.strictEqual(w.getCaseworkers().find((c) => c.id === 'cw_A').notes, 'ab',
    'the save of "a" cleared the flag while "ab" was unsaved, so the reload reverted it');
});

test('client: a save that finishes while newer text is pending leaves it flagged', async () => {
  const w = app();
  clientPane(w);
  w.saveProfilesLS({ Alice: { clientName: 'Alice', clientNotes: '', _dbId: 7, invoices: [] } });
  let release = null, hold = true;
  w.fetch = (url, opt) => {
    const m = (opt && opt.method) || 'GET';
    if (m === 'POST' && /homecare-clients/.test(String(url))) {
      if (hold) return new Promise((res) => { release = () => res(ok({ id: 7, row_version: 'bbbbbbbbbbbbbbbb' })); });
      return ok({ id: 7, row_version: 'bbbbbbbbbbbbbbbb' });
    }
    return ok([]);
  };
  w.activeProfileName = 'Alice';
  w.renderNotesPane();
  const ta = w.document.getElementById('clientNotesArea');
  ta.value = 'a'; ta.dispatchEvent(new w.Event('input'));
  w._flushPendingNoteSaves();
  await settle();
  ta.value = 'ab'; ta.dispatchEvent(new w.Event('input'));
  hold = false; if (release) release();
  await settle();
  assert.strictEqual(w.getProfiles().Alice._unsaved, true,
    'the save of "a" cleared the flag while "ab" was unsaved — the next load would revert it');
  assert.strictEqual(w.getProfiles().Alice.clientNotes, 'ab');
});
