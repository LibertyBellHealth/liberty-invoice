'use strict';
// A whole family of bugs shipped from ONE mistake, repeated: the app tracks "who you are working
// on" in a global (activeProfileName / activeCgId / activeCwId), and anything asynchronous — a
// fetch, or a confirm dialog waiting on a click — opens a gap in which that global can change.
// Code that read it AFTER the gap wrote to whoever was current then, not who the action started
// for. Because the DOM is reused rather than rebuilt (cg-milogin-pass is a static element cleared
// between caregivers), the write landed in a LIVE field belonging to someone else.
//
// It produced, among others: one client's documents under another's name (#159), a MI Login
// password saved onto the wrong caregiver (#179), an invoice marked Paid on the wrong client
// (#167), and an authorization cleared from the wrong client (#176).
//
// This test is the guard. It fails when a NEW callback reads one of those globals and writes,
// so the family cannot quietly grow back. If a new hit is legitimate, capture the record BEFORE
// the async gap and verify it is still current before writing — then add it here with the reason.
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8').split('\n');

// Reviewed and safe — each reads a global only in a way that cannot write to the wrong record.
const ALLOWED = {
  _clearAuth: 'captures forClient; the global read only decides whether to re-render',
  uploadCgDocAzure: 'clientId is captured BEFORE the fetch; the refresh is guarded by _docListStillCurrent',
  handleCgDocScan: 'same — the upload target is captured before the request',
  deleteCaseworker: 'the callback filters by the captured id, not by the global',
};

function offenders() {
  const fns = []; let cur = null, buf = [], start = 0;
  src.forEach((line, i) => {
    const m = /^(?:async )?function ([_A-Za-z0-9]+)/.exec(line);
    if (m) { if (cur) fns.push([cur, start, buf.join('\n')]); cur = m[1]; buf = [line]; start = i + 1; }
    else if (cur !== null) { buf.push(line); if (line === '}') { fns.push([cur, start, buf.join('\n')]); cur = null; buf = []; } }
  });
  const GLOBALS = /\bactive(ProfileName|CgId|CwId)\b/;
  const WRITE = /save[A-Z][a-zA-Z]*LS\(|saveProfileSP\(|\.value\s*=|innerHTML\s*=/;
  const out = [];
  fns.forEach(([name, line, body]) => {
    const re = /(\.then\(function[^\n]*|showConfirm\([^\n]*function[^\n]*)/g;
    let m;
    while ((m = re.exec(body))) {
      const tail = body.slice(m.index + m[0].length, m.index + m[0].length + 700);
      if (GLOBALS.test(tail) && WRITE.test(tail)) { out.push({ name, line }); break; }
    }
  });
  return out;
}

test('no NEW code writes using a "current record" global read after an async gap', () => {
  const unexpected = offenders().filter((o) => !Object.prototype.hasOwnProperty.call(ALLOWED, o.name));
  assert.deepStrictEqual(unexpected.map((o) => o.name + ' (app.js:' + o.line + ')'), [],
    'This is the bug family that put one client\'s data on another client. Capture the record ' +
    'BEFORE the dialog or fetch, verify it is still current before writing, then add it to ALLOWED.');
});

test('the known-safe list has not silently grown', () => {
  const names = offenders().map((o) => o.name).sort();
  assert.deepStrictEqual(names, Object.keys(ALLOWED).sort(),
    'a reviewed exception disappeared or a new one appeared — re-check both');
});
