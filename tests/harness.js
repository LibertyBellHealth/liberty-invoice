'use strict';
// Loads the real app.js into a jsdom window so its functions can be smoke-tested in Node.
// No copying of logic — we exercise the SAME code that ships.
const { JSDOM } = require('jsdom');
const fs = require('fs');
const path = require('path');

let _win = null;

function loadApp() {
  if (_win) return _win;
  const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
  // The INIT block partway through app.js touches real page elements. When it throws, execution of
  // the REST of the file stops — function declarations still hoist, but every `var x = {...}` after
  // that point (e.g. _ASST_UPDATABLE, STATE_FORM_* maps) is left undefined, which silently hid whole
  // areas from the tests. Providing the handful of elements INIT needs lets the file finish.
  const initDom = `
    <input id="dateSubmitted"><input id="sigDate1"><input id="sigDate2">
    <input id="billingPeriod"><input id="billingPeriod2">
    <table><tbody id="svcBody"></tbody><tbody id="cplxBody"></tbody>
           <tfoot><tr id="svcAllRow"><td></td><td></td></tr>
                  <tr id="cplxAllRow"><td></td><td></td></tr></tfoot></table>
    <div id="activityFeed"></div><span id="taskBadge"></span>
    <div id="page-home" class="page"></div><div id="breadcrumb"></div>
    <div id="topbarActions"></div><div id="clientTableBody"></div>
    <div id="sidebarClients"></div><div id="attentionPanel"></div>
    <div id="sbClientList"></div><div id="clientGrid"></div><div id="clientCount"></div>
    <div id="statTotal"></div><div id="statActive"></div><div id="statInvoices"></div>
    <div id="statOutstanding"></div><div id="lastSyncedLabel"></div><div id="activityList"></div>
    <div id="taskBadgeWrap"></div><div id="undoBanner"></div><div id="oneDriveBanner"></div>`;
  const dom = new JSDOM('<!doctype html><html><head></head><body>' + initDom + '</body></html>', {
    url: 'http://localhost/',        // hostname=localhost → app picks the dev API base (never called here)
    runScripts: 'outside-only',
    pretendToBeVisual: true,
  });
  const { window } = dom;
  // CDN globals app.js references aren't loaded in jsdom; stub the few its startup might touch.
  window.msal = window.msal || {};
  window.appInsights = window.appInsights || {
    trackEvent() {}, trackException() {}, trackPageView() {}, trackTrace() {},
  };
  // Run app.js in global scope. Its INIT block at the very end touches page elements that don't
  // exist here and throws — swallow it. Every function is declared BEFORE init, so they're all
  // defined regardless.
  try { window.eval(src); } catch (e) { /* expected: INIT block DOM access */ }
  _win = window;
  return window;
}

// Fresh localStorage per test group. Also null the tick-scoped profiles read-cache so a test
// never gets a stale {} from a previous test's read (the cache normally clears on a microtask).
function resetStorage(win) { win.localStorage.clear(); win._profilesCache = null; }

// Realm-safe deep compare. Objects/arrays returned from app.js live in the jsdom realm, so
// assert.deepStrictEqual fails on prototype identity vs a Node-realm literal. Compare by value.
function jsonEqual(a, b) { return JSON.stringify(a) === JSON.stringify(b); }

module.exports = { loadApp, resetStorage, jsonEqual };
