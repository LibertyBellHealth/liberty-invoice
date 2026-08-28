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
  // Run app.js in global scope, and let ANY top-level exception fail the suite.
  //
  // This used to be wrapped in `try { ... } catch (e) {}` on the theory that the INIT block always
  // throws in jsdom and every function is hoisted anyway. Neither half held up. app.js evaluates
  // cleanly against the initDom above (its own init paths are internally guarded), so the catch
  // caught nothing real — it just stood ready to swallow the next genuine startup regression. And
  // hoisting only covers function declarations: a `var X = {...}` after the throw point is left
  // undefined, which is exactly how whole areas of the file silently stopped being tested before.
  // If app.js cannot be loaded, that IS the failure — do not hide it behind 249 green checkmarks.
  window.eval(src);
  // Belt and braces: a sentinel from the LAST few hundred lines of the file. If evaluation ever
  // stops early without throwing, this is undefined and every suite fails loudly instead of
  // quietly testing a half-loaded app.
  if (typeof window._ASST_UPDATABLE === 'undefined') {
    throw new Error('harness: app.js did not evaluate to completion — late declarations are missing');
  }
  _win = window;
  return window;
}

// Fresh localStorage per test group. Also null the tick-scoped profiles read-cache so a test
// never gets a stale {} from a previous test's read (the cache normally clears on a microtask).
function resetStorage(win) {
  win.localStorage.clear();
  win._profilesCache = null;
  // loadApp() caches ONE window, so module-scope state survives between tests. The SSN overlays are
  // the dangerous ones: a value cached by an earlier test reappears through getProfiles/getCaregivers
  // and can make a later test pass for the wrong reason (this masked two vacuous tests).
  try { win.eval('_ssnMem = Object.create(null); _clientSyncedMem = Object.create(null); if (typeof _cgSsnMem !== "undefined") _cgSsnMem = Object.create(null);'); } catch (e) {}
}

// Realm-safe deep compare. Objects/arrays returned from app.js live in the jsdom realm, so
// assert.deepStrictEqual fails on prototype identity vs a Node-realm literal. Compare by value.
function jsonEqual(a, b) { return JSON.stringify(a) === JSON.stringify(b); }

module.exports = { loadApp, resetStorage, jsonEqual };
