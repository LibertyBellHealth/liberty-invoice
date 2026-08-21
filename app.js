// ============================================================
//  STATE
// ============================================================
var SVC=15,CPLX=9,active=31,clipboard={};
var activeProfileName=null,lastLoadedStates=null;
var pendingSigTarget=null,sigDrawing=false,sigCanvas=null,sigCtx=null;
var unsavedChanges=false;

// ============================================================
//  AZURE FUNCTIONS API CONFIG
// ============================================================
// Dev sandbox: when served from localhost, talk to the DEV backend (liberty-crm-db-dev,
// fake data) so local testing never touches production. Every DEPLOYED hostname (the SWA
// prod site + PR previews) always uses the prod backend — this switch is localhost-only,
// so it is safe to ship. Run locally on a FIXED port (4280) that is registered as an
// Azure AD redirect URI, e.g.:  cd liberty-invoice-site && python3 -m http.server 4280
var _IS_LOCAL   = (location.hostname === 'localhost' || location.hostname === '127.0.0.1');
// The FIXED staging URL is a throwaway sandbox on the DEV backend (fake data) so destructive testing
// there can NEVER touch a real client record. localhost also uses dev. Every OTHER deployed host
// (the prod site + random PR previews) uses prod. Prod's own hostname never matches this, so prod is
// unaffected; its CSP stays tight (only the staging branch's CSP allows the dev backend host).
var _IS_STAGING = (location.hostname === 'zealous-forest-01e406a10-2.centralus.7.azurestaticapps.net');
var API_BASE    = (_IS_LOCAL || _IS_STAGING)
  ? 'https://liberty-crm-api-dev.azurewebsites.net/api'
  : 'https://liberty-crm-api-cyb3dkhnd2e7a3cy.centralus-01.azurewebsites.net/api';
var API_APP_ID  = '0c1627c1-c186-4e46-b919-e4a12f2f3952'; // Easy Auth app registration
var _apiToken   = null; // cached Bearer token, refreshed automatically
var _savesInFlight = 0; // count of in-flight entity saves (client + roster); blocks a background reload from clobbering them
var _apiTokenTimer = null; // handle for the self-rescheduling refresh; cleared on sign-out/idle

// Microsoft identity — kept for Sign In and Outlook email
var SP_CLIENT_ID = (window._SP_CLIENT_ID || '1be40fcb-4db4-45b0-8c12-99d945eb78e7');
var SP_TENANT_ID = (window._SP_TENANT_ID || '12be0d3c-3e63-429f-bf46-1a2f746aa25f');

// ── ALLOWED USERS — add/remove emails here to control access ──
// All addresses must be @mybellcare.com accounts you have added in Azure AD
var ALLOWED_USERS = [
  'tommy@mybellcare.com',
  'paul@mybellcare.com',
  'rob@mybellcare.com'
];
// Use current origin so it works across desktop and mobile (any URL the app is loaded from).
// The redirect URI must also be registered in the AAD app registration's "Single-page application" platform.
var REDIRECT_URI = window.location.origin;
var API_SCOPE    = 'api://' + API_APP_ID + '/user_impersonation';
var GRAPH_SCOPES = ['openid', 'profile', 'https://graph.microsoft.com/Mail.Send', 'https://graph.microsoft.com/Files.ReadWrite']; // Graph only — never mix with API_SCOPE
var msalInstance = null, spToken = null;

function apiHeaders() {
  var h = { 'Content-Type': 'application/json' };
  if (_apiToken) h['Authorization'] = 'Bearer ' + _apiToken;
  return h;
}
// For multipart/form-data uploads — no Content-Type (browser sets boundary automatically)
function authUploadHeaders() {
  return _apiToken ? { 'Authorization': 'Bearer ' + _apiToken } : {};
}
async function refreshApiToken() {
  if (!msalInstance) return;
  var accounts = msalInstance.getAllAccounts();
  if (!accounts.length) return;
  try {
    var res = await msalInstance.acquireTokenSilent({
      scopes: [API_SCOPE],
      account: accounts[0]
    });
    _apiToken = res.accessToken;
    // Auto-refresh 10 min before expiry (tokens last ~1 hour)
    var ttl = res.expiresOn ? (res.expiresOn.getTime() - Date.now() - 600000) : 3000000;
    _apiTokenTimer = setTimeout(refreshApiToken, Math.max(ttl, 60000));
  } catch(e) {
    console.warn('API token silent refresh failed, falling back to redirect:', e);
    // Redirect works on mobile (popup is blocked on iOS Safari)
    try {
      await msalInstance.acquireTokenRedirect({ scopes: [API_SCOPE] });
    } catch(e2) { console.warn('API token redirect failed:', e2); }
  }
}

// ── Reactive auth-expiry handling ────────────────────────────────────────────
// The proactive timer above refreshes ~10min before expiry, but it MISSES when a device sleeps
// (laptop lid closed) and wakes with an already-expired token: setTimeout doesn't fire on a
// sleeping tab, so the very next save/upload 401s silently while the UI still looks signed-in.
// Unattended workers on their own devices hit exactly this. So we also react to any 401 from OUR
// API: attempt a silent refresh (fixes the token for the next call) AND show a visible banner so a
// worker never keeps typing into saves that are quietly failing. A later successful API call
// clears the banner.
var _authBannerShown = false, _authRecovering = false;
function _showAuthExpiredBanner(){
  if(_authBannerShown) return; _authBannerShown = true;
  var bar = document.getElementById('authExpiredBar');
  if(!bar){
    bar = document.createElement('div');
    bar.id = 'authExpiredBar';
    bar.style.cssText = 'position:fixed;top:0;left:0;right:0;z-index:99999;background:#b03030;color:#fff;'+
      'font-size:13px;font-weight:600;text-align:center;padding:10px 14px;box-shadow:0 2px 8px rgba(0,0,0,.25);';
    bar.innerHTML = 'Your sign-in expired — a recent change may not have saved. '+
      '<a href="#" style="color:#fff;text-decoration:underline;" onclick="_reauthNow();return false;">Sign in again</a>';
    (document.body||document.documentElement).appendChild(bar);
  }
  bar.style.display = 'block';
}
function _hideAuthExpiredBanner(){ _authBannerShown = false; var b=document.getElementById('authExpiredBar'); if(b) b.style.display='none'; }
function _reauthNow(){
  try{ if(msalInstance){ msalInstance.acquireTokenRedirect({ scopes:[API_SCOPE] }); return; } }
  catch(e){ console.warn('reauth redirect failed:', e); }
  try{ location.reload(); }catch(e2){}
}
// Called when an API response returns 401. Shows the banner and kicks ONE silent refresh so the
// next call can succeed; we never auto-redirect (that would yank a worker mid-edit) — the banner's
// link does that on demand.
function _onApiAuthFail(){
  _showAuthExpiredBanner();
  if(_authRecovering) return; _authRecovering = true;
  Promise.resolve().then(function(){
    if(!msalInstance) return;
    var accounts = msalInstance.getAllAccounts();
    if(!accounts.length) return;
    return msalInstance.acquireTokenSilent({ scopes:[API_SCOPE], account:accounts[0] })
      .then(function(res){ if(res && res.accessToken) _apiToken = res.accessToken; })
      .catch(function(){ /* leave the banner up; the worker must click Sign in again */ });
  }).catch(function(){}).then(function(){ _authRecovering = false; });
}
// One-time fetch wrapper: watch ONLY our API responses. On 401 → surface auth expiry; on any OK
// response → clear a stale banner. Everything is passed through untouched (the body is never read).
(function installApiAuthWatch(){
  if(typeof window==='undefined' || window.__apiAuthWatch) return;
  var _origFetch = window.fetch;
  if(typeof _origFetch !== 'function') return; // jsdom test env has no fetch — no-op there
  window.__apiAuthWatch = true;
  window.fetch = function(input, init){
    var p = _origFetch.apply(this, arguments);
    try{
      var url = (typeof input === 'string') ? input : (input && input.url) || '';
      if(url && typeof API_BASE === 'string' && API_BASE && url.indexOf(API_BASE) === 0){
        p.then(function(resp){
          if(!resp) return;
          if(resp.status === 401) _onApiAuthFail();
          else if(resp.ok && _authBannerShown) _hideAuthExpiredBanner();
        }, function(){});
      }
    }catch(e){}
    return p;
  };
})();

function getIdMap() {
  try { return JSON.parse(localStorage.getItem('lhca_id_map') || '{}'); } catch(e) { return {}; }
}
// Live-format an SSN input as the user types: digits only, capped at 9, auto-dashed
// to XXX-XX-XXXX. Wired to oninput on every SSN field.
function formatSSN(el){
  if(!el)return;
  var d=(el.value||'').replace(/\D/g,'').slice(0,9);
  if(d.length>5) el.value=d.slice(0,3)+'-'+d.slice(3,5)+'-'+d.slice(5);
  else if(d.length>3) el.value=d.slice(0,3)+'-'+d.slice(3);
  else el.value=d;
}
// DB2: keep a money field numeric — digits + a single decimal point, max 2 decimals.
// So "$27/hr" or "twenty" can't be typed into pay_rate. Wired to oninput on rate fields.
function formatRate(el){
  if(!el)return;
  var v=(el.value||'').replace(/[^0-9.]/g,'');
  var parts=v.split('.');
  el.value=parts[0]+(parts.length>1 ? '.'+parts.slice(1).join('').slice(0,2) : '');
}
// C4: one place for CLIENT status labels — 'inactive' displays as "In Progress"
// (onboarding, not yet billing). Keeps every client-status badge consistent.
function clientStatusLabel(s){ s=s||'active'; return s==='inactive' ? 'In Progress' : s.charAt(0).toUpperCase()+s.slice(1); }
// A small pill showing the client's program: "CHAMPS" (green) or the carrier name (blue) for
// managed-care. Self-contained inline styles so no CSS file change is needed. Absent = CHAMPS.
function _programBadge(prof){
  if(!prof)return '';
  var carrier=prof.program==='carrier';
  var label=carrier?(prof.carrier||'Managed Care'):'CHAMPS';
  var bg=carrier?'#e8f0fe':'#eaf5ec', fg=carrier?'#1a56b8':'#1a7740';
  return '<span style="display:inline-block;font-size:10px;font-weight:700;padding:1px 7px;border-radius:10px;background:'+bg+';color:'+fg+';vertical-align:middle;">'+esc(label)+'</span>';
}
// Caseworker Org pill: MDHHS (green) vs a managed-care carrier (blue). Empty org → no badge.
function _cwOrgBadge(cw){
  var org=(cw&&cw.org||'').trim();
  if(!org)return '';
  var isMdhhs=/mdhhs/i.test(org);
  var bg=isMdhhs?'#eaf5ec':'#e8f0fe', fg=isMdhhs?'#1a7740':'#1a56b8';
  return '<span style="display:inline-block;font-size:10px;font-weight:700;padding:1px 7px;border-radius:10px;background:'+bg+';color:'+fg+';vertical-align:middle;">'+esc(org)+'</span>';
}

// ============================================================
//  NAVIGATION
// ============================================================
function showPage(p){
  document.querySelectorAll('.page').forEach(function(el){el.classList.remove('active');});
  document.getElementById('page-'+p).classList.add('active');
  ['sb-home','sb-caregivers','sb-settings','sb-tasks','sb-reports','sb-caseworkers','sb-forms'].forEach(function(id){
    var el=document.getElementById(id);if(el)el.classList.remove('active');
  });
  var map={'home':'sb-home','caregivers':'sb-caregivers','settings':'sb-settings','tasks':'sb-tasks','reports':'sb-reports','caseworkers':'sb-caseworkers','forms':'sb-forms','form-fill':'sb-forms'};
  if(map[p]){var el=document.getElementById(map[p]);if(el)el.classList.add('active');}
  var db=document.getElementById('draftBadge');if(db)db.style.display='none';
}
function bc(crumbs){
  var el=document.getElementById('breadcrumb');el.innerHTML='';
  crumbs.forEach(function(c,i){
    if(i>0){var s=document.createElement('span');s.className='tb-sep';s.textContent='›';el.appendChild(s);}
    var s=document.createElement('span');
    if(i<crumbs.length-1){s.className='tb-crumb';s.textContent=c.l;s.onclick=c.fn;}
    else{s.className='tb-active';s.textContent=c.l;}
    el.appendChild(s);
  });
}
function navHome(){showPage('home');bc([{l:'Clients'}]);document.getElementById('topbarActions').innerHTML='';unsavedChanges=false;renderClientTable();updateStats();renderSidebarClients();if(typeof revalidate==='function')revalidate();}

// ============================================================
//  HASH ROUTER — enables right-click "Open in new tab" for records
//  URL format: #/client/<id>, #/caregiver/<id>, #/caseworker/<id>,
//              #/forms, #/tasks, #/caregivers, #/caseworkers, #/settings
// ============================================================
// Route clients by opaque dbId so the patient NAME never lands in the URL / browser history.
// Falls back to the name only for a client not yet saved to the server (no dbId yet).
function buildClientUrl(name){
  var id=(getIdMap()||{})[name];
  return '#/client/'+encodeURIComponent(id!=null&&id!==''?id:name);
}
function buildCaregiverUrl(id){return '#/caregiver/'+encodeURIComponent(id);}
function buildCaseworkerUrl(id){return '#/caseworker/'+encodeURIComponent(id);}

function routeFromHash(){
  var hash=(window.location.hash||'').replace(/^#\/?/, '');
  if(!hash){if(typeof navHome==='function')navHome();return;}
  var parts=hash.split('/').map(decodeURIComponent);
  var route=parts[0];
  try{
    if(route==='client'&&parts[1]){
      var seg=parts[1], nm=seg;
      // seg is an opaque dbId (all digits) → resolve back to the client name. A non-numeric
      // seg is a legacy #/client/<name> link (old bookmark / open tab) — use it as-is.
      if(/^\d+$/.test(seg)){
        var idMap=getIdMap()||{}; nm=null;
        for(var k in idMap){ if(String(idMap[k])===seg){ nm=k; break; } }
        if(!nm){ if(typeof navHome==='function')navHome(); return; }
      }
      activeProfileName=nm;
      if(typeof navDetail==='function')navDetail(nm,parts[2]||null);
    } else if(route==='caregiver'&&parts[1]){
      if(typeof navCaregivers==='function')navCaregivers();
      setTimeout(function(){if(typeof openCgDetail==='function')openCgDetail(parts[1]);},80);
    } else if(route==='caseworker'&&parts[1]){
      if(typeof navCaseworkers==='function')navCaseworkers();
      setTimeout(function(){if(typeof openCwDetail==='function')openCwDetail(parts[1]);},80);
    } else if(route==='caregivers'){if(typeof navCaregivers==='function')navCaregivers();}
    else if(route==='caseworkers'){if(typeof navCaseworkers==='function')navCaseworkers();}
    else if(route==='forms'){if(typeof navForms==='function')navForms();}
    else if(route==='tasks'){if(typeof navTasks==='function')navTasks();}
    else if(route==='reports'){if(typeof navReports==='function')navReports();}
    else if(route==='settings'){if(typeof navSettings==='function')navSettings();}
    else if(typeof navHome==='function')navHome();
  }catch(e){console.warn('Route error:',e);}
}
window.addEventListener('hashchange',routeFromHash);
// On a cold load / hard refresh the router runs before the roster has downloaded (and a refresh
// wipes the local roster for HIPAA), so a deep link like #/client/5 can't resolve yet and falls
// back to the list. Re-run the router once a roster loader finishes — but only while no detail is
// open, so a background revalidation never yanks a user off a record they're viewing.
function _restoreRouteAfterLoad(){
  var h=window.location.hash||'';
  if(!h || h==='#/' || h==='#') return;
  var onDetail=(typeof activeProfileName!=='undefined'&&activeProfileName)
             ||(typeof activeCgId!=='undefined'&&activeCgId)
             ||(typeof activeCwId!=='undefined'&&activeCwId);
  if(onDetail) return;
  if(typeof routeFromHash==='function') routeFromHash();
}

// Explicit nav click helper — guarantees the target page renders even when the
// current URL already matches the target hash (in which case hashchange doesn't
// fire and the router would silently skip the click). Cmd/Ctrl/Shift-click still
// falls through to the browser's default behavior so "open in new tab" works.
function navClick(e, targetHash, targetFn){
  if(e && (e.metaKey||e.ctrlKey||e.shiftKey||e.button===1))return true;
  if(e && typeof e.preventDefault==='function')e.preventDefault();
  var cur=window.location.hash||'';
  if(cur===targetHash){
    // Already on target — force a re-render since hashchange won't fire. Record links pass no
    // targetFn, so fall back to re-running the router for the current hash (fixes: re-clicking the
    // client/caregiver/caseworker already in the URL did nothing).
    if(typeof targetFn==='function')targetFn();
    else if(typeof routeFromHash==='function')routeFromHash();
  } else {
    // hashchange will fire and routeFromHash will call the appropriate nav function
    window.location.hash=targetHash;
  }
  return false;
}

// ============================================================
//  MOBILE SIDEBAR — open via hamburger, close via backdrop tap or nav
// ============================================================
function toggleMobileSidebar(){
  var sb=document.getElementById('sidebar');
  if(!sb)return;
  var opening=!sb.classList.contains('mobile-open');
  sb.classList.toggle('mobile-open');
  var existing=document.getElementById('sbBackdrop');
  if(opening){
    if(!existing){
      var bd=document.createElement('div');
      bd.id='sbBackdrop';
      bd.style.cssText='position:fixed;inset:0;background:rgba(0,0,0,0.35);z-index:999;cursor:pointer;';
      bd.addEventListener('click',closeMobileSidebar);
      document.body.appendChild(bd);
    }
  } else if(existing){
    existing.remove();
  }
}
function closeMobileSidebar(){
  var sb=document.getElementById('sidebar');
  if(sb)sb.classList.remove('mobile-open');
  var bd=document.getElementById('sbBackdrop');
  if(bd)bd.remove();
}
// Auto-close when any sidebar nav link or client row is tapped
document.addEventListener('click',function(e){
  var sb=document.getElementById('sidebar');
  if(!sb||!sb.classList.contains('mobile-open'))return;
  // If the tap landed inside the sidebar and on an <a> or .sb-btn, close it after navigation
  if(sb.contains(e.target)){
    var link=e.target.closest('a,.sb-btn');
    if(link)closeMobileSidebar();
  }
});

function navNewClient(){
  showPage('new-client');bc([{l:'Clients',fn:navHome},{l:'New Client'}]);document.getElementById('topbarActions').innerHTML='';
  ['nc-first','nc-middle','nc-last','nc-nickname','nc-program','nc-carrier','nc-member','nc-medicaid','nc-medicare','nc-rate','nc-dl','nc-ssn','nc-phone','nc-cemail','nc-street','nc-city','nc-state','nc-zip','nc-county','nc-start-date','nc-worker-search','nc-worker-val','nc-caregiver-search','nc-caregiver-val'].forEach(function(id){var e=document.getElementById(id);if(e)e.value='';});
  if(typeof ncProgramToggle==='function')ncProgramToggle();   // re-hide carrier fields until a program is chosen
  var drop=document.getElementById('nc-worker-drop');if(drop)drop.style.display='none';
  var cgDrop=document.getElementById('nc-caregiver-drop');if(cgDrop)cgDrop.style.display='none';
}
function navDetail(name,tab){
  aiTrack('ClientRecordOpened',{client:name,tab:tab||'info'});
  activeProfileName=name;var prof=getProfiles()[name];if(!prof)return;
  showPage('detail');
  var ini=name.split(' ').map(function(w){return w[0]||'';}).slice(0,2).join('').toUpperCase();
  document.getElementById('detailAvatar').textContent=ini;
  document.getElementById('detailName').textContent=name+(prof.nickname?' ('+prof.nickname+')':'');
  var st=prof.clientStatus||'active';
  document.getElementById('detailMeta').innerHTML=(prof.medicaidId?'Medicaid: '+esc(prof.medicaidId):'No Medicaid ID')+(prof.phone?' &nbsp;·&nbsp; '+esc(prof.phone):'')+
    ' &nbsp;<span class="cs-badge cs-'+st+'">'+clientStatusLabel(st)+'</span>'+' &nbsp;'+_programBadge(prof);
  // Carrier (managed-care) clients aren't invoiced in the CRM — hide the + New Invoice action.
  var _nib=document.getElementById('detailNewInvoiceBtn'); if(_nib)_nib.style.display=isCarrierClient(prof)?'none':'';
  bc([{l:'Clients',fn:navHome},{l:name}]);
  document.getElementById('topbarActions').innerHTML='';
  switchTab(tab||'overview');renderSidebarClients();
}
function navInvoice(loadSpecific){
  if(!activeProfileName){showAlert('Select a client first.');return;}
  var prof=getProfiles()[activeProfileName];
  if(isCarrierClient(prof)){showAlert('This is a managed-care (carrier) client — billing goes through the carrier’s software, so invoices aren’t created here.');return;}
  if(loadSpecific){
    showPage('invoice');
    document.getElementById('invClientTag').textContent=activeProfileName;
    document.getElementById('saveInvoiceBtn').style.display='inline-block';
    bc([{l:'Clients',fn:navHome},{l:activeProfileName,fn:function(){navDetail(activeProfileName);}},{l:'Invoice'}]);
    document.getElementById('topbarActions').innerHTML='';
    document.getElementById('dupWarning').style.display='none';
    applyFullInvoice(loadSpecific);
    syncBillingPeriodFields();
    return;
  }
  var invs=prof.invoices||[];
  var copyBtn=document.getElementById('copyLastInvBtn');
  var copyMeta=document.getElementById('copyLastInvMeta');
  if(invs.length){
    copyBtn.disabled=false;copyBtn.style.opacity='1';
    copyMeta.textContent='Copy tasks from: '+invs[0].billingPeriod+' ('+invs[0].savedAt.split(',')[0]+')';
  } else {
    copyBtn.disabled=true;copyBtn.style.opacity='0.45';
    copyMeta.textContent='No previous invoices for this client';
  }
  document.getElementById('newInvChoiceModal').classList.add('open');
}
function confirmNewInvoice(mode){
  var prof=getProfiles()[activeProfileName];
  if(isCarrierClient(prof)){showAlert('This is a managed-care (carrier) client — billing goes through the carrier’s software, so no invoice is created here.');return;}
  document.getElementById('newInvChoiceModal').classList.remove('open');
  showPage('invoice');
  document.getElementById('invClientTag').textContent=activeProfileName;
  document.getElementById('saveInvoiceBtn').style.display='inline-block';
  bc([{l:'Clients',fn:navHome},{l:activeProfileName,fn:function(){navDetail(activeProfileName);}},{l:'Invoice'}]);
  document.getElementById('topbarActions').innerHTML='';
  document.getElementById('dupWarning').style.display='none';
  // Kill any stale draft (legacy cleanup — draft-autosave was removed)
  try{localStorage.removeItem('lhca_draft_'+activeProfileName);}catch(e){}
  if(mode==='copy'&&prof.invoices&&prof.invoices.length){
    // Copy tasks AND total hours from last invoice; auto-advance to next month
    var last=prof.invoices[0];
    applyFullInvoice(last.data);
    // Compute next month's billing period from the last invoice's period
    var nextBP='';
    var lastBP=last.billingPeriod||'';
    var bpParts=lastBP.split('/');
    if(bpParts.length===2&&bpParts[1].length===4){
      var lm=parseInt(bpParts[0]),ly=parseInt(bpParts[1]);
      if(!isNaN(lm)&&!isNaN(ly)){
        lm++;if(lm>12){lm=1;ly++;}
        nextBP=String(lm).padStart(2,'0')+'/'+ly;
      }
    }
    document.getElementById('billingPeriod').value=nextBP;
    document.getElementById('billingPeriod2').value=nextBP;
    var T=today();
    document.getElementById('dateSubmitted').value=T;
    document.getElementById('sigDate1').value=T;
    document.getElementById('sigDate2').value=T;
    // Rebuild grid for the new month BEFORE applying states (so day count matches)
    if(nextBP){
      var partsNew=nextBP.split('/');
      var savedStates=captureStates();
      rebuild(daysIn(partsNew[0],partsNew[1]));
      applyStates(savedStates);
      checkDuplicatePeriod(nextBP);
    }
    // Keep hours fields populated from last invoice (do NOT clear them)
    document.getElementById('dupWarning').style.display='none';
  } else {
    loadProfileIntoForm(prof);
    // Default billing period to the previous month (invoices are billed in arrears)
    // instead of leaving it blank — saves re-typing MM/YYYY every cycle.
    var _pm=new Date();_pm=new Date(_pm.getFullYear(),_pm.getMonth()-1,1);
    var _mm=String(_pm.getMonth()+1).padStart(2,'0'),_yy=String(_pm.getFullYear());
    var _defBP=_mm+'/'+_yy;
    document.getElementById('billingPeriod').value=_defBP;
    document.getElementById('billingPeriod2').value=_defBP;
    var T=today();
    document.getElementById('dateSubmitted').value=T;
    document.getElementById('sigDate1').value=T;
    document.getElementById('sigDate2').value=T;
    ['svcHH','svcMM','cplxHH','cplxMM','p1HH','p1MM','grandHH','grandMM'].forEach(function(id){var e=document.getElementById(id);if(e)e.value='';});
    rebuild(daysIn(_mm,_yy));
    checkDuplicatePeriod(_defBP);
    resetSigArea(1);resetSigArea(2);
  }
}
function loadInvFromHistory(idx){
  // Open invoice as a file — no confirmation needed
  var inv=getProfiles()[activeProfileName]&&getProfiles()[activeProfileName].invoices[idx];
  if(!inv)return;
  navInvoice(inv.data);
}
function navCaregivers(){
  showPage('caregivers');bc([{l:'Caregivers'}]);document.getElementById('topbarActions').innerHTML='';
  hideCgForm();
  document.getElementById('cgDetailView').style.display='none';
  document.getElementById('cgGridView').style.display='';
  activeCgId=null;
  renderCaregiverGrid();
  // Cold cache (nothing stored yet): fetch directly, bypassing the 30s revalidate throttle.
  if(typeof spToken!=='undefined'&&spToken&&Object.keys(getCaregivers()).length===0&&typeof loadCaregiversAPI==='function')loadCaregiversAPI();
  if(typeof revalidate==='function')revalidate();
}
function navSettings(){showPage('settings');bc([{l:'Settings'}]);document.getElementById('topbarActions').innerHTML='';renderSigSettings();updateSettingsAuth();renderEmailAuditTable();if(typeof loadSigningTemplates==='function')loadSigningTemplates();if(typeof revalidate==='function')revalidate();var sr=document.getElementById('stateRateInput');if(sr)sr.value=stateRate();var srs=document.getElementById('stateRateStatus');if(srs)srs.textContent='';if(typeof renderAgencySettings==='function')renderAgencySettings();}
function navTasks(){showPage('tasks');bc([{l:'Tasks'}]);document.getElementById('topbarActions').innerHTML='';populateTodoClientSelect();renderTodos();if(typeof revalidate==='function')revalidate();}
function navReports(){showPage('reports');bc([{l:'Reports'}]);document.getElementById('topbarActions').innerHTML='';renderReports();}

// ============================================================
//  SIDEBAR
// ============================================================
function toggleSidebar(){document.getElementById('sidebar').classList.toggle('collapsed');}
function renderSidebarClients(){
  var list=document.getElementById('sbClientList'),profiles=getProfiles();
  // Active clients only in sidebar — terminated/inactive/lost stay accessible via main Clients page filter
  var keys=Object.keys(profiles).filter(function(k){return (profiles[k].clientStatus||'active')==='active';});
  list.innerHTML='';
  if(!keys.length){list.innerHTML='<div style="padding:8px 14px;font-size:11px;color:#435f7a;">No clients yet.</div>';return;}
  keys.slice(0,15).forEach(function(name){
    var ini=name.split(' ').map(function(w){return w[0]||'';}).slice(0,2).join('').toUpperCase();
    // Use <a href> so right-click → "Open in new tab" works natively
    var row=document.createElement('a');
    row.className='sb-client-row'+(name===activeProfileName?' active':'');
    row.href=buildClientUrl(name);
    row.style.textDecoration='none';
    var sbDisplay=name+(profiles[name].nickname?' ('+profiles[name].nickname+')':'');
    row.innerHTML='<div class="sb-avatar">'+ini+'</div><div class="sb-client-info"><div class="sb-client-name">'+esc(sbDisplay)+'</div><div class="sb-client-meta">'+esc(profiles[name].medicaidId||'No ID')+'</div></div>';
    list.appendChild(row);
  });
  if(keys.length>15){
    var m=document.createElement('a');
    m.href='#/';
    m.style.cssText='padding:5px 14px;font-size:10px;color:#435f7a;cursor:pointer;display:block;text-decoration:none;';
    m.textContent='+'+(keys.length-15)+' more — view all';
    list.appendChild(m);
  }
}

// ============================================================
//  HOME
// ============================================================
function updateStats(){
  var p=getProfiles(),keys=Object.keys(p),ti=0,outstanding=0;
  var activeKeys=keys.filter(function(k){return !p[k].clientStatus||p[k].clientStatus==='active';});
  // Only count outstanding invoices for ACTIVE clients (terminated/inactive/lost shouldn't appear in dashboard alerts)
  keys.forEach(function(k){
    var invs=(p[k].invoices)||[];ti+=invs.length;
    var st=p[k].clientStatus||'active';
    if(st!=='active')return;
    invs.forEach(function(inv){if(!inv.status||inv.status==='draft'||inv.status==='submitted')outstanding++;});
  });
  var s=document.getElementById('statTotal');if(s)s.textContent=activeKeys.length;
  var si=document.getElementById('statInvoices');if(si)si.textContent=ti;
  var sc=document.getElementById('statCaregivers');if(sc)sc.textContent=Object.keys(getCaregivers()).length;
  var so=document.getElementById('statOutstanding');if(so)so.textContent=outstanding;
  var d=new Date();var sm=document.getElementById('statMonth');if(sm)sm.textContent=String(d.getMonth()+1).padStart(2,'0')+'/'+d.getFullYear();
  var sel=document.getElementById('filterCaregiver');
  if(sel){var cur=sel.value;sel.innerHTML='<option value="">All Caregivers</option>';var cgs=getCaregivers();Object.keys(cgs).forEach(function(id){var o=document.createElement('option');o.value=id;o.textContent=cgs[id].name;if(id===cur)o.selected=true;sel.appendChild(o);});}
  renderAttentionPanel();
  if(typeof renderUndoBanner==='function')renderUndoBanner();
}
// Returns true if the given billing period (MM/YYYY) is on or after the client's startDate
function clientWasActiveInPeriod(prof,period){
  if(!prof.startDate)return true; // no start date set — assume always active
  if(!period||period.length<7)return true;
  var pp=period.split('/');if(pp.length!==2)return true;
  var pYear=parseInt(pp[1]),pMonth=parseInt(pp[0]);
  // Period start = first day of that month
  var periodEnd=new Date(pYear,pMonth,0); // last day of period month
  var startD=new Date(prof.startDate);
  return periodEnd>=startD;
}
// A client should only be flagged as "missing an invoice" for a period if they have a service
// start date on/before that period. No start date -> we can't say they were active, so don't nag
// (bug #10: clients with no start date, or a start date after the period, were being flagged).
// Managed-care (carrier) clients are billed through the carrier's own software — they never get a
// CRM invoice, so they're excluded from every invoicing/"missing invoice" surface.
function isCarrierClient(prof){ return !!prof && prof.program==='carrier'; }
// Owner decision (2026-08-21): only an ACTIVE client is invoiced. "In Progress" (stored as
// 'inactive'), Lost and Terminated clients are excluded from every invoicing surface — generation,
// the missing-invoice report, the monthly preview and the batch send. This lives here, in the single
// predicate, because the four surfaces each carried their own inline status test and had drifted:
// updateMissingReport relied on clientDueForInvoice, which never checked status at all, so it nagged
// about clients who were not being served yet.
function isInvoiceableStatus(prof){ return ((prof&&prof.clientStatus)||'active')==='active'; }
function clientDueForInvoice(prof, period){
  return !!(prof && prof.startDate) && !isCarrierClient(prof) && isInvoiceableStatus(prof) &&
         clientWasActiveInPeriod(prof, period);
}
// True when a client has a real DHS-1210 authorization on file (hours, tasks, or an effective date).
// A DHS-1210 is what officially makes you the client's agency, so it gates the Active status.
function hasAuthorization(prof){
  var a=prof&&prof.authorization; if(!a)return false;
  return a.hours!=null || (a.tasks && a.tasks.length>0) || !!a.effectiveDate;
}
// Stricter: the authorization actually carries APPROVED HOURS, so an invoice built from it has a
// real Total Time. hasAuthorization() passes on an effective date alone (a partial OCR read), and
// generating from that produced a blank-total invoice which then SATISFIED the "already has an
// invoice for this period" check — so the client silently stopped showing as missing one.
function hasBillableAuthorization(prof){
  var a=prof&&prof.authorization; if(!a)return false;
  var h=a.hours; if(h==null||String(h).trim()==='')return false;
  return !isNaN(parseFloat(h));
}
function renderAttentionPanel(){
  var panel=document.getElementById('attentionPanel');if(!panel)return;
  var p=getProfiles(),items=[];
  var d=new Date();
  // Flag PREVIOUS month — current month is generated on the 1st of next month, so it's not "missing" yet
  var prev=new Date(d.getFullYear(),d.getMonth()-1,1);
  var prevPeriod=String(prev.getMonth()+1).padStart(2,'0')+'/'+prev.getFullYear();
  var active=Object.keys(p).filter(function(k){return !p[k].clientStatus||p[k].clientStatus==='active';});
  var missingPrev=active.filter(function(k){return clientDueForInvoice(p[k],prevPeriod) && !((p[k].invoices)||[]).some(function(i){return i.billingPeriod===prevPeriod;});});
  if(missingPrev.length){
    items.push({cls:'attn-warn',count:missingPrev.length,label:missingPrev.length+' active client'+(missingPrev.length>1?'s':'')+' missing invoice for '+prevPeriod,fn:'showMissingInvoicesModal(\''+prevPeriod+'\')'});
  }
  var stale=[];Object.keys(p).forEach(function(k){(p[k].invoices||[]).forEach(function(inv){if(inv.status==='submitted'){var age=(new Date()-new Date(inv.savedAt))/(1000*60*60*24);if(age>30)stale.push(1);}});});
  if(stale.length)items.push({cls:'attn-danger',count:stale.length,label:stale.length+' submitted invoice'+(stale.length>1?'s':'')+' pending 30+ days — follow up on payment',fn:'openAllInvoicesModal("outstanding")'});
  var overdueTasks=getTodos().filter(function(t){return !t.done&&t.due&&new Date(t.due)<new Date();});
  if(overdueTasks.length)items.push({cls:'attn-warn',count:overdueTasks.length,label:overdueTasks.length+' overdue task'+(overdueTasks.length>1?'s':''),fn:'navTasks()'});
  if(!items.length)items.push({cls:'attn-ok',count:'',label:'No items require attention today',fn:null});
  panel.innerHTML=items.map(function(it){
    return '<div class="attn-item '+it.cls+'"'+(it.fn?' onclick="'+it.fn+'"':'')+'>'+
      (it.count?'<span class="attn-count">'+it.count+'</span>':'')+
      '<span class="attn-label">'+it.label+'</span>'+(it.fn?'<span class="attn-arrow">→</span>':'')+
    '</div>';
  }).join('');
}
var bulkSelected={};
// Sort state for the Clients table — column-header clicks toggle this
var _clientSort={key:'name',dir:'asc'};
function sortClientBy(key){
  if(_clientSort.key===key)_clientSort.dir=_clientSort.dir==='asc'?'desc':'asc';
  else {_clientSort.key=key;_clientSort.dir='asc';}
  renderClientTable();
}
function _clientSortCompare(a,b,profiles,cgs,cwsArr){
  var pa=profiles[a],pb=profiles[b];
  var key=_clientSort.key,dir=_clientSort.dir==='asc'?1:-1;
  function val(k,p,name){
    if(k==='name')return (name||'').toLowerCase();
    if(k==='status')return (p.clientStatus||'active');
    if(k==='caregiver'){var cg=p.caregiverId&&cgs[p.caregiverId]?cgs[p.caregiverId].name:'';return (cg||'').toLowerCase();}
    if(k==='caseworker'){
      var cw=null;
      if(p.caseworkerId && Array.isArray(cwsArr))cw=cwsArr.find(function(c){return c.id===p.caseworkerId;});
      return (cw?cw.name:(p.worker||'')).toLowerCase();
    }
    if(k==='county')return (p.county||'').toLowerCase();
    if(k==='phone')return (p.phone||'').toLowerCase();
    if(k==='lastInvoice'){var inv=(p.invoices||[])[0];return inv?new Date(inv.savedAt).getTime():0;}
    return 0;
  }
  var va=val(key,pa,a),vb=val(key,pb,b);
  if(typeof va==='string'){return va.localeCompare(vb)*dir;}
  return (va<vb?-1:va>vb?1:0)*dir;
}

// ── Client-table column widths + pagination — persisted per-device in localStorage ──
// Widths are % of table width (excluding the 32px checkbox column). Should sum to ~95-100.
// Status is small because "• Active" is short; Client is largest because it stacks name + Medicaid ID.
// Column widths sum to 100 so every column gets exactly the space it deserves —
// no browser auto-expansion. Name column trimmed from 22 → 18 per user feedback
// (it was reading as bulky against the other columns).
var _clientColDefaults={name:18,status:7,phone:14,caregiver:16,caseworker:16,county:11,lastInvoice:14};
var _clientColumnWidths=null;
var _clientPage=1;
var _clientPageSize=25;
(function loadClientPrefs(){
  try{
    var w=localStorage.getItem('lhca_client_col_widths');if(w)_clientColumnWidths=JSON.parse(w);
    var ps=localStorage.getItem('lhca_client_page_size');
    if(ps==='all')_clientPageSize=Infinity;
    else if(ps){var n=parseInt(ps);if(n>0)_clientPageSize=n;}
  }catch(e){}
})();
function applyClientColWidths(){
  var headers=document.querySelectorAll('#clientTable thead th[data-col]');
  headers.forEach(function(th){
    var col=th.dataset.col;
    var w=(_clientColumnWidths&&_clientColumnWidths[col])||_clientColDefaults[col];
    if(w)th.style.width=w+'%';
  });
}
function startColResize(e,col){
  e.preventDefault();e.stopPropagation();
  var th=e.target.closest('th');if(!th)return;
  var startX=e.pageX;
  var startWidth=th.offsetWidth;
  var tableWidth=th.closest('table').offsetWidth;
  document.body.style.cursor='col-resize';document.body.style.userSelect='none';
  function onMove(ev){
    var delta=ev.pageX-startX;
    var newPx=Math.max(60,startWidth+delta);
    var newPct=(newPx/tableWidth)*100;
    th.style.width=newPct+'%';
    if(!_clientColumnWidths)_clientColumnWidths={};
    _clientColumnWidths[col]=newPct;
  }
  function onUp(){
    document.removeEventListener('mousemove',onMove);
    document.removeEventListener('mouseup',onUp);
    document.body.style.cursor='';document.body.style.userSelect='';
    try{localStorage.setItem('lhca_client_col_widths',JSON.stringify(_clientColumnWidths||{}));}catch(e){}
  }
  document.addEventListener('mousemove',onMove);
  document.addEventListener('mouseup',onUp);
}
function clientPageSizeChange(){
  var sel=document.getElementById('clientPageSize');
  var v=sel.value;
  if(v==='all'){_clientPageSize=Infinity;}
  else {_clientPageSize=parseInt(v)||25;}
  _clientPage=1;
  try{localStorage.setItem('lhca_client_page_size',v);}catch(e){}
  renderClientTable();
}
function goToClientPage(n){_clientPage=n;renderClientTable();}
function renderClientTable(forceStatus){
  var profiles=getProfiles();
  var q=((document.getElementById('clientSearch')?document.getElementById('clientSearch').value:'')||'').toLowerCase();
  var filterStatus=forceStatus||(document.getElementById('filterStatus')?document.getElementById('filterStatus').value:'active');
  var filterCg=(document.getElementById('filterCaregiver')&&document.getElementById('filterCaregiver').value)||'';
  var cgs=getCaregivers();
  var cwsArr=getCaseworkers();
  var cwById={};cwsArr.forEach(function(c){cwById[c.id]=c;});
  var keys=Object.keys(profiles).filter(function(k){
    var st=profiles[k].clientStatus||'active';
    var matchStatus=filterStatus==='all'||st===filterStatus;
    var matchQ=!q||k.toLowerCase().includes(q)||(profiles[k].medicaidId||'').toLowerCase().includes(q);
    var matchCg=!filterCg||profiles[k].caregiverId===filterCg;
    return matchStatus&&matchQ&&matchCg;
  });
  keys.sort(function(a,b){return _clientSortCompare(a,b,profiles,cgs,cwsArr);});

  var headers=document.querySelectorAll('#clientTable thead th.sortable');
  headers.forEach(function(h){
    h.classList.remove('sort-asc','sort-desc');
    if(h.dataset.sortkey===_clientSort.key){h.classList.add(_clientSort.dir==='asc'?'sort-asc':'sort-desc');}
  });
  // Restore user-set column widths on every render (headers rebuild if the page reloads)
  applyClientColWidths();

  var totalMatched=keys.length;
  var totalPages=Math.max(1,Math.ceil(totalMatched/_clientPageSize));
  if(_clientPage>totalPages)_clientPage=totalPages;
  var startIdx=(_clientPage-1)*_clientPageSize;
  var endIdx=Math.min(startIdx+_clientPageSize,totalMatched);
  var visibleKeys=keys.slice(startIdx,endIdx);

  var tbody=document.getElementById('clientTableBody'),empty=document.getElementById('clientTableEmpty');
  var countEl=document.getElementById('clientTableCount');
  if(countEl){
    if(!totalMatched){countEl.textContent='';}
    else if(_clientPageSize===Infinity||totalMatched<=_clientPageSize){
      countEl.textContent=totalMatched+' client'+(totalMatched===1?'':'s');
    } else {
      countEl.textContent='Showing '+(startIdx+1)+'–'+endIdx+' of '+totalMatched;
    }
  }
  var pageControls=document.getElementById('clientPageControls');
  if(pageControls){
    if(totalPages<=1){pageControls.innerHTML='';}
    else {
      var html='';
      html+='<button class="pg-btn" onclick="goToClientPage('+(_clientPage-1)+')"'+(_clientPage===1?' disabled':'')+'>‹</button>';
      var startPg=Math.max(1,_clientPage-2);
      var endPg=Math.min(totalPages,startPg+4);
      if(endPg-startPg<4)startPg=Math.max(1,endPg-4);
      for(var p=startPg;p<=endPg;p++){
        html+='<button class="pg-btn'+(p===_clientPage?' active':'')+'" onclick="goToClientPage('+p+')">'+p+'</button>';
      }
      html+='<button class="pg-btn" onclick="goToClientPage('+(_clientPage+1)+')"'+(_clientPage===totalPages?' disabled':'')+'>›</button>';
      pageControls.innerHTML=html;
    }
  }

  if(!tbody)return;tbody.innerHTML='';
  if(!keys.length){if(empty)empty.style.display='block';if(tbody)tbody.innerHTML='';return;}
  if(empty)empty.style.display='none';
  visibleKeys.forEach(function(name){
    var prof=profiles[name];
    var st=prof.clientStatus||'active';
    var stLabel=st==='inactive'?'In Progress':st.charAt(0).toUpperCase()+st.slice(1);
    var cgName=prof.caregiverId&&cgs[prof.caregiverId]?cgs[prof.caregiverId].name:'—';
    var cw=prof.caseworkerId?cwById[prof.caseworkerId]:null;
    var cwName=cw?cw.name:(prof.worker||'—');
    var county=prof.county||'—';
    var phone=prof.phone||'—';
    var invs=prof.invoices||[];
    var lastInv=invs.length?invs[0].billingPeriod:'—';
    var checked=bulkSelected[name]?'checked':'';
    var tr=document.createElement('tr');
    var hrefCl=buildClientUrl(name);
    tr.innerHTML=
      '<td style="width:26px;" onclick="event.stopPropagation()"><input type="checkbox" '+checked+' onchange="toggleBulkClient(\''+escJsAttr(name)+'\',this)" style="width:12px;height:12px;cursor:pointer;"></td>'+
      '<td><a href="'+hrefCl+'" class="link-plain" style="display:block;" onclick="return navClick(event,this.getAttribute(\'href\'))"><div class="ct-name">'+esc(name)+(prof.nickname?'<span style="font-weight:normal;color:var(--text-subtle);"> ('+esc(prof.nickname)+')</span>':'')+'</div><div class="ct-id">'+esc(prof.medicaidId||'No Medicaid ID')+' '+_programBadge(prof)+'</div></a></td>'+
      '<td><span class="status-inline"><span class="status-dot '+st+'"></span>'+stLabel+'</span></td>'+
      '<td style="color:var(--text-muted);font-size:12px;">'+esc(phone)+'</td>'+
      '<td style="color:var(--text-muted);font-size:12px;">'+esc(cgName)+'</td>'+
      '<td style="color:var(--text-muted);font-size:12px;">'+esc(cwName)+'</td>'+
      '<td style="color:var(--text-muted);font-size:12px;">'+esc(county)+'</td>'+
      '<td style="font-size:12px;color:var(--text-muted);">'+esc(lastInv)+'</td>';
    tr.addEventListener('click',function(e){
      if(e.target.closest('a')||e.target.tagName==='BUTTON'||e.target.tagName==='INPUT')return;
      navDetail(name);
    });
    tbody.appendChild(tr);
  });

  // Sync the page-size selector with saved value on first render (dropdown may still show default)
  var sel=document.getElementById('clientPageSize');
  if(sel){
    var savedPs=_clientPageSize===Infinity?'all':String(_clientPageSize);
    if(sel.value!==savedPs)sel.value=savedPs;
  }
}
function renderClientGrid(){renderClientTable();}
function toggleBulkClient(name,cb){
  if(cb.checked)bulkSelected[name]=true;else delete bulkSelected[name];
  var count=Object.keys(bulkSelected).length;
  var bar=document.getElementById('ctBulkBar');
  if(bar){bar.classList.toggle('visible',count>0);var lbl=document.getElementById('ctBulkCount');if(lbl)lbl.textContent=count+' selected';}
}
function clearBulkSelect(){bulkSelected={};var bar=document.getElementById('ctBulkBar');if(bar)bar.classList.remove('visible');renderClientTable();}
// F3: run a batch of saves as QUIET saves and report ONE summary status — so a failure in
// the middle of a batch can't be visually overwritten by a later "Saved ✓" (which left the
// failed record un-synced in LS, to be clobbered on the next reload). thunks each return a
// save promise. On any failure the persistent red toast tells the user how many + which.
function batchSaveWithSummary(label, thunks){
  if(!thunks.length)return Promise.resolve({ok:0,fail:0});
  var ok=0,failed=[];
  return Promise.all(thunks.map(function(t){
    return Promise.resolve(t()).then(function(){ok++;},function(){failed.push(t._label||'');});
  })).then(function(){
    if(failed.length){
      var some=failed.filter(Boolean).slice(0,3).join(', ');
      _showSaveStatus('failed', label+': '+ok+' saved, '+failed.length+' FAILED'+(some?' ('+some+')':'')+' — re-check and save again', null);
    } else { _showSaveStatus('saved', label+' ('+ok+')'); }
    return {ok:ok,fail:failed.length};
  });
}
function bulkSetClientStatus(status){
  // saveClientInfo and createClient both refuse Active without a DHS-1210 for a CHAMPS client; this
  // path wrote the status directly, so bulk "Mark Active" was a way around the rule — and an Active
  // client with no authorization is then picked up by invoice generation.
  if(status==='active'){
    var _p0=getProfiles();
    var _blocked=Array.from(bulkSelected).filter(function(n){
      var pr=_p0[n]; return pr && pr.program!=='carrier' && !hasAuthorization(pr);
    });
    if(_blocked.length){
      showAlert('These clients have no DHS-1210 authorization on file, so they can’t be set Active:\n\n• '+
        _blocked.slice(0,10).join('\n• ')+(_blocked.length>10?('\n• …and '+(_blocked.length-10)+' more'):'')+
        '\n\nImport each authorization on the client’s Authorization tab first.',{title:'Authorization required'});
      return;
    }
  }
  var names=Object.keys(bulkSelected);if(!names.length)return;
  var p=getProfiles(),changed=0,thunks=[];
  names.forEach(function(name){
    if(!p[name])return;
    p[name].clientStatus=status;
    addAuditEntry(name,'Client status changed to '+status+' (bulk action)');
    var th=function(){return saveProfileSP(name,p[name],true);}; th._label=name; thunks.push(th);
    changed++;
  });
  saveProfilesLS(p);
  batchSaveWithSummary('Status update', thunks);
  logActivity('status','Bulk client status update: '+changed+' set to '+status);
  clearBulkSelect();updateStats();renderClientTable();renderSidebarClients();
  showAlert(changed+' client'+(changed!==1?'s':'')+' set to '+status+'.');
}
function bulkDelete(){
  var names=Object.keys(bulkSelected);if(!names.length)return;
  var preview=names.slice(0,5).map(function(n){return '• '+n;}).join('\n')+(names.length>5?'\n• …and '+(names.length-5)+' more':'');
  showConfirm(
    'Delete '+names.length+' client'+(names.length>1?'s':'')+' and ALL their data (invoices, notes, audit log)?\n\n'+preview+'\n\nThis cannot be undone.',
    function(){
      var p=getProfiles(),deleted=0;
      names.forEach(function(name){
        if(!p[name])return;
        try{deleteProfileSP(name);}catch(e){console.error('Backend delete failed',e);}
        delete p[name];
        try{localStorage.removeItem('lhca_draft_'+name);}catch(e){}
        deleted++;
      });
      saveProfilesLS(p);
      logActivity('delete','Bulk delete: '+deleted+' client'+(deleted!==1?'s':'')+' removed');
      if(activeProfileName && !p[activeProfileName])activeProfileName=null;
      clearBulkSelect();updateStats();renderSidebarClients();
      showAlert(deleted+' client'+(deleted!==1?'s':'')+' deleted.');
    },
    {title:'Delete '+names.length+' Clients?',okText:'Delete '+names.length+(names.length>1?' Clients':' Client'),danger:true}
  );
}

// ============================================================
//  CLIENT DETAIL TABS
// ============================================================
function switchTab(tab){
  if(unsavedChanges&&tab!=='info'){
    showConfirm('You have unsaved changes on the Profile tab. Leave anyway?',function(){
      unsavedChanges=false;
      _doSwitchTab(tab);
    },{title:'Unsaved Changes',okText:'Discard & Leave',danger:true});
    return;
  }
  _doSwitchTab(tab);
}
function _doSwitchTab(tab){
  ['overview','info','auth','history','notes','docs','audit'].forEach(function(t){
    var dtab=document.getElementById('dtab-'+t);
    var dpane=document.getElementById('dpane-'+t);
    if(dtab){dtab.classList.toggle('active',t===tab);dtab.setAttribute('aria-selected',t===tab?'true':'false');}
    if(dpane)dpane.classList.toggle('active',t===tab);
  });
  if(tab==='overview')renderOverviewPane();
  if(tab==='info')renderInfoPane();
  if(tab==='auth')renderAuthPane();
  if(tab==='history')renderInvHistory();
  if(tab==='notes')renderNotesPane();
  if(tab==='docs')renderDocsPane();
  if(tab==='audit')renderAuditPane();
}

function renderOverviewPane(){
  if(!activeProfileName)return;
  var prof=getProfiles()[activeProfileName];
  var pane=document.getElementById('dpane-overview');
  var invoices=prof.invoices||[];
  var cgName='';if(prof.caregiverId){var cgs=getCaregivers();if(cgs[prof.caregiverId])cgName=cgs[prof.caregiverId].name;}
  var paid=invoices.filter(function(i){return i.status==='paid';}).length;
  var submitted=invoices.filter(function(i){return i.status==='submitted';}).length;
  var draft=invoices.filter(function(i){return !i.status||i.status==='draft';}).length;
  var notesPreview=prof.clientNotes?prof.clientNotes.slice(0,200):'No notes yet.';
  var cwRec=getCaseworkers().find(function(c){return c.id===prof.caseworkerId||c.name===prof.worker;})||null;
  var cwName=cwRec?cwRec.name:(prof.worker||'');
  var addrStr=(prof.street||'')+( prof.city?', '+prof.city:'')+(prof.state?' '+prof.state:'')+(prof.zip?' '+prof.zip:'');
  var clientTasks=getTodos().filter(function(t){return t.client===activeProfileName&&!t.done;});
  var tasksHtml='';
  if(clientTasks.length){
    tasksHtml=clientTasks.slice(0,4).map(function(t){
      var overdue=t.due&&new Date(t.due)<new Date();
      return '<div class="ov-inv-row" style="cursor:pointer;" onclick="navTasks()" title="Go to Tasks">'+
        '<span style="font-size:11px;'+(overdue?'color:#b03030;font-weight:600;':'color:#1a2b45;')+'">'+esc(t.text)+'</span>'+
        (t.due?'<span style="font-size:10px;color:'+(overdue?'#b03030':'#5c7590')+';">'+t.due+'</span>':'')+
      '</div>';
    }).join('');
  } else {
    tasksHtml='<div style="color:#5c7590;font-size:12px;padding:4px 0;">No open tasks.</div>';
  }
  var recentInvHtmlRO='';
  if(!invoices.length){recentInvHtmlRO='<div style="color:#5c7590;font-size:12px;padding:6px 0;">No invoices yet.</div>';}
  else{invoices.slice(0,4).forEach(function(inv){
    var st=inv.status||'draft';
    var daysSince=Math.floor((Date.now()-new Date(inv.savedAt))/86400000);
    var overdueFlag=(st==='submitted'&&daysSince>30)?'<span style="background:#ff8c00;color:#fff;font-size:10px;padding:1px 6px;border-radius:8px;margin-left:4px;">'+daysSince+'d</span>':'';
    recentInvHtmlRO+='<div class="ov-inv-row" style="cursor:pointer;" onclick="switchTab(\'history\')" title="Open Invoices tab">'+
      '<span class="inv-badge" style="min-width:60px;text-align:center;">'+esc(inv.billingPeriod)+'</span>'+
      '<span class="inv-badge st-'+st+'" style="min-width:70px;text-align:center;">'+st.charAt(0).toUpperCase()+st.slice(1)+overdueFlag+'</span>'+
      '<span style="flex:1;color:#5c7590;font-size:11px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">'+esc(inv.invoiceNote||'')+'</span>'+
      '<span style="font-size:11px;color:#4d6c88;white-space:nowrap;">'+esc(inv.savedAt.split(',')[0])+'</span></div>';
  });}
  // Authorization (DHS-1210) card — only when the client has an imported authorization.
  var auth=prof.authorization, authCardHtml='';
  if(auth){
    var dueTxt='', dueColor='#1a2b45';
    if(auth.reassessDate){
      var dp=auth.reassessDate.split('/');
      if(dp.length===3){var dd=new Date(+dp[2],(+dp[0])-1,+dp[1]);var days=Math.floor((dd-new Date())/86400000);
        if(days<0){dueTxt=' · overdue';dueColor='#b03030';}else if(days<=30){dueTxt=' · '+days+'d';dueColor='#c67605';}}
    }
    // Slim read-only glance — the editable version lives in the Profile tab. Click to go there.
    authCardHtml='<div class="ov-card" style="cursor:pointer;" onclick="switchTab(\'auth\')" title="Open Authorization tab">'+
      '<h4>Authorization (DHS-1210)</h4>'+
      '<div class="ov-row"><span class="ov-label">Approved / mo</span><span class="ov-value">'+(auth.hours!=null?auth.hours+'h '+(auth.minutes||0)+'m':'—')+'</span></div>'+
      '<div class="ov-row"><span class="ov-label">Reassess due</span><span class="ov-value" style="color:'+dueColor+';">'+esc(auth.reassessDate||'—')+dueTxt+'</span></div>'+
      '<div style="margin-top:6px;font-size:10px;color:#94a7bd;">Click to open the Authorization tab</div>'+
    '</div>';
  }
  pane.innerHTML='<div class="overview-grid">'+
    '<div class="ov-card"><h4>Client Info</h4>'+
      '<div class="ov-row"><span class="ov-label">Medicaid ID</span><span class="ov-value">'+esc(prof.medicaidId||'—')+'</span></div>'+
      (prof.medicare?'<div class="ov-row"><span class="ov-label">Medicare #</span><span class="ov-value">'+esc(prof.medicare)+'</span></div>':'')+
      '<div class="ov-row"><span class="ov-label">Hourly Rate</span><span class="ov-value">'+(prof.hourlyRate?'$'+esc(prof.hourlyRate)+'/hr':'—')+'</span></div>'+
      '<div class="ov-row"><span class="ov-label">Phone</span><span class="ov-value">'+esc(prof.phone||'—')+'</span></div>'+
      (addrStr.trim()?'<div class="ov-row"><span class="ov-label">Address</span><span class="ov-value">'+esc(addrStr.trim())+'</span></div>':'')+
      '<div class="ov-row"><span class="ov-label">Caregiver</span>'+(cgName&&prof.caregiverId?'<span class="ov-value" style="color:#185FA5;cursor:pointer;text-decoration:underline;" onclick="navCaregivers();setTimeout(function(){openCgDetail(\''+escJsAttr(prof.caregiverId)+'\');},50)">'+esc(cgName)+'</span>':'<span class="ov-value">'+esc(cgName||'Unassigned')+'</span>')+'</div>'+
      (prof.liveIn?'<div class="ov-row"><span class="ov-label">Live-In</span><span class="ov-value" style="display:inline-block;background:#fff3cd;color:#856404;font-size:11px;font-weight:600;padding:2px 10px;border-radius:10px;border:1px solid #ffeaa7;">YES</span></div>':'')+
      '<div class="ov-row"><span class="ov-label">Caseworker</span>'+(cwRec?'<span class="ov-value" style="color:#185FA5;cursor:pointer;text-decoration:underline;" onclick="navCaseworkers();setTimeout(function(){openCwDetail(\''+escJsAttr(cwRec.id)+'\');},50)">'+esc(cwName)+'</span>':'<span class="ov-value">'+esc(cwName||'—')+'</span>')+'</div>'+
      (prof.startDate?'<div class="ov-row"><span class="ov-label">Service Start</span><span class="ov-value">'+esc(prof.startDate)+'</span></div>':'')+
    '</div>'+
    authCardHtml+
    '<div class="ov-card"><h4>Invoice Summary</h4>'+
      '<div class="ov-row"><span class="ov-label">Total Saved</span><span class="ov-value">'+invoices.length+'</span></div>'+
      '<div class="ov-row"><span class="ov-label">Paid</span><span class="ov-value" style="color:#1e7e34;">'+paid+'</span></div>'+
      '<div class="ov-row"><span class="ov-label">Submitted</span><span class="ov-value" style="color:#1565a0;">'+submitted+'</span></div>'+
      '<div class="ov-row"><span class="ov-label">Draft / Unsent</span><span class="ov-value" style="color:#666;">'+draft+'</span></div>'+
      '<div style="margin-top:10px;border-top:1px solid #f0f3f7;padding-top:10px;">'+recentInvHtmlRO+'</div>'+
    '</div>'+
    '<div class="ov-card"><h4>Tasks <span style="font-size:10px;color:#5c7590;font-weight:normal;text-transform:none;letter-spacing:0;">('+clientTasks.length+' open)</span>'+
      '<div style="margin-left:auto;display:flex;gap:6px;">'+
        '<button class="btn btn-secondary btn-sm" style="font-size:10px;" onclick="addTaskForClient(\''+escJsAttr(activeProfileName)+'\')">+ Add Task</button>'+
        '<button class="btn btn-secondary btn-sm" style="font-size:10px;" onclick="openWorkflowModal()">Workflow</button>'+
      '</div></h4>'+
      tasksHtml+
      '<div style="margin-top:8px;"><button class="btn btn-secondary btn-sm" onclick="navTasks()">View All Tasks</button></div>'+
    '</div>'+
    '<div class="ov-card"><h4>Notes</h4>'+
      '<div class="ov-notes-preview">'+esc(notesPreview)+'</div>'+
    '</div>'+
  '</div>';
}

function openAllInvoicesModal(filter){
  var p=getProfiles(),rows=[];
  Object.keys(p).forEach(function(name){
    (p[name].invoices||[]).forEach(function(inv,idx){
      if(filter==='outstanding'&&inv.status==='paid')return;
      rows.push({client:name,inv:inv,idx:idx});
    });
  });
  rows.sort(function(a,b){return new Date(b.inv.savedAt)-new Date(a.inv.savedAt);});
  document.getElementById('allInvModalTitle').textContent=filter==='outstanding'?'Outstanding Invoices':'All Saved Invoices';
  document.getElementById('allInvModalSubtitle').textContent=rows.length+' invoice'+(rows.length!==1?'s':'')+(filter==='outstanding'?' not yet paid':'');
  var list=document.getElementById('allInvModalList');list.innerHTML='';
  if(!rows.length){list.innerHTML='<div class="af-empty">'+(filter==='outstanding'?'No outstanding invoices. ':'No invoices saved yet.')+'</div>';document.getElementById('allInvoicesModal').classList.add('open');return;}
  rows.forEach(function(r){
    var st=r.inv.status||'draft';
    var row=document.createElement('div');row.className='inv-file-card';row.style.marginBottom='6px';
    row.innerHTML=
      ''+
      '<div class="inv-file-info">'+
        '<div style="display:flex;align-items:center;gap:8px;">'+
          '<span class="inv-file-period">'+esc(r.inv.billingPeriod)+'</span>'+
          '<span style="font-size:11px;color:#185FA5;font-weight:600;">'+esc(r.client)+'</span>'+
        '</div>'+
        '<div class="inv-file-meta">'+esc(r.inv.savedAt)+(r.inv.invoiceNote?' · '+esc(r.inv.invoiceNote):'')+'</div>'+
      '</div>'+
      '<div class="inv-file-actions">'+
        '<span class="status-select st-'+st+'" style="cursor:default;">'+st.charAt(0).toUpperCase()+st.slice(1)+'</span>'+
        '<button class="btn btn-secondary btn-sm" onclick="event.stopPropagation();navDetail(\''+escJsAttr(r.client)+'\')">Profile</button>'+
        '<button class="btn btn-secondary btn-sm" onclick="event.stopPropagation();openInvFromModal(\''+escJsAttr(r.client)+'\','+r.idx+')">Open Invoice →</button>'+
      '</div>';
    list.appendChild(row);
  });
  document.getElementById('allInvoicesModal').classList.add('open');
}
function openInvFromModal(clientName,idx){
  document.getElementById('allInvoicesModal').classList.remove('open');
  activeProfileName=clientName;
  var inv=getProfiles()[clientName].invoices[idx];if(!inv)return;
  navInvoice(inv.data);
}

// ── Hover-copy on detail-form fields (ported from the Health CRM's wireCopyableFields) ──
var _SVG_COPY='<svg viewBox="0 0 24 24" width="12" height="12" stroke="currentColor" stroke-width="1.8" fill="none" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>';
var _SVG_CHECK='<svg viewBox="0 0 24 24" width="12" height="12" stroke="currentColor" stroke-width="1.8" fill="none" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg>';
// Format the value for copying: SSN / Medicaid / Medicare / Driver's-License → digits only (dash-free
// for pasting into state portals); date fields → MM/DD/YYYY; everything else copied as shown.
function _copyFmt(inp){
  var v=(inp.value||'').trim(); if(!v)return '';
  var id=inp.id||'';
  // ID numbers → strip separators only (dashes/spaces), keeping letters — Medicare MBI and many
  // driver's licenses are alphanumeric, so digits-only would corrupt them.
  if(/ssn|medicaid|medicare/i.test(id) || /-dl$/i.test(id)) return v.replace(/[^0-9A-Za-z]/g,'');
  if(inp.type==='date'){ var m=v.match(/^(\d{4})-(\d{2})-(\d{2})$/); return m?(m[2]+'/'+m[3]+'/'+m[1]):v; }
  return v;
}
// Attach a hover-only copy button to every text/date input in a detail form (client/caregiver use
// .info-field, caseworker uses .ff). Idempotent via data-copyable; skips checkboxes/hidden/buttons
// and fields with no id. The SSN double-click shortcut still works too.
function wireCopyableFields(c){
  var root=(typeof c==='string')?document.getElementById(c):c; if(!root)return;
  root.querySelectorAll('input:not([type=checkbox]):not([type=hidden]):not([type=button]):not([data-copyable])').forEach(function(inp){
    if(!inp.id)return;
    inp.setAttribute('data-copyable','1');
    var field=inp.closest('.info-field, .ff'); if(!field)return;
    field.classList.add('field-copyable');
    var btn=document.createElement('button');
    btn.type='button'; btn.className='copy-hover-btn'; btn.title='Copy'; btn.tabIndex=-1; btn.innerHTML=_SVG_COPY;
    btn.addEventListener('click',function(e){
      e.preventDefault(); e.stopPropagation();
      var v=_copyFmt(inp); if(!v)return;
      try{ if(navigator.clipboard&&navigator.clipboard.writeText)navigator.clipboard.writeText(v); }catch(_){}
      btn.classList.add('copied'); btn.innerHTML=_SVG_CHECK;
      setTimeout(function(){ btn.classList.remove('copied'); btn.innerHTML=_SVG_COPY; },1200);
    });
    field.appendChild(btn);
  });
}
function renderInfoPane(){
  if(!activeProfileName)return;
  var prof=getProfiles()[activeProfileName],g=document.getElementById('infoGrid');g.innerHTML='';

  // Derive first/last from stored fields or split activeProfileName as fallback
  var storedFirst=prof.firstName||'';
  var storedLast=prof.lastName||'';
  if(!storedFirst&&!storedLast){
    var parts=activeProfileName.split(' ');
    storedFirst=parts[0]||'';storedLast=parts.slice(1).join(' ')||'';
  }

  function mkField(id,label,val,full){
    var d=document.createElement('div');d.className='info-field'+(full?' full':'');
    // autocomplete=off: keep PHI (DL, phone, email, address, medicaid id) out of the
    // browser's form-autofill store, which lives outside the app and survives the idle-wipe.
    d.innerHTML='<label for="'+id+'">'+label+'</label><input id="'+id+'" value="'+esc(val)+'" autocomplete="off" oninput="unsavedChanges=true;">';
    g.appendChild(d);
  }
  function mkDivider(label){
    var div=document.createElement('div');div.className='form-section-divider full';div.innerHTML='<span>'+label+'</span>';g.appendChild(div);
  }

  var dName=document.createElement('div');dName.className='info-field-row full';dName.style.gridTemplateColumns='1fr 1fr 1fr';
  dName.innerHTML='<div class="info-field"><label for="ei-first">First Name *</label><input id="ei-first" value="'+esc(storedFirst)+'" oninput="unsavedChanges=true;"></div>'+
    '<div class="info-field"><label for="ei-middle">Middle Name</label><input id="ei-middle" value="'+esc(prof.middleName||'')+'" oninput="unsavedChanges=true;"></div>'+
    '<div class="info-field"><label for="ei-last">Last Name *</label><input id="ei-last" value="'+esc(storedLast)+'" oninput="unsavedChanges=true;"></div>';
  g.appendChild(dName);

  var dNickSt=document.createElement('div');dNickSt.className='info-field-row full';
  dNickSt.innerHTML='<div class="info-field"><label for="ei-nickname">Nickname / Goes By</label><input id="ei-nickname" value="'+esc(prof.nickname||'')+'" oninput="unsavedChanges=true;"></div>'+
    '<div class="info-field"><label for="ei-status">Client Status</label><select id="ei-status">'+
      ['active','inactive','lost','terminated'].map(function(s){return '<option value="'+s+'"'+((prof.clientStatus||'active')===s?' selected':'')+'>'+(s==='inactive'?'In Progress':s.charAt(0).toUpperCase()+s.slice(1))+'</option>';}).join('')+
    '</select></div>';
  g.appendChild(dNickSt);

  mkField('ei-medicaid','Medicaid ID',prof.medicaidId||'',false);
  mkField('ei-medicare','Medicare #',prof.medicare||'',false);
  // Program (CHAMPS vs Managed-Care carrier). Absent = champs. Carrier shows carrier + member #.
  var _pg=prof.program||'champs';
  var dProg=document.createElement('div');dProg.className='info-field-row full';dProg.style.gridTemplateColumns='1fr 1fr 1fr';
  dProg.innerHTML='<div class="info-field"><label for="ei-program">Program</label><select id="ei-program" onchange="unsavedChanges=true;eiProgramToggle();">'+
      '<option value="champs"'+(_pg!=='carrier'?' selected':'')+'>MDHHS (CHAMPS)</option>'+
      '<option value="carrier"'+(_pg==='carrier'?' selected':'')+'>Carrier (Managed Care)</option>'+
    '</select></div>'+
    '<div class="info-field" id="ei-carrier-wrap"><label for="ei-carrier">Carrier</label><select id="ei-carrier" onchange="unsavedChanges=true;">'+
      ['','Humana','Priority Health','Aetna','UnitedHealthcare','HAP','Molina','Meridian','Other'].map(function(c){return '<option value="'+esc(c)+'"'+((prof.carrier||'')===c?' selected':'')+'>'+(c||'— Select carrier —')+'</option>';}).join('')+
    '</select></div>'+
    '<div class="info-field" id="ei-member-wrap"><label for="ei-member">Member #</label><input id="ei-member" value="'+esc(prof.memberId||'')+'" autocomplete="off" oninput="unsavedChanges=true;"></div>';
  g.appendChild(dProg);
  eiProgramToggle();
  // Date of Birth — needed on DHS-390 / MDHHS-6200 / MSA-4676 state forms
  var dDob=document.createElement('div');dDob.className='info-field';
  dDob.innerHTML='<label for="ei-dob">Date of Birth <span style="font-weight:400;font-size:11px;color:#5c7590;">(double-click to copy)</span></label><input id="ei-dob" type="date" value="'+esc(prof.dob||'')+'" autocomplete="off" oninput="unsavedChanges=true;" ondblclick="_copyField(this,\'mdy\')" title="Double-click to copy as MM/DD/YYYY">';
  g.appendChild(dDob);
  var dGender=document.createElement('div');dGender.className='info-field';
  var gv=prof.gender||'';
  dGender.innerHTML='<label for="ei-gender">Gender</label><select id="ei-gender" onchange="unsavedChanges=true;">'+
    '<option value=""'+(gv===''?' selected':'')+'>—</option>'+
    '<option value="Male"'+(gv==='Male'?' selected':'')+'>Male</option>'+
    '<option value="Female"'+(gv==='Female'?' selected':'')+'>Female</option>'+
  '</select>';
  g.appendChild(dGender);
  // Hourly Rate: numeric-guarded (formatRate) like caregiver pay, so it can't hold
  // free text (defense-in-depth alongside the esc() on the overview render).
  var dRate=document.createElement('div');dRate.className='info-field';
  dRate.innerHTML='<label for="ei-rate">Hourly Rate</label><input id="ei-rate" value="'+esc(prof.hourlyRate||'')+'" inputmode="decimal" oninput="formatRate(this);unsavedChanges=true;">';
  g.appendChild(dRate);
  mkField('ei-dl',"Driver's License #",prof.driversLicense||'',false);
  // SSN masked by default; reveals on focus, re-masks a few seconds after blur (#9).
  // Double-click copies the 9 digits with no dashes (#6).
  var dSsn=document.createElement('div');dSsn.className='info-field';
  dSsn.innerHTML='<label for="ei-ssn">Social Security # <span style="font-weight:400;font-size:11px;color:#5c7590;">(double-click to copy, no dashes)</span></label>'+
    '<input id="ei-ssn" type="password" autocomplete="off" maxlength="11" placeholder="XXX-XX-XXXX" value="'+esc(prof.ssn||'')+'" oninput="formatSSN(this);unsavedChanges=true;" onfocus="_revealSecret(this)" onblur="_maskSecretSoon(this)" ondblclick="_copyField(this,\'digits\')" title="Double-click to copy without dashes">';
  g.appendChild(dSsn);
  mkField('ei-phone','Client Phone',prof.phone||'',false);
  mkField('ei-home-phone','Home Phone',prof.homePhone||'',false);
  mkField('ei-cemail','Client Email',prof.clientEmail||'',false);
  mkField('ei-street','Street',prof.street||'',true);

  var dCityState=document.createElement('div');dCityState.className='info-field-row full';
  dCityState.innerHTML='<div class="info-field"><label for="ei-city">City</label><input id="ei-city" value="'+esc(prof.city||'')+'" oninput="unsavedChanges=true;"></div>'+
    '<div class="info-field"><label for="ei-state">State</label><input id="ei-state" value="'+esc(prof.state||'')+'" oninput="unsavedChanges=true;"></div>';
  g.appendChild(dCityState);

  var dZipCounty=document.createElement('div');dZipCounty.className='info-field-row full';
  dZipCounty.innerHTML='<div class="info-field"><label for="ei-zip">ZIP</label><input id="ei-zip" value="'+esc(prof.zip||'')+'" oninput="unsavedChanges=true;lookupZip(\'ei-zip\',\'ei-city\',\'ei-state\',\'ei-county\')"></div>'+
    '<div class="info-field"><label for="ei-county">County</label><input id="ei-county" value="'+esc(prof.county||'')+'" oninput="unsavedChanges=true;"></div>';
  g.appendChild(dZipCounty);

  mkDivider('Assignments');

  // Service Start Date — moved ABOVE caregiver per user
  var dStart=document.createElement('div');dStart.className='info-field full';
  dStart.innerHTML='<label for="ei-start-date">Service Start Date <span style="font-weight:400;font-size:11px;color:#5c7590;">(prevents missing-invoice warnings for months before this date)</span></label><input type="date" id="ei-start-date" value="'+esc(prof.startDate||'')+'" oninput="unsavedChanges=true;" ondblclick="_copyField(this,\'mdy\')" title="Double-click to copy as MM/DD/YYYY">';
  g.appendChild(dStart);

  // Assigned Caregiver — full width, with Open button (now matches the Caseworker layout below)
  var cgName='';var cgsMap=getCaregivers();if(prof.caregiverId&&cgsMap[prof.caregiverId])cgName=cgsMap[prof.caregiverId].name;
  var dCg=document.createElement('div');dCg.className='info-field full';
  dCg.innerHTML='<label for="ei-caregiver-search">Assigned Caregiver</label>'+
    '<div style="display:flex;align-items:center;gap:6px;position:relative;">'+
      '<div style="flex:1;position:relative;">'+
        '<input id="ei-caregiver-search" placeholder="Click to browse, or type to search…" maxlength="80" autocomplete="off" value="'+esc(cgName)+'" oninput="cgSearch(this,\'ei-caregiver-val\',\'ei-caregiver-drop\');unsavedChanges=true;" onfocus="cgSearch(this,\'ei-caregiver-val\',\'ei-caregiver-drop\')" onblur="setTimeout(function(){var d=document.getElementById(\'ei-caregiver-drop\');if(d)d.style.display=\'none\';},200)" style="width:100%;padding:7px 10px;border:1px solid #d0d8e4;border-radius:5px;font-size:13px;font-family:Arial,sans-serif;outline:none;">'+
        '<input type="hidden" id="ei-caregiver-val" value="'+esc(prof.caregiverId||'')+'">'+
        '<div id="ei-caregiver-drop" style="display:none;position:absolute;top:100%;left:0;right:0;background:#fff;border:1px solid #d0d8e4;border-radius:0 0 6px 6px;z-index:200;max-height:180px;overflow-y:auto;box-shadow:0 4px 12px rgba(0,0,0,0.1);"></div>'+
      '</div>'+
      (prof.caregiverId?'<button class="btn-open" onclick="navCaregivers();setTimeout(function(){openCgDetail(\''+escJsAttr(prof.caregiverId)+'\');},50)">Open ↗</button>':'')+
    '</div>';
  g.appendChild(dCg);
  // Live-In on its own row (was cramped beside the caregiver box, making it narrower)
  var dLiveIn=document.createElement('div');dLiveIn.className='info-field full';
  dLiveIn.innerHTML='<label style="display:flex;align-items:center;gap:6px;cursor:pointer;font-weight:500;margin:0;"><input type="checkbox" id="ei-live-in" '+(prof.liveIn?'checked':'')+' onchange="unsavedChanges=true;" style="width:14px;height:14px;cursor:pointer;"> Live-In Caregiver</label>';
  g.appendChild(dLiveIn);

  // Caseworker — full width. Open resolves by id OR by matching name, so a caseworker stored as a
  // plain name (legacy data with no caseworkerId) still gets an Open button.
  var cwName=prof.worker||'';
  var cwRecOpen=getCaseworkers().find(function(c){return c.id===prof.caseworkerId || (cwName && (c.name||'').toLowerCase()===cwName.toLowerCase());});
  var cwOpenId=cwRecOpen?cwRecOpen.id:'';
  var dCw=document.createElement('div');dCw.className='info-field full';
  dCw.innerHTML='<label for="ei-worker-search">Caseworker</label>'+
    '<div style="display:flex;align-items:center;gap:6px;position:relative;">'+
      '<div style="flex:1;position:relative;">'+
        '<input id="ei-worker-search" placeholder="Click to browse, or type to search…" maxlength="80" autocomplete="off" value="'+esc(cwName)+'" oninput="cwSearch(this,\'ei-worker-val\',\'ei-worker-drop\');unsavedChanges=true;" onfocus="cwSearch(this,\'ei-worker-val\',\'ei-worker-drop\')" onblur="setTimeout(function(){var d=document.getElementById(\'ei-worker-drop\');if(d)d.style.display=\'none\';},200)" style="width:100%;padding:7px 10px;border:1px solid #d0d8e4;border-radius:5px;font-size:13px;font-family:Arial,sans-serif;outline:none;">'+
        '<input type="hidden" id="ei-worker-val" value="'+esc(prof.caseworkerId||'')+'">'+
        '<div id="ei-worker-drop" style="display:none;position:absolute;top:100%;left:0;right:0;background:#fff;border:1px solid #d0d8e4;border-radius:0 0 6px 6px;z-index:200;max-height:180px;overflow-y:auto;box-shadow:0 4px 12px rgba(0,0,0,0.1);"></div>'+
      '</div>'+
      (cwOpenId?'<button class="btn-open" onclick="navCaseworkers();setTimeout(function(){openCwDetail(\''+escJsAttr(cwOpenId)+'\');},50)">Open ↗</button>':'')+
    '</div>';
  g.appendChild(dCw);
  wireCopyableFields('infoGrid');
}
// Editable task rows for the Authorization section.
function _renderAuthTaskRows(tasks){
  var host=document.getElementById('ei-auth-tasks'); if(!host)return; host.innerHTML='';
  if(!tasks||!tasks.length){ _authAddTaskRow(); return; }
  tasks.forEach(function(t){_authAddTaskRow(t);});
}
function _authAddTaskRow(t){
  t=t||{};
  var host=document.getElementById('ei-auth-tasks'); if(!host)return;
  var row=document.createElement('div');
  row.className='ei-auth-task-row';
  row.style.cssText='display:grid;grid-template-columns:1.7fr 0.9fr 1.4fr 0.9fr 0.9fr 24px;gap:6px;margin-bottom:5px;align-items:center;';
  // min-width:0 lets each input shrink to its grid track instead of forcing the row wider than
  // the card (the cause of the overflow + header misalignment).
  var s=' style="min-width:0;width:100%;box-sizing:border-box;"';
  row.innerHTML=
    '<input class="ei-at-task" placeholder="Task" value="'+esc(t.task||'')+'"'+s+'>'+
    '<input class="ei-at-perDay" placeholder="HH:MM" value="'+esc(t.perDay||'')+'"'+s+'>'+
    '<input class="ei-at-freq" placeholder="e.g. 7 days per week" value="'+esc(t.freq||'')+'"'+s+'>'+
    '<input class="ei-at-perMonth" placeholder="HH:MM" value="'+esc(t.perMonth||'')+'"'+s+'>'+
    '<input class="ei-at-amount" placeholder="$" value="'+esc(t.amount!=null?String(t.amount):'')+'" inputmode="decimal"'+s+'>'+
    '<button type="button" class="btn btn-secondary btn-sm" title="Remove" onclick="this.parentNode.remove();" style="padding:2px 4px;min-width:0;">×</button>';
  host.appendChild(row);
}
// "Set to 6 months after effective" button — fill reassessment = effective + 6 months.
// Reassessment due = effective + 6 months, advanced in 6-month steps to the NEXT date on/after
// today (MDHHS reviews every 6 months). So an OLD form (effective years ago) yields the next
// upcoming reassessment, not a long-past one. refDate defaults to now; pass one for tests.
function _nextReassessment(effMdy, refDate){
  var p=String(effMdy||'').split('/'); if(p.length!==3)return '';
  var y=+p[2], mi=(+p[0])-1, day=+p[1];
  if(isNaN(y)||isNaN(mi)||isNaN(day))return '';
  var ref=refDate||new Date(); ref=new Date(ref.getFullYear(),ref.getMonth(),ref.getDate());
  var d=new Date(y,mi+6,day);            // first reassessment (effective + 6 months)
  var guard=0;
  while(d<ref && guard++<400){ d=new Date(d.getFullYear(),d.getMonth()+6,d.getDate()); }
  return ('0'+(d.getMonth()+1)).slice(-2)+'/'+('0'+d.getDate()).slice(-2)+'/'+d.getFullYear();
}
function _fillReassess(){
  var eff=(document.getElementById('ei-auth-eff').value||'').trim();
  var next=_nextReassessment(eff);
  if(!next){showAlert('Enter the effective date first (MM/DD/YYYY).');return false;}
  document.getElementById('ei-auth-reassess').value=next;
  unsavedChanges=true; return false;
}
// Auto-fill reassessment = effective + 6 months when the user finishes entering the effective
// date, but only if reassessment is still blank (never clobber a value they typed).
function _autoReassessFromEff(){
  var eff=(document.getElementById('ei-auth-eff').value||'').trim();
  var re=document.getElementById('ei-auth-reassess');
  var p=eff.split('/');
  if(re && !re.value.trim() && p.length===3 && p[2].length===4){
    var next=_nextReassessment(eff);
    if(next) re.value=next;
  }
}
function _mdyToYmd(mdy){var p=String(mdy||'').split('/');if(p.length!==3)return '';return p[2]+'-'+('0'+p[0]).slice(-2)+'-'+('0'+p[1]).slice(-2);}
// 1st of the effective month as YYYY-MM-01 (the usual service start date on a DHS-1210).
function _firstOfEffectiveMonth(effMdy){var p=String(effMdy||'').split('/');if(p.length!==3)return '';return p[2]+'-'+('0'+p[0]).slice(-2)+'-01';}
// Approved time as HH:MM (e.g. 29:47), matching the DHS-1210. Empty when no hours set.
function _authHM(a){ if(!a||a.hours==null)return ''; return a.hours+':'+('0'+(a.minutes||0)).slice(-2); }
// Parse "HH:MM" (or plain hours) back to {hours,minutes}.
function _parseHM(v){ v=String(v||'').trim(); if(!v)return {hours:null,minutes:null}; var p=v.split(':'); var h=parseInt(p[0]); var m=p.length>1?parseInt(p[1]):0; return {hours:isNaN(h)?null:h, minutes:isNaN(m)?0:m}; }
// ── Authorization tab — read-only display by default; Edit reveals the form ──
function renderAuthPane(edit){
  if(!activeProfileName)return;
  var host=document.getElementById('authContent'); if(!host)return;
  var prof=getProfiles()[activeProfileName]; if(!prof)return;
  // Managed-care (carrier) clients don't use the DHS-1210 / CRM invoicing path — no import here.
  if(isCarrierClient(prof)){
    host.innerHTML='<div class="form-card" style="max-width:640px;text-align:center;padding:28px;">'+
      '<div style="font-size:13px;color:#5c7590;">This is a <b>managed-care</b> client'+(prof.carrier?(' ('+esc(prof.carrier)+')'):'')+'. DHS-1210 authorization and CRM invoicing aren’t used — authorizations and billing go through the carrier.</div>'+
    '</div>';
    return;
  }
  var a=prof.authorization;
  if(edit){ host.innerHTML=_authEditHtml(a||{}); _renderAuthTaskRows((a&&a.tasks)||[]); return; }
  if(!a || (a.hours==null && !(a.tasks&&a.tasks.length) && !a.effectiveDate)){
    host.innerHTML='<div class="form-card" style="max-width:640px;text-align:center;padding:28px;">'+
      '<div style="font-size:13px;color:#5c7590;margin-bottom:12px;">No DHS-1210 authorization on file for this client yet.</div>'+
      '<button class="btn btn-primary" onclick="importDHS1210()">Import DHS-1210 / MDHHS-6064</button>'+
      '<div style="font-size:11px;color:#94a7bd;margin-top:10px;">or <a href="#" onclick="renderAuthPane(true);return false;">enter it manually</a></div>'+
    '</div>';
    return;
  }
  var dueTxt='',dueColor='#1a2b45';
  if(a.reassessDate){var dp=a.reassessDate.split('/');if(dp.length===3){var dd=new Date(+dp[2],(+dp[0])-1,+dp[1]);var days=Math.floor((dd-new Date())/86400000);
    if(days<0){dueTxt=' · overdue';dueColor='#b03030';}else if(days<=30){dueTxt=' · '+days+' days';dueColor='#c67605';}}}
  var kv=function(l,v,color){return '<div style="display:flex;justify-content:space-between;gap:16px;padding:6px 0;border-bottom:1px solid #f0f3f7;"><span style="color:#5c7590;font-size:13px;">'+l+'</span><span style="font-weight:600;color:'+(color||'#1a2b45')+';font-size:13px;">'+v+'</span></div>';};
  var rows=(a.tasks||[]).map(function(t){
    return '<tr><td style="padding:4px 8px;">'+esc(t.task||'')+'</td><td style="padding:4px 8px;color:#5c7590;">'+esc(t.perDay||'—')+'</td><td style="padding:4px 8px;color:#5c7590;">'+esc(t.freq||'')+'</td><td style="padding:4px 8px;text-align:right;">'+esc(t.perMonth||'')+'</td><td style="padding:4px 8px;text-align:right;">'+(t.amount!=null?'$'+Number(t.amount).toFixed(2):'—')+'</td></tr>';
  }).join('');
  // First invoice: offer it when there's no invoice yet for the authorization's effective month.
  var _effP=(a.effectiveDate||'').split('/'); var _effPeriod=(_effP.length===3)?(_effP[0]+'/'+_effP[2]):'';
  var _canCreateInv=_effPeriod && !((prof.invoices||[]).some(function(i){return i.billingPeriod===_effPeriod;}));
  host.innerHTML='<div class="form-card" style="max-width:720px;">'+
    '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:14px;gap:10px;flex-wrap:wrap;">'+
      '<h3 style="margin:0;">Authorization ('+esc(a.formType||'DHS-1210')+')</h3>'+
      '<div style="display:flex;gap:8px;flex-wrap:wrap;">'+
        (_canCreateInv?'<button class="btn btn-primary btn-sm" onclick="createFirstInvoiceFromAuth()">Create '+esc(_effPeriod)+' invoice</button>':'')+
        ((a.tasks&&a.tasks.length)?'<button class="btn btn-secondary btn-sm" onclick="shareCaregiverTaskImage()" title="Copy the task chart (with times) as an image — paste into Messages or Mail">Task chart image</button>':'')+
        ((a.tasks&&a.tasks.length)?'<button class="btn btn-secondary btn-sm" onclick="exportCaregiverTaskSheet()" title="Open the printable task sheet (Print / Save PDF)">Task sheet (print/PDF)</button>':'')+
        '<button class="btn btn-secondary btn-sm" onclick="importDHS1210()">Import DHS-1210 / MDHHS-6064</button>'+
        '<button class="btn btn-secondary btn-sm" onclick="renderAuthPane(true)">Edit</button>'+
      '</div>'+
    '</div>'+
    '<div style="display:grid;grid-template-columns:1fr 1fr;gap:0 24px;">'+
      kv('Approved / month', _authHM(a)||'—')+
      kv('Provider rate', (a.rate!=null&&a.rate!=='')?'$'+esc(String(a.rate))+'/hr':'—')+
      kv('Effective', esc(a.effectiveDate||'—'))+
      kv('Reassessment due', esc(a.reassessDate||'—')+dueTxt, dueColor)+
      kv('Monthly total', (a.total!=null&&a.total!=='')?'$'+esc(Number(a.total).toFixed(2)):'—')+
      kv('Caseworker (ASW)', esc(a.aswName||prof.worker||'—'))+
    '</div>'+
    (rows?'<div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:#4d6c88;margin:18px 0 4px;">Tasks</div>'+
      '<table style="width:100%;border-collapse:collapse;font-size:12px;"><thead><tr style="color:#8296ab;text-align:left;font-size:9px;text-transform:uppercase;letter-spacing:.3px;"><th style="padding:4px 8px;">Task</th><th style="padding:4px 8px;">Time/Day</th><th style="padding:4px 8px;">Number of Days</th><th style="padding:4px 8px;text-align:right;">Time/Month</th><th style="padding:4px 8px;text-align:right;">Amount</th></tr></thead><tbody>'+rows+'</tbody></table>':'')+
    (a.importedAt?'<div style="font-size:10px;color:#94a7bd;margin-top:14px;">Imported '+esc(a.importedAt)+(a.sourceFile?' · '+esc(a.sourceFile):'')+'</div>':'')+
  '</div>';
}
function _authEditHtml(a){
  return '<div class="form-card" style="max-width:720px;">'+
    '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;gap:10px;flex-wrap:wrap;">'+
      '<h3 style="margin:0;">Edit Authorization</h3>'+
      '<button type="button" class="btn btn-secondary btn-sm" onclick="importDHS1210()">Import DHS-1210 / MDHHS-6064</button>'+
    '</div>'+
    '<div class="info-field-row" style="grid-template-columns:1fr 1fr;margin-bottom:16px;">'+
      '<div class="info-field"><label for="ei-auth-permonth">Approved Hours / Month <span style="font-weight:400;font-size:10px;color:#5c7590;">(HH:MM)</span></label><input id="ei-auth-permonth" placeholder="HH:MM" value="'+esc(_authHM(a))+'"></div>'+
      '<div class="info-field"><label for="ei-auth-rate">Rate ($/hr)</label><input id="ei-auth-rate" value="'+esc(a.rate!=null?String(a.rate):'')+'" inputmode="decimal"></div>'+
    '</div>'+
    '<div class="info-field-row" style="grid-template-columns:1fr 1fr;margin-bottom:16px;">'+
      '<div class="info-field"><label for="ei-auth-eff">Effective <span style="font-weight:400;font-size:10px;color:#5c7590;">(MM/DD/YYYY)</span></label><input id="ei-auth-eff" placeholder="MM/DD/YYYY" value="'+esc(a.effectiveDate||'')+'" onchange="_autoReassessFromEff()"></div>'+
      '<div class="info-field"><label for="ei-auth-reassess">Reassessment due <span style="font-weight:400;font-size:10px;color:#5c7590;">(MM/DD/YYYY)</span></label><input id="ei-auth-reassess" placeholder="MM/DD/YYYY" value="'+esc(a.reassessDate||'')+'">'+
        '<button type="button" onclick="_fillReassess()" style="margin-top:5px;font-size:10px;background:none;border:none;color:#185FA5;text-decoration:underline;cursor:pointer;padding:0;">Set to 6 months after effective</button></div>'+
    '</div>'+
    '<div class="info-field" style="margin-bottom:16px;"><label for="ei-auth-total">Monthly Total ($)</label><input id="ei-auth-total" value="'+esc(a.total!=null?String(a.total):'')+'" inputmode="decimal"></div>'+
    '<div style="margin-top:14px;">'+
      '<div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:#4d6c88;margin-bottom:6px;">Tasks</div>'+
      '<div style="display:grid;grid-template-columns:1.7fr 0.9fr 1.4fr 0.9fr 0.9fr 24px;gap:6px;margin-bottom:5px;font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.3px;color:#8296ab;">'+
        '<span style="min-width:0;overflow:hidden;">Task</span><span style="min-width:0;overflow:hidden;">Time/Day</span><span style="min-width:0;overflow:hidden;">Number of Days</span><span style="min-width:0;overflow:hidden;">Time/Month</span><span style="min-width:0;overflow:hidden;">Amount</span><span></span>'+
      '</div>'+
      '<div id="ei-auth-tasks"></div>'+
      '<button type="button" class="btn btn-secondary btn-sm" onclick="_authAddTaskRow()" style="margin-top:6px;font-size:11px;">+ Add task</button>'+
    '</div>'+
    '<div style="display:flex;gap:8px;margin-top:16px;align-items:center;">'+
      '<button class="btn btn-primary" onclick="saveAuthPane()">Save</button>'+
      '<button class="btn btn-secondary" onclick="renderAuthPane(false)">Cancel</button>'+
      ((a&&a.hours!=null)?'<button class="btn btn-danger btn-sm" onclick="_clearAuth()" style="margin-left:auto;">Remove authorization</button>':'')+
    '</div>';
}
function saveAuthPane(){
  if(!activeProfileName)return;
  var p=getProfiles(); var rec=p[activeProfileName]; if(!rec)return;
  var tasks=[];
  document.querySelectorAll('#ei-auth-tasks .ei-auth-task-row').forEach(function(r){
    var tk=(r.querySelector('.ei-at-task').value||'').trim();
    var pd=(r.querySelector('.ei-at-perDay').value||'').trim();
    var fq=(r.querySelector('.ei-at-freq').value||'').trim();
    var pm=(r.querySelector('.ei-at-perMonth').value||'').trim();
    var am=(r.querySelector('.ei-at-amount').value||'').trim();
    if(tk||pd||fq||pm||am)tasks.push({task:tk,perDay:pd,freq:fq,perMonth:pm,amount:am!==''?parseFloat(am):null});
  });
  var hm=_parseHM(document.getElementById('ei-auth-permonth').value);
  var re=(document.getElementById('ei-auth-reassess').value||'').trim();
  rec.authorization=Object.assign({},rec.authorization||{},{
    hours:hm.hours, minutes:hm.minutes,
    effectiveDate:(document.getElementById('ei-auth-eff').value||'').trim(),
    reassessDate:re, rate:(document.getElementById('ei-auth-rate').value||'').trim(),
    total:(function(v){return v!==''?parseFloat(v):'';})((document.getElementById('ei-auth-total').value||'').trim()),
    tasks:tasks,
  });
  saveProfilesLS(p); saveProfileSP(activeProfileName,rec);
  _syncReassessTask(activeProfileName, re);
  if(typeof addAuditEntry==='function')addAuditEntry(activeProfileName,'Authorization (DHS-1210) updated');
  renderAuthPane(false);
  if(typeof showToast==='function')showToast('✓ Authorization saved');
}
function _clearAuth(){
  showConfirm('Remove the DHS-1210 authorization from this client? The filed PDF stays in Documents.',function(){
    var p=getProfiles(); var rec=p[activeProfileName]; if(!rec)return;
    rec.authorization=null; saveProfilesLS(p); saveProfileSP(activeProfileName,rec);
    renderAuthPane(false);
  },{title:'Remove Authorization',okText:'Remove'});
}
// Keep a single "reassessment due" task per client in sync with the authorization date.
function _syncReassessTask(clientName, reassessMdy){
  if(!reassessMdy)return; // don't auto-delete existing tasks if the date is cleared
  var todos=getTodos();
  var existing=todos.find(function(t){return t.client===clientName && /^DHS-1210 reassessment/.test(t.text||'');});
  var dueYmd=_mdyToYmd(reassessMdy);
  var text='DHS-1210 reassessment due ('+reassessMdy+')';
  if(existing){
    if(existing.done)return; // leave a completed one alone
    existing.due=dueYmd; existing.text=text; saveTodos(todos);
    if(typeof saveTaskAPI==='function')saveTaskAPI(existing);
  } else {
    var nt={id:todoId(),text:text,client:clientName,due:dueYmd,note:'Auto-added from the DHS-1210 authorization. Review services with the client + provider.',done:false,kind:'dhs-reassess',created:new Date().toLocaleString()};
    todos.unshift(nt); saveTodos(todos);
    if(typeof saveTaskAPI==='function')saveTaskAPI(nt);
  }
  if(typeof updateTaskBadge==='function')updateTaskBadge();
}
function saveClientInfo(){
  if(!activeProfileName)return;
  var p=getProfiles();
  var rec=p[activeProfileName];
  if(((document.getElementById('ei-status')||{}).value||'')==='active' && !((document.getElementById('ei-start-date')||{}).value||'').trim()){showAlert('Service Start Date is required when the status is Active.');return;}
  // Require a DHS-1210 only when TURNING a client active — not on every edit of one that's already
  // active. Existing active clients (from before this rule, or before their PDF was imported) must
  // stay editable without being forced to import an authorization first.
  var _newStatus=((document.getElementById('ei-status')||{}).value||'');
  var _wasActive=(rec && rec.clientStatus==='active');
  // The DHS-1210 rule is CHAMPS-only — managed-care (carrier) clients are authorized by the carrier
  // and never have one, so without this exception (which createClient already has) a carrier client
  // created "In Progress" could never be flipped Active from this pane.
  var _newProgram=((document.getElementById('ei-program')||{}).value)||(rec&&rec.program)||'';
  if(_newStatus==='active' && !_wasActive && _newProgram!=='carrier' && !hasAuthorization(rec)){showAlert('A DHS-1210 authorization is required before a client can be Active.\n\nImport one on the Authorization tab first — that’s what officially makes you their agency.');return;}
  var first=(document.getElementById('ei-first').value||'').trim();
  var middle=(document.getElementById('ei-middle').value||'').trim();
  var last=(document.getElementById('ei-last').value||'').trim();
  var nickname=(document.getElementById('ei-nickname').value||'').trim();
  var newName=((first+' '+last).trim())||activeProfileName;
  rec.firstName=first;rec.middleName=middle;rec.lastName=last;rec.nickname=nickname;
  rec.clientName=newName;rec.medicaidId=document.getElementById('ei-medicaid').value;
  rec.medicare=(document.getElementById('ei-medicare')||{}).value||'';
  var _pgEl=document.getElementById('ei-program');if(_pgEl)rec.program=_pgEl.value;
  var _caEl=document.getElementById('ei-carrier');if(_caEl)rec.carrier=_caEl.value;
  var _meEl=document.getElementById('ei-member');if(_meEl)rec.memberId=_meEl.value;
  // Carrier/member # are HIDDEN for a CHAMPS client — clear them when the program switches so a
  // stale carrier doesn't linger on the record (inert for billing, but it shows up in exports).
  if(rec.program!=='carrier'){ rec.carrier=''; rec.memberId=''; }
  var dobEl=document.getElementById('ei-dob');if(dobEl)rec.dob=dobEl.value||'';
  var genderEl=document.getElementById('ei-gender');if(genderEl)rec.gender=genderEl.value||'';
  rec.hourlyRate=document.getElementById('ei-rate').value;
  var dlEl=document.getElementById('ei-dl');if(dlEl)rec.driversLicense=dlEl.value;
  var ssnEl=document.getElementById('ei-ssn');if(ssnEl)rec.ssn=ssnEl.value;
  rec.phone=document.getElementById('ei-phone').value;rec.clientEmail=document.getElementById('ei-cemail').value;
  rec.homePhone=(document.getElementById('ei-home-phone')||{}).value||'';
  rec.street=document.getElementById('ei-street').value;
  rec.city=document.getElementById('ei-city').value;
  rec.state=document.getElementById('ei-state').value;
  rec.zip=document.getElementById('ei-zip').value;
  rec.county=document.getElementById('ei-county').value;
  var cgValEl=document.getElementById('ei-caregiver-val');rec.caregiverId=cgValEl?cgValEl.value:'';
  var liveInEl=document.getElementById('ei-live-in');if(liveInEl)rec.liveIn=liveInEl.checked;
  var startDateEl=document.getElementById('ei-start-date');if(startDateEl)rec.startDate=startDateEl.value||'';
  var cwValEl=document.getElementById('ei-worker-val');var cwSearchEl=document.getElementById('ei-worker-search');
  rec.caseworkerId=cwValEl?cwValEl.value:'';rec.worker=cwSearchEl?cwSearchEl.value:'';
  var statusEl=document.getElementById('ei-status');if(statusEl)rec.clientStatus=statusEl.value;
  // (Authorization lives in its own tab now — see renderAuthPane/saveAuthPane.)
  if(newName!==activeProfileName){
    p[newName]=rec;delete p[activeProfileName];
    // lhca_id_map is keyed by client NAME and is how syncNewInvoices finds the client's db id.
    // Leaving it under the old name made every later invoice save silently no-op (no dbId → early return).
    try{ var _m=getIdMap(); if(_m[activeProfileName]!==undefined){ _m[newName]=_m[activeProfileName]; delete _m[activeProfileName]; localStorage.setItem('lhca_id_map',JSON.stringify(_m)); } }catch(e){}
    activeProfileName=newName;
  }
  saveProfilesLS(p);saveProfileSP(activeProfileName,p[activeProfileName]);
  unsavedChanges=false;
  logActivity('edit','Profile updated for '+activeProfileName);
  addAuditEntry(activeProfileName,'Profile information updated');
  var displayName=activeProfileName+(rec.nickname?' ('+rec.nickname+')':'');
  document.getElementById('detailName').textContent=displayName;
  var st=rec.clientStatus||'active';
  document.getElementById('detailMeta').innerHTML=(rec.medicaidId?'Medicaid: '+esc(rec.medicaidId):'No Medicaid ID')+(rec.phone?' &nbsp;·&nbsp; '+esc(rec.phone):'')+' &nbsp;<span class="cs-badge cs-'+st+'">'+clientStatusLabel(st)+'</span>'+' &nbsp;'+_programBadge(rec);
  renderSidebarClients();
  var btn=document.getElementById('saveInfoBtn');btn.textContent='Saved';setTimeout(function(){btn.textContent='Save Changes';},1800);
}
function deleteClient(){
  if(!activeProfileName)return;
  var name=activeProfileName;
  showConfirm('Delete "'+name+'" and all their data? This cannot be undone.',function(){
    aiTrack('ClientDeleted',{client:name});
    addAuditEntry(name,'CLIENT RECORD DELETED by '+currentUserEmail());   // durable record — aiTrack is scrubbed telemetry, not an audit trail
    var p=getProfiles();deleteProfileSP(name);delete p[name];
    try{localStorage.removeItem('lhca_draft_'+name);}catch(e){}
    saveProfilesLS(p);activeProfileName=null;logActivity('delete','Client deleted: '+name);navHome();
  },{title:'Delete Client',okText:'Delete'});
}

// ============================================================
//  INVOICE HISTORY
// ============================================================
function renderInvHistory(){
  if(!activeProfileName)return;
  var prof=getProfiles()[activeProfileName],invoices=(prof&&prof.invoices)?prof.invoices:[],c=document.getElementById('invHistoryContent');
  if(!invoices.length){
    c.innerHTML='<div class="empty-state"><h3>No invoices yet</h3><p style="font-size:13px;">Click "+ New Invoice" to create one.</p></div>';
    return;
  }
  c.innerHTML='<div style="margin-bottom:12px;display:flex;align-items:center;justify-content:space-between;"><span style="font-size:13px;color:#4d6c88;">'+invoices.length+' invoice'+(invoices.length!==1?'s':'')+' saved</span></div>';
  invoices.forEach(function(inv,idx){
    var st=inv.status||'draft';
    var daysSince=Math.floor((Date.now()-new Date(inv.savedAt))/86400000);
    var overdueFlag=(st==='submitted'&&daysSince>30)?'<span style="background:#ff8c00;color:#fff;font-size:10px;padding:1px 6px;border-radius:8px;margin-left:6px;">'+daysSince+'d overdue</span>':'';
    var card=document.createElement('div');card.className='inv-file-card';
    card.innerHTML=
      ''+
      '<div class="inv-file-info">'+
        '<div class="inv-file-period">'+esc(inv.billingPeriod)+overdueFlag+'</div>'+
        '<div class="inv-file-meta">Saved '+esc(inv.savedAt)+'</div>'+
        (inv.invoiceNote?'<div class="inv-file-note">'+esc(inv.invoiceNote)+'</div>':'')+
      '</div>'+
      '<div class="inv-file-actions">'+
        '<select class="status-select st-'+st+'" data-idx="'+idx+'" onchange="changeInvStatus(this)" onclick="event.stopPropagation()" title="Change status">'+
          '<option value="draft"'+(st==='draft'?' selected':'')+'>Draft</option>'+
          '<option value="submitted"'+(st==='submitted'?' selected':'')+'>Submitted</option>'+
          '<option value="paid"'+(st==='paid'?' selected':'')+'>Paid</option>'+
        '</select>'+
        '<button class="btn btn-secondary btn-sm" onclick="event.stopPropagation();promptInvNote('+idx+')" title="Add or edit a note for this invoice">'+(inv.invoiceNote?'Edit Note':'+ Note')+'</button>'+
        '<button class="btn btn-secondary btn-sm" onclick="event.stopPropagation();downloadInvoice('+idx+')" title="Download as PDF">Download</button>'+
        '<button class="btn btn-danger btn-sm" onclick="event.stopPropagation();deleteInv('+idx+')" title="Delete this invoice">Delete</button>'+
      '</div>';
    card.addEventListener('click',function(){loadInvFromHistory(idx);});
    c.appendChild(card);
  });
}
function changeInvStatus(sel){
  var idx=parseInt(sel.dataset.idx),next=sel.value;
  var p=getProfiles();if(!p[activeProfileName]||!p[activeProfileName].invoices[idx])return;
  var inv=p[activeProfileName].invoices[idx];
  var period=inv.billingPeriod;
  var prev=inv.status||'draft';
  if(prev===next)return;
  // Risky transitions need confirmation
  var needsConfirm=null;
  if(next==='paid'){
    needsConfirm={
      title:'Mark as Paid?',
      message:'Mark invoice '+period+' as PAID?\n\nPaid invoices are locked from further edits. Only mark Paid after you have actually received payment.',
      okText:'Mark as Paid'
    };
  } else if(next==='submitted'&&prev==='draft'){
    needsConfirm={
      title:'Mark as Submitted?',
      message:'Mark invoice '+period+' as SUBMITTED?\n\nThis indicates the invoice has been sent to the caseworker. Only do this if you have actually emailed/delivered it.',
      okText:'Mark as Submitted'
    };
  } else if(prev==='paid'&&next!=='paid'){
    needsConfirm={
      title:'Unlock Paid Invoice?',
      message:'Invoice '+period+' is currently marked Paid. Changing it back to '+next+' will allow edits — only do this if the prior Paid status was a mistake.',
      okText:'Change to '+next.charAt(0).toUpperCase()+next.slice(1),
      danger:true
    };
  }
  function applyChange(){
    sel.className='status-select st-'+next;
    var p2=getProfiles();var inv2=p2[activeProfileName].invoices[idx];if(!inv2)return;
    inv2.status=next;
    // saveProfileSP already persists status via the invoice upsert; a separate status
    // PATCH here would write the same row twice and (with optimistic concurrency) make
    // the save conflict with itself, so it was removed.
    saveProfilesLS(p2);saveProfileSP(activeProfileName,p2[activeProfileName]);
    logActivity('status','Invoice '+period+' for '+activeProfileName+' marked '+next);
    updateStats();renderOverviewPane();
  }
  if(needsConfirm){
    showConfirm(needsConfirm.message,applyChange,{title:needsConfirm.title,okText:needsConfirm.okText,danger:!!needsConfirm.danger,onCancel:function(){sel.value=prev;}});
  } else {
    applyChange();
  }
}
function promptInvNote(idx){
  var p=getProfiles();if(!p[activeProfileName]||!p[activeProfileName].invoices[idx])return;
  var cur=p[activeProfileName].invoices[idx].invoiceNote||'';
  showPrompt('Note for this invoice:',cur,function(note){
    var p2=getProfiles();if(!p2[activeProfileName]||!p2[activeProfileName].invoices[idx])return;
    // invoice_note is NVARCHAR(1000). An over-length note used to fail the ENTIRE client save with a
    // generic 500 that named nothing — say so here instead of letting it become a failed save.
    if((note||'').length>1000){
      showAlert('That note is '+note.length+' characters. The limit is 1000 — shorten it and save again.',{title:'Note too long'});
      return;
    }
    p2[activeProfileName].invoices[idx].invoiceNote=note||'';
    saveProfilesLS(p2);saveProfileSP(activeProfileName,p2[activeProfileName]);renderInvHistory();renderOverviewPane();
  },{title:'Invoice Note',okText:'Save Note'});
}
function saveInvNote(input){
  var p=getProfiles(),idx=parseInt(input.dataset.idx);
  var clientName=activeProfileName; // capture the client now; activeProfileName may change before the flush
  if(p[clientName]&&p[clientName].invoices&&p[clientName].invoices[idx]){
    p[clientName].invoices[idx].invoiceNote=input.value;
    saveProfilesLS(p);
    // Debounce the backend save (per-client key so switching clients can't cancel it), and
    // register it so a tab-close can flush the pending save (F4).
    _scheduleNoteSave('inv:'+clientName, function(){
      // Re-read the CURRENT store at flush time (not the keystroke-time snapshot) so a
      // concurrent status/note/reload write in the 600ms window isn't reverted, and use the
      // captured clientName in case the user navigated away.
      var pNow=getProfiles(); if(!pNow[clientName])return;
      // D8: surface a failure toast instead of silently saving to LS only.
      var doSave=function(){return surfaceSaveFailure(saveProfileSP(clientName,pNow[clientName],true),'Invoice note',doSave);};
      doSave();
    });
  }
}
function deleteInv(idx){
  var p=getProfiles(),inv=p[activeProfileName].invoices[idx];
  if(!inv)return;
  showConfirm('Permanently delete the '+inv.billingPeriod+' invoice for '+activeProfileName+'? This cannot be undone.',function(){doDeleteInv(idx);},{title:'Delete Invoice',okText:'Delete'});
}
function doDeleteInv(idx){
  var p=getProfiles(),inv=p[activeProfileName].invoices[idx];
  if(!inv)return;
  var clientName=activeProfileName;
  addAuditEntry(clientName,'Invoice DELETED for '+(inv.billingPeriod||'(no period)')+' by '+currentUserEmail());
  // Remove from LS by INVOICE IDENTITY (dbId, else period+savedAt) — the array index can
  // shift, and this runs later (after the server delete confirms).
  var removeLocally=function(){
    var p2=getProfiles(); if(!p2[clientName]||!p2[clientName].invoices)return;
    var i = inv.dbId
      ? p2[clientName].invoices.findIndex(function(x){return x.dbId===inv.dbId;})
      : p2[clientName].invoices.findIndex(function(x){return !x.dbId && (x.billingPeriod||'')===(inv.billingPeriod||'') && x.savedAt===inv.savedAt;});
    if(i>=0){ p2[clientName].invoices.splice(i,1); saveProfilesLS(p2); }
    renderInvHistory();
  };
  // Delete server-side FIRST; remove locally ONLY on a confirmed delete. A no-dbId invoice
  // was never synced, so just remove it locally.
  deleteInvoiceAPI(inv.dbId,clientName,inv.billingPeriod,removeLocally);
}
async function downloadInvoice(idx){
  var inv=getProfiles()[activeProfileName].invoices[idx];if(!inv)return;
  var clientName=activeProfileName;
  var period=inv.billingPeriod||'';
  // Use the same vector PDF path as Print/PDF and Email Worker — works whether or not the invoice page is currently open
  var savedActive=document.querySelector('.page.active')&&document.querySelector('.page.active').id;
  var invPage=document.getElementById('page-invoice');
  invPage.style.position='fixed';invPage.style.left='-9999px';invPage.style.top='0';invPage.style.zIndex='-1';
  invPage.classList.add('active');
  try{
    await loadInvoiceForCapture(clientName,inv,period);
    var base64=await captureInvoicePDF();
    var bin=atob(base64);
    var bytes=new Uint8Array(bin.length);
    for(var i=0;i<bin.length;i++)bytes[i]=bin.charCodeAt(i);
    var blob=new Blob([bytes],{type:'application/pdf'});
    var url=URL.createObjectURL(blob);
    var fname=(clientName.replace(/[^a-z0-9]/gi,'_'))+'_'+(period.replace('/','_'))+'.pdf';
    var a=document.createElement('a');a.href=url;a.download=fname;
    document.body.appendChild(a);a.click();document.body.removeChild(a);
    setTimeout(function(){URL.revokeObjectURL(url);},2000);
  }catch(e){showAlert('Download failed: '+e.message);console.error(e);}
  finally{
    invPage.classList.remove('active');
    invPage.style.position='';invPage.style.left='';invPage.style.top='';invPage.style.zIndex='';
    document.querySelectorAll('.page').forEach(function(el){el.classList.remove('active');});
    if(savedActive){var p=document.getElementById(savedActive);if(p)p.classList.add('active');}
  }
}

// ============================================================
//  NOTES
// ============================================================
function renderNotesPane(){
  if(!activeProfileName)return;
  var prof=getProfiles()[activeProfileName],ta=document.getElementById('clientNotesArea');
  ta.value=(prof&&prof.clientNotes)?prof.clientNotes:'';
  if(ta._nl)ta.removeEventListener('input',ta._nl);
  ta._nl=function(){
    // Capture the client NOW — activeProfileName may change before the 600ms flush. And write
    // LS SYNCHRONOUSLY (not inside the timer) so closing the tab / switching clients within the
    // debounce can't lose the note. Only the backend save is debounced. (Parity with saveInvNote.)
    var clientName=activeProfileName, val=ta.value;
    var pNow=getProfiles();
    if(pNow[clientName]){ pNow[clientName].clientNotes=val; saveProfilesLS(pNow); }
    // Debounce the backend save, but register it so a tab-close can flush it (F4).
    _scheduleNoteSave('client:'+clientName, function(){
      var p2=getProfiles(); if(!p2[clientName])return;
      // D8: only claim "Saved ✓" after the API resolves; surface failure otherwise.
      var doSave=function(){return flashQuietSave(saveProfileSP(clientName,p2[clientName],true),'notesSavedFlash','Client note',doSave);};
      doSave();
    });
  };
  ta.addEventListener('input',ta._nl);
  var c=document.getElementById('invNotesContent'),invoices=(prof&&prof.invoices)?prof.invoices:[];
  if(!invoices.length){c.innerHTML='<div style="color:#5c7590;font-size:13px;">No saved invoices yet.</div>';return;}
  c.innerHTML='';
  invoices.forEach(function(inv,idx){
    var row=document.createElement('div');row.style.cssText='display:flex;align-items:center;gap:10px;margin-bottom:8px;';
    row.innerHTML='<span style="min-width:72px;font-size:12px;font-weight:600;color:#185FA5;">'+esc(inv.billingPeriod)+'</span>'+
      '<input maxlength="1000" style="flex:1;padding:6px 10px;border:1px solid #d0d8e4;border-radius:5px;font-size:12px;font-family:Arial,sans-serif;outline:none;color:#1a2b45;" value="'+esc(inv.invoiceNote||'')+'" placeholder="Note…" data-idx="'+idx+'">';
    c.appendChild(row);
  });
  c.querySelectorAll('input[data-idx]').forEach(function(inp){
    inp.addEventListener('change',function(){
      var p2=getProfiles(),i=parseInt(inp.dataset.idx);
      if(p2[activeProfileName]&&p2[activeProfileName].invoices[i]){
        p2[activeProfileName].invoices[i].invoiceNote=inp.value;
        saveProfilesLS(p2);
        // D8: surface a failure toast instead of silently saving to LS only.
        var doSave=function(){return surfaceSaveFailure(saveProfileSP(activeProfileName,p2[activeProfileName],true),'Invoice note',doSave);};
        doSave();
      }
    });
  });
}

// ============================================================
//  DOCUMENTS (Azure Blob Storage)
// ============================================================
function getHcClientId(){
  var prof=getProfiles()[activeProfileName];
  return prof&&prof._dbId?prof._dbId:null;
}
// ═══════════════════════════════════════════════════════════════════════════
// DHS-1210 reader — extract authorized hours/tasks from the MDHHS approval packet.
// Runs entirely in-browser via pdf.js (self-hosted); the PHI on the form never
// leaves the machine. Output feeds a review/confirm step, then the client profile.
// ═══════════════════════════════════════════════════════════════════════════
function _dhsReady(){ return typeof pdfjsLib!=='undefined'; }

// Group one page's text items into ordered lines (top→bottom, then left→right).
async function _dhsPageLines(page){
  var tc=await page.getTextContent();
  var buckets={};
  tc.items.forEach(function(it){
    if(!it.str||!it.str.trim())return;
    var x=it.transform[4], y=Math.round(it.transform[5]);
    var key=Object.keys(buckets).find(function(k){return Math.abs(k-y)<=3;}); // merge a row's items
    if(key===undefined)key=y;
    (buckets[key]=buckets[key]||[]).push({x:x,s:it.str});
  });
  return Object.keys(buckets).map(Number).sort(function(a,b){return b-a;}).map(function(y){
    return buckets[y].sort(function(a,b){return a.x-b.x;}).map(function(o){return o.s;}).join(' ').replace(/\s+/g,' ').trim();
  }).filter(Boolean);
}

// Parse the grouped lines into a structured authorization. Self-checks the extraction
// against the form's own printed totals (task $ = total; task minutes = approved hours).
// Handles both DHS-1210 (older) and MDHHS-6064-P (newer, effective 9-25) — the two
// forms share the same Section 3 task table structure, only the header/labels differ.
function parseDHS1210(pages){
  var lines=[].concat.apply([],pages);
  var flat=lines.join(' ').replace(/\s+/g,' ');
  var out={warnings:[]};
  // Detect which form this is so downstream UI can label it correctly
  out.formType = /MDHHS-?6064/i.test(flat) ? 'MDHHS-6064' : 'DHS-1210';
  // Read "approved for 29 Hours and 47 Minutes per month" (DHS-1210 style).
  // MDHHS-6064 doesn't have this line — it uses the "Total per month" row instead,
  // which we back-fill from a sibling regex below if the primary miss.
  var m=flat.match(/approved for\s+(\d+)\s*Hours?\s+(?:and\s+)?(\d+)\s*Minutes?/i)
       || flat.match(/(\d+)\s*Hours?\s+(?:and\s+)?(\d+)\s*Minutes?\s+per\s+month/i);
  if(m){out.hours=+m[1];out.minutes=+m[2];}
  // MDHHS-6064 fallback: "Total per month  HH:MM  $NNN.NN" — take the HH:MM as approved hours.
  // Tried BEFORE the unlabeled fallback below: it's anchored to a real label, so it can't be a task row.
  if(out.hours==null){
    var tm=flat.match(/Total\s*per\s*month[^0-9]*(\d{1,3}):(\d{2})/i);
    if(tm){out.hours=+tm[1]; out.minutes=+tm[2];}
  }
  // Last resort: ANY "N Hours M Minutes" on the form. That can be a single task's time rather than
  // the monthly authorization, so it's flagged — the review modal shows the warning, and the
  // task-minutes-vs-approved reconciliation check is what confirms it.
  if(out.hours==null){
    var lm=flat.match(/(\d+)\s*Hours?\s+(?:and\s+)?(\d+)\s*Minutes?/i);
    if(lm){ out.hours=+lm[1]; out.minutes=+lm[2]; out.hoursGuessed=true;
      out.warnings.push('approved hours (read from an unlabeled line — verify the total)'); }
  }
  if(out.hours==null) out.warnings.push('approved hours');
  // Effective date. DHS-1210 says "effective MM/DD/YYYY". MDHHS-6064 uses a plain
  // "Date  MM/DD/YYYY" in Section 2. Try both.
  var e=flat.match(/effective\s+(\d{2}\/\d{2}\/\d{4})/i);
  if(!e && out.formType==='MDHHS-6064'){
    // Grab a Date field that appears in Section 2 near the ASW block
    e=flat.match(/(?:^|\s)Date\s+(\d{2}\/\d{2}\/\d{4})/i);
  }
  if(e)out.effectiveDate=e[1]; else out.warnings.push('effective date');
  var em=flat.match(/([A-Za-z][A-Za-z.]*@michigan\.gov)/i); if(em)out.aswEmail=em[1];
  var ph=flat.match(/(\d{3}-\d{3}-\d{4})/); if(ph)out.aswPhone=ph[1];
  var co=flat.match(/\b(\d{2}-[A-Z]{3,})\b/); if(co)out.county=co[1];
  // Client ID (= Medicaid #). Anchored to the "Client ID" label so it doesn't grab the Case Number.
  var mid=flat.match(/Client ID(?:\s*Number)?\s*:?\s*([0-9]{6,12})/i); if(mid)out.medicaidId=mid[1];
  pages.forEach(function(ls){ls.forEach(function(ln,i){
    if(out.aswName)return;
    if(/Adult Services Worker.*Name/i.test(ln)){var nxt=(ls[i+1]||'').match(/^([A-Z]\.?\s?[A-Z][a-z]+)/); if(nxt)out.aswName=nxt[1].trim();}
  });});
  if(!out.aswName){var a2=flat.match(/\b([A-Z]\s[A-Z][a-z]+)\b\s+\d{3}-\d{3}-\d{4}/); if(a2)out.aswName=a2[1];}
  // Provider name + pay rate: a "<name> $NN.NN" line that isn't a task row.
  pages.forEach(function(ls){ls.forEach(function(ln){
    var r=ln.match(/^([A-Za-z][A-Za-z .,&'-]+?)\s+\$\s*(\d{1,3}\.\d{2})$/);
    if(r && !/\d{2}:\d{2}/.test(ln) && out.rate===undefined){out.providerName=r[1].trim();out.rate=+r[2];}
  });});
  var t=flat.match(/Total per month[\s\S]*?\$\s*([\d,]+\.\d{2})/i); if(t)out.printedTotal=+t[1].replace(/,/g,'');
  // Task rows: <task> <HH:MM/day> <frequency> <HH:MM/month> [$amount]
  var tasks=[],seen={};
  pages.forEach(function(ls){ls.forEach(function(ln){
    var r=ln.match(/^(.+?)\s+(\d{2}:\d{2})\s+(\d+ days? per week|Once per (?:week|month)|\d+ days? per month)\s+(\d{2}:\d{2})(?:\s+\$?([\d,]+\.\d{2}))?/i);
    if(r){var nm=r[1].trim(); if(!seen[nm]){seen[nm]=1;tasks.push({task:nm,perDay:r[2],freq:r[3],perMonth:r[4],amount:r[5]?+r[5].replace(/,/g,''):null});}}
  });});
  out.tasks=tasks;
  if(tasks.length){
    var sum=Math.round(tasks.filter(function(x){return x.amount!=null;}).reduce(function(a,x){return a+x.amount;},0)*100)/100;
    out.taskAmountSum=sum;
    out.total=out.printedTotal!=null?out.printedTotal:sum; // bill from summed line items; printed total cross-checks
    out.amountReconciles=out.printedTotal!=null?Math.abs(sum-out.printedTotal)<0.02:null;
  }
  if(out.hours!=null){
    var mins=tasks.reduce(function(a,x){var p=(x.perMonth||'0:0').split(':');return a+(+p[0])*60+(+p[1]);},0);
    out.taskMinuteSum=mins; out.approvedTotalMin=out.hours*60+out.minutes;
    out.timeReconciles=Math.abs(mins-out.approvedTotalMin)<=1;
  }
  // Reassessment due = effective + 6 months, advanced to the next date on/after today (so an old
  // form doesn't show a long-past reassessment — see _nextReassessment).
  if(out.effectiveDate){ out.reassessDate=_nextReassessment(out.effectiveDate); }
  return out;
}

function readDHS1210File(file){
  return file.arrayBuffer().then(function(buf){
    if(!pdfjsLib.GlobalWorkerOptions.workerSrc) pdfjsLib.GlobalWorkerOptions.workerSrc='vendor/pdf.worker.min.js';
    // isEvalSupported:false is the documented mitigation for CVE-2024-4367 (arbitrary JS execution
    // from a crafted PDF, fixed upstream in 4.2.67; this bundle is 3.11.174). DHS-1210 files arrive
    // from outside the agency — emailed, scanned, downloaded — and this origin holds the signed-in
    // user's tokens and the whole roster, so the font-rendering eval path must stay closed.
    return pdfjsLib.getDocument({data:buf, isEvalSupported:false}).promise;
  }).then(function(pdf){
    var jobs=[]; for(var p=1;p<=pdf.numPages;p++)jobs.push(pdf.getPage(p).then(_dhsPageLines));
    return Promise.all(jobs);
  }).then(function(pages){
    // A clean digital PDF has a text layer — parse it here (free + instant). A SCANNED/photographed
    // PDF has essentially no extractable text, so pdf.js returns almost nothing; fall back to server
    // OCR (Azure Document Intelligence), which reads scans and returns the same pages shape.
    var totalText=pages.reduce(function(a,p){return a+p.join('').length;},0);
    if(totalText>=200) return parseDHS1210(pages);
    return _dhsOcrFile(file).then(function(ocrPages){
      return parseDHS1210(ocrPages&&ocrPages.length?ocrPages:pages);
    });
  });
}
// Read a File to base64 (no data: prefix) for the OCR endpoint.
function _fileToBase64(file){
  return new Promise(function(res,rej){
    var fr=new FileReader();
    fr.onload=function(){ var s=String(fr.result||''); var i=s.indexOf(','); res(i>=0?s.slice(i+1):s); };
    fr.onerror=function(){rej(new Error('read failed'));};
    fr.readAsDataURL(file);
  });
}
// Send a scanned/photo authorization to the backend OCR (Document Intelligence) and return
// pages = [[line,...],...] — the same shape pdf.js produces — for parseDHS1210().
function _dhsOcrFile(file){
  var status=document.getElementById('hcDocStatus'); if(status)status.textContent='Reading scanned authorization (OCR)…';
  return _fileToBase64(file).then(function(b64){
    return fetch(API_BASE+'/dhs-ocr',{method:'POST',headers:apiHeaders(),body:JSON.stringify({fileBase64:b64})});
  }).then(function(r){
    if(r.status===401) throw new Error('Sign in first (Settings) to read a scanned authorization.');
    if(!r.ok) throw new Error('OCR service error (HTTP '+r.status+')');
    return r.json();
  }).then(function(j){ return (j&&Array.isArray(j.pages))?j.pages:[]; });
}

// Entry point — triggered from the client Documents pane.
function importDHS1210(){
  if(!activeProfileName){showAlert('Open a client first.');return;}
  if(!_dhsReady()){showAlert('PDF reader is still loading — try again in a second.');return;}
  var inp=document.getElementById('dhsImportFile'); if(!inp)return;
  inp.value=''; inp.click();
}
function handleDhsImport(input){
  var file=input&&input.files&&input.files[0]; if(!file)return;
  var isPdf=/\.pdf$/i.test(file.name)||file.type==='application/pdf';
  var isImg=/\.(jpe?g|png|tiff?|bmp|heic|webp)$/i.test(file.name)||/^image\//.test(file.type);
  if(!isPdf&&!isImg){showAlert('Please choose the DHS-1210 / MDHHS-6064 PDF, or a photo/scan of it.');return;}
  var status=document.getElementById('hcDocStatus'); if(status)status.textContent='Reading authorization…';
  // PDFs: try pdf.js first, OCR-fallback for scans. Images (phone photos): straight to OCR.
  var work=isPdf?readDHS1210File(file):_dhsOcrFile(file).then(parseDHS1210);
  work.then(function(res){
    if(status)status.textContent='';
    showDhsReview(file,res);
  }).catch(function(e){
    if(status)status.textContent='';
    showAlert('Could not read that file: '+((e&&e.message)||e));
  });
}

function _dhsChip(ok){return ok?'<span style="color:#1a7f4b;font-weight:700;">✓</span>':(ok===false?'<span style="color:#c0392b;font-weight:700;">✗</span>':'<span style="color:#5c7590;">–</span>');}

// #5: fields a DHS-1210 suggests updating on the client / matched caseworker. Only when the form
// has a value that DIFFERS from what's stored — never suggests blanking a field. Client NAME is
// intentionally excluded (renaming a client changes its record key — too risky to auto-suggest).
function _dhsSuggestedUpdates(res, prof, cw){
  var out=[];
  // `verify` = this value was read heuristically off the form (the ASW phone/email are simply the
  // first phone/email found anywhere on it — which can be a county main line or a fax). Combined
  // with `from` being non-empty (an OVERWRITE), the review renders it UNCHECKED so nothing already
  // on file is replaced by an un-confirmed OCR read unless the user ticks it deliberately.
  var add=function(label,target,field,from,to,id,verify){
    to=(to==null?'':String(to)).trim(); from=(from==null?'':String(from)).trim();
    if(to && to!==from) out.push({label:label,target:target,field:field,from:from,to:to,id:id||null,
      overwrite:!!from, verify:!!verify, shared:(target==='caseworker')});
  };
  if(cw){
    add('Caseworker email','caseworker','email',cw.email,res.aswEmail,cw.id,true);
    // Compare phones by DIGITS only — a stored "3135051660" and a form "313-505-1660" are the same
    // number, so a formatting-only difference must not be flagged as a change (mirrors the rate
    // normalization below). Only suggest when the actual digits differ.
    var _digits=function(s){return (s==null?'':String(s)).replace(/\D/g,'');};
    if(_digits(res.aswPhone) && _digits(res.aswPhone)!==_digits(cw.phone)){
      add('Caseworker phone','caseworker','phone',cw.phone,res.aswPhone,cw.id,true);
    }
  }
  add('Client Medicaid ID','client','medicaidId',prof&&prof.medicaidId,res.medicaidId,null,false);
  // The client "Hourly Rate" field is NOT the invoice rate — invoices always bill the state rate
  // ($27) — so the DHS provider rate is no longer pushed into it (that mismatch was confusing).
  return out;
}
function showDhsReview(file,res){
  var ex=document.getElementById('dhsReviewModal'); if(ex)ex.remove();
  // Capture the client this review is FOR. activeProfileName is a mutable global: a background
  // load/sync (or the user clicking another client) can change or null it while the modal is open,
  // which used to drop the import or write it to the WRONG client at Apply time.
  var targetName=activeProfileName;
  var prof=getProfiles()[targetName]||{};
  var cws=getCaseworkers();
  var match=res.aswEmail?cws.find(function(c){return (c.email||'').toLowerCase()===res.aswEmail.toLowerCase();}):null;
  var hasCw=!!prof.caseworkerId;
  var row=function(l,v){return '<div style="display:flex;justify-content:space-between;gap:12px;padding:3px 0;font-size:13px;"><span style="color:#5c7590;">'+l+'</span><span style="font-weight:600;color:#1a2b45;text-align:right;">'+v+'</span></div>';};
  var warn=(res.warnings&&res.warnings.length)?'<div style="background:#fff3cd;border:1px solid #ffeaa7;color:#856404;font-size:12px;padding:6px 10px;border-radius:6px;margin:8px 0;">Couldn\'t read: '+esc(res.warnings.join(', '))+' — verify before saving.</div>':'';
  var tasksHtml=(res.tasks||[]).map(function(t){
    return '<tr><td style="padding:2px 6px;">'+esc(t.task)+'</td><td style="padding:2px 6px;color:#5c7590;">'+esc(t.freq)+'</td><td style="padding:2px 6px;text-align:right;">'+esc(t.perMonth)+'</td><td style="padding:2px 6px;text-align:right;">'+(t.amount!=null?'$'+t.amount.toFixed(2):'—')+'</td></tr>';
  }).join('');
  var cwSection='';
  if(match){
    cwSection='<label style="display:flex;gap:8px;align-items:flex-start;font-size:13px;margin-top:4px;"><input type="checkbox" id="dhs-link-cw" '+(!hasCw?'checked':'')+' style="margin-top:3px;"><span>Set this client\'s caseworker to <b>'+esc(match.name)+'</b> <span style="color:#1a7f4b;">(matched by email)</span></span></label>';
  } else if(res.aswName){
    cwSection='<label style="display:flex;gap:8px;align-items:flex-start;font-size:13px;margin-top:4px;"><input type="checkbox" id="dhs-add-cw" checked style="margin-top:3px;"><span>Add <b>'+esc(res.aswName)+'</b> ('+esc(res.aswEmail||'no email')+') as a caseworker and assign to this client</span></label>';
  }
  // #5: the caseworker to potentially update is the client's current one (by id or name), or the
  // email-matched one — NOT just the email match (a differing email wouldn't match at all).
  var cwForUpdate=cws.find(function(c){return c.id===prof.caseworkerId;}) || match ||
    (res.aswName?cws.find(function(c){return (c.name||'').toLowerCase()===String(res.aswName).toLowerCase();}):null);
  var updates=_dhsSuggestedUpdates(res, prof, cwForUpdate);
  var updatesSection='';
  if(updates.length){
    updatesSection='<div style="border-top:1px solid #e8ecf0;margin-top:12px;padding-top:10px;">'+
      '<div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:#4d6c88;margin-bottom:6px;">Update from this form?</div>'+
      updates.map(function(u,i){
        // Filling a BLANK is pre-checked; REPLACING a value already on file is not — the stored one
        // may well be the correct one, and this is OCR output, not a verified read.
        var notes=[];
        if(u.overwrite)notes.push('replaces the value on file — verify against the form');
        if(u.verify)notes.push('read from anywhere on the form (could be a county main line/fax)');
        if(u.shared)notes.push('shared caseworker record — changes it for every client assigned to them');
        return '<label style="display:flex;gap:8px;align-items:flex-start;font-size:13px;margin-bottom:6px;"><input type="checkbox" id="dhs-upd-'+i+'"'+(u.overwrite?'':' checked')+' style="margin-top:3px;"><span>'+esc(u.label)+': '+
          (u.from?'<span style="color:#b03030;text-decoration:line-through;">'+esc(u.from)+'</span> &rarr; ':'<span style="color:#94a7bd;">(empty) &rarr; </span>')+
          '<b>'+esc(u.to)+'</b>'+
          (notes.length?'<span style="display:block;color:#b8860b;font-size:11px;font-weight:600;">⚠ '+esc(notes.join(' · '))+'</span>':'')+
          '</span></label>';
      }).join('')+
    '</div>';
  }
  // Render whenever a check actually RAN, not only when one passed — the case that matters most (OCR
  // misread BOTH the hours and the total, so both checks fail) used to render nothing at all. The
  // wording now follows the outcome instead of always saying "match".
  var _recRan=(res.timeReconciles!=null)||(res.amountReconciles!=null);
  var _recFail=(res.timeReconciles===false)||(res.amountReconciles===false);
  var reconcileNote=_recRan
    ? '<div style="font-size:12px;color:'+(_recFail?'#c0392b':'#1a7f4b')+';margin-top:6px;'+(_recFail?'font-weight:700;':'')+'">'+
        (res.timeReconciles!=null?(_dhsChip(res.timeReconciles)+' Task times '+(res.timeReconciles?'match':'DO NOT match')+' the approved hours'):'')+
        (res.amountReconciles!=null?(' &nbsp; '+_dhsChip(res.amountReconciles)+' Task $ '+(res.amountReconciles?'match':'DO NOT match')+' the printed total'):'')+
      '</div>'
    : '';
  var ov=document.createElement('div');
  ov.id='dhsReviewModal'; ov.className='modal-overlay open';
  ov.innerHTML='<div class="modal-box" style="max-width:560px;max-height:88vh;overflow:auto;">'+
    '<h3>'+esc(res.formType||'DHS-1210')+' — review before saving</h3>'+
    '<p style="font-size:12px;color:#5c7590;margin:-4px 0 8px;">Read from <b>'+esc(file.name)+'</b>. Nothing was sent anywhere — parsed in your browser.</p>'+
    warn+
    '<div style="background:#f4f6f9;border-radius:8px;padding:10px 12px;">'+
      row('Approved / month', (res.hours!=null?res.hours+'h '+res.minutes+'m':'—'))+
      row('Effective', esc(res.effectiveDate||'—'))+
      row('Reassessment due', esc(res.reassessDate||'—'))+
      row('Provider rate', (res.rate!=null?'$'+res.rate+'/hr':'—'))+
      row('Monthly total', (res.total!=null?'$'+res.total.toFixed(2):'—'))+
      row('Adult Services Worker', esc((res.aswName||'—')+(res.aswEmail?' · '+res.aswEmail:'')))+
    '</div>'+
    reconcileNote+
    (tasksHtml?'<div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:#4d6c88;margin:12px 0 4px;">Tasks ('+res.tasks.length+')</div>'+
      '<table style="width:100%;border-collapse:collapse;font-size:12px;"><thead><tr style="color:#5c7590;text-align:left;"><th style="padding:2px 6px;">Task</th><th style="padding:2px 6px;">Frequency</th><th style="padding:2px 6px;text-align:right;">Time/mo</th><th style="padding:2px 6px;text-align:right;">Amount</th></tr></thead><tbody>'+tasksHtml+'</tbody></table>':'')+
    (cwSection?'<div style="border-top:1px solid #e8ecf0;margin-top:12px;padding-top:10px;">'+cwSection+'</div>':'')+
    updatesSection+
    // #6: a DHS-1210 makes you the client's agency → offer to set them Active with a start date.
    '<div style="border-top:1px solid #e8ecf0;margin-top:12px;padding-top:10px;">'+
      '<label style="display:flex;gap:8px;align-items:flex-start;font-size:13px;"><input type="checkbox" id="dhs-set-active" '+((prof.clientStatus||'active')!=='active'?'checked':'')+' style="margin-top:3px;"><span>Set <b>'+esc(targetName)+'</b> to <b>Active</b> — this DHS-1210 officially makes you their agency</span></label>'+
      '<div style="display:flex;align-items:center;gap:8px;margin-top:8px;padding-left:24px;flex-wrap:wrap;">'+
        '<label for="dhs-start-date" style="font-size:12px;color:#5c7590;">Service start date</label>'+
        '<input type="date" id="dhs-start-date" value="'+esc(_firstOfEffectiveMonth(res.effectiveDate))+'" style="padding:5px 8px;border:1px solid #d0d8e4;border-radius:5px;font-size:13px;">'+
        '<span style="font-size:11px;color:#94a7bd;">defaults to the 1st of the effective month — edit if needed</span>'+
      '</div>'+
    '</div>'+
    '<div class="modal-row" style="justify-content:flex-end;gap:8px;margin-top:14px;">'+
      '<button class="btn btn-secondary" onclick="document.getElementById(\'dhsReviewModal\').remove()">Cancel</button>'+
      '<button class="btn btn-primary" id="dhs-confirm-btn">Save to '+esc(targetName)+'</button>'+
    '</div>'+
  '</div>';
  document.body.appendChild(ov);
  document.getElementById('dhs-confirm-btn').onclick=function(){
    var opts={link:!!(document.getElementById('dhs-link-cw')&&document.getElementById('dhs-link-cw').checked),
              add:!!(document.getElementById('dhs-add-cw')&&document.getElementById('dhs-add-cw').checked),
              setActive:!!(document.getElementById('dhs-set-active')&&document.getElementById('dhs-set-active').checked),
              startDate:(document.getElementById('dhs-start-date')&&document.getElementById('dhs-start-date').value)||'',
              updates:updates.filter(function(u,i){var el=document.getElementById('dhs-upd-'+i);return el&&el.checked;}),
              match:match, target:targetName};
    _applyDhsImport(file,res,opts);
    ov.remove();
  };
}

function _applyDhsImport(file,res,opts){
  // Apply to the client captured when the review opened — never to whatever is active NOW.
  var name=(opts&&opts.target)||activeProfileName;
  var p=getProfiles(); var prof=name?p[name]:null;
  if(!prof){ showAlert('“'+(name||'That client')+'” is no longer available, so the authorization was NOT saved. Re-open the client and import the form again.'); return; }
  prof.authorization={
    hours:res.hours!=null?res.hours:null, minutes:res.minutes!=null?res.minutes:null,
    effectiveDate:res.effectiveDate||'', reassessDate:res.reassessDate||'',
    rate:res.rate!=null?res.rate:'', total:res.total!=null?res.total:'',
    providerName:res.providerName||'', county:res.county||'',
    tasks:res.tasks||[], aswName:res.aswName||'', aswEmail:res.aswEmail||'', aswPhone:res.aswPhone||'',
    importedAt:new Date().toLocaleString(), sourceFile:file.name,
    formType:res.formType||'DHS-1210',   // tracks which form (older DHS-1210 vs newer MDHHS-6064)
  };
  // Caseworker verify/match/add (the "verify and match; if not there, add it" ask).
  if(opts.match && opts.link){ prof.caseworkerId=opts.match.id; prof.worker=opts.match.name; }
  else if(!opts.match && opts.add && res.aswName){
    var parts=res.aswName.split(/\s+/); var first=parts[0]||res.aswName; var last=parts.slice(1).join(' ')||'';
    var cw={id:cwId(),name:res.aswName,first_name:first,last_name:last,agency:'MDHHS',org:'MDHHS',email:res.aswEmail||'',phone:res.aswPhone||'',supervisor_id:''};
    var cws=getCaseworkers(); cws.push(cw); saveCaseworkersLS(cws);
    if(typeof saveCaseworkerAPI==='function') saveCaseworkerAPI(cw);
    prof.caseworkerId=cw.id; prof.worker=cw.name;
  }
  // #5: apply the field updates the user confirmed (caseworker email/phone, client Medicaid ID).
  if(opts.updates && opts.updates.length){
    var cwDirty=null, allCws=null;
    opts.updates.forEach(function(u){
      if(u.target==='client'){ prof[u.field]=u.to; }
      else if(u.target==='caseworker' && u.id){
        if(!allCws){ allCws=getCaseworkers(); }
        var rec=allCws.find(function(c){return c.id===u.id;});
        if(rec){ rec[u.field]=u.to; cwDirty=rec; }
      }
    });
    if(cwDirty){ saveCaseworkersLS(allCws); if(typeof saveCaseworkerAPI==='function') saveCaseworkerAPI(cwDirty); }
  }
  // #6: activate the client from the DHS-1210 (the authorization we just set satisfies the
  // Active-requires-authorization rule). Needs a start date; default to the 1st of the effective month.
  if(opts.setActive){
    var sd=opts.startDate||_firstOfEffectiveMonth(res.effectiveDate);
    if(sd){ prof.clientStatus='active'; prof.startDate=sd; }
  }
  saveProfilesLS(p); saveProfileSP(name,prof);
  if(typeof _syncReassessTask==='function') _syncReassessTask(name, prof.authorization.reassessDate);
  if(typeof renderSidebarClients==='function') renderSidebarClients(); // reflect a newly-active client in the sidebar
  // File the PDF into Documents (Authorization category) — best effort; needs a saved client.
  // Use the TARGET client's own db id (getHcClientId() reads the active client, which may have moved on).
  var cid=prof._dbId||null;
  if(cid){
    var fd=new FormData();
    fd.append('clientType','homecare'); fd.append('clientId',cid); fd.append('category','Authorization');
    fd.append('file',new File([file],'Authorization__'+file.name,{type:file.type||'application/pdf'}));
    // Authorization DATA is already persisted above; the PDF file is a separate best-effort upload.
    // The toast must reflect what actually happened — never claim "PDF filed" before the POST resolves.
    showToast('✓ Authorization saved for '+name+' — filing PDF…');
    fetch(API_BASE+'/documents',{method:'POST',headers:authUploadHeaders(),body:fd})
      .then(function(r){ if(!r.ok)throw new Error('HTTP '+r.status); if(name===activeProfileName&&document.getElementById('hcDocList'))loadHcDocs(cid); showToast('✓ Authorization PDF filed for '+name); })
      .catch(function(e){ console.error('DHS PDF upload failed',e);
        showAlert('The authorization was saved, but the PDF could not be filed ('+((e&&e.message)||'connection error')+'). Your data is safe — re-upload the PDF from the Documents tab'+(/\b401\b/.test(String((e&&e.message)||''))?' after signing in again.':'.')); });
  } else {
    showToast('✓ Authorization saved for '+name+' (save the client to file the PDF)');
  }
  // Repaint only if the imported client is still the one on screen (the panes render the ACTIVE
  // client — repainting them for a different client would show its data under the wrong header).
  if(name===activeProfileName){
    if(typeof renderOverviewPane==='function')renderOverviewPane();
    // If the Authorization tab is the one on screen, refresh it to show the imported data.
    var ap=document.getElementById('dpane-auth');
    if(ap && ap.classList.contains('active') && typeof renderAuthPane==='function')renderAuthPane(false);
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// Caregiver task sheet — export the authorized tasks (NO dollar amounts) as a
// printable / emailable one-pager so caregivers know exactly what's approved.
// Everything runs locally in the browser; opens a new tab with a print-ready
// layout, plus buttons for Print / Copy / Email so Row can send it however
// works best (AirDrop, iMessage, email attachment, printed handout).
// ═══════════════════════════════════════════════════════════════════════════
// PHI-minimized client label for caregiver-facing outputs: first name + last initial (e.g.
// "Darnelle D.") instead of the full name. Reduces the identifier that leaves the system when the
// sheet is texted/emailed. Prefers the structured first/last fields; falls back to parsing the name.
function _caregiverClientLabel(prof, fullName){
  var first=((prof&&prof.firstName)||'').trim();
  var last=((prof&&prof.lastName)||'').trim();
  if(!first && !last){
    var parts=String(fullName||'').trim().split(/\s+/).filter(Boolean);
    first=parts[0]||''; last=parts.length>1?parts[parts.length-1]:'';
  }
  return (first+(last?(' '+last.charAt(0).toUpperCase()+'.'):'')).trim() || String(fullName||'').trim();
}
function exportCaregiverTaskSheet(){
  if(!activeProfileName){showAlert('Open a client first.');return;}
  var prof=getProfiles()[activeProfileName]||{};
  var a=prof.authorization;
  if(!a || !a.tasks || !a.tasks.length){showAlert('No authorized tasks to export. Import a DHS-1210 / MDHHS-6064 first.');return;}

  var clientName=_caregiverClientLabel(prof, activeProfileName);  // first name + last initial (PHI-minimized)
  var effective=a.effectiveDate||'';
  var reassess=a.reassessDate||'';
  var totalHours=(a.hours!=null)? a.hours+'h '+(a.minutes||0)+'m' : '';

  // Build rows: task · time per day · number of days · time per month
  // Intentionally omits Amount and any $/hr — the caregiver never needs to
  // see the rate, and Row asked for this specifically.
  var rows=(a.tasks||[]).map(function(t){
    return '<tr>'+
      '<td>'+_escHtml(t.task||'')+'</td>'+
      '<td class="c">'+_escHtml(t.perDay||'—')+'</td>'+
      '<td class="c">'+_escHtml(t.freq||'')+'</td>'+
      '<td class="c">'+_escHtml(t.perMonth||'')+'</td>'+
    '</tr>';
  }).join('');

  var today=new Date();
  var todayStr=(today.getMonth()+1)+'/'+today.getDate()+'/'+today.getFullYear();

  // Purpose-built PLAIN-TEXT body for the "Email" button. The old button emailed
  // document.body.innerText, which flattened the task table into run-on lines and swept in the
  // on-screen action buttons ("Print / Save PDF", "Copy text", "Email"). This keeps it readable —
  // each task on its own line — with no UI text.
  var _emailLines=[
    'Liberty Home Care Assistance',
    'Authorized Tasks', '',
    'Client: '+clientName,
  ];
  if(effective)  _emailLines.push('Effective: '+effective);
  if(reassess)   _emailLines.push('Reassessment due: '+reassess);
  if(totalHours) _emailLines.push('Approved per month: '+totalHours);
  if(a.aswName)  _emailLines.push('Caseworker (ASW): '+a.aswName+(a.aswPhone?' — '+a.aswPhone:''));
  _emailLines.push('', 'Authorized tasks (perform during each scheduled visit):');
  (a.tasks||[]).forEach(function(t){
    _emailLines.push('• '+(t.task||'')+' — '+(t.perDay||'—')+'/day · '+(t.freq||'')+(t.perMonth?(' · '+t.perMonth+'/month'):''));
  });
  _emailLines.push('', 'Note: Complete each authorized task during each scheduled visit. If a task cannot be performed on a given day, note the reason in your visit log. Do not perform tasks outside this authorization list without checking with the office first.');
  var _plainBody=_emailLines.join('\n');  // used by the "Copy text" button (clean, no page chrome)

  var html='<!doctype html><html><head><meta charset="utf-8">'+
    '<title>Authorized Tasks — '+_escHtml(clientName)+'</title>'+
    '<meta name="viewport" content="width=device-width,initial-scale=1">'+
    '<style>'+
      '*{box-sizing:border-box}'+
      'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Helvetica,Arial,sans-serif;'+
        'color:#1a2b45;margin:0;padding:32px 40px;background:#fff;line-height:1.4}'+
      'h1{margin:0 0 4px;font-size:22px;color:#1a2b45}'+
      '.sub{color:#5c7590;font-size:13px;margin-bottom:20px}'+
      '.brand{display:flex;justify-content:space-between;align-items:baseline;'+
        'border-bottom:2px solid #1a2b45;padding-bottom:10px;margin-bottom:18px}'+
      '.brand .co{font-size:14px;font-weight:700;color:#1a2b45}'+
      '.brand .doc{font-size:12px;color:#5c7590;text-transform:uppercase;letter-spacing:.06em}'+
      '.meta{display:grid;grid-template-columns:1fr 1fr;gap:6px 24px;margin-bottom:18px;font-size:13px}'+
      '.meta .l{color:#5c7590}'+
      '.meta .v{font-weight:600}'+
      'table{width:100%;border-collapse:collapse;font-size:13px;margin-top:6px}'+
      'th{background:#eef4fb;color:#2b4a6b;text-align:left;padding:8px 10px;'+
        'font-size:11px;text-transform:uppercase;letter-spacing:.06em;border-bottom:2px solid #d5e4f3}'+
      'td{padding:9px 10px;border-bottom:1px solid #edf1f6;vertical-align:top}'+
      'td.c{text-align:center;color:#334a68}'+
      'tr:last-child td{border-bottom:1px solid #edf1f6}'+
      '.notes{margin-top:16px;padding:10px 12px;background:#fff9e6;'+
        'border:1px solid #f0e0a0;border-radius:6px;font-size:12px;color:#5b4a1f}'+
      '.actions{position:fixed;bottom:14px;right:14px;display:flex;gap:8px}'+
      '.actions button,.actions a{background:#1a2b45;color:#fff;border:0;padding:9px 14px;'+
        'border-radius:6px;font-size:13px;font-weight:600;cursor:pointer;text-decoration:none;'+
        'display:inline-flex;align-items:center;box-shadow:0 2px 8px rgba(0,0,0,.18)}'+
      '.actions .sec{background:#4d6c88}'+
      '@media print{.actions{display:none}}'+
    '</style></head><body>'+
    '<div class="brand">'+
      '<div class="co">Liberty Home Care Assistance</div>'+
      '<div class="doc">Authorized Tasks</div>'+
    '</div>'+
    '<h1>'+_escHtml(clientName)+'</h1>'+
    '<div class="sub">Authorized services approved by the Michigan Department of Health &amp; Human Services (MDHHS). Perform these tasks during each visit as scheduled below.</div>'+
    '<div class="meta">'+
      (effective ? '<div><span class="l">Effective:</span> <span class="v">'+_escHtml(effective)+'</span></div>':'')+
      (reassess  ? '<div><span class="l">Reassessment due:</span> <span class="v">'+_escHtml(reassess)+'</span></div>':'')+
      (totalHours? '<div><span class="l">Approved per month:</span> <span class="v">'+_escHtml(totalHours)+'</span></div>':'')+
      (a.aswName ? '<div><span class="l">Adult Services Worker:</span> <span class="v">'+_escHtml(a.aswName)+(a.aswPhone?' · '+_escHtml(a.aswPhone):'')+'</span></div>':'')+
      '<div><span class="l">Agency Manager:</span> <span class="v">Thomas Jaboro · 248-291-4106</span></div>'+
      '<div><span class="l">Prepared:</span> <span class="v">'+_escHtml(todayStr)+'</span></div>'+
    '</div>'+
    '<table><thead><tr>'+
      '<th>Authorized Task</th>'+
      '<th style="text-align:center;">Time / Day</th>'+
      '<th style="text-align:center;">Frequency</th>'+
      '<th style="text-align:center;">Time / Month</th>'+
    '</tr></thead><tbody>'+rows+'</tbody></table>'+
    '<div class="notes"><b>Caregiver note:</b> Complete each authorized task during each scheduled visit. '+
      'If a task cannot be performed on a given day, note the reason in your visit log. '+
      'Do not perform tasks outside this authorization list without checking with the office first.</div>'+
    // Clean text used by Copy / Email / Text (hidden; escaped so quotes/brackets can't break markup).
    '<pre id="plainBody" style="display:none;white-space:pre-wrap;">'+_escHtml(_plainBody)+'</pre>'+
    '<div class="actions">'+
      '<button onclick="window.print()">Print / Save PDF</button>'+
      '<button class="sec" onclick="var t=document.getElementById(\'plainBody\').textContent;navigator.clipboard&&navigator.clipboard.writeText(t);this.textContent=\'Copied ✓\';setTimeout(()=>this.textContent=\'Copy text\',1400)">Copy text</button>'+
    '</div>'+
    '</body></html>';

  var w=window.open('','_blank');
  // Track it: the sheet carries the client's name, authorized hours and worker contact, and the
  // 45-minute idle wipe only clears THIS document — a popup left open kept PHI on screen after the
  // session ended. clearPHIFromStorage closes anything still open here.
  try{ if(w) _phiWindows.push(w); }catch(e){}
  if(!w){showAlert('Pop-up was blocked. Allow pop-ups from this site to open the task sheet.');return;}
  w.document.write(html); w.document.close(); w.focus();
}

// tiny escaper used by exportCaregiverTaskSheet so the generated HTML tab
// stays safe even if a task name or client name contains angle brackets/quotes
function _escHtml(s){s=(s==null?'':String(s));return s.replace(/[&<>"']/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c];});}

// Render the authorized-task chart (WITH times) to a PNG the user can paste into Messages/Mail —
// the format Row wanted: a real chart, not a wall of text, and easy to read on a phone. Copies to
// the clipboard when the browser allows (secure context + this click gesture); otherwise downloads
// the PNG to attach manually. Uses html2canvas (already loaded for invoice PDFs).
async function shareCaregiverTaskImage(){
  if(!activeProfileName){showAlert('Open a client first.');return;}
  var prof=getProfiles()[activeProfileName]||{};
  var a=prof.authorization;
  if(!a||!a.tasks||!a.tasks.length){showAlert('No authorized tasks to export. Import a DHS-1210 / MDHHS-6064 first.');return;}
  if(typeof html2canvas!=='function'){showAlert('The image tool is still loading — give it a second and try again.');return;}
  var clientName=_caregiverClientLabel(prof, activeProfileName), esc=_escHtml;  // first name + last initial (PHI-minimized)
  var totalHours=(a.hours!=null)?(a.hours+'h '+(a.minutes||0)+'m'):'';
  var td='padding:7px 9px;border-bottom:1px solid #edf1f6;', tdc=td+'text-align:center;color:#334a68;';
  var rows=(a.tasks||[]).map(function(t){
    return '<tr><td style="'+td+'">'+esc(t.task||'')+'</td>'+
      '<td style="'+tdc+'">'+esc(t.perDay||'—')+'</td>'+
      '<td style="'+tdc+'">'+esc(t.freq||'')+'</td>'+
      '<td style="'+tdc+'">'+esc(t.perMonth||'')+'</td></tr>';
  }).join('');
  // Caregivers don't need the reassessment date, so it's left off. Contacts are spelled out (no
  // "ASW" abbreviation a caregiver wouldn't recognize) and include the agency manager to call.
  var metaLines=[];
  if(a.effectiveDate) metaLines.push('Effective '+esc(a.effectiveDate));
  if(a.aswName) metaLines.push('Adult Services Worker: '+esc(a.aswName)+(a.aswPhone?' · '+esc(a.aswPhone):''));
  metaLines.push('Agency Manager: Thomas Jaboro · 248-291-4106');
  var metaBits=metaLines.join('<br>');
  var approvedBox=totalHours
    ? '<div style="background:#eef4fb;border:1px solid #d5e4f3;border-radius:7px;padding:4px 11px;text-align:center;white-space:nowrap;">'+
        '<div style="font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:#5c7590;">Approved / month</div>'+
        '<div style="font-size:15px;font-weight:700;color:#1a3a5c;line-height:1.1;">'+esc(totalHours)+'</div>'+
      '</div>'
    : '';
  var host=document.createElement('div');
  host.style.cssText='position:fixed;left:-99999px;top:0;width:680px;background:#fff;color:#1a2b45;'+
    'font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Helvetica,Arial,sans-serif;padding:26px 30px;line-height:1.4;';
  host.innerHTML=
    '<div style="display:flex;justify-content:space-between;align-items:baseline;border-bottom:2px solid #1a2b45;padding-bottom:8px;margin-bottom:12px;">'+
      '<div style="font-size:15px;font-weight:700;">Liberty Home Care Assistance</div>'+
      '<div style="font-size:11px;color:#5c7590;text-transform:uppercase;letter-spacing:.06em;">Authorized Tasks</div></div>'+
    '<div style="display:flex;justify-content:space-between;align-items:center;gap:12px;">'+
      '<div style="font-size:20px;font-weight:700;">'+esc(clientName)+'</div>'+
      approvedBox+
    '</div>'+
    (metaBits?'<div style="font-size:12px;color:#5c7590;margin:3px 0 14px;line-height:1.55;">'+metaBits+'</div>':'<div style="height:12px;"></div>')+
    '<table style="width:100%;border-collapse:collapse;font-size:13px;">'+
      '<thead><tr style="background:#eef4fb;color:#2b4a6b;">'+
        '<th style="text-align:left;padding:7px 9px;border-bottom:2px solid #d5e4f3;">Authorized Task</th>'+
        '<th style="padding:7px 9px;border-bottom:2px solid #d5e4f3;">Time/Day</th>'+
        '<th style="padding:7px 9px;border-bottom:2px solid #d5e4f3;">Frequency</th>'+
        '<th style="padding:7px 9px;border-bottom:2px solid #d5e4f3;">Time/Month</th>'+
      '</tr></thead><tbody>'+rows+'</tbody></table>'+
    '<div style="margin-top:12px;font-size:11px;color:#5c7590;">Perform each authorized task during each scheduled visit. Do not perform tasks outside this list without checking with the office first.</div>';
  document.body.appendChild(host);
  try{
    var canvas=await html2canvas(host,{scale:2,backgroundColor:'#ffffff',useCORS:true,logging:false});
    var dataUrl=canvas.toDataURL('image/png');
    canvas.toBlob(function(blob){ _showTaskImagePreview(dataUrl, blob, clientName); }, 'image/png');
  }catch(e){ console.error('Task image failed',e); showAlert('Could not generate the image: '+((e&&e.message)||'error')); }
  finally{ host.remove(); }
}
// Show the generated chart image so the user can SEE it, then copy / download / drag it into a
// message. (The first version copied to the clipboard invisibly — nothing appeared on screen, so
// it looked like nothing happened.)
function _showTaskImagePreview(dataUrl, blob, clientName){
  var ex=document.getElementById('taskImgModal'); if(ex)ex.remove();
  var ov=document.createElement('div');
  ov.id='taskImgModal'; ov.className='modal-overlay open';
  ov.innerHTML='<div class="modal-box" style="max-width:640px;max-height:92vh;overflow:auto;">'+
    '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;">'+
      '<h3 style="margin:0;">Task chart image</h3>'+
      '<button class="btn btn-secondary btn-sm" onclick="var m=document.getElementById(\'taskImgModal\');if(m)m.remove();">Close</button>'+
    '</div>'+
    '<div style="font-size:12px;color:#5c7590;margin-bottom:10px;">Drag this image straight into a text or email — or use the buttons below.</div>'+
    '<img src="'+dataUrl+'" alt="Authorized tasks for '+esc(clientName)+'" style="width:100%;border:1px solid #e2e8f0;border-radius:8px;display:block;">'+
    '<div style="display:flex;gap:8px;margin-top:12px;flex-wrap:wrap;">'+
      '<button class="btn btn-primary btn-sm" id="taskImgCopyBtn">Copy image</button>'+
      '<button class="btn btn-secondary btn-sm" id="taskImgDlBtn">Download</button>'+
    '</div></div>';
  document.body.appendChild(ov);
  var copyBtn=document.getElementById('taskImgCopyBtn');
  copyBtn.onclick=async function(){
    try{
      if(!blob || !navigator.clipboard || !window.ClipboardItem) throw new Error('clipboard unavailable');
      await navigator.clipboard.write([new ClipboardItem({'image/png':blob})]);
      copyBtn.textContent='Copied ✓'; setTimeout(function(){copyBtn.textContent='Copy image';},1500);
    }catch(e){ showToast("Copy isn't available here — use Download, or drag the image into your message."); }
  };
  document.getElementById('taskImgDlBtn').onclick=function(){
    var url=(blob?URL.createObjectURL(blob):dataUrl);
    var link=document.createElement('a'); link.href=url; link.download='Authorized Tasks - '+clientName+'.png';
    document.body.appendChild(link); link.click(); link.remove();
    if(blob) setTimeout(function(){URL.revokeObjectURL(url);},5000);
  };
}

function renderDocsPane(){
  var c=document.getElementById('docsContent');c.innerHTML='';
  if(!activeProfileName)return;
  var clientId=getHcClientId();
  var _dprof=getProfiles()[activeProfileName]||{};
  c.innerHTML=
    (isCarrierClient(_dprof)?'':   // no DHS-1210 import prompt for managed-care clients
    '<div style="display:flex;justify-content:space-between;align-items:center;gap:12px;background:#eef4fb;border:1px solid #d5e4f3;border-radius:8px;padding:10px 14px;margin-bottom:12px;flex-wrap:wrap;">'+
      '<div style="font-size:12px;color:#2b4a6b;"><b>📄 DHS-1210 or MDHHS-6064?</b> Import it to auto-fill authorized hours &amp; tasks — and file the PDF here.</div>'+
      '<button class="btn btn-primary btn-sm" onclick="importDHS1210()" style="white-space:nowrap;">Import DHS-1210 / MDHHS-6064</button>'+
    '</div>')+
    docUploaderHtml({title:'Client Documents',subtitle:"SSN cards, driver's licenses, insurance cards, authorizations, etc.",hasCategory:true,catId:'hcDocCategory',fileId:'hcDocFileInput',scanId:'docScanInput',scanFn:'handleDocScan(this)',uploadFn:'uploadHcDoc()',statusId:'hcDocStatus',listId:'hcDocList'});
  if(clientId){loadHcDocs(clientId);}
  else{document.getElementById('hcDocList').innerHTML='<div style="color:#5c7590;font-size:12px;">Save this client to the database first before uploading documents.</div>';}
}
// Render a document-load FAILURE — never a false "No documents yet". The files are safe on the
// server; a failed GET (expired sign-in / connection) must show a retry, not an empty pane, so a
// worker doesn't think the docs vanished and re-upload duplicates. retryExpr re-runs the loader.
function _renderDocLoadError(listId, retryExpr, err){
  var el=document.getElementById(listId); if(!el)return;
  el.className='';
  var is401=/\b401\b/.test(String((err&&err.message)||''));
  el.innerHTML='<div style="color:#b03030;font-size:12px;padding:8px 2px;line-height:1.5;">'+
    (is401?'Your sign-in expired — <b>sign in again</b>, then ':'Couldn’t load documents (connection issue) — ')+
    'your files are safe on the server. <a href="#" style="color:#185FA5;font-weight:600;" onclick="'+retryExpr+';return false;">Retry</a></div>';
}
function loadHcDocs(clientId){
  fetch(API_BASE+'/documents?clientType=homecare&clientId='+clientId,{headers:apiHeaders()})
  .then(function(r){ if(!r.ok)throw new Error('HTTP '+r.status); return r.json(); })
  .then(function(docs){ renderHcDocList(clientId, Array.isArray(docs)?docs:[]); })
  .catch(function(err){ _renderDocLoadError('hcDocList', "loadHcDocs('"+clientId+"')", err); });
}
// ── Shared modern document grid (client / caregiver / caseworker) ──
var DOC_CATS=[['Other','Other'],['SSN_Card','SSN Card'],['Drivers_License',"Driver's License"],['Insurance_Card','Insurance Card'],['Medicare_Card','Medicare Card'],['Medicaid_Card','Medicaid Card'],['Authorization','Authorization'],['Certification','Certification'],['Background_Check','Background Check'],['I9_W4','I-9 / W-4']];
function _docCatLabel(c){for(var i=0;i<DOC_CATS.length;i++){if(DOC_CATS[i][0]===c)return DOC_CATS[i][1];}return c||'Other';}
// Category values are free-text (chips insert a LABEL, and users can type their own). Resolve any
// value — a DOC_CATS key, its label, or custom text — to the canonical KEY so key-based gating
// (_idCat, _CARD_TYPE) matches no matter how the document was tagged. Unknown values pass through.
function _catKey(c){for(var i=0;i<DOC_CATS.length;i++){if(DOC_CATS[i][0]===c||DOC_CATS[i][1]===c)return DOC_CATS[i][0];}return c;}
function _fmtDocSize(n){if(!n)return '';return n<1048576?Math.round(n/1024)+' KB':(n/1048576).toFixed(1)+' MB';}
var _docEditCtx=null;
// Emailing a client document to a CAREGIVER: two gates, both enforced again inside the handler.
//  • _DOC_EMAIL_BLOCKED — never emailable: ID cards plus the SSN-bearing HR forms (I-9/W-4 has the
//    SSN on its face; a background check carries SSN + DOB). One click used to attach either.
//  • _DOC_EMAIL_SAFE   — known caregiver-safe categories that send with no extra prompt. Anything
//    else (Other, or a category the user typed) still sends, but only after an explicit warning —
//    "Other" is the DEFAULT upload bucket, so blocking it outright would break the 4676 workflow.
var _DOC_EMAIL_BLOCKED={SSN_Card:1,Drivers_License:1,Insurance_Card:1,Medicare_Card:1,Medicaid_Card:1,I9_W4:1,Background_Check:1};
var _DOC_EMAIL_SAFE={Authorization:1,Certification:1};
function renderDocGrid(list,docs,opts){
  if(!list)return; opts=opts||{};
  if(!docs||!docs.length){list.className='';list.innerHTML='<div class="doc-empty">No documents yet.</div>';return;}
  list.className='doc-grid';
  _docEditCtx={docs:docs,clientType:opts.clientType,clientId:opts.clientId,refresh:opts.refresh};
  list.innerHTML=docs.map(function(d,i){
    var display=d.displayName||d.name||'Document';
    var ext=(display.split('.').pop()||'').toLowerCase();
    var isImg=['jpg','jpeg','png','gif','webp','heic','bmp'].indexOf(ext)>=0;
    var isPdf=ext==='pdf';
    // Email is offered on client docs except the categories that carry an SSN or an ID card (see
    // _DOC_EMAIL_BLOCKED); the handlers re-check, so this only hides the buttons. Two destinations,
    // because a document has two legs: to the caregiver (e.g. a 4676 to sign) and on to the
    // caseworker (e.g. that same 4676 once it's signed).
    var _catk=_catKey(d.category);   // normalize label/free-text → canonical key before gating
    var canEmail=(opts.clientType==='homecare')&&!_DOC_EMAIL_BLOCKED[_catk];
    // Extract (OCR → fill client fields) is offered only on the ID/benefit cards we can read.
    var canExtract=(opts.clientType==='homecare')&&_CARD_TYPE[_catk];
    var thumb=isImg
      ?'<span class="doc-thumb" style="background-image:url(\''+escJsAttr(d.url)+'\')"></span>'
      :'<span class="doc-thumb doc-thumb-file">'+_docFileIcon(ext)+'</span>';
    return '<div class="doc-card">'+
      '<a class="doc-card-preview" href="'+esc(d.url||'#')+'" target="_blank" rel="noopener" onclick="return openDocPreview('+i+')" title="Preview">'+thumb+'</a>'+
      '<div class="doc-card-body">'+
        '<a class="doc-card-name" href="'+esc(d.url||'#')+'" target="_blank" rel="noopener" onclick="return openDocPreview('+i+')" title="'+esc(display)+'">'+esc(display)+'</a>'+
        '<div class="doc-card-meta"><span class="doc-cat-pill">'+esc(_docCatLabel(d.category))+'</span>'+(d.size?'<span class="doc-size">'+_fmtDocSize(d.size)+'</span>':'')+'</div>'+
      '</div>'+
      '<div class="doc-card-actions">'+
        (canExtract?'<button class="doc-act" title="Read this card &amp; fill client fields" onclick="extractCardFields('+i+')">🔍</button>':'')+
        (canEmail?'<button class="doc-act" title="Email to caregiver (e.g. a 4676 to sign)" onclick="emailDocToCaregiver('+i+')">✉</button>':'')+
        (canEmail?'<button class="doc-act" title="Email to caseworker (e.g. the signed 4676)" onclick="emailDocToCaseworker('+i+')">📨</button>':'')+
        '<a class="doc-act" href="'+esc(d.downloadUrl||d.url||'#')+'" title="Download" download>⬇</a>'+
        '<button class="doc-act" title="Rename / relabel" onclick="openDocEditModal('+i+')">✎</button>'+
        '<button class="doc-act doc-act-del" title="Delete" onclick="deleteDocAt('+i+')">✕</button>'+
      '</div>'+
    '</div>';
  }).join('');
}
// Delete the document at grid index i via the shared render context. Index-based on purpose: no
// filename is ever interpolated into an inline handler — an apostrophe in the name (e.g. from a
// "Driver's License" category prefix) used to break the onclick with a syntax error, and a crafted
// filename was an XSS vector. Works for all three panes (homecare / caregiver / caseworker).
function deleteDocAt(i){
  var ctx=_docEditCtx; if(!ctx||!ctx.docs||!ctx.docs[i])return;
  var d=ctx.docs[i];
  showConfirm('Delete this document?',function(){
    fetch(API_BASE+'/documents?clientType='+encodeURIComponent(ctx.clientType)+'&clientId='+encodeURIComponent(ctx.clientId),
      {method:'DELETE',headers:apiHeaders(),body:JSON.stringify({name:d.name})})
      .then(function(r){if(!r.ok)throw new Error('HTTP '+r.status);if(typeof ctx.refresh==='function')ctx.refresh();})
      .catch(function(e){showAlert('Delete failed: '+((e&&e.message)||e));});
  },{title:'Delete Document',okText:'Delete'});
}
// The document-email gate, enforced HERE rather than only in the grid: the button is a UI
// affordance, and a document can be re-categorized (or these called directly) after it was drawn.
// Blocked categories never send; anything not known safe (including the default "Other" bucket)
// takes one explicit confirmation naming the file. `who` is the word used in the prompts.
function _docEmailGate(d, who, onOk){
  var cat=_catKey(d.category);
  if(_DOC_EMAIL_BLOCKED[cat]){
    showAlert('“'+_docCatLabel(d.category)+'” documents can’t be emailed to a '+who+' — they contain a Social Security number or ID card.');
    return;
  }
  if(_DOC_EMAIL_SAFE[cat]){ onOk(); return; }
  var display=d.displayName||d.name||'the document';
  showConfirm('“'+display+'” is filed as “'+_docCatLabel(d.category)+'”, which isn’t a known safe category to send.\n\nOpen it first and confirm it contains no Social Security number, ID card, or other sensitive identifiers before emailing it to a '+who+'.',
    onOk, {title:'Send this to the '+who+'?',okText:'I’ve checked — continue',danger:true});
}
// Email a client document (e.g. the MSA-4676) to the client's assigned caregiver, with a default
// draft tuned to the document + a first-time self-introduction. Everything is editable before send.
// The intro is NOT tracked in localStorage (that wouldn't survive across devices/accounts) — it's
// just seeded into the draft with a hint to delete it if you've emailed the caregiver before.
// Step 1 of the MSA-4676 workflow: the caregiver signs it; emailDocToCaseworker sends it on.
function emailDocToCaregiver(index){
  var ctx=_docEditCtx; if(!ctx||ctx.clientType!=='homecare')return;
  var d=ctx.docs[index]; if(!d)return;
  var prof=getProfiles()[activeProfileName]||{};
  var clientName=prof.clientName||activeProfileName||'the client';
  var cgs=getCaregivers(); var cg=(prof.caregiverId&&cgs[prof.caregiverId])?cgs[prof.caregiverId]:null;
  var cgEmail=cg?(cg.email||''):'';
  var cgFirst=cg?((cg.firstName||(cg.name||'').split(' ')[0])||''):'';
  var ag=getAgencyInfo()||{}; var agName=ag.agency_name||ag.agency_provider_name||'our agency'; var agPhone=ag.agency_phone||'';
  var display=d.displayName||d.name||'the document';
  var is4676=/4676/i.test(display)||/4676/i.test(_docCatLabel(d.category));
  var greet='Hello'+(cgFirst?(' '+cgFirst):'')+',';
  var subject, body;
  if(is4676){
    subject='MSA-4676 Home Help Services Agreement — '+clientName;
    body=greet+'\n\n'+
      'My name is [your name] with '+agName+(agPhone?(' ('+agPhone+')'):'')+'. I am reaching out about '+clientName+'’s Home Help services.\n\n'+
      'Attached is the MSA-4676 Home Help Services Agreement for '+clientName+'. Please review and sign it so we can transition '+clientName+'’s Home Help services to '+agName+'.\n\n'+
      'Once signed, please reply to this email with the completed form, or let me know if you have any questions.\n\n'+
      'Thank you,\n'+agName;
  } else {
    subject=display+' — '+clientName;
    body=greet+'\n\n'+
      'Attached is '+display+' for '+clientName+' from '+agName+'. Please review and let me know if you have any questions.\n\n'+
      'Thank you,\n'+agName;
  }
  _docEmailGate(d, 'caregiver', function(){ _openDocEmailModal(d, cgEmail, subject, body); });
}
// Step 2 of the MSA-4676 workflow: the caregiver has signed it and the signed PDF is filed on the
// client — send it on to the caseworker. Also used for any other client document the caseworker
// needs (an authorization, a certification). Same gate as the caregiver path.
function emailDocToCaseworker(index){
  var ctx=_docEditCtx; if(!ctx||ctx.clientType!=='homecare')return;
  var d=ctx.docs[index]; if(!d)return;
  var prof=getProfiles()[activeProfileName]||{};
  var clientName=prof.clientName||activeProfileName||'the client';
  var cws=getCaseworkers();
  var cw=(prof.caseworkerId?cws.find(function(c){return c.id===prof.caseworkerId;}):null) ||
         (prof.worker?cws.find(function(c){return (c.name||'').trim().toLowerCase()===String(prof.worker).trim().toLowerCase();}):null);
  var cwEmail=cw?(cw.email||''):'';
  if(!cwEmail){
    showAlert('No email on file for '+clientName+'’s caseworker'+(cw&&cw.name?(' ('+cw.name+')'):'')+'.\n\nAdd it on the Caseworkers page (or assign a caseworker to this client), then try again.');
    return;
  }
  var cwFirst=cw?((cw.first_name||(cw.name||'').split(' ')[0])||''):'';
  var ag=getAgencyInfo()||{}; var agName=ag.agency_name||ag.agency_provider_name||'our agency'; var agPhone=ag.agency_phone||'';
  var display=d.displayName||d.name||'the document';
  var is4676=/4676/i.test(display)||/4676/i.test(_docCatLabel(d.category));
  var greet='Hello'+(cwFirst?(' '+cwFirst):'')+',';
  var subject, body;
  if(is4676){
    subject='Signed MSA-4676 — '+clientName;
    body=greet+'\n\n'+
      'Attached is the signed MSA-4676 Home Help Services Agreement for '+clientName+', authorizing '+agName+(agPhone?(' ('+agPhone+')'):'')+' as their Home Help provider.\n\n'+
      'Please let me know if you need anything else to complete the transition.\n\n'+
      'Thank you,\n'+agName;
  } else {
    subject=display+' — '+clientName;
    body=greet+'\n\n'+
      'Attached is '+display+' for '+clientName+' from '+agName+'. Please let me know if you have any questions.\n\n'+
      'Thank you,\n'+agName;
  }
  _docEmailGate(d, 'caseworker', function(){ _openDocEmailModal(d, cwEmail, subject, body); });
}
// d may be null → plain email with no attachment (the Assistant uses this for a note with no form).
// opts.auditClient names the client to log the send under (defaults to the active profile) — the
// Assistant emails for a client that isn't necessarily the one currently open.
function _openDocEmailModal(d, to, subject, body, opts){
  opts=opts||{};
  var _auditClient=opts.auditClient||activeProfileName;
  var _lbl=d?(d.displayName||d.name||'document'):'email';
  var ov=document.createElement('div');ov.className='modal-overlay open';
  ov.innerHTML='<div class="modal-box" style="max-width:520px;">'+
    '<h3>'+(d?'Email document':'Compose email')+'</h3>'+
    '<p style="font-size:12px;color:#5c7590;margin-top:-4px;">'+(d?('Sends “'+esc(_lbl)+'” as an attachment from your Microsoft account. '):'Sends from your Microsoft account. ')+'Edit anything below before sending.</p>'+
    '<label class="qc-l" for="dem-to">To</label><input id="dem-to" class="qc-i" value="'+esc(to)+'" placeholder="caregiver@email.com">'+
    '<label class="qc-l" for="dem-subj">Subject</label><input id="dem-subj" class="qc-i" value="'+esc(subject)+'">'+
    '<label class="qc-l" for="dem-body">Message <span style="font-weight:400;color:#5c7590;">(delete the intro line if you’ve emailed them before)</span></label>'+
    '<textarea id="dem-body" class="qc-i" style="min-height:180px;resize:vertical;font-family:Arial,sans-serif;">'+esc(body)+'</textarea>'+
    '<div id="dem-status" style="font-size:12px;color:#c0392b;min-height:16px;margin-top:4px;"></div>'+
    '<div class="modal-row"><button class="btn btn-primary" id="dem-send">Send Email</button><button class="btn btn-secondary" id="dem-cancel">Cancel</button></div>'+
  '</div>';
  document.body.appendChild(ov);
  var close=function(){if(ov.parentNode)ov.parentNode.removeChild(ov);};
  ov.querySelector('#dem-cancel').addEventListener('click',close);
  ov.addEventListener('mousedown',function(e){if(e.target===ov)close();});
  ov.querySelector('#dem-send').addEventListener('click',function(){
    var toV=(ov.querySelector('#dem-to').value||'').trim();
    var subjV=(ov.querySelector('#dem-subj').value||'').trim();
    var bodyText=ov.querySelector('#dem-body').value||'';
    var status=ov.querySelector('#dem-status');
    if(!toV||!/@/.test(toV)){status.style.color='#c0392b';status.textContent='Enter a valid recipient email.';return;}
    if(typeof spToken==='undefined'||!spToken){status.style.color='#c0392b';status.textContent='Sign in first (Settings) to send email.';return;}
    var btn=ov.querySelector('#dem-send');btn.disabled=true;btn.textContent='Sending…';status.style.color='#5c7590';status.textContent=(d?'Attaching document…':'Sending…');
    var bodyHtml=esc(bodyText).replace(/\n/g,'<br>');
    var attachP=d?_docToBase64(d.url||d.downloadUrl).then(function(b64){return [{name:(d.displayName||d.name||'document.pdf'),base64:b64}];}):Promise.resolve([]);
    attachP.then(function(atts){
      status.textContent='Sending…';
      return sendMailWithPDF(toV,subjV,bodyHtml,atts);
    }).then(function(){
      if(typeof showToast==='function')showToast('✓ Emailed '+_lbl+' to '+toV,3500);
      if(typeof addAuditEntry==='function')addAuditEntry(_auditClient,'Emailed “'+_lbl+'” to '+toV);
      // _aiScrub drops keys matching name/email/recipient/file/… — `to` and `doc` did NOT match, so
      // the recipient's real address and the document's (client-named) filename were being sent to
      // App Insights. Renamed to keys the scrubber recognizes: the event still records that a
      // document was emailed, without recording WHO or WHICH.
      if(typeof aiTrack==='function')aiTrack('DocEmailedToCaregiver',{recipient:toV,docName:_lbl});
      close();
    }).catch(function(err){
      btn.disabled=false;btn.textContent='Send Email';status.style.color='#c0392b';
      status.textContent='Could not send: '+((err&&err.message)||'error')+'. Try again.';
    });
  });
}
// For OCR: fetch an image, downscale to <= maxDim on the long edge, and return JPEG base64. Azure's
// free OCR tier caps a document at 4 MB and rejects HEIC — a full-res phone photo blows past both, so
// re-encoding to a right-sized JPEG in the browser makes any photo work (and reads more accurately).
// PDFs, or images the browser can't decode, fall back to the raw bytes.
function _imageToOcrBase64(url, maxDim){
  maxDim = maxDim || 2200;
  return fetch(url).then(function(r){ if(!r.ok)throw new Error('HTTP '+r.status); return r.blob(); })
    .then(function(blob){
      if((blob.type&&blob.type.indexOf('pdf')>=0) || /\.pdf(\?|$)/i.test(url)) return _fileToBase64(blob);
      return new Promise(function(res,rej){
        var objUrl=URL.createObjectURL(blob), img=new Image();
        img.onload=function(){
          try{
            var sc=Math.min(1, maxDim/Math.max(img.naturalWidth||1, img.naturalHeight||1));
            var cw=Math.max(1,Math.round((img.naturalWidth||1)*sc)), ch=Math.max(1,Math.round((img.naturalHeight||1)*sc));
            var cv=document.createElement('canvas'); cv.width=cw; cv.height=ch;
            cv.getContext('2d').drawImage(img,0,0,cw,ch);
            var dataUrl=cv.toDataURL('image/jpeg',0.85);
            URL.revokeObjectURL(objUrl);
            res(dataUrl.slice(dataUrl.indexOf(',')+1));
          }catch(e){ URL.revokeObjectURL(objUrl); _fileToBase64(blob).then(res); }
        };
        img.onerror=function(){   // e.g. HEIC on a browser that can't decode it
          URL.revokeObjectURL(objUrl);
          // Azure's OCR rejects HEIC outright, so sending the raw bytes just returns an opaque 502.
          // Fail here with an instruction the user can act on.
          if(/heic|heif/i.test(blob.type||'') || /\.hei[cf](\?|$)/i.test(url)){
            rej(new Error('This photo is in Apple’s HEIC format, which this browser can’t convert and the reader can’t accept. Re-save or re-take it as JPEG/PNG (iPhone: Settings → Camera → Formats → Most Compatible) and upload again.'));
            return;
          }
          _fileToBase64(blob).then(res);
        };
        img.src=objUrl;
      });
    });
}
// Fetch a stored document (blob-storage URL) and return its base64 (no data: prefix) for attaching.
function _docToBase64(url){
  if(!url)return Promise.reject(new Error('no document URL'));
  return fetch(url).then(function(r){ if(!r.ok)throw new Error('HTTP '+r.status); return r.blob(); })
    .then(function(blob){ return new Promise(function(res,rej){
      var fr=new FileReader();
      fr.onload=function(){ var s=String(fr.result||''); var i=s.indexOf(','); res(i>=0?s.slice(i+1):s); };
      fr.onerror=function(){rej(new Error('read failed'));};
      fr.readAsDataURL(blob);
    }); });
}
// ── ID / benefit card autofill (Driver's License, SSN, Medicaid) ─────────────
// Read a labeled card via backend OCR and offer to apply the fields — comparing against what's
// already on the client: blanks are offered to fill, matches are shown verified, and differences
// are flagged for you to choose. Never auto-applies; you click Apply. Name fields are intentionally
// NOT applied here (the client's name is its record key — renaming is its own Save-Changes flow).
var _CARD_TYPE={Drivers_License:'drivers_license',SSN_Card:'ssn',Medicaid_Card:'medicaid',Medicare_Card:'medicare'};
var _CARD_FIELD_LABEL={dob:'Date of Birth',driversLicense:"Driver's License #",street:'Street',city:'City',state:'State',zip:'ZIP',ssn:'Social Security #',medicaidId:'Medicaid ID',medicare:'Medicare #'};
var _CARD_FIELD_INPUT={dob:'ei-dob',driversLicense:'ei-dl',street:'ei-street',city:'ei-city',state:'ei-state',zip:'ei-zip',ssn:'ei-ssn',medicaidId:'ei-medicaid',medicare:'ei-medicare'};
function extractCardFields(index){
  var ctx=_docEditCtx; if(!ctx||ctx.clientType!=='homecare')return;
  var d=ctx.docs[index]; if(!d)return;
  var targetName=activeProfileName;   // capture now — OCR is async and the active client can change
  var type=_CARD_TYPE[_catKey(d.category)];
  if(!type){showAlert('This document type can’t be auto-read.');return;}
  if(typeof spToken==='undefined'||!spToken){showAlert('Sign in first (Settings) to read a card.');return;}
  if(typeof showToast==='function')showToast('Reading card…',1500);
  _imageToOcrBase64(d.url||d.downloadUrl).then(function(b64){
    return fetch(API_BASE+'/id-extract',{method:'POST',headers:apiHeaders(),body:JSON.stringify({fileBase64:b64,cardType:type})});
  }).then(function(r){
    if(r.ok)return r.json();
    return r.json().then(function(j){throw new Error((j&&j.error)||('HTTP '+r.status));},function(){throw new Error('HTTP '+r.status);});
  }).then(function(j){
    var fields=(j&&j.fields)||{};
    var conf=(j&&j.confidence)||{};
    if(!Object.keys(fields).length){showAlert('Couldn’t read any fields off that card. Try a clearer photo/scan.');return;}
    _showCardReview(fields,conf,targetName);
  }).catch(function(e){showAlert('Could not read the card: '+((e&&e.message)||e));});
}
// Normalize for comparison: numbers by their digits/letters only (so 123-45-6789 == 123456789),
// everything else case-insensitively.
function _cardNorm(field,v){
  v=(v==null?'':String(v)).trim();
  if(field==='ssn'||field==='medicaidId'||field==='medicare'||field==='driversLicense') return v.replace(/[^0-9A-Za-z]/g,'').toLowerCase();
  return v.toLowerCase();
}
function _showCardReview(fields,conf,targetName){
  conf=conf||{};
  targetName=targetName||activeProfileName;
  var prof=getProfiles()[targetName]||{};
  var rows=[];
  Object.keys(fields).forEach(function(k){
    if(!_CARD_FIELD_LABEL[k])return;                      // only known, applyable fields (skips name)
    var ext=String(fields[k]||'').trim(); if(!ext)return;
    var cur=String(prof[k]||'').trim();
    var same=cur && _cardNorm(k,cur)===_cardNorm(k,ext);
    // low-confidence = the number was only a format-fallback (not found next to its label). It is
    // flagged and NOT pre-checked, so a possibly-wrong ID can't be silently applied.
    rows.push({field:k,label:_CARD_FIELD_LABEL[k],cur:cur,ext:ext,lowconf:conf[k]==='guess',state:(!cur?'fill':(same?'match':'differ'))});
  });
  if(!rows.length){showAlert('Nothing new to apply from that card.');return;}
  var warnTag='<span style="color:#b8860b;font-weight:600;font-size:11px;"> ⚠ verify — not clearly labeled on the card</span>';
  var bodyHtml=rows.map(function(r,i){
    if(r.state==='match') return '<div style="padding:6px 0;font-size:13px;"><span style="color:#1a7740;font-weight:700;">✓</span> <b>'+esc(r.label)+'</b> matches the card ('+esc(r.ext)+')</div>';
    if(r.state==='fill') return '<label style="display:flex;gap:8px;align-items:flex-start;padding:6px 0;font-size:13px;"><input type="checkbox" data-idx="'+i+'" class="cardfill"'+(r.lowconf?'':' checked')+' style="margin-top:3px;"><span>Fill <b>'+esc(r.label)+'</b> → <span style="color:#185FA5;font-weight:600;">'+esc(r.ext)+'</span>'+(r.lowconf?warnTag:'')+'</span></label>';
    return '<div style="padding:6px 0;font-size:13px;border-top:1px dashed #e1e5ea;">'+
      '<div style="color:#b8860b;font-weight:700;">⚠ '+esc(r.label)+' differs from the card</div>'+
      '<label style="display:block;margin-top:3px;"><input type="radio" name="card'+i+'" value="keep" checked> Keep on file: <b>'+esc(r.cur)+'</b></label>'+
      '<label style="display:block;"><input type="radio" name="card'+i+'" value="use"> Use card: <span style="color:#185FA5;font-weight:600;">'+esc(r.ext)+'</span>'+(r.lowconf?warnTag:'')+'</label>'+
    '</div>';
  }).join('');
  var ov=document.createElement('div');ov.className='modal-overlay open';ov.id='cardReviewModal';
  ov.innerHTML='<div class="modal-box" style="max-width:480px;">'+
    '<h3>Apply card details to '+esc(targetName||'client')+'?</h3>'+
    '<p style="font-size:12px;color:#5c7590;margin-top:-4px;">Blanks are filled, matches are confirmed, and differences are yours to choose. Nothing changes until you click Apply.</p>'+
    '<div style="max-height:340px;overflow:auto;">'+bodyHtml+'</div>'+
    '<div class="modal-row"><button class="btn btn-primary" id="card-apply">Apply</button><button class="btn btn-secondary" id="card-cancel">Cancel</button></div>'+
  '</div>';
  document.body.appendChild(ov);
  var close=function(){if(ov.parentNode)ov.parentNode.removeChild(ov);};
  ov.querySelector('#card-cancel').addEventListener('click',close);
  ov.addEventListener('mousedown',function(e){if(e.target===ov)close();});
  ov.querySelector('#card-apply').addEventListener('click',function(){
    var apply={};
    rows.forEach(function(r,i){
      if(r.state==='fill'){ var cb=ov.querySelector('.cardfill[data-idx="'+i+'"]'); if(cb&&cb.checked)apply[r.field]=r.ext; }
      else if(r.state==='differ'){ var rd=ov.querySelector('input[name="card'+i+'"][value="use"]'); if(rd&&rd.checked)apply[r.field]=r.ext; }
    });
    close();
    if(Object.keys(apply).length)_applyCardFields(apply,targetName);
  });
}
function _applyCardFields(apply,targetName){
  // Write to the client the review was opened for (captured at extract time), not to whatever is
  // active now — a background sync can change activeProfileName while the OCR/review is on screen.
  var name=targetName||activeProfileName;
  var p=getProfiles(); var prof=name?p[name]:null;
  if(!prof){showAlert('“'+(name||'That client')+'” is no longer available — nothing was applied.');return;}
  var sameClient=(name===activeProfileName);
  Object.keys(apply).forEach(function(k){
    prof[k]=apply[k];
    var inputId=_CARD_FIELD_INPUT[k]; var el=sameClient&&inputId&&document.getElementById(inputId);
    if(el)el.value=apply[k];                              // reflect in the Profile form if it's open
  });
  saveProfilesLS(p);
  if(typeof saveProfileSP==='function')saveProfileSP(name,prof);
  if(typeof addAuditEntry==='function')addAuditEntry(name,'Applied card details: '+Object.keys(apply).map(function(k){return _CARD_FIELD_LABEL[k]||k;}).join(', '));
  if(typeof showToast==='function')showToast('✓ Applied '+Object.keys(apply).length+' field(s) from the card',3000);
}
function openDocEditModal(index){
  var ctx=_docEditCtx;if(!ctx)return; var d=ctx.docs[index];if(!d)return;
  var ov=document.createElement('div');ov.className='modal-overlay open';
  ov.innerHTML='<div class="modal-box" style="max-width:380px;">'+
    '<h3>Edit Document</h3>'+
    '<p>Rename this document or change its category.</p>'+
    '<label class="qc-l" for="de-name">Name</label><input id="de-name" class="qc-i" maxlength="120" value="'+esc(d.displayName||d.name||'')+'">'+
    docCategoryFieldHtml('de-cat',_docCatLabel(d.category))+
    '<div class="modal-row"><button class="btn btn-primary" id="de-save">Save</button><button class="btn btn-secondary" id="de-cancel">Cancel</button></div>'+
  '</div>';
  document.body.appendChild(ov);
  var close=function(){if(ov.parentNode)ov.parentNode.removeChild(ov);};
  ov.querySelector('#de-cancel').addEventListener('click',close);
  ov.addEventListener('mousedown',function(e){if(e.target===ov)close();});
  ov.querySelector('#de-save').addEventListener('click',function(){
    var name=(ov.querySelector('#de-name').value||'').trim();var cat=ov.querySelector('#de-cat').value;
    if(!name){ov.querySelector('#de-name').focus();return;}
    var btn=ov.querySelector('#de-save');btn.textContent='Saving…';btn.disabled=true;
    fetch(API_BASE+'/documents/meta',{method:'PATCH',headers:apiHeaders(),body:JSON.stringify({clientType:ctx.clientType,clientId:ctx.clientId,name:d.name,displayName:name,category:cat})})
      .then(function(r){if(!r.ok)throw new Error('Save failed ('+r.status+')');close();if(typeof ctx.refresh==='function')ctx.refresh();})
      .catch(function(e){btn.textContent='Save';btn.disabled=false;showAlert(String(e));});
  });
  setTimeout(function(){var n=ov.querySelector('#de-name');if(n){n.focus();}},30);
}
// In-app document preview (lightbox) — images inline, PDFs in a frame, else open in a tab.
function openDocPreview(index){
  var ctx=_docEditCtx;if(!ctx)return true; var d=ctx.docs[index];if(!d)return true;
  var name=d.displayName||d.name||'Document';
  var ext=(name.split('.').pop()||'').toLowerCase();
  var isImg=['jpg','jpeg','png','gif','webp','bmp'].indexOf(ext)>=0;
  var isPdf=ext==='pdf';
  if(!isImg&&!isPdf){window.open(d.url,'_blank');return false;}
  var ov=document.createElement('div');ov.className='modal-overlay open doc-lightbox';
  var body=isImg?'<img src="'+esc(d.url)+'" alt="'+esc(name)+'" class="doc-lb-img">':'<iframe src="'+esc(d.url)+'" class="doc-lb-frame" title="'+esc(name)+'"></iframe>';
  ov.innerHTML='<div class="doc-lb-box">'+
    '<div class="doc-lb-head"><span class="doc-lb-title" title="'+esc(name)+'">'+esc(name)+'</span>'+
      '<span class="doc-lb-actions">'+
        '<a class="btn btn-secondary btn-sm" href="'+esc(d.downloadUrl||d.url)+'" download>Download</a>'+
        '<a class="btn btn-secondary btn-sm" href="'+esc(d.url)+'" target="_blank" rel="noopener">Open in tab</a>'+
        '<button class="btn btn-secondary btn-sm doc-lb-close">Close</button>'+
      '</span></div>'+
    '<div class="doc-lb-body">'+body+'</div>'+
  '</div>';
  document.body.appendChild(ov);
  var close=function(){if(ov.parentNode)ov.parentNode.removeChild(ov);};
  ov.querySelector('.doc-lb-close').addEventListener('click',close);
  ov.addEventListener('mousedown',function(e){if(e.target===ov)close();});
  return false;
}
// Clean colour-coded file icon for non-image documents.
function _docFileIcon(ext){
  var colors={pdf:'#e2574c',doc:'#2b7cd3',docx:'#2b7cd3',txt:'#5a7290',xls:'#1e7e34',xlsx:'#1e7e34',csv:'#1e7e34'};
  var col=colors[ext]||'#5a7290';
  return '<span class="doc-fileicon" style="color:'+col+'">'+
    '<svg viewBox="0 0 24 24" width="38" height="38" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>'+
    '<span class="doc-fileicon-ext">'+esc((ext||'FILE').toUpperCase().slice(0,4))+'</span>'+
  '</span>';
}
// Shared modern document uploader (drop zone + category combobox + scan) for all 3 panes.
function docCategoryFieldHtml(id,value){
  var chips=DOC_CATS.map(function(c){return '<button type="button" class="doc-chip" onclick="var e=document.getElementById(\''+id+'\');if(e){e.value=this.textContent;e.focus();}">'+esc(c[1])+'</button>';}).join('');
  return '<div class="doc-field"><label class="doc-flabel" for="'+id+'">Category</label>'+
    '<input id="'+id+'" class="doc-input" value="'+esc(value||'Other')+'" placeholder="Type a category, or tap one below" autocomplete="off">'+
    '<div class="doc-chips">'+chips+'</div>'+
  '</div>';
}
function docUploaderHtml(o){
  return '<div class="doc-uploader">'+
    '<div class="doc-uploader-head"><h4>'+esc(o.title)+'</h4><p>'+esc(o.subtitle)+'</p></div>'+
    (o.hasCategory?docCategoryFieldHtml(o.catId,'Other'):'')+
    '<label class="doc-drop" id="'+o.fileId+'-drop" for="'+o.fileId+'" ondragover="event.preventDefault();this.classList.add(\'dragover\')" ondragleave="this.classList.remove(\'dragover\')" ondrop="_docDrop(event,\''+o.fileId+'\')">'+
      '<svg class="doc-drop-icon" viewBox="0 0 24 24" width="26" height="26" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 9 12 4 17 9"/><line x1="12" y1="4" x2="12" y2="15"/></svg>'+
      '<span class="doc-drop-text"><b>Choose a file</b> or drag it here</span>'+
      '<span class="doc-drop-file" id="'+o.fileId+'-name"></span>'+
      '<input type="file" id="'+o.fileId+'" accept=".pdf,.jpg,.jpeg,.png,.doc,.docx,.heic,.webp" multiple style="position:absolute;width:1px;height:1px;opacity:0;overflow:hidden;" onchange="_docShowPick(this)">'+
    '</label>'+
    '<div class="doc-uploader-actions">'+
      '<button class="btn btn-primary" onclick="'+o.uploadFn+'">Upload</button>'+
      '<input type="file" id="'+o.scanId+'" accept="image/*" capture="environment" style="display:none" onchange="'+o.scanFn+'">'+
      '<button class="btn btn-secondary" onclick="document.getElementById(\''+o.scanId+'\').click()">📷 Scan / Photo</button>'+
    '</div>'+
    '<span id="'+o.statusId+'" class="doc-upload-status"></span>'+
  '</div>'+
  '<div id="'+o.listId+'"><div class="doc-empty">Loading…</div></div>';
}
function _docShowPick(input){
  var el=document.getElementById(input.id+'-name');
  if(el)el.textContent=input.files&&input.files.length?(input.files.length>1?input.files.length+' files selected':input.files[0].name):'';
  var drop=document.getElementById(input.id+'-drop');if(drop)drop.classList.toggle('has-file',!!(input.files&&input.files.length));
}
function _docDrop(e,fileId){
  e.preventDefault();
  var inp=document.getElementById(fileId);var drop=document.getElementById(fileId+'-drop');
  if(drop)drop.classList.remove('dragover');
  if(inp&&e.dataTransfer&&e.dataTransfer.files&&e.dataTransfer.files.length){
    try{inp.files=e.dataTransfer.files;}catch(err){}
    _docShowPick(inp);
  }
}
function renderHcDocList(clientId,docs){
  renderDocGrid(document.getElementById('hcDocList'),docs,{clientType:'homecare',clientId:clientId,deleteExpr:function(name){return 'deleteHcDoc('+clientId+',\''+encodeURIComponent(name)+'\')';},refresh:function(){loadHcDocs(clientId);}});
}
function uploadHcDoc(){
  var clientId=getHcClientId();
  if(!clientId){showAlert('Save this client first before uploading documents.');return;}
  var input=document.getElementById('hcDocFileInput');
  if(!input||!input.files||!input.files.length){showAlert('Please select a file first.');return;}
  var cat=_catKey((document.getElementById('hcDocCategory')&&document.getElementById('hcDocCategory').value)||'Other');
  var status=document.getElementById('hcDocStatus');
  status.textContent='Uploading...';
  var fd=new FormData();
  fd.append('clientType','homecare');fd.append('clientId',clientId);fd.append('category',cat);
  Array.from(input.files).forEach(function(f){
    var prefixedFile=new File([f],cat+'__'+f.name,{type:f.type});
    fd.append('file',prefixedFile);
  });
  var fileNames=Array.from(input.files).map(function(f){return f.name;}).join(', ');
  fetch(API_BASE+'/documents',{method:'POST',headers:authUploadHeaders(),body:fd})
  .then(function(r){if(!r.ok)throw new Error('Upload failed ('+r.status+')');return r;})
  .then(function(){
    aiTrack('DocumentUploaded',{clientType:'homecare',clientId:clientId,category:cat,files:fileNames});
    status.textContent='';input.value='';loadHcDocs(clientId);
  })
  .catch(function(e){status.textContent='Upload failed: '+e;});
}
function handleDocScan(input){
  var clientId=getHcClientId();
  if(!clientId){showAlert('Save this client first before uploading documents.');return;}
  if(!input||!input.files||!input.files.length)return;
  var cat=_catKey((document.getElementById('hcDocCategory')&&document.getElementById('hcDocCategory').value)||'Other');
  var status=document.getElementById('hcDocStatus');
  if(status)status.textContent='Uploading scanned image…';
  var fd=new FormData();
  fd.append('clientType','homecare');fd.append('clientId',clientId);fd.append('category',cat);
  var f=input.files[0];
  var prefixedFile=new File([f],cat+'__'+f.name,{type:f.type});
  fd.append('file',prefixedFile);
  fetch(API_BASE+'/documents',{method:'POST',headers:authUploadHeaders(),body:fd})
  .then(function(r){if(!r.ok)throw new Error('Upload failed ('+r.status+')');return r;})
  .then(function(){
    aiTrack('DocumentUploaded',{clientType:'homecare',clientId:clientId,category:cat,files:f.name,source:'scan'});
    if(status)status.textContent='';input.value='';loadHcDocs(clientId);
  })
  .catch(function(e){if(status)status.textContent='Upload failed: '+e;});
}
function deleteHcDoc(clientId,encodedName){
  showConfirm('Delete this document?',function(){
    fetch(API_BASE+'/documents?clientType=homecare&clientId='+clientId,{method:'DELETE',headers:apiHeaders(),body:JSON.stringify({name:decodeURIComponent(encodedName)})})
    .then(function(r){if(!r.ok)throw new Error('HTTP '+r.status);loadHcDocs(clientId);}).catch(function(e){showAlert('Delete failed: '+e);});
  },{title:'Delete Document',okText:'Delete'});
}

// ── CAREGIVER DOCUMENTS ──────────────────────────────────────
function loadCgDocs(cgId){
  var sec=document.getElementById('cgDocsSection');if(!sec)return;
  sec.innerHTML='<div style="font-weight:600;font-size:12px;margin-bottom:8px;color:#1a3a5c;">Documents <span style="font-weight:normal;color:#5c7590;font-size:11px;">(SSN card, license, certs)</span></div>'+
    '<div id="cgDocList" style="margin-bottom:8px;"><div style="color:#5c7590;font-size:12px;">Loading...</div></div>'+
    '<div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;">'+
    '<input type="file" id="cgDocFileInput" accept=".pdf,.jpg,.jpeg,.png,.doc,.docx" multiple style="font-size:11px;flex:1;min-width:0;">'+
    '<button class="btn btn-primary btn-sm" onclick="uploadCgDoc(\''+cgId+'\')">Upload</button>'+
    '</div>'+
    '<span id="cgDocStatus" style="font-size:11px;color:#666;"></span>';
  fetch(API_BASE+'/documents?clientType=caregiver&clientId='+cgId,{headers:apiHeaders()})
  .then(function(r){ if(!r.ok)throw new Error('HTTP '+r.status); return r.json(); })
  .then(function(docs){
    var list=document.getElementById('cgDocList');if(!list)return;
    if(!Array.isArray(docs)||!docs.length){list.innerHTML='<div style="color:#5c7590;font-size:12px;">No documents yet.</div>';return;}
    list.innerHTML='';
    docs.forEach(function(d){
      var kb=d.size?Math.round(d.size/1024)+'KB':'';
      var div=document.createElement('div');
      div.style.cssText='display:flex;align-items:center;gap:6px;padding:5px 0;border-bottom:1px solid #f0f0f0;font-size:12px;';
      div.innerHTML='<a href="'+esc(d.url)+'" target="_blank" style="flex:1;color:#1a3a5c;text-decoration:none;word-break:break-all;">'+esc(d.name)+'</a>'+
        '<span style="color:#5c7590;font-size:11px;">'+kb+'</span>'+
        '<button class="btn btn-danger btn-sm" style="padding:2px 8px;font-size:10px;" onclick="deleteCgDoc(\''+cgId+'\',\''+encodeURIComponent(d.name)+'\')">✕</button>';
      list.appendChild(div);
    });
  }).catch(function(err){ _renderDocLoadError('cgDocList', "loadCgDocs('"+cgId+"')", err); });
}
function uploadCgDoc(cgId){
  var input=document.getElementById('cgDocFileInput');
  if(!input||!input.files||!input.files.length){showAlert('Please select a file first.');return;}
  var status=document.getElementById('cgDocStatus');status.textContent='Uploading...';
  var fd=new FormData();fd.append('clientType','caregiver');fd.append('clientId',cgId);
  var cgFileNames=Array.from(input.files).map(function(f){return f.name;}).join(', ');
  Array.from(input.files).forEach(function(f){fd.append('file',f);});
  fetch(API_BASE+'/documents',{method:'POST',headers:authUploadHeaders(),body:fd})
  .then(function(r){if(!r.ok)throw new Error('Upload failed ('+r.status+')');return r;})
  .then(function(){
    aiTrack('DocumentUploaded',{clientType:'caregiver',clientId:cgId,files:cgFileNames});
    status.textContent='';input.value='';loadCgDocs(cgId);
  })
  .catch(function(e){status.textContent='Upload failed: '+e;});
}
function deleteCgDoc(cgId,encodedName){
  showConfirm('Delete this document?',function(){
    fetch(API_BASE+'/documents?clientType=caregiver&clientId='+cgId,{method:'DELETE',headers:apiHeaders(),body:JSON.stringify({name:decodeURIComponent(encodedName)})})
    .then(function(r){if(!r.ok)throw new Error('HTTP '+r.status);loadCgDocs(cgId);}).catch(function(e){showAlert('Delete failed: '+e);});
  },{title:'Delete Document',okText:'Delete'});
}

// ============================================================
//  NEW CLIENT
// ============================================================
// Show the carrier + member fields only when Program = Carrier (managed care). Used on both the
// new-client form (nc-*) and by navNewClient's reset.
function ncProgramToggle(){
  var prog=(document.getElementById('nc-program')||{}).value||'';
  var show=prog==='carrier';
  var cw=document.getElementById('nc-carrier-wrap'), mw=document.getElementById('nc-member-wrap');
  if(cw)cw.style.display=show?'':'none';
  if(mw)mw.style.display=show?'':'none';
}
// Same show/hide for the client EDIT form (ei-*): carrier + member appear only for a Carrier client.
function eiProgramToggle(){
  var prog=(document.getElementById('ei-program')||{}).value||'champs';
  var show=prog==='carrier';
  var cw=document.getElementById('ei-carrier-wrap'), mw=document.getElementById('ei-member-wrap');
  if(cw)cw.style.display=show?'':'none';
  if(mw)mw.style.display=show?'':'none';
}
function createClient(){
  var first=(document.getElementById('nc-first').value||'').trim();
  var middle=(document.getElementById('nc-middle').value||'').trim();
  var last=(document.getElementById('nc-last').value||'').trim();
  var nickname=(document.getElementById('nc-nickname').value||'').trim();
  if(!first||!last){showAlert('First Name and Last Name are required.');return;}
  // Program is REQUIRED at creation so a client can never be mislabeled MDHHS vs Carrier.
  var ncProgram=(document.getElementById('nc-program')||{}).value||'';
  if(!ncProgram){showAlert('Please choose a Program — MDHHS (CHAMPS) or Carrier (Managed Care) — for this client.');return;}
  var ncStatus=(document.getElementById('nc-status')||{}).value||'active';
  if(ncStatus==='active' && !((document.getElementById('nc-start-date')||{}).value||'').trim()){showAlert('Service Start Date is required when the status is Active.');return;}
  // The "must import a DHS-1210 to go Active" rule is CHAMPS-only — carrier clients don't use that form.
  if(ncStatus==='active' && ncProgram!=='carrier'){showAlert('A new client can’t start as Active — a DHS-1210 authorization is required first.\n\nCreate them as “In Progress”, then import the DHS-1210 on their Authorization tab to activate them.');return;}
  var name=(first+' '+last).trim();
  var p=getProfiles();
  if(p[name]){
    showConfirm('"'+name+'" already exists. Overwrite their existing data?',function(){_doCreateClient(name,first,middle,last,nickname);},{title:'Client Exists',okText:'Overwrite'});
    return;
  }
  _doCreateClient(name,first,middle,last,nickname);
}
function _doCreateClient(name,first,middle,last,nickname){
  var p=getProfiles();
  aiTrack('ClientCreated',{client:name});
  if(!p[name])p[name]={invoices:[],clientNotes:''};
  p[name].clientName=name;p[name].firstName=first;p[name].middleName=middle;p[name].lastName=last;p[name].nickname=nickname;
  p[name].program=(document.getElementById('nc-program')||{}).value||'';
  p[name].carrier=(document.getElementById('nc-carrier')||{}).value||'';
  p[name].memberId=(document.getElementById('nc-member')||{}).value.trim()||'';
  if(p[name].program!=='carrier'){ p[name].carrier=''; p[name].memberId=''; }   // hidden fields keep stale values
  p[name].medicaidId=document.getElementById('nc-medicaid').value.trim();
  p[name].medicare=(document.getElementById('nc-medicare')||{}).value.trim()||'';
  p[name].hourlyRate=document.getElementById('nc-rate').value.trim();
  p[name].driversLicense=document.getElementById('nc-dl').value.trim();
  p[name].ssn=document.getElementById('nc-ssn').value.trim();
  p[name].phone=document.getElementById('nc-phone').value.trim();p[name].clientEmail=document.getElementById('nc-cemail').value.trim();
  p[name].street=document.getElementById('nc-street').value.trim();
  p[name].city=document.getElementById('nc-city').value.trim();
  p[name].state=document.getElementById('nc-state').value.trim();
  p[name].zip=document.getElementById('nc-zip').value.trim();
  p[name].county=document.getElementById('nc-county').value.trim();
  p[name].worker=document.getElementById('nc-worker-search').value.trim();
  p[name].caseworkerId=document.getElementById('nc-worker-val').value.trim();
  p[name].caregiverId=(document.getElementById('nc-caregiver-val').value||'').trim();
  p[name].liveIn=document.getElementById('nc-live-in').checked;
  p[name].startDate=document.getElementById('nc-start-date').value||'';
  p[name].clientStatus=document.getElementById('nc-status').value||'active';
  p[name].hasComplex=false;if(!p[name].tasks)p[name].tasks=captureStates();
  saveProfilesLS(p);saveProfileSP(name,p[name]);logActivity('client','New client added: '+name);navDetail(name);
}

// ============================================================
//  CAREGIVERS
// ============================================================
var _cgSsnMem = Object.create(null);
function getCaregivers(){
  try{
    var cg=JSON.parse(localStorage.getItem('lhca_caregivers')||'{}');
    for(var id in cg){ if(cg[id] && _cgSsnMem[id]!==undefined) cg[id].ssn=_cgSsnMem[id]; }
    return cg;
  }catch(e){return{};}
}
function saveCaregiversLS(cg){
  // HIPAA/S8: never persist MI Login passwords OR SSN to localStorage. Strip on a shallow
  // copy so the in-memory record still has them for the API save; SSN is cached in memory
  // (_cgSsnMem) and overlaid back by getCaregivers.
  var clean={};
  Object.keys(cg).forEach(function(k){
    var c=Object.assign({},cg[k]);
    if(c.ssn) _cgSsnMem[k]=c.ssn;
    delete c.miloginPassword;delete c.milogin_password;delete c.ssn;
    clean[k]=c;
  });
  localStorage.setItem('lhca_caregivers',JSON.stringify(clean));
}
function cgId(){return 'cg_'+Date.now()+'_'+Math.random().toString(36).slice(2,7);}

var cgBulkSelected={};
// Sort / pagination / column-width state for Caregivers table — mirrors Clients pattern
var _cgSort={key:'name',dir:'asc'};
// Widths sum to ~95; Caregiver column trimmed from 24 → 20 for balance.
var _cgColDefaults={name:20,status:9,phone:16,role:16,hireDate:15,clients:19};
var _cgColumnWidths=null;
var _cgPage=1;
var _cgPageSize=25;
(function loadCgPrefs(){
  try{
    var w=localStorage.getItem('lhca_cg_col_widths');if(w)_cgColumnWidths=JSON.parse(w);
    var ps=localStorage.getItem('lhca_cg_page_size');
    if(ps==='all')_cgPageSize=Infinity;else if(ps){var n=parseInt(ps);if(n>0)_cgPageSize=n;}
  }catch(e){}
})();
function sortCgBy(key){if(_cgSort.key===key)_cgSort.dir=_cgSort.dir==='asc'?'desc':'asc';else{_cgSort.key=key;_cgSort.dir='asc';}renderCaregiverGrid();}
function _cgSortCompare(a,b,cgs,profiles){
  var ca=cgs[a],cb=cgs[b];
  var key=_cgSort.key,dir=_cgSort.dir==='asc'?1:-1;
  function clientCount(id){return Object.keys(profiles).filter(function(k){return profiles[k].caregiverId===id && (profiles[k].clientStatus||'active')==='active';}).length;}
  function val(k,c,id){
    if(k==='name')return (c.name||'').toLowerCase();
    if(k==='status')return c.status||'active';
    if(k==='phone')return (c.phone||'').toLowerCase();
    if(k==='role')return (c.emptype||'').toLowerCase();
    if(k==='hireDate')return c.hireDate?new Date(c.hireDate).getTime():0;
    if(k==='clients')return clientCount(id);
    return 0;
  }
  var va=val(key,ca,a),vb=val(key,cb,b);
  if(typeof va==='string')return va.localeCompare(vb)*dir;
  return (va<vb?-1:va>vb?1:0)*dir;
}
function applyCgColWidths(){
  var headers=document.querySelectorAll('#cgTable thead th[data-col]');
  headers.forEach(function(th){var col=th.dataset.col;var w=(_cgColumnWidths&&_cgColumnWidths[col])||_cgColDefaults[col];if(w)th.style.width=w+'%';});
}
function startCgColResize(e,col){
  e.preventDefault();e.stopPropagation();
  var th=e.target.closest('th');if(!th)return;
  var startX=e.pageX,startWidth=th.offsetWidth,tableWidth=th.closest('table').offsetWidth;
  document.body.style.cursor='col-resize';document.body.style.userSelect='none';
  function onMove(ev){var d=ev.pageX-startX;var nx=Math.max(60,startWidth+d);var pct=(nx/tableWidth)*100;th.style.width=pct+'%';if(!_cgColumnWidths)_cgColumnWidths={};_cgColumnWidths[col]=pct;}
  function onUp(){document.removeEventListener('mousemove',onMove);document.removeEventListener('mouseup',onUp);document.body.style.cursor='';document.body.style.userSelect='';try{localStorage.setItem('lhca_cg_col_widths',JSON.stringify(_cgColumnWidths||{}));}catch(e){}}
  document.addEventListener('mousemove',onMove);document.addEventListener('mouseup',onUp);
}
function cgPageSizeChange(){var s=document.getElementById('cgPageSize');var v=s.value;if(v==='all')_cgPageSize=Infinity;else _cgPageSize=parseInt(v)||25;_cgPage=1;try{localStorage.setItem('lhca_cg_page_size',v);}catch(e){}renderCaregiverGrid();}
function goToCgPage(n){_cgPage=n;renderCaregiverGrid();}

function renderCaregiverGrid(){
  var cgs=getCaregivers();
  var q=(document.getElementById('cgSearch')?document.getElementById('cgSearch').value:'').toLowerCase();
  var filterStatus=(document.getElementById('cgFilterStatus')&&document.getElementById('cgFilterStatus').value)||'active';
  var profiles=getProfiles();

  var allIds=Object.keys(cgs);
  var ids=allIds.filter(function(id){
    var cg=cgs[id];
    var st=cg.status||'active';
    var matchStatus=filterStatus==='all'||st===filterStatus;
    var matchQ=!q||(cg.name||'').toLowerCase().includes(q)||(cg.phone||'').includes(q);
    return matchStatus&&matchQ;
  });
  ids.sort(function(a,b){return _cgSortCompare(a,b,cgs,profiles);});

  var headers=document.querySelectorAll('#cgTable thead th.sortable');
  headers.forEach(function(h){h.classList.remove('sort-asc','sort-desc');if(h.dataset.sortkey===_cgSort.key)h.classList.add(_cgSort.dir==='asc'?'sort-asc':'sort-desc');});
  applyCgColWidths();

  var totalMatched=ids.length;
  var totalPages=Math.max(1,Math.ceil(totalMatched/_cgPageSize));
  if(_cgPage>totalPages)_cgPage=totalPages;
  var startIdx=(_cgPage-1)*_cgPageSize;
  var endIdx=Math.min(startIdx+_cgPageSize,totalMatched);
  var visibleIds=ids.slice(startIdx,endIdx);
  var countEl=document.getElementById('cgTableCount');
  if(countEl){
    if(!totalMatched)countEl.textContent='';
    else if(_cgPageSize===Infinity||totalMatched<=_cgPageSize)countEl.textContent=totalMatched+' caregiver'+(totalMatched===1?'':'s');
    else countEl.textContent='Showing '+(startIdx+1)+'–'+endIdx+' of '+totalMatched;
  }
  var pageControls=document.getElementById('cgPageControls');
  if(pageControls){
    if(totalPages<=1)pageControls.innerHTML='';
    else {
      var html='<button class="pg-btn" onclick="goToCgPage('+(_cgPage-1)+')"'+(_cgPage===1?' disabled':'')+'>‹</button>';
      var startPg=Math.max(1,_cgPage-2),endPg=Math.min(totalPages,startPg+4);
      if(endPg-startPg<4)startPg=Math.max(1,endPg-4);
      for(var p=startPg;p<=endPg;p++)html+='<button class="pg-btn'+(p===_cgPage?' active':'')+'" onclick="goToCgPage('+p+')">'+p+'</button>';
      html+='<button class="pg-btn" onclick="goToCgPage('+(_cgPage+1)+')"'+(_cgPage===totalPages?' disabled':'')+'>›</button>';
      pageControls.innerHTML=html;
    }
  }

  var tbody=document.getElementById('cgTableBody');if(!tbody)return;tbody.innerHTML='';
  var empty=document.getElementById('cgTableEmpty');
  if(!ids.length){if(empty)empty.style.display='block';return;}
  if(empty)empty.style.display='none';
  visibleIds.forEach(function(id){
    var cg=cgs[id];
    var st=cg.status||'active';
    var stLabel=st.charAt(0).toUpperCase()+st.slice(1);
    var clientCount=Object.keys(profiles).filter(function(k){return profiles[k].caregiverId===id && (profiles[k].clientStatus||'active')==='active';}).length;
    var displayName=(cg.firstName&&cg.lastName)?(cg.firstName+(cg.middleName?' '+cg.middleName:'')+' '+cg.lastName).trim():(cg.name||id);
    var tr=document.createElement('tr');
    var hrefCg=buildCaregiverUrl(id);
    var checked=cgBulkSelected[id]?'checked':'';
    tr.innerHTML=
      '<td style="width:26px;" onclick="event.stopPropagation()"><input type="checkbox" class="cg-select" data-id="'+esc(id)+'" '+checked+' onchange="toggleBulkCaregiver(\''+escJsAttr(id)+'\',this)" style="width:12px;height:12px;cursor:pointer;"></td>'+
      '<td><a href="'+hrefCg+'" class="link-plain" style="display:block;" onclick="return navClick(event,this.getAttribute(\'href\'))"><div class="ct-name">'+esc(displayName)+(cg.nickname?'<span style="font-weight:normal;color:var(--text-subtle);"> ('+esc(cg.nickname)+')</span>':'')+'</div><div class="ct-id">'+esc(cg.email||'No email')+'</div></a></td>'+
      '<td><span class="status-inline"><span class="status-dot '+st+'"></span>'+stLabel+'</span></td>'+
      '<td style="color:var(--text-muted);font-size:12px;">'+esc(cg.phone||'—')+'</td>'+
      '<td style="color:var(--text-muted);font-size:12px;">'+esc(cg.emptype||'—')+'</td>'+
      '<td style="color:var(--text-muted);font-size:12px;">'+esc(cg.hireDate||'—')+'</td>'+
      '<td style="color:var(--text);font-size:12px;">'+clientCount+'</td>';
    tr.addEventListener('click',function(e){
      if(e.target.closest('a')||e.target.tagName==='BUTTON'||e.target.tagName==='INPUT')return;
      openCgDetail(id);
    });
    tbody.appendChild(tr);
  });

  var sel=document.getElementById('cgPageSize');
  if(sel){var savedPs=_cgPageSize===Infinity?'all':String(_cgPageSize);if(sel.value!==savedPs)sel.value=savedPs;}
}
function toggleBulkCaregiver(id,cb){
  if(cb.checked)cgBulkSelected[id]=true;else delete cgBulkSelected[id];
  var count=Object.keys(cgBulkSelected).length;
  var bar=document.getElementById('cgBulkBar');
  if(bar){bar.classList.toggle('visible',count>0);var lbl=document.getElementById('cgBulkCount');if(lbl)lbl.textContent=count+' selected';}
}
function clearCgBulkSelect(){cgBulkSelected={};var bar=document.getElementById('cgBulkBar');if(bar)bar.classList.remove('visible');var sa=document.getElementById('cgSelectAll');if(sa)sa.checked=false;renderCaregiverGrid();}
function toggleAllCaregivers(cb){
  document.querySelectorAll('.cg-select').forEach(function(c){
    c.checked=cb.checked;
    var id=c.dataset.id;
    if(cb.checked)cgBulkSelected[id]=true;else delete cgBulkSelected[id];
  });
  var count=Object.keys(cgBulkSelected).length;
  var bar=document.getElementById('cgBulkBar');
  if(bar){bar.classList.toggle('visible',count>0);var lbl=document.getElementById('cgBulkCount');if(lbl)lbl.textContent=count+' selected';}
}
function bulkSetCaregiverStatus(status){
  var ids=Object.keys(cgBulkSelected);if(!ids.length)return;
  var cgs=getCaregivers(),changed=0;
  ids.forEach(function(id){if(cgs[id]){cgs[id].status=status;changed++;}});
  saveCaregiversLS(cgs);
  batchSaveWithSummary('Caregiver status', ids.filter(function(id){return cgs[id];}).map(function(id){var th=function(){return saveCaregiverAPI(id,cgs[id],true);};th._label=(cgs[id]&&cgs[id].name)||id;return th;}));
  logActivity('status','Bulk caregiver update: '+changed+' set to '+status);
  clearCgBulkSelect();updateStats();
  showAlert(changed+' caregiver'+(changed!==1?'s':'')+' updated to '+status+'.');
}
function bulkDeleteCaregivers(){
  var ids=Object.keys(cgBulkSelected);if(!ids.length){showAlert('No caregivers selected.');return;}
  var cgs=getCaregivers();
  var preview=ids.slice(0,5).map(function(id){return '• '+(cgs[id]?cgs[id].name||id:id);}).join('\n')+(ids.length>5?'\n• …and '+(ids.length-5)+' more':'');
  showConfirm(
    'Delete '+ids.length+' caregiver'+(ids.length>1?'s':'')+'?\n\n'+preview+'\n\nThis cannot be undone.',
    function(){
      var deleted=0;
      ids.forEach(function(id){if(cgs[id]){delete cgs[id];try{deleteCaregiverAPI(id);}catch(e){}deleted++;}});
      saveCaregiversLS(cgs);
      logActivity('delete','Bulk caregiver delete: '+deleted+' removed');
      clearCgBulkSelect();updateStats();
      showAlert(deleted+' caregiver'+(deleted!==1?'s':'')+' deleted.');
    },
    {title:'Delete '+ids.length+' Caregivers?',okText:'Delete '+ids.length+(ids.length>1?' Caregivers':' Caregiver'),danger:true}
  );
}
function showNewCaregiverForm(){
  document.getElementById('cgFormWrap').style.display='block';
  document.getElementById('cgFormTitle').textContent='New Caregiver';
  document.getElementById('cg-editing-id').value='';
  ['cg-first','cg-middle','cg-last','cg-nickname','cg-phone','cg-email','cg-dl','cg-ssn','cg-street','cg-city','cg-state','cg-zip','cg-county','cg-dob','cg-gender','cg-hire','cg-pay','cg-champs','cg-milogin-user','cg-milogin-pass','cg-notes'].forEach(function(id){var e=document.getElementById(id);if(e)e.value='';});
  document.getElementById('cg-status').value='active';document.getElementById('cg-emptype').value='full-time';
  document.getElementById('cgDeleteBtn').style.display='none';
  var cgDocSec=document.getElementById('cgDocsSection');if(cgDocSec)cgDocSec.style.display='none';
  document.getElementById('cgFormWrap').scrollIntoView({behavior:'smooth'});
}
// (Removed) editCaregiver() populated the OLD inline caregiver form for an existing
// caregiver, but caregiver editing moved to the detail view (openCgDetail via the
// #/caregiver/<id> route + saveCgInfoPane). It had zero callers (not even a dynamic
// onclick). The inline form now only serves "New Caregiver".
// Build a caregiver's display name from its parts. BOTH save paths must agree: they used to differ
// (this one dropped the middle name, the detail pane included it), so the first save from the detail
// pane silently renamed the record — and AuditLog rows are keyed on the exact name string, so every
// entry written under the old name became unreachable from that caregiver's Audit tab.
function _cgDisplayName(first,middle,last){
  return ((first||'')+((middle||'')?(' '+middle):'')+' '+(last||'')).trim();
}
function saveCaregiver(){
  var first=(document.getElementById('cg-first').value||'').trim();
  var middle=(document.getElementById('cg-middle').value||'').trim();
  var last=(document.getElementById('cg-last').value||'').trim();
  var nickname=(document.getElementById('cg-nickname').value||'').trim();
  if(!first||!last){showAlert('First Name and Last Name are required.');return;}
  var name=(first+' '+last).trim();
  aiTrack('CaregiverSaved',{caregiver:name});
  var cgs=getCaregivers();
  var id=document.getElementById('cg-editing-id').value||cgId();
  cgs[id]={name:name,firstName:first,middleName:middle,lastName:last,nickname:nickname,
    status:document.getElementById('cg-status').value,phone:document.getElementById('cg-phone').value,
    email:document.getElementById('cg-email').value,
    driversLicense:(document.getElementById('cg-dl')||{}).value||'',
    ssn:(document.getElementById('cg-ssn')||{}).value||'',
    street:document.getElementById('cg-street').value,
    city:document.getElementById('cg-city').value,
    state:document.getElementById('cg-state').value,
    zip:document.getElementById('cg-zip').value,
    county:document.getElementById('cg-county').value,
    address:[document.getElementById('cg-street').value,document.getElementById('cg-city').value,document.getElementById('cg-state').value,document.getElementById('cg-zip').value].filter(Boolean).join(', '),
    dob:(document.getElementById('cg-dob')||{}).value||'',
    hireDate:document.getElementById('cg-hire').value,emptype:document.getElementById('cg-emptype').value,
    payRate:document.getElementById('cg-pay').value,
    champsId:(document.getElementById('cg-champs')||{}).value||'',
    miloginUsername:(document.getElementById('cg-milogin-user')||{}).value||'',
    miloginPassword:(document.getElementById('cg-milogin-pass')||{}).value||'',
    gender:(document.getElementById('cg-gender')||{}).value||'',
    notes:document.getElementById('cg-notes').value};
  saveCaregiversLS(cgs);saveCaregiverAPI(id, cgs[id]);hideCgForm();updateStats();
  // Return to detail view if we were editing an existing caregiver
  if(id&&document.getElementById('cg-editing-id').value){
    openCgDetail(id);
  } else {
    document.getElementById('cgGridView').style.display='';renderCaregiverGrid();
  }
}
function deleteCaregiver(){
  var id=document.getElementById('cg-editing-id').value;if(!id)return;
  var cgs=getCaregivers();var cgName=(cgs[id]&&cgs[id].name)||id;
  showConfirm('Delete caregiver "'+cgName+'"? This cannot be undone.',function(){_doDeleteCaregiver(id);},{title:'Delete Caregiver',okText:'Delete'});
}
function _doDeleteCaregiver(id){
  try{ var _n=(getCaregivers()[id]||{}).name||id; addAuditEntry('',"Caregiver DELETED: "+_n+' by '+currentUserEmail()); }catch(e){}
  var cgs=getCaregivers();var cgName=(cgs[id]&&cgs[id].name)||id;
  delete cgs[id];saveCaregiversLS(cgs);deleteCaregiverAPI(id);
  aiTrack('CaregiverDeleted',{caregiverId:id,caregiverName:cgName});
  hideCgForm();activeCgId=null;
  document.getElementById('cgDetailView').style.display='none';
  document.getElementById('cgGridView').style.display='';
  renderCaregiverGrid();updateStats();
}
function hideCgForm(){
  document.getElementById('cgFormWrap').style.display='none';
  // If neither detail nor grid is visible, show grid
  var dv=document.getElementById('cgDetailView');
  var gv=document.getElementById('cgGridView');
  if(gv&&gv.style.display==='none'&&(!dv||dv.style.display==='none')){
    gv.style.display='';
    renderCaregiverGrid();
  }
}


// --- Caregiver detail view ---
var activeCgId=null;
function openCgDetail(id){
  var cg=getCaregivers()[id];if(!cg)return;
  activeCgId=id;
  // Match Client detail's topbar breadcrumb pattern: "Caregivers > [name]" with the
  // parent clickable to go back to the list. Was previously left as just "Caregivers"
  // from the parent page.
  bc([{l:'Caregivers',fn:navCaregivers},{l:cg.name||''}]);
  document.getElementById('cgGridView').style.display='none';
  document.getElementById('cgFormWrap').style.display='none';
  document.getElementById('cgDetailView').style.display='block';
  var ini=(cg.name||'?').split(' ').map(function(w){return w[0]||'';}).slice(0,2).join('').toUpperCase();
  document.getElementById('cgDetailAvatar').textContent=ini;
  document.getElementById('cgDetailName').textContent=cg.name||'';
  var st=cg.status||'active';
  document.getElementById('cgDetailMeta').innerHTML=esc(cg.emptype||'')+(cg.payRate?' · $'+cg.payRate+'/hr':'')+
    ' &nbsp;<span class="cs-badge cs-'+st+'">'+st.charAt(0).toUpperCase()+st.slice(1)+'</span>';
  switchCgTab('overview');
}
function showCgGrid(){
  activeCgId=null;
  document.getElementById('cgDetailView').style.display='none';
  document.getElementById('cgGridView').style.display='';
  renderCaregiverGrid();
}

// ── SIGNING TEMPLATES (admin uploader in Settings) ──────────────
async function loadSigningTemplates(){
  var host=document.getElementById('signTplList');if(!host)return;
  try{
    var resp=await fetch(API_BASE+'/signing/templates',{headers:apiHeaders()});
    if(!resp.ok)throw new Error('HTTP '+resp.status);
    var arr=await resp.json();
    if(!arr.length){host.innerHTML='<div style="padding:14px;color:#5c7590;text-align:center;">No templates uploaded yet. Pick a PDF above to add your first one.</div>';return;}
    host.innerHTML=arr.map(function(t){
      var when=new Date(t.created_at).toLocaleDateString();
      return '<div style="display:flex;align-items:center;gap:10px;padding:10px 14px;border-bottom:1px solid #f0f3f7;">'+
        '<div style="flex:1;min-width:0;">'+
          '<div style="font-weight:600;color:#1a2b45;">'+esc(t.name)+(t.version?' <span style="color:#5c7590;font-weight:400;font-size:11px;">('+esc(t.version)+')</span>':'')+'</div>'+
          '<div style="font-size:11px;color:#5c7590;">Uploaded '+when+(t.is_active?'':' · Inactive')+'</div>'+
        '</div>'+
      '</div>';
    }).join('');
  }catch(e){
    host.innerHTML='<div style="padding:14px;color:#a05a00;font-size:12px;">Could not load templates: '+esc(e.message||'unknown')+'</div>';
  }
}
async function uploadSigningTemplate(){
  var fileInp=document.getElementById('signTplFile');var nameInp=document.getElementById('signTplName');var verInp=document.getElementById('signTplVersion');
  if(!fileInp.files||!fileInp.files[0]){showAlert('Pick a PDF file first.');return;}
  var file=fileInp.files[0];
  if(file.type!=='application/pdf'&&!/\.pdf$/i.test(file.name)){showAlert('Only PDF files are accepted as signing templates.');return;}
  var name=(nameInp.value||'').trim()||file.name.replace(/\.pdf$/i,'');
  var version=(verInp.value||'').trim()||'v1';
  try{
    showToast('Uploading '+file.name+'…',2000);
    // Step 1: upload PDF bytes
    var upResp=await fetch(API_BASE+'/signing/templates/upload?name='+encodeURIComponent(name.replace(/\s+/g,'_'))+'&version='+encodeURIComponent(version),{
      method:'POST',headers:Object.assign({},apiHeaders(),{'Content-Type':'application/pdf'}),body:file
    });
    var upData=await upResp.json();
    if(!upResp.ok)throw new Error(upData.error||'Upload failed ('+upResp.status+')');
    // Step 2: register the template
    var regResp=await fetch(API_BASE+'/signing/templates',{
      method:'POST',headers:apiHeaders(),body:JSON.stringify({name:name,version:version,blobPath:upData.blobPath})
    });
    var regData=await regResp.json();
    if(!regResp.ok)throw new Error(regData.error||'Register failed ('+regResp.status+')');
    showToast('✓ Template uploaded',3000);
    fileInp.value='';nameInp.value='';verInp.value='';
    loadSigningTemplates();
  }catch(e){
    showAlert('Upload failed: '+(e.message||e));
  }
}

// ── SIGNING REQUESTS (per-caregiver panel) ──────────────────────
async function loadCgSigningRequests(){
  var host=document.getElementById('cgSigList');if(!host||!activeCgId)return;
  host.innerHTML='<div style="padding:14px;color:#5c7590;text-align:center;font-size:12px;">Loading…</div>';
  try{
    var resp=await fetch(API_BASE+'/signing/list?caregiverId='+encodeURIComponent(activeCgId),{headers:apiHeaders()});
    if(!resp.ok)throw new Error('HTTP '+resp.status);
    var arr=await resp.json();
    if(!arr.length){host.innerHTML='<div style="padding:14px;color:#5c7590;text-align:center;font-size:12px;">No signing requests yet. Use the Send for Signature button above to create one.</div>';return;}
    var now=Date.now();
    host.innerHTML=arr.map(function(r){
      var status=r.status||'sent';
      var locked=!!r.locked_at;
      var color=status==='signed'?'#1a7740':status==='revoked'?'#888':locked?'#a72e2e':status==='expired'?'#888':'#185FA5';
      var label=locked?'LOCKED':status.toUpperCase();
      var expDate=new Date(r.expires_at);
      var expDays=Math.round((expDate-now)/86400000);
      var expText=status==='signed'?'Signed '+new Date(r.signed_at).toLocaleDateString()
                 :status==='revoked'?'Revoked'
                 :status==='expired'||now>expDate?'Expired '+expDate.toLocaleDateString()
                 :locked?'Locked — needs resend'
                 :('Expires in '+expDays+'d ('+expDate.toLocaleDateString()+')');
      var actions='';
      if(status==='signed')actions+='<button class="btn btn-secondary btn-sm" style="font-size:11px;padding:3px 8px;" onclick="downloadSignedDoc('+r.id+')">Download</button>';
      if(status!=='signed'&&status!=='revoked')actions+='<button class="btn btn-secondary btn-sm" style="font-size:11px;padding:3px 8px;" onclick="revokeSigningRequest('+r.id+')">Revoke</button>';
      if(locked||status==='expired'||status==='revoked'||(now>expDate&&status!=='signed'))actions+='<button class="btn btn-primary btn-sm" style="font-size:11px;padding:3px 8px;" onclick="resendSigningRequest('+r.id+')">Resend</button>';
      actions+='<button class="btn btn-secondary btn-sm" style="font-size:11px;padding:3px 8px;" onclick="viewSigningAudit('+r.id+')">Audit</button>';
      return '<div style="padding:10px 14px;border-bottom:1px solid #f0f3f7;display:flex;align-items:center;gap:10px;flex-wrap:wrap;">'+
        '<div style="flex:1;min-width:200px;">'+
          '<div style="font-weight:600;font-size:12px;color:#1a2b45;">'+esc(r.template_name||'Document')+(r.template_version?' <span style="color:#5c7590;font-weight:400;">('+esc(r.template_version)+')</span>':'')+'</div>'+
          '<div style="font-size:11px;color:#5a7296;">'+esc(expText)+(locked?' · '+(r.verification_attempts||0)+' failed DOB attempts':'')+'</div>'+
        '</div>'+
        '<span style="font-size:10px;font-weight:700;color:#fff;background:'+color+';padding:2px 7px;border-radius:3px;">'+label+'</span>'+
        '<div style="display:flex;gap:5px;">'+actions+'</div>'+
      '</div>';
    }).join('');
  }catch(e){
    host.innerHTML='<div style="padding:14px;color:#a05a00;font-size:12px;">Could not load requests: '+esc(e.message||'unknown')+'</div>';
  }
}
async function downloadSignedDoc(id){
  try{
    var resp=await fetch(API_BASE+'/signing/'+id+'/signed-url',{headers:apiHeaders()});
    var data=await resp.json();
    if(!resp.ok)throw new Error(data.error||'HTTP '+resp.status);
    window.open(data.url,'_blank');
  }catch(e){showAlert('Could not open signed copy: '+(e.message||e));}
}
function revokeSigningRequest(id){
  showConfirm('Revoke this signing request? The recipient\'s link will stop working immediately. They can be sent a new one with the Resend button.',async function(){
    try{
      var resp=await fetch(API_BASE+'/signing/'+id+'/revoke',{method:'POST',headers:apiHeaders()});
      if(!resp.ok)throw new Error('HTTP '+resp.status);
      showToast('Request revoked',3000);
      loadCgSigningRequests();
    }catch(e){showAlert('Failed: '+(e.message||e));}
  },{title:'Revoke',okText:'Revoke',danger:true});
}
// Shared, professional signing-request email — used by both the initial send and
// the reminder. Table layout + inline styles for email-client compatibility, and
// no external images (blocked images look broken/spammy). The clean layout, real
// footer, and de-emphasized fallback link keep it out of spam and out of the
// "is this phishing?" pile.
function buildSigningEmail(o){
  o=o||{};
  var brand='Liberty Home Care Assistance';
  var accent='#185FA5';
  var logoUrl=((typeof location!=='undefined'&&location.origin)?location.origin:'https://app.libertybellhealth.com')+'/email-logo.png';
  var docName=esc((o.docName||'').trim()||'your document');
  var name=esc((o.name||'').trim()||'there');
  var url=o.signUrl||'#';
  var urlEsc=esc(url);
  var expiry='';
  try{ if(o.expiresAt) expiry=new Date(o.expiresAt).toLocaleDateString(undefined,{month:'long',day:'numeric',year:'numeric'}); }catch(e){}
  var subject=o.isReminder ? ('Reminder — please sign '+docName) : ('Signature requested: '+docName);
  var lead=o.isReminder
    ? ('This is a friendly reminder that '+brand+' still needs your signature on <strong>'+docName+'</strong>. Your previous link has expired, so here is a fresh one.')
    : (brand+' has a document ready for your signature: <strong>'+docName+'</strong>. It only takes a minute to review and sign online — no account or app needed.');
  var html=''
    +'<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#f4f6f8;padding:24px 0;margin:0;font-family:-apple-system,BlinkMacSystemFont,\'Segoe UI\',Roboto,Helvetica,Arial,sans-serif;">'
    +'<tr><td align="center">'
    +'<table role="presentation" width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;background:#ffffff;border-radius:10px;overflow:hidden;border:1px solid #e6eaee;">'
    +'<tr><td align="center" style="background:#000000;padding:22px 32px;">'
    +'<img src="'+esc(logoUrl)+'" width="280" alt="'+brand+'" style="display:block;margin:0 auto;border:0;outline:none;text-decoration:none;width:280px;max-width:72%;height:auto;color:#d4af6a;font-size:20px;font-weight:700;">'
    +'</td></tr>'
    +'<tr><td style="padding:32px 32px 28px;color:#243b53;font-size:15px;line-height:1.6;">'
    +'<p style="margin:0 0 16px;">Hi '+name+',</p>'
    +'<p style="margin:0 0 26px;">'+lead+'</p>'
    +'<table role="presentation" cellpadding="0" cellspacing="0" style="margin:0 auto 26px;"><tr><td style="border-radius:6px;background:'+accent+';">'
    +'<a href="'+urlEsc+'" style="display:inline-block;padding:14px 32px;color:#ffffff;font-size:15px;font-weight:600;text-decoration:none;border-radius:6px;">Review &amp; sign your document</a>'
    +'</td></tr></table>'
    +'<p style="margin:0 0 20px;font-size:13px;color:#486581;background:#eef3f8;border-radius:6px;padding:12px 14px;">🔒 For your protection, you\'ll be asked to confirm your date of birth before the document opens — so your information stays private even if this email is forwarded.</p>'
    +(expiry?'<p style="margin:0 0 4px;font-size:13px;color:#627d98;">This secure link expires on <strong>'+esc(expiry)+'</strong> and can only be used once.</p>':'')
    +'<p style="margin:18px 0 0;font-size:12px;color:#9fb3c8;">If the button doesn\'t work, copy and paste this link into your browser:<br><a href="'+urlEsc+'" style="color:'+accent+';word-break:break-all;">'+urlEsc+'</a></p>'
    +'</td></tr>'
    +'<tr><td style="padding:18px 32px;background:#f7f9fb;border-top:1px solid #e6eaee;color:#829ab1;font-size:12px;line-height:1.5;">'
    +'<p style="margin:0 0 6px;color:#486581;font-weight:600;">'+brand+'</p>'
    +'<p style="margin:0 0 6px;">Need help or have a question? Just reply to this email, or call us at <a href="tel:+12482914106" style="color:#486581;font-weight:600;text-decoration:none;">(248) 291-4106</a>.</p>'
    +'<p style="margin:0;color:#9fb3c8;">You received this email because '+brand+' requested your signature on a document.</p>'
    +'</td></tr>'
    +'</table>'
    +'</td></tr></table>';
  return {subject:subject, html:html};
}
async function resendSigningRequest(id){
  if(!spToken){showAlert('Sign in with Microsoft first to send the email.');return;}
  try{
    var cg=getCaregivers()[activeCgId]||{};
    var resp=await fetch(API_BASE+'/signing/'+id+'/resend',{method:'POST',headers:apiHeaders()});
    var data=await resp.json();
    if(!resp.ok)throw new Error(data.error||'HTTP '+resp.status);
    var _mail=buildSigningEmail({name:data.recipientName||cg.name,docName:data.templateName||data.documentName||'',signUrl:data.signUrl,expiresAt:data.expiresAt,isReminder:true});
    var subject=_mail.subject;
    var body=_mail.html;
    var emailResp=await sendMailWithPDF(data.recipientEmail,subject,body,[]);
    if(!emailResp.ok){showAlert('Created the new link but email send failed: '+(emailResp.err||emailResp.status||'unknown')+'\n\nManual link:\n'+data.signUrl);loadCgSigningRequests();return;}
    showToast('✓ New link emailed',4000);
    logActivity('signing','Resent signing request to '+data.recipientName);
    loadCgSigningRequests();
  }catch(e){showAlert('Resend failed: '+(e.message||e));}
}
async function viewSigningAudit(id){
  try{
    var resp=await fetch(API_BASE+'/signing/'+id+'/audit',{headers:apiHeaders()});
    if(!resp.ok)throw new Error('HTTP '+resp.status);
    var arr=await resp.json();
    var rows=arr.map(function(e){
      var when=new Date(e.logged_at).toLocaleString();
      return when+'  ·  '+e.event_type+(e.ip?'  ·  '+e.ip:'')+(e.detail_json?'\n   '+e.detail_json:'');
    }).join('\n\n');
    showAlert(rows||'No audit events yet.',{title:'Audit Trail — Request #'+id});
  }catch(e){showAlert('Could not load audit log: '+(e.message||e));}
}

// ── SEND FOR SIGNATURE ──────────────────────────────────────────
// Picks a template, creates a SigningRequest server-side, then uses
// Graph (existing spToken) to email the recipient the secure link.
async function openSendForSignatureModal(){
  if(!activeCgId){showAlert('Open a caregiver first.');return;}
  var cg=getCaregivers()[activeCgId];if(!cg){showAlert('Caregiver not found.');return;}
  if(!cg.email){showAlert('This caregiver has no email on file. Add one in their Profile tab first.',{title:'Email Required'});return;}

  var templates=[];
  try{
    var resp=await fetch(API_BASE+'/signing/templates',{headers:apiHeaders()});
    if(resp.ok)templates=await resp.json();
  }catch(e){console.error('Templates fetch failed:',e);}
  templates=templates.filter(function(t){return t.is_active!==false;});

  var existing=document.getElementById('sendSigModal');if(existing)existing.remove();
  var ov=document.createElement('div');
  ov.id='sendSigModal';ov.className='modal-overlay open';
  var tplOptions=templates.length
    ? templates.map(function(t){return '<option value="'+t.id+'">'+esc(t.name)+(t.version?' ('+esc(t.version)+')':'')+'</option>';}).join('')
    : '<option value="">No templates yet — use Settings → Signing Templates to upload one</option>';
  // Pre-fill DOB if we have it on the caregiver record (multiple possible field names)
  var prefillDob=(cg.dob||cg.dateOfBirth||cg.birthDate||'').trim();
  ov.innerHTML='<div class="modal-box" style="max-width:540px;">'+
    '<h3>📝 Send for Signature</h3>'+
    '<div style="font-size:13px;color:#4a5d7a;margin:8px 0 14px;">Sends <b>'+esc(cg.name||'')+'</b> a secure link to '+esc(cg.email)+'. They verify their DOB, then sign — no CRM login needed.</div>'+
    '<label style="display:block;font-size:11px;color:#4d6c88;font-weight:600;text-transform:uppercase;letter-spacing:.4px;margin-bottom:4px;" for="sendSigTemplate">Document</label>'+
    '<select id="sendSigTemplate" style="width:100%;padding:8px 10px;border:1px solid #d0d8e4;border-radius:5px;font-size:13px;outline:none;background:#fff;margin-bottom:12px;">'+tplOptions+'</select>'+
    '<label style="display:block;font-size:11px;color:#4d6c88;font-weight:600;text-transform:uppercase;letter-spacing:.4px;margin-bottom:4px;" for="sendSigDob">Recipient Date of Birth <span style="font-weight:400;text-transform:none;letter-spacing:0;color:#a05a00;">(used to verify identity at signing — 3 wrong attempts locks the link)</span></label>'+
    '<input id="sendSigDob" type="date" value="'+esc(prefillDob)+'" style="width:200px;padding:8px 10px;border:1px solid #d0d8e4;border-radius:5px;font-size:13px;outline:none;margin-bottom:14px;">'+
    '<div style="font-size:11px;color:#5c7590;margin-bottom:14px;">Link expires in 14 days. A copy of every signed document is auto-CC\'d to <b>tommy@mybellcare.com</b>.</div>'+
    '<div id="sendSigError" style="display:none;background:#fdeaea;border:1px solid #e7a8a8;border-radius:5px;padding:10px;font-size:12px;color:#7a1f1f;margin-bottom:12px;"></div>'+
    '<div class="modal-row" style="display:flex;gap:8px;justify-content:flex-end;margin-top:8px;">'+
      '<button class="btn btn-secondary" onclick="closeSendSigModal()">Cancel</button>'+
      '<button id="sendSigBtn" class="btn btn-primary"'+(templates.length?'':' disabled')+' onclick="doSendForSignature()">Send Link</button>'+
    '</div>'+
  '</div>';
  document.body.appendChild(ov);
}
function closeSendSigModal(){var m=document.getElementById('sendSigModal');if(m)m.remove();}

async function doSendForSignature(){
  var btn=document.getElementById('sendSigBtn');var errEl=document.getElementById('sendSigError');
  errEl.style.display='none';
  var cg=getCaregivers()[activeCgId];
  var tplId=parseInt(document.getElementById('sendSigTemplate').value,10);
  var dob=(document.getElementById('sendSigDob').value||'').trim();
  if(!tplId){errEl.textContent='Pick a document template.';errEl.style.display='block';return;}
  if(!/^\d{4}-\d{2}-\d{2}$/.test(dob)){errEl.textContent='Enter the recipient\'s date of birth (used for identity verification).';errEl.style.display='block';return;}
  btn.disabled=true;btn.textContent='Creating link…';
  try{
    // 1. Backend creates the request and returns the sign URL
    var resp=await fetch(API_BASE+'/signing/send',{
      method:'POST',
      headers:apiHeaders(),
      body:JSON.stringify({
        templateId:tplId,
        caregiverId:activeCgId,
        recipientName:cg.name||'',
        recipientEmail:cg.email,
        recipientDob:dob,
        ccEmails:'tommy@mybellcare.com'
      })
    });
    var data=await resp.json();
    if(!resp.ok){throw new Error(data.error||'HTTP '+resp.status);}
    // 2. Frontend uses existing Graph email infra to deliver it
    btn.textContent='Sending email…';
    var _mail=buildSigningEmail({name:cg.name,docName:(document.getElementById('sendSigTemplate').selectedOptions[0].textContent||'').trim(),signUrl:data.signUrl,expiresAt:data.expiresAt,isReminder:false});
    var subject=_mail.subject;
    var body=_mail.html;
    if(spToken){
      var emailResp=await sendMailWithPDF(cg.email,subject,body,[]);
      if(!emailResp.ok){
        // Backend already created the request — show the URL so admin can copy/paste
        errEl.innerHTML='Created the request, but email failed to send: '+esc(emailResp.err||emailResp.status||'unknown')+'<br><br>You can manually share this link:<br><a href="'+esc(data.signUrl)+'" target="_blank">'+esc(data.signUrl)+'</a>';
        errEl.style.display='block';
        btn.textContent='Send Link';btn.disabled=false;
        return;
      }
    } else {
      errEl.innerHTML='Sign in with Microsoft first to email the link automatically.<br><br>For now, you can manually share this link:<br><a href="'+esc(data.signUrl)+'" target="_blank">'+esc(data.signUrl)+'</a>';
      errEl.style.display='block';
      btn.textContent='Send Link';btn.disabled=false;
      return;
    }
    closeSendSigModal();
    var _expTxt=data.expiresAt?('The link expires '+new Date(data.expiresAt).toLocaleDateString()+'.'):'The link expires in 14 days.';
    showAlert('✓ Signing link emailed to '+cg.email+'.\n\n'+_expTxt+' You can track status from the caregiver detail page.',{title:'Sent'});
    logActivity('signing','Sent signing request to '+cg.name+' for template #'+tplId);
  } catch(e){
    errEl.textContent='Error: '+(e.message||e);
    errEl.style.display='block';
    btn.textContent='Send Link';btn.disabled=false;
  }
}
function deleteCaregiverFromDetail(){
  var cg=getCaregivers()[activeCgId];
  if(!cg)return;
  showConfirm('Delete caregiver "'+cg.name+'"? This cannot be undone.',function(){_doDeleteCaregiverFromDetail();},{title:'Delete Caregiver',okText:'Delete'});
}
function _doDeleteCaregiverFromDetail(){
  var cgs=getCaregivers();
  var cgName=(cgs[activeCgId]&&cgs[activeCgId].name)||activeCgId;   // read BEFORE the delete; `cg` is the caller's local
  delete cgs[activeCgId];saveCaregiversLS(cgs);deleteCaregiverAPI(activeCgId);
  aiTrack('CaregiverDeleted',{caregiverId:activeCgId,caregiverName:cgName});
  showCgGrid();
}
var cgUnsavedChanges=false;
function switchCgTab(tab){
  // Guard against losing edits when leaving the Info tab with unsaved changes
  if(cgUnsavedChanges && tab!=='info'){
    showConfirm('You have unsaved changes on the caregiver Info tab. Leave anyway?',function(){
      cgUnsavedChanges=false;
      _doSwitchCgTab(tab);
    },{title:'Unsaved Changes',okText:'Discard & Leave',danger:true});
    return;
  }
  _doSwitchCgTab(tab);
}
function _doSwitchCgTab(tab){
  ['overview','info','clients','notes','docs','audit'].forEach(function(t){
    var tb=document.getElementById('cgtab-'+t);
    var pn=document.getElementById('cgpane-'+t);
    if(tb)tb.classList.toggle('active',t===tab);
    if(pn)pn.classList.toggle('active',t===tab);
  });
  if(tab==='overview')renderCgOverviewPane();
  if(tab==='info')renderCgInfoPane();
  if(tab==='clients')renderCgClientsPane();
  if(tab==='notes')renderCgNotesPane();
  if(tab==='docs'){
    renderCgDocsPane();
    // Signatures used to have its own tab; now they live at the bottom of Documents.
    // Kick off the fetch so the list populates when the tab opens.
    if(typeof loadCgSigningRequests==='function')loadCgSigningRequests();
  }
  if(tab==='audit')renderCgAuditPane();
}
function renderCgOverviewPane(){
  if(!activeCgId)return;
  var cg=getCaregivers()[activeCgId];
  var pane=document.getElementById('cgpane-overview');
  if(!cg||!pane)return;
  var profiles=getProfiles();
  // Active assignments only — terminated clients shouldn't count toward caregiver's current workload
  var assigned=Object.keys(profiles).filter(function(k){return profiles[k].caregiverId===activeCgId && (profiles[k].clientStatus||'active')==='active';});
  var addrStr=(cg.street||cg.address||'').trim();
  pane.innerHTML='<div class="overview-grid">'+
    '<div class="ov-card"><h4>Contact Info</h4>'+
      (cg.phone?'<div class="ov-row"><span class="ov-label">Phone</span><span class="ov-value">'+esc(cg.phone)+'</span></div>':'')+
      (cg.email?'<div class="ov-row"><span class="ov-label">Email</span><span class="ov-value">'+esc(cg.email)+'</span></div>':'')+
      (addrStr?'<div class="ov-row"><span class="ov-label">Address</span><span class="ov-value">'+esc(addrStr)+'</span></div>':'')+
    '</div>'+
    '<div class="ov-card"><h4>Employment</h4>'+
      (cg.emptype?'<div class="ov-row"><span class="ov-label">Type</span><span class="ov-value">'+esc(cg.emptype)+'</span></div>':'')+
      (cg.hireDate?'<div class="ov-row"><span class="ov-label">Hire Date</span><span class="ov-value">'+esc(cg.hireDate)+'</span></div>':'')+
      (cg.payRate?'<div class="ov-row"><span class="ov-label">Pay Rate</span><span class="ov-value">$'+esc(cg.payRate)+'/hr</span></div>':'')+
    '</div>'+
    '<div class="ov-card"><h4>Assigned Clients</h4>'+
      '<div class="ov-row" style="cursor:pointer;" onclick="switchCgTab(\'clients\')">'+
        '<span class="ov-label">Clients</span>'+
        '<span class="ov-value" style="color:#185FA5;font-size:18px;font-weight:700;">'+assigned.length+'</span>'+
      '</div>'+
      (assigned.length?'<div style="font-size:12px;color:#4a6a8a;margin-top:6px;">'+assigned.slice(0,3).map(function(n){return esc(n);}).join(', ')+(assigned.length>3?' + '+(assigned.length-3)+' more':'')+'</div>':'<div style="font-size:12px;color:#5c7590;">No clients assigned.</div>')+
    '</div>'+
  '</div>';
}
function renderCgNotesPane(){
  var cg=getCaregivers()[activeCgId];
  var c=document.getElementById('cgNotesContent');
  if(!c||!cg)return;
  c.innerHTML='<div style="margin-bottom:6px;font-size:11px;color:#5c7590;">Auto-saves as you type <span id="cgNotesSavedFlash" style="display:none;color:#1a7740;font-weight:600;">· Saved ✓</span></div>'+
    '<textarea id="cgNotesArea" style="width:100%;min-height:200px;padding:12px;border:1px solid #d0d8e4;border-radius:6px;font-size:13px;font-family:Arial,sans-serif;outline:none;resize:vertical;max-width:620px;">'+esc(cg.notes||'')+'</textarea>';
  var ta=document.getElementById('cgNotesArea');
  var t=null;
  ta.addEventListener('input',function(){
    clearTimeout(t);
    t=setTimeout(function(){
      var cgs=getCaregivers();
      if(!cgs[activeCgId])return;
      cgs[activeCgId].notes=ta.value;
      saveCaregiversLS(cgs);
      // D8: only claim "Saved ✓" after the API resolves; surface failure otherwise.
      var doSave=function(){return flashQuietSave(saveCaregiverAPI(activeCgId,cgs[activeCgId],true),'cgNotesSavedFlash','Caregiver note',doSave);};
      doSave();
    },600);
  });
}
function renderCgAuditPane(){
  var pane=document.getElementById('cgpane-audit');
  var cg=getCaregivers()[activeCgId];
  if(!pane||!cg)return;
  pane.innerHTML='<div style="color:#5c7590;font-size:13px;padding:8px 0;">Loading…</div>';
  if(spToken){
    fetch(API_BASE+'/audit/search',{method:'POST',headers:apiHeaders(),body:JSON.stringify({client:cg.name||activeCgId,limit:100})})
      .then(function(r){return r.ok?r.json():Promise.reject();})
      .then(function(rows){
        pane.innerHTML='';
        if(!rows.length){pane.innerHTML='<div style="color:#5c7590;font-size:13px;padding:16px 0;">No audit entries yet.</div>';return;}
        var wrap=document.createElement('div');wrap.style.cssText='max-width:600px;';
        rows.forEach(function(e){
          var row=document.createElement('div');row.className='audit-row';
          var ts=e.created_at?new Date(e.created_at).toLocaleString():e.ts||'';
          row.innerHTML='<span class="audit-icon">—</span><div><div class="audit-text">'+esc(e.action)+'</div><div class="audit-who">'+esc(e.who)+' · '+esc(ts)+'</div></div>';
          wrap.appendChild(row);
        });
        pane.appendChild(wrap);
      })
      .catch(function(){pane.innerHTML='<div style="color:#5c7590;font-size:13px;padding:16px 0;">No audit entries yet.</div>';});
  } else {
    pane.innerHTML='<div style="color:#5c7590;font-size:13px;padding:16px 0;">Sign in to view audit log.</div>';
  }
}
function renderCgInfoPane(){
  if(!activeCgId)return;
  var cg=getCaregivers()[activeCgId];
  var c=document.getElementById('cgInfoContent');c.innerHTML='';
  if(!cg)return;
  // Re-rendering the pane means the prior unsaved state is being replaced by a fresh form load
  cgUnsavedChanges=false;
  // Capture any subsequent input/change as "unsaved" — delegated listener on the pane
  c.oninput=function(){cgUnsavedChanges=true;};
  c.onchange=function(){cgUnsavedChanges=true;};

  var g=document.createElement('div');g.className='info-grid';
  c.appendChild(g);

  function mkF(id,label,val,full){
    var d=document.createElement('div');d.className='info-field'+(full?' full':'');
    d.innerHTML='<label for="'+id+'">'+label+'</label><input id="'+id+'" value="'+esc(val||'')+'">';
    g.appendChild(d);
  }
  function mkDiv(label){
    var d=document.createElement('div');d.className='form-section-divider full';d.innerHTML='<span>'+label+'</span>';g.appendChild(d);
  }
  function mkRow(children){
    var d=document.createElement('div');d.className='info-field-row full';d.innerHTML=children;g.appendChild(d);
  }

  var dName=document.createElement('div');dName.className='info-field-row full';dName.style.gridTemplateColumns='1fr 1fr 1fr';
  var storedFirst=cg.firstName||cg.first_name||(cg.name||'').split(' ')[0]||'';
  var storedLast=cg.lastName||cg.last_name||(cg.name||'').split(' ').slice(1).join(' ')||'';
  dName.innerHTML='<div class="info-field"><label for="cgi-first">First Name *</label><input id="cgi-first" value="'+esc(storedFirst)+'"></div>'+
    '<div class="info-field"><label for="cgi-middle">Middle Name</label><input id="cgi-middle" value="'+esc(cg.middleName||cg.middle_name||'')+'"></div>'+
    '<div class="info-field"><label for="cgi-last">Last Name *</label><input id="cgi-last" value="'+esc(storedLast)+'"></div>';
  g.appendChild(dName);

  // DOB + Gender row (moved up — needed for state forms / identity)
  mkRow('<div class="info-field"><label for="cgi-dob">Date of Birth <span style="font-weight:400;font-size:11px;color:#5c7590;">(double-click to copy)</span></label><input id="cgi-dob" type="date" value="'+esc(cg.dob||cg.dateOfBirth||'')+'" autocomplete="off" ondblclick="_copyField(this,\'mdy\')" title="Double-click to copy as MM/DD/YYYY"></div>'+
    '<div class="info-field"><label for="cgi-gender">Gender</label><select id="cgi-gender"><option value=""'+(!cg.gender?' selected':'')+'>—</option><option value="Male"'+(cg.gender==='Male'?' selected':'')+'>Male</option><option value="Female"'+(cg.gender==='Female'?' selected':'')+'>Female</option></select></div>');

  mkRow('<div class="info-field"><label for="cgi-nickname">Nickname</label><input id="cgi-nickname" value="'+esc(cg.nickname||'')+'"></div>'+
    '<div class="info-field"><label for="cgi-status">Status</label><select id="cgi-status"><option value="active"'+((!cg.status||cg.status==="active")?" selected":"")+'>Active</option><option value="inactive"'+(cg.status==="inactive"?" selected":"")+'>Inactive</option><option value="terminated"'+(cg.status==="terminated"?" selected":"")+'>Terminated</option></select></div>');

  mkF('cgi-phone','Phone',cg.phone,false);
  mkF('cgi-email','Email',cg.email,false);
  mkF('cgi-dl',"Driver's License #",cg.driversLicense,false);
  // SSN masked; reveals on focus, re-masks a few seconds after blur (#9). Double-click copies
  // the 9 digits with no dashes (#6).
  var cgSsnDiv=document.createElement('div');cgSsnDiv.className='info-field';
  cgSsnDiv.innerHTML='<label for="cgi-ssn">Social Security # <span style="font-weight:400;font-size:11px;color:#5c7590;">(double-click to copy, no dashes)</span></label>'+
    '<input id="cgi-ssn" type="password" autocomplete="off" maxlength="11" placeholder="'+(cg.ssnLast4?('•••-••-'+esc(cg.ssnLast4)+' — click to show'):'XXX-XX-XXXX')+'" value="'+esc(cg.ssn||'')+'" oninput="formatSSN(this);" onfocus="_revealCgSsnField(this,activeCgId)" onblur="_maskSecretSoon(this)" ondblclick="_copyField(this,\'digits\')" title="Click to reveal; double-click to copy without dashes">';
  g.appendChild(cgSsnDiv);
  mkF('cgi-street','Street',cg.street||cg.address,true);
  mkRow('<div class="info-field"><label for="cgi-city">City</label><input id="cgi-city" value="'+esc(cg.city||'')+'"></div>'+
    '<div class="info-field"><label for="cgi-state">State</label><input id="cgi-state" value="'+esc(cg.state||'')+'"></div>');
  mkRow('<div class="info-field"><label for="cgi-zip">ZIP</label><input id="cgi-zip" value="'+esc(cg.zip||'')+'" oninput="lookupZip(\'cgi-zip\',\'cgi-city\',\'cgi-state\',\'cgi-county\')"></div>'+
    '<div class="info-field"><label for="cgi-county">County</label><input id="cgi-county" value="'+esc(cg.county||'')+'"></div>');

  mkRow('<div class="info-field full"><label for="cgi-champs">CHAMPS Provider ID <span style="font-weight:400;color:#5c7590;">(used on BPHASA-2421)</span></label><input id="cgi-champs" value="'+esc(cg.champsId||cg.champs_id||'')+'" placeholder="e.g. 6221933"></div>');
  mkRow(
    '<div class="info-field"><label for="cgi-milogin-user">MI Login Username <span style="font-weight:400;color:#5c7590;">(state portal)</span></label><input id="cgi-milogin-user" value="'+esc(cg.miloginUsername||cg.milogin_username||'')+'" placeholder="username" autocomplete="off"></div>'+
    '<div class="info-field"><label for="cgi-milogin-pass">MI Login Password <span style="font-weight:400;font-size:11px;color:#5c7590;">(click to reveal)</span></label>'+
      '<input id="cgi-milogin-pass" type="password" value="" placeholder="•••••• (click to reveal)" autocomplete="off" onfocus="_revealMiloginField(this,\''+escJsAttr(activeCgId)+'\')" onblur="_maskSecretSoon(this)" style="width:100%;">'+
    '</div>'
  );

  mkDiv('Employment');
  mkRow('<div class="info-field"><label for="cgi-hire">Hire Date <span style="font-weight:400;font-size:11px;color:#5c7590;">(double-click to copy)</span></label><input id="cgi-hire" type="date" value="'+esc(cg.hireDate||'')+'" ondblclick="_copyField(this,\'mdy\')" title="Double-click to copy as MM/DD/YYYY"></div>'+
    '<div class="info-field"><label for="cgi-emptype">Employment Type</label><select id="cgi-emptype"><option value="full-time"'+(cg.emptype==='full-time'?' selected':'')+'>Full-Time</option><option value="part-time"'+(cg.emptype==='part-time'?' selected':'')+'>Part-Time</option><option value="per-diem"'+(cg.emptype==='per-diem'?' selected':'')+'>Per Diem</option></select></div>');
  mkRow('<div class="info-field"><label for="cgi-pay">Pay Rate ($/hr)</label><input id="cgi-pay" value="'+esc(cg.payRate||'')+'" inputmode="decimal" oninput="formatRate(this)"></div>');

  var actions=document.createElement('div');actions.style.cssText='margin-top:16px;display:flex;gap:8px;';
  actions.innerHTML='<button class="btn btn-primary" id="cgSaveInfoBtn" onclick="saveCgInfoPane()">Save Changes</button>'+
    '<button class="btn btn-danger btn-sm" onclick="deleteCaregiverFromDetail()" style="padding:6px 14px;">Delete Caregiver</button>';
  c.appendChild(actions);
  wireCopyableFields('cgInfoContent');
}
function saveCgInfoPane(){
  if(!activeCgId)return;
  var cgs=getCaregivers();var cg=cgs[activeCgId];if(!cg)return;
  var first=(document.getElementById('cgi-first').value||'').trim();
  var last=(document.getElementById('cgi-last').value||'').trim();
  if(!first||!last){showAlert('First and last name are required.');return;}
  cg.firstName=first;cg.middleName=document.getElementById('cgi-middle').value;cg.lastName=last;
  cg.name=_cgDisplayName(first,document.getElementById('cgi-middle').value,last);
  cg.nickname=document.getElementById('cgi-nickname').value;
  cg.status=document.getElementById('cgi-status').value;
  cg.phone=document.getElementById('cgi-phone').value;cg.email=document.getElementById('cgi-email').value;
  var cgiDl=document.getElementById('cgi-dl');if(cgiDl)cg.driversLicense=cgiDl.value;
  // Only take the SSN when the field actually holds one. An untouched (never-focused) field is empty
  // now that the roster doesn't carry the SSN — writing that through would blank the stored value.
  // saveCaregiverAPI omits an empty ssn and the backend keeps what it has.
  var cgiSsn=document.getElementById('cgi-ssn');if(cgiSsn&&cgiSsn.value.trim())cg.ssn=cgiSsn.value;
  cg.street=document.getElementById('cgi-street').value;cg.city=document.getElementById('cgi-city').value;
  cg.state=document.getElementById('cgi-state').value;cg.zip=document.getElementById('cgi-zip').value;cg.county=document.getElementById('cgi-county').value;
  var cgiDob=document.getElementById('cgi-dob');if(cgiDob)cg.dob=cgiDob.value;
  var cgiGender=document.getElementById('cgi-gender');if(cgiGender)cg.gender=cgiGender.value;
  cg.hireDate=document.getElementById('cgi-hire').value;cg.emptype=document.getElementById('cgi-emptype').value;
  cg.payRate=document.getElementById('cgi-pay').value;
  var cgiChamps=document.getElementById('cgi-champs');if(cgiChamps)cg.champsId=cgiChamps.value;
  var cgiMiu=document.getElementById('cgi-milogin-user');if(cgiMiu)cg.miloginUsername=cgiMiu.value;
  var cgiMip=document.getElementById('cgi-milogin-pass');if(cgiMip)cg.miloginPassword=cgiMip.value;
  saveCaregiversLS(cgs);saveCaregiverAPI(activeCgId,cg);
  document.getElementById('cgDetailName').textContent=cg.name;
  var st=cg.status||'active';
  document.getElementById('cgDetailMeta').innerHTML=esc(cg.emptype||'')+(cg.payRate?' · $'+cg.payRate+'/hr':'')+' &nbsp;<span class="cs-badge cs-'+st+'">'+st.charAt(0).toUpperCase()+st.slice(1)+'</span>';
  var btn=document.getElementById('cgSaveInfoBtn');if(btn){btn.textContent='Saved ✓';setTimeout(function(){btn.textContent='Save Changes';},1800);}
  addAuditEntry(cg.name,'Caregiver profile updated');
  cgUnsavedChanges=false;
  renderCaregiverGrid();
}
function renderCgClientsPane(){
  if(!activeCgId)return;
  var c=document.getElementById('cgClientsContent');
  var profiles=getProfiles();
  // Active clients only — keeps the caregiver's current workload view clean
  var assigned=Object.keys(profiles).filter(function(k){return profiles[k].caregiverId===activeCgId && (profiles[k].clientStatus||'active')==='active';});
  if(!assigned.length){c.innerHTML='<div class="empty-state"><h3>No clients assigned</h3><p style="font-size:13px;">Assign this caregiver from the client\'s Profile tab.</p></div>';return;}
  c.innerHTML='';
  assigned.forEach(function(name){
    var prof=profiles[name];
    var ini=name.split(' ').map(function(w){return w[0]||'';}).slice(0,2).join('').toUpperCase();
    var row=document.createElement('div');
    row.style.cssText='display:flex;align-items:center;gap:12px;padding:11px 14px;background:#fff;border:1px solid #e1e5ea;border-radius:8px;margin-bottom:8px;cursor:pointer;transition:border-color 0.15s;max-width:500px;';
    var liveBadge=prof.liveIn?'<span style="display:inline-block;background:#fff3cd;color:#856404;font-size:10px;font-weight:600;padding:2px 8px;border-radius:10px;border:1px solid #ffeaa7;margin-left:6px;">LIVE-IN</span>':'';
    row.innerHTML='<div class="cc-avatar" style="width:36px;height:36px;font-size:13px;">'+ini+'</div>'+
      '<div style="flex:1;min-width:0;">'+
        '<div style="font-size:13px;font-weight:600;color:#1a2b45;">'+esc(name)+liveBadge+'</div>'+
        '<div style="font-size:11px;color:#4d6c88;">'+esc(prof.medicaidId||'No Medicaid ID')+(prof.phone?' · '+esc(prof.phone):'')+'</div>'+
      '</div>'+
      '<span style="font-size:11px;color:#185FA5;font-weight:500;">Open →</span>';
    row.addEventListener('mouseenter',function(){this.style.borderColor='#b0c8e8';});
    row.addEventListener('mouseleave',function(){this.style.borderColor='#e1e5ea';});
    row.addEventListener('click',function(){navDetail(name);});
    c.appendChild(row);
  });
}
function renderCgDocsPane(){
  if(!activeCgId)return;
  var cgName=getCaregivers()[activeCgId]?getCaregivers()[activeCgId].name:'Caregiver';
  var c=document.getElementById('cgDocsContent');
  c.innerHTML=docUploaderHtml({title:'Documents for '+esc(cgName),subtitle:"SSN card, driver's license, certifications, etc.",hasCategory:true,catId:'cgDocCategory',fileId:'cgDocFileInput2',scanId:'cgDocScanInput',scanFn:'handleCgDocScan(this)',uploadFn:'uploadCgDocAzure()',statusId:'cgDocUploadStatus',listId:'cgDocListAzure'});
  loadCgDocsAzure(activeCgId);
}
function loadCgDocsAzure(cgId){
  fetch(API_BASE+'/documents?clientType=caregiver&clientId='+cgId,{headers:apiHeaders()})
  .then(function(r){ if(!r.ok)throw new Error('HTTP '+r.status); return r.json(); })
  .then(function(docs){ renderCgDocListAzure(cgId, Array.isArray(docs)?docs:[]); })
  .catch(function(err){ _renderDocLoadError('cgDocListAzure', "loadCgDocsAzure('"+cgId+"')", err); });
}
function renderCgDocListAzure(cgId,docs){
  renderDocGrid(document.getElementById('cgDocListAzure'),docs,{clientType:'caregiver',clientId:cgId,deleteExpr:function(name){return 'deleteCgDocAzure(\''+cgId+'\',\''+encodeURIComponent(name)+'\')';},refresh:function(){loadCgDocsAzure(cgId);}});
}
function uploadCgDocAzure(){
  var input=document.getElementById('cgDocFileInput2');
  if(!input||!input.files||!input.files.length){showAlert('Please select a file first.');return;}
  var cat=document.getElementById('cgDocCategory').value||'Other';
  var status=document.getElementById('cgDocUploadStatus');status.textContent='Uploading...';
  var fd=new FormData();fd.append('clientType','caregiver');fd.append('clientId',activeCgId);fd.append('category',cat);
  Array.from(input.files).forEach(function(f){
    var prefixedFile=new File([f],cat+'__'+f.name,{type:f.type});
    fd.append('file',prefixedFile);
  });
  fetch(API_BASE+'/documents',{method:'POST',headers:authUploadHeaders(),body:fd})
  .then(function(r){if(!r.ok)throw new Error('Upload failed ('+r.status+')');return r;})
  .then(function(){status.textContent='';input.value='';loadCgDocsAzure(activeCgId);})
  .catch(function(e){status.textContent='Upload failed: '+e;});
}
function deleteCgDocAzure(cgId,encodedName){
  showConfirm('Delete this document?',function(){
    fetch(API_BASE+'/documents?clientType=caregiver&clientId='+cgId,{method:'DELETE',headers:apiHeaders(),body:JSON.stringify({name:decodeURIComponent(encodedName)})})
    .then(function(r){if(!r.ok)throw new Error('HTTP '+r.status);loadCgDocsAzure(cgId);}).catch(function(e){showAlert('Delete failed: '+e);});
  },{title:'Delete Document',okText:'Delete'});
}
function handleCgDocScan(input){
  if(!activeCgId){showAlert('Open a caregiver first.');return;}
  if(!input||!input.files||!input.files.length)return;
  var cat=(document.getElementById('cgDocCategory')&&document.getElementById('cgDocCategory').value)||'Other';
  var status=document.getElementById('cgDocUploadStatus');
  if(status)status.textContent='Uploading scanned image…';
  var fd=new FormData();fd.append('clientType','caregiver');fd.append('clientId',activeCgId);fd.append('category',cat);
  var f=input.files[0];
  var prefixedFile=new File([f],cat+'__'+f.name,{type:f.type});
  fd.append('file',prefixedFile);
  fetch(API_BASE+'/documents',{method:'POST',headers:authUploadHeaders(),body:fd})
  .then(function(r){if(!r.ok)throw new Error('Upload failed ('+r.status+')');return r;})
  .then(function(){if(status)status.textContent='';input.value='';loadCgDocsAzure(activeCgId);})
  .catch(function(e){if(status)status.textContent='Upload failed: '+e;});
}

// ============================================================
//  SETTINGS
// ============================================================
// ── Signatures: DB-backed, LS used as a local display cache ──
function getSigs(){try{return JSON.parse(localStorage.getItem('lhca_signatures')||'[]');}catch(e){return[];}}
function saveSigsLS(arr){try{localStorage.setItem('lhca_signatures',JSON.stringify(arr));}catch(e){}}
function sigId(){return 'sig_'+Date.now()+'_'+Math.random().toString(36).slice(2,6);}
// Merge a freshly-loaded id-array over the local one WITHOUT dropping unsynced local additions —
// the same cold-device race that hit the rosters applies to tasks and signatures. Server wins for
// shared ids; a local-only item is kept UNLESS it carries `syncedProp` (it was previously confirmed
// by the server and is now gone → deleted elsewhere → drop). Pass no syncedProp to always keep
// local-only items (signatures have no per-row sync marker).
function _mergeByIdKeepUnsynced(serverArr, localArr, syncedProp){
  var have={}; (serverArr||[]).forEach(function(x){ if(x&&x.id!=null) have[x.id]=true; });
  var out=(serverArr||[]).slice();
  (localArr||[]).forEach(function(x){
    if(!x || x.id==null || have[x.id]) return;
    if(syncedProp && x[syncedProp]) return;   // was synced, now absent from server → deleted elsewhere
    out.push(x);
  });
  return out;
}
function loadSignaturesAPI(){
  if(!spToken)return;
  fetch(API_BASE+'/signatures',{headers:apiHeaders()})
    .then(function(r){return r.ok?r.json():Promise.reject(r.status);})
    .then(function(rows){
      // rows: [{id,label,data_url,created_at}]
      var sigs=(Array.isArray(rows)?rows:[]).map(function(r){return {id:r.id,label:r.label,data:r.data_url};});
      // Merge so a signature saved on this device while the load was in flight isn't wiped.
      saveSigsLS(_mergeByIdKeepUnsynced(sigs, getSigs()));
      renderSigSettings();
    })
    .catch(function(e){console.error('Signatures load error:',e);});
}
function renderSigSettings(){
  var sigs=getSigs(),list=document.getElementById('sigStoredList');list.innerHTML='';
  if(!sigs.length){list.innerHTML='<div style="color:#5c7590;font-size:13px;">No signatures saved yet.</div>';return;}
  sigs.forEach(function(s,i){
    var item=document.createElement('div');item.className='sig-stored-item';
    // Build via DOM API so quoted IDs don't break the onclick attribute
    var img=document.createElement('img');img.className='sig-stored-img';img.src=s.data;img.alt='sig';
    var lbl=document.createElement('span');lbl.className='sig-stored-label';lbl.textContent=s.label||'Signature '+(i+1);
    var btn=document.createElement('button');btn.className='btn btn-danger btn-sm';btn.textContent='Remove';
    (function(target){btn.addEventListener('click',function(){deleteSig(target);});})(s.id||i);
    item.appendChild(img);item.appendChild(lbl);item.appendChild(btn);
    list.appendChild(item);
  });
}
// In-website confirm modal — replaces native browser confirm() for cleaner UX
function showConfirm(message,onConfirm,opts){
  opts=opts||{};
  var modal=document.getElementById('confirmModal');
  document.getElementById('confirmTitle').textContent=opts.title||'Confirm';
  document.getElementById('confirmMessage').textContent=message||'Are you sure?';
  var okBtn=document.getElementById('confirmOkBtn'),cancelBtn=document.getElementById('confirmCancelBtn'),extraBtn=document.getElementById('confirmExtraBtn');
  // Alert mode: hide cancel button (used by showAlert)
  cancelBtn.style.display=opts.hideCancel?'none':'';
  // If okText is explicitly empty, hide the OK button entirely (e.g., extra-only choice)
  if(opts.okText===''||opts.okText===null){
    okBtn.style.display='none';
  } else {
    okBtn.style.display='';
    okBtn.textContent=opts.okText||'Confirm';
    okBtn.className='btn '+(opts.danger!==false?'btn-danger':'btn-primary');
  }
  cancelBtn.textContent=opts.cancelText||'Cancel';
  // Optional third button (e.g., 'Send Anyway')
  if(opts.extraText){
    extraBtn.textContent=opts.extraText;
    extraBtn.style.display='';
    extraBtn.className='btn '+(opts.extraDanger?'btn-danger':'btn-secondary');
  } else {
    extraBtn.style.display='none';
  }
  // Replace handlers freshly each time
  var newOk=okBtn.cloneNode(true),newCancel=cancelBtn.cloneNode(true),newExtra=extraBtn.cloneNode(true);
  okBtn.parentNode.replaceChild(newOk,okBtn);
  cancelBtn.parentNode.replaceChild(newCancel,cancelBtn);
  extraBtn.parentNode.replaceChild(newExtra,extraBtn);
  newOk.addEventListener('click',function(){modal.classList.remove('open');if(typeof onConfirm==='function')onConfirm();});
  newCancel.addEventListener('click',function(){modal.classList.remove('open');if(typeof opts.onCancel==='function')opts.onCancel();});
  newExtra.addEventListener('click',function(){modal.classList.remove('open');if(typeof opts.onExtra==='function')opts.onExtra();});
  modal.classList.add('open');
}

// Single-button alert modal — replaces native showAlert()
function showAlert(message,opts){
  opts=opts||{};
  showConfirm(message,function(){if(typeof opts.onOK==='function')opts.onOK();},{title:opts.title||'Notice',okText:opts.okText||'OK',danger:!!opts.danger,hideCancel:true});
}

// In-website prompt modal — replaces native browser prompt() for cleaner UX
function showPrompt(message,initialValue,onSave,opts){
  opts=opts||{};
  var modal=document.getElementById('promptModal');
  document.getElementById('promptTitle').textContent=opts.title||'Enter value';
  document.getElementById('promptMessage').textContent=message||'';
  var inp=document.getElementById('promptInput');
  inp.value=initialValue||'';
  var okBtn=document.getElementById('promptOkBtn'),cancelBtn=document.getElementById('promptCancelBtn');
  okBtn.textContent=opts.okText||'Save';
  var newOk=okBtn.cloneNode(true),newCancel=cancelBtn.cloneNode(true);
  okBtn.parentNode.replaceChild(newOk,okBtn);cancelBtn.parentNode.replaceChild(newCancel,cancelBtn);
  newOk.addEventListener('click',function(){modal.classList.remove('open');if(typeof onSave==='function')onSave(inp.value);});
  newCancel.addEventListener('click',function(){modal.classList.remove('open');});
  modal.classList.add('open');
  setTimeout(function(){inp.focus();inp.select();},50);
}

// Fetch a caregiver's MI Login password on demand (it is never preloaded or cached),
// then reveal it via the normal mask toggle (which auto-re-masks after 8s).
// ── Secret-field reveal / copy helpers (SSN, MI Login password, dates) ──────────
// Reveal a masked field and cancel any pending auto re-mask.
function _revealSecret(inp){ if(!inp)return; if(inp._maskTimeout){clearTimeout(inp._maskTimeout);inp._maskTimeout=null;} inp.type='text'; }
// Re-mask a few seconds AFTER the field loses focus — not instantly — so you can alt-tab to
// another window/tab and paste the value before it snaps back to dots (#9). Cancelled if the
// field is re-focused (via _revealSecret) before the timer fires.
function _maskSecretSoon(inp,ms){ if(!inp)return; if(inp._maskTimeout)clearTimeout(inp._maskTimeout); inp._maskTimeout=setTimeout(function(){ if(inp&&inp.type==='text')inp.type='password'; if(inp)inp._maskTimeout=null; },ms||4000); }
// MI Login password: fetch on FOCUS (never preloaded/cached) then reveal — replaces the old
// "Show" button (#8). Empty-on-save is safe (backend pwProvided guard preserves the stored value).
// Click into the SSN field and it fetches the real value, then reveals it — same shape as MI Login,
// no button. The roster load only carries the last 4, so an unfetched field shows a •••• placeholder
// rather than the number; focusing it is what pulls the SSN, and that read is audited server-side.
function _revealCgSsnField(inp,id){
  if(!inp)return;
  _revealSecret(inp);
  if(inp.value||!id)return;                 // already fetched, or the user is typing a new one
  if(inp.dataset.fetching==='1')return;
  inp.dataset.fetching='1';
  fetch(API_BASE+'/caregivers/'+encodeURIComponent(id)+'/ssn',{headers:apiHeaders()})
    .then(function(r){return r.ok?r.json():null;})
    .then(function(d){
      if(d&&typeof d.ssn==='string'&&!inp.value){
        inp.value=d.ssn;
        try{ var cgs=getCaregivers(); if(cgs[id]){ cgs[id].ssn=d.ssn; saveCaregiversLS(cgs); } }catch(e){}
      }
    })
    .catch(function(){})
    .then(function(){ inp.dataset.fetching=''; });
}
function _revealMiloginField(inp,id){
  if(!inp)return;
  _revealSecret(inp);
  if(inp.value||!id)return;                 // already fetched / user typed one / new caregiver
  fetch(API_BASE+'/caregivers/'+encodeURIComponent(id)+'/milogin',{headers:apiHeaders()})
    .then(function(r){return r.ok?r.json():null;})
    .then(function(d){ if(d&&typeof d.milogin_password==='string'&&!inp.value)inp.value=d.milogin_password; })
    .catch(function(){});
}
// Copy a field's value to the clipboard for quick paste into state portals/forms, with a brief
// "Copied ✓" flash. transform: 'digits' → strip non-digits (SSN without dashes, #6);
// 'mdy' → YYYY-MM-DD to MM/DD/YYYY (date inputs, #7). Wired to ondblclick.
function _copyField(el,transform){
  if(!el)return;
  var v=(el.value!=null&&el.value!=='')?el.value:(el.textContent||'');
  if(transform==='digits') v=String(v).replace(/\D/g,'');
  else if(transform==='mdy'){ var m=String(v).match(/^(\d{4})-(\d{2})-(\d{2})$/); if(m)v=m[2]+'/'+m[3]+'/'+m[1]; }
  v=String(v).trim(); if(!v)return;
  try{ if(navigator.clipboard&&navigator.clipboard.writeText)navigator.clipboard.writeText(v); }catch(e){}
  _flashCopied(el);
}
function _flashCopied(el){
  try{
    var tip=document.createElement('span'); tip.textContent='Copied ✓';
    tip.style.cssText='position:absolute;background:#1a7740;color:#fff;font-size:11px;font-weight:600;padding:2px 7px;border-radius:4px;z-index:99999;pointer-events:none;box-shadow:0 1px 4px rgba(0,0,0,.2);';
    var r=el.getBoundingClientRect();
    tip.style.left=(r.left+window.scrollX)+'px'; tip.style.top=(r.top+window.scrollY-24)+'px';
    document.body.appendChild(tip); setTimeout(function(){if(tip&&tip.parentNode)tip.parentNode.removeChild(tip);},1000);
  }catch(e){}
}
function revealMilogin(inputId,btn,cgIdArg){
  var inp=document.getElementById(inputId);if(!inp)return;
  // Already revealed, or the user has typed a value — just toggle visibility.
  if(inp.type==='text'||inp.value){toggleMask(inputId,btn);return;}
  var idEl=document.getElementById('cg-editing-id');
  var id=cgIdArg||(idEl&&idEl.value)||activeCgId||'';
  if(!id){toggleMask(inputId,btn);return;} // new caregiver — nothing stored yet
  var orig=btn?btn.textContent:'';if(btn)btn.textContent='…';
  fetch(API_BASE+'/caregivers/'+encodeURIComponent(id)+'/milogin',{headers:apiHeaders()})
    .then(function(r){return r.ok?r.json():null;})
    .then(function(d){
      if(d&&typeof d.milogin_password==='string')inp.value=d.milogin_password;
      if(btn)btn.textContent=orig||'Show';
      toggleMask(inputId,btn);
    })
    .catch(function(){if(btn)btn.textContent=orig||'Show';});
}
function toggleMask(inputId,btn){
  var inp=document.getElementById(inputId);if(!inp)return;
  if(inp.type==='password'){
    inp.type='text';
    if(btn){btn.textContent='Hide';btn.classList.remove('btn-secondary');btn.classList.add('btn-danger');}
    // Auto re-mask after 8 seconds for safety
    if(inp._maskTimeout)clearTimeout(inp._maskTimeout);
    inp._maskTimeout=setTimeout(function(){
      if(inp.type==='text'){
        inp.type='password';
        if(btn){btn.textContent='Show';btn.classList.remove('btn-danger');btn.classList.add('btn-secondary');}
      }
    },8000);
  } else {
    inp.type='password';
    if(btn){btn.textContent='Show';btn.classList.remove('btn-danger');btn.classList.add('btn-secondary');}
    if(inp._maskTimeout){clearTimeout(inp._maskTimeout);inp._maskTimeout=null;}
  }
}

function deleteSig(idOrIdx){
  showConfirm('Remove this signature? This cannot be undone.',function(){doDeleteSig(idOrIdx);},{title:'Remove Signature',okText:'Remove'});
}
function doDeleteSig(idOrIdx){
  var sigs=getSigs();
  // idOrIdx may be a string ID or a legacy numeric index
  var sigId_=typeof idOrIdx==='number'?null:(idOrIdx||null);
  if(sigId_){
    fetch(API_BASE+'/signatures/'+encodeURIComponent(sigId_),{method:'DELETE',headers:apiHeaders()})
      .catch(function(e){console.error('Sig delete error:',e);});
    saveSigsLS(sigs.filter(function(s){return s.id!==sigId_;}));
  } else {
    sigs.splice(idOrIdx,1);saveSigsLS(sigs);
  }
  renderSigSettings();
}
function openAddSigModal(){
  document.getElementById('sigLabel').value='';
  var tn=document.getElementById('sigTypeName');if(tn)tn.value='';
  window._sigUploadOriginal=null;window._sigUploadProcessed=null;
  var uf=document.getElementById('sigUploadFile');if(uf)uf.value='';
  var uh=document.getElementById('sigUploadHint');if(uh)uh.textContent='Click to choose file or drag & drop';
  var uc=document.getElementById('sigUploadCanvas');if(uc){var uctx=uc.getContext('2d');uctx.clearRect(0,0,uc.width,uc.height);uctx.fillStyle='#fff';uctx.fillRect(0,0,uc.width,uc.height);}
  document.getElementById('sigModal').classList.add('open');
  if(!sigCanvas)initSigCanvas();
  clearSigPad();
  switchSigTab('draw');
}
function updateSettingsAuth(){
  var s=document.getElementById('settingsAuthStatus'),b=document.getElementById('settingsAuthBtn');
  if(spToken){if(s){s.textContent='✓ Signed in — Azure SQL sync & email active';s.style.color='#1e7e34';}if(b){b.textContent='Sign Out';b.onclick=signOut;}}
  else{if(s){s.textContent='Authentication required';s.style.color='#c0392b';}if(b){b.textContent='Sign In';b.onclick=signIn;}}
}
function clearAllDrafts(){
  showConfirm('Clear all draft invoices? Saved invoices are not affected.',function(){
    Object.keys(localStorage).filter(function(k){return k.startsWith('lhca_draft_');}).forEach(function(k){localStorage.removeItem(k);});
    showConfirm('All drafts cleared.',function(){},{title:'Done',okText:'OK',danger:false});
  },{title:'Clear Drafts',okText:'Clear All Drafts'});
}
function clearAllData(){
  showConfirm(
    'This clears the LOCAL browser cache (clients, caregivers, caseworkers, drafts, signatures, tasks, audit log).\n\n'+
    'IMPORTANT: This does NOT delete data from the Azure SQL database. After you reload, the app will re-download everything from the cloud.\n\n'+
    'Use this only when you want to force a fresh sync — for example, if local data looks out-of-date or corrupted.',
    function(){
      showConfirm('Last chance — clear local cache now?\n\n(Server data is safe; it will re-download on next reload.)',function(){
        var keys=['lhca_profiles','lhca_caregivers','lhca_caseworkers','lhca_signatures','lhca_sig','lhca_todos','lhca_email_audit','lhca_activity','lhca_id_map','lhca_last_synced'];
        keys=keys.concat(Object.keys(localStorage).filter(function(k){return k.startsWith('lhca_draft_');}));
        keys.forEach(function(k){try{localStorage.removeItem(k);}catch(e){}});
        showConfirm('Local cache cleared. Reload the page now to re-sync from the cloud.',function(){location.reload();},{title:'Cache Cleared',okText:'Reload Now',danger:false});
      },{title:'Confirm Clear',okText:'Clear Local Cache'});
    },
    {title:'Clear Local Cache',okText:'Continue',danger:false}
  );
}

// ============================================================
//  EMAIL AUTOCOMPLETE
// ============================================================

// ============================================================
//  ZIP LOOKUP
// ============================================================
function lookupZip(zipId, cityId, stateId, countyId) {
  var zip = document.getElementById(zipId).value.trim();
  if (zip.length !== 5 || !/^\d{5}$/.test(zip)) return;
  fetch('https://api.zippopotam.us/us/' + zip)
    .then(function(r){ return r.ok ? r.json() : Promise.reject(); })
    .then(function(data){
      if (data.places && data.places.length) {
        var p = data.places[0];
        var cEl = document.getElementById(cityId); if(cEl) cEl.value = p['place name'] || '';
        var sEl = document.getElementById(stateId); if(sEl) sEl.value = p['state abbreviation'] || '';
        var coEl = document.getElementById(countyId);
        // FCC block/find requires latitude+longitude; the &zip= variant doesn't work.
        // Zippopotam already gave us lat/lon so we use that.
        if (coEl && p.latitude && p.longitude) {
          fetch('https://geo.fcc.gov/api/census/block/find?format=json&latitude='+p.latitude+'&longitude='+p.longitude+'&showall=true')
            .then(function(r2){return r2.ok?r2.json():Promise.reject();})
            .then(function(d2){
              var countyName = d2 && d2.County && d2.County.name;
              if (countyName && coEl) coEl.value = countyName.replace(/ County$/i,'').trim();
            })
            .catch(function(e){console.warn('County lookup failed:',e);});
        }
      }
    })
    .catch(function(e){console.warn('ZIP lookup failed:',e);});
}

// ============================================================
//  CASEWORKER SEARCH AUTOCOMPLETE
// ============================================================
function cwSearch(input, hiddenId, dropId) {
  var val = input.value.trim().toLowerCase();
  var drop = document.getElementById(dropId);
  var hidden = document.getElementById(hiddenId);
  var cws = getCaseworkers().slice().sort(function(a,b){return (a.name||'').localeCompare(b.name||'');});
  // When empty: show ALL caseworkers (scrollable). When typing: filter.
  var matches = val ? cws.filter(function(c){ return (c.name||'').toLowerCase().includes(val); }) : cws.slice();
  drop.innerHTML = '';
  matches.slice(0, 50).forEach(function(c) {
    var item = document.createElement('div');
    item.style.cssText = 'padding:8px 12px;cursor:pointer;font-size:13px;border-bottom:1px solid #f0f3f7;';
    item.textContent = c.name + (c.agency ? ' — ' + c.agency : '');
    item.addEventListener('mousedown', function(e){ e.preventDefault(); input.value = c.name; if(hidden)hidden.value = c.id; drop.style.display = 'none'; });
    item.addEventListener('mouseover', function(){ this.style.background='#f0f5fb'; });
    item.addEventListener('mouseout', function(){ this.style.background=''; });
    drop.appendChild(item);
  });
  // "Add new" option only when typing AND no exact match exists
  if (val && !matches.some(function(c){ return (c.name||'').toLowerCase() === val; })) {
    var addItem = document.createElement('div');
    addItem.style.cssText = 'padding:8px 12px;cursor:pointer;font-size:13px;color:#185FA5;font-weight:600;';
    addItem.textContent = '+ Add "' + input.value.trim() + '" as new caseworker';
    addItem.addEventListener('mousedown', function(e){
      e.preventDefault();
      drop.style.display = 'none';
      openQuickCreate('caseworker', input.value.trim(), function(id, name){
        input.value = name; if(hidden) hidden.value = id;
      });
    });
    drop.appendChild(addItem);
  }
  if(!matches.length && !val){
    var noItem=document.createElement('div');noItem.style.cssText='padding:8px 12px;font-size:13px;color:#5c7590;';noItem.textContent='No caseworkers yet';drop.appendChild(noItem);
  }
  drop.style.display = 'block';
}

// ============================================================
//  CAREGIVER SEARCH AUTOCOMPLETE
// ============================================================
function cgSearch(input, hiddenId, dropId) {
  var val = input.value.trim().toLowerCase();
  var drop = document.getElementById(dropId);
  var hidden = document.getElementById(hiddenId);
  var cgs = getCaregivers();
  // When empty: show ALL caregivers (scrollable). When typing: filter.
  var ids = Object.keys(cgs).sort(function(a,b){return (cgs[a].name||'').localeCompare(cgs[b].name||'');});
  if (val) ids = ids.filter(function(id){ return (cgs[id].name||'').toLowerCase().includes(val); });
  drop.innerHTML = '';
  // Show up to 50 (dropdown is scrollable beyond initial 180px)
  ids.slice(0, 50).forEach(function(id) {
    var cg = cgs[id];
    var item = document.createElement('div');
    item.style.cssText = 'padding:8px 12px;cursor:pointer;font-size:13px;border-bottom:1px solid #f0f3f7;';
    item.textContent = cg.name + (cg.status !== 'active' ? ' (' + cg.status + ')' : '');
    item.addEventListener('mousedown', function(e){ e.preventDefault(); input.value = cg.name; if(hidden)hidden.value = id; drop.style.display = 'none'; });
    item.addEventListener('mouseover', function(){ this.style.background='#f0f5fb'; });
    item.addEventListener('mouseout', function(){ this.style.background=''; });
    drop.appendChild(item);
  });
  // "+ Add new" option only when typing AND no exact name match exists
  if (val && !ids.some(function(id){ return (cgs[id].name||'').toLowerCase() === val; })) {
    var addItem = document.createElement('div');
    addItem.style.cssText = 'padding:8px 12px;cursor:pointer;font-size:13px;color:#185FA5;font-weight:600;';
    addItem.textContent = '+ Add "' + input.value.trim() + '" as new caregiver';
    addItem.addEventListener('mousedown', function(e){
      e.preventDefault();
      drop.style.display = 'none';
      openQuickCreate('caregiver', input.value.trim(), function(id, name){
        input.value = name; if(hidden) hidden.value = id;
      });
    });
    drop.appendChild(addItem);
  }
  if (!ids.length && !val) {
    var noItem = document.createElement('div');
    noItem.style.cssText = 'padding:8px 12px;font-size:13px;color:#5c7590;';
    noItem.textContent = 'No caregivers yet';
    drop.appendChild(noItem);
  }
  drop.style.display = 'block';
}

// Inline "quick create" popup for a new caregiver/caseworker from an assignment
// field. Captures just the essentials (name + contact); the full record is
// filled in later from their own page. On save, calls onCreated(id, name).
function openQuickCreate(kind, prefillName, onCreated){
  var isCg = kind === 'caregiver';
  var ov = document.createElement('div');
  ov.className = 'modal-overlay open';
  ov.innerHTML =
    '<div class="modal-box" style="max-width:380px;">'+
      '<h3>New ' + (isCg ? 'Caregiver' : 'Caseworker') + '</h3>'+
      '<p>Just the essentials — add the rest later on their page.</p>'+
      '<label class="qc-l" for="qc-name">Name</label><input id="qc-name" class="qc-i" maxlength="80" value="'+esc(prefillName||'')+'">'+
      (isCg ? '' : '<label class="qc-l" for="qc-agency">Agency</label><input id="qc-agency" class="qc-i" maxlength="80">')+
      '<label class="qc-l" for="qc-phone">Phone</label><input id="qc-phone" class="qc-i" maxlength="20" placeholder="(555) 555-5555">'+
      '<label class="qc-l" for="qc-email">Email</label><input id="qc-email" class="qc-i" maxlength="80" placeholder="name@email.com">'+
      '<div class="modal-row">'+
        '<button class="btn btn-primary" id="qc-save">Create &amp; select</button>'+
        '<button class="btn btn-secondary" id="qc-cancel">Cancel</button>'+
      '</div>'+
    '</div>';
  document.body.appendChild(ov);
  var close = function(){ if(ov.parentNode) ov.parentNode.removeChild(ov); };
  ov.querySelector('#qc-cancel').addEventListener('click', close);
  ov.addEventListener('mousedown', function(e){ if(e.target === ov) close(); });
  ov.querySelector('#qc-save').addEventListener('click', function(){
    var name = (ov.querySelector('#qc-name').value||'').trim();
    if(!name){ ov.querySelector('#qc-name').focus(); return; }
    var phone = (ov.querySelector('#qc-phone').value||'').trim();
    var email = (ov.querySelector('#qc-email').value||'').trim();
    var newId;
    if(isCg){
      newId = cgId();
      var cg = { id:newId, name:name, phone:phone, email:email, status:'active' };
      var all = getCaregivers(); all[newId] = cg; saveCaregiversLS(all);
      if(typeof saveCaregiverAPI === 'function') saveCaregiverAPI(newId, cg);
    } else {
      newId = cwId();
      var agency = (ov.querySelector('#qc-agency').value||'').trim();
      var cw = { id:newId, name:name, first_name:'', last_name:'', agency:agency, phone:phone, email:email, fax:'', street:'', city:'', state:'', zip:'', county:'', notes:'' };
      var arr = getCaseworkers(); arr.push(cw); saveCaseworkersLS(arr);
      if(typeof saveCaseworkerAPI === 'function') saveCaseworkerAPI(cw);
    }
    close();
    if(typeof onCreated === 'function') onCreated(newId, name);
  });
  setTimeout(function(){ var n = ov.querySelector('#qc-name'); if(n){ n.focus(); n.setSelectionRange(n.value.length, n.value.length); } }, 30);
}

// ============================================================
//  ACTIVITY LOG
// ============================================================
function getActivity(){try{return JSON.parse(localStorage.getItem('lhca_activity')||'[]');}catch(e){return[];}}
function logActivity(type,text){
  var log=getActivity();
  log.unshift({type:type,text:text,ts:new Date().toLocaleString()});
  if(log.length>40)log=log.slice(0,40);
  try{localStorage.setItem('lhca_activity',JSON.stringify(log));}catch(e){}
  _postAuditRecord({event_type:type,client_name:'',action:text,who:currentUserEmail()});
  renderActivityFeed();
}
function renderActivityFeed(){
  var list=document.getElementById('activityList');if(!list)return;
  var log=getActivity();
  if(!log.length){list.innerHTML='<div class="af-empty">No recent activity.</div>';return;}
  var icons={invoice:'inv',edit:'edit',client:'client',caregiver:'cg',status:'status',delete:'del'};
  list.innerHTML='';
  log.slice(0,10).forEach(function(e){
    var item=document.createElement('div');item.className='af-item';
    // Try to extract client name from activity text for navigation
    var clientLink='';
    var profiles=getProfiles();
    Object.keys(profiles).forEach(function(name){
      if(e.text.indexOf(name)!==-1&&!clientLink)clientLink=name;
    });
    var clickable=clientLink&&(e.type==='invoice'||e.type==='edit'||e.type==='status'||e.type==='client');
    item.style.cssText=clickable?'cursor:pointer;':'';
    item.innerHTML='<span class="af-icon">'+(icons[e.type]||'')+'</span>'+
      '<div style="flex:1;"><div class="af-text"'+(clickable?' style="color:#185FA5;"':'')+'>'+esc(e.text)+'</div><div class="af-time">'+esc(e.ts)+'</div></div>'+
      (clickable?'<span style="font-size:11px;color:#185FA5;">→</span>':'');
    if(clickable){
      (function(n,t){item.addEventListener('click',function(){
        navDetail(n);
        // If it's an invoice activity, also open the invoices tab
        if(t==='invoice'||t==='status'){setTimeout(function(){switchTab('history');},50);}
      });})(clientLink,e.type);
    }
    list.appendChild(item);
  });
}
// ============================================================
//  AUDIT TRAIL
// ============================================================
function currentUserEmail(){
  return (msalInstance&&msalInstance.getAllAccounts().length?msalInstance.getAllAccounts()[0].username:null)||'Local User';
}
// Only this account may generate invoices or send the monthly emails.
var INVOICE_ADMIN_EMAIL='tommy@mybellcare.com';
function isInvoiceAdmin(){return (currentUserEmail()||'').toLowerCase()===INVOICE_ADMIN_EMAIL;}
// Hide the Generate Invoices + Monthly Emails controls for everyone else.
function applyInvoiceAdminVisibility(){
  var admin=isInvoiceAdmin();
  ['genInvoicesBtn','sb-monthly'].forEach(function(id){var e=document.getElementById(id);if(e)e.style.display=admin?'':'none';});
}
function getAuditLog(){try{return JSON.parse(localStorage.getItem('lhca_audit')||'[]');}catch(e){return[];}}
function addAuditEntry(clientName,action){
  var who=currentUserEmail();
  var log=getAuditLog();
  log.unshift({client:clientName,action:action,who:who,ts:new Date().toLocaleString()});
  if(log.length>200)log=log.slice(0,200);
  try{localStorage.setItem('lhca_audit',JSON.stringify(log));}catch(e){}
  // Persist to DB. This is the HIPAA §164.312(b) record, so a failure must be VISIBLE: it used to
  // be fire-and-forget with no r.ok check, and was skipped entirely when the token had expired —
  // both of which stop the trail silently while the app carries on.
  _postAuditRecord({event_type:'audit',client_name:clientName,action:action,who:who});
}
// Shared writer for both audit paths. Surfaces a failure through the save-status toast rather than
// swallowing it, and says plainly when the record could not be written at all.
function _postAuditRecord(body){
  if(typeof spToken==='undefined'||!spToken){
    if(typeof _showSaveStatus==='function')_showSaveStatus('failed','Not signed in — this action was NOT written to the audit log.');
    return;
  }
  fetch(API_BASE+'/audit',{method:'POST',headers:apiHeaders(),body:JSON.stringify(body)})
    .then(function(r){
      if(!r.ok){
        console.error('Audit save failed: HTTP '+r.status);
        if(typeof _showSaveStatus==='function')_showSaveStatus('failed','Audit log write failed ('+r.status+') — this action was not recorded.');
      }
    })
    .catch(function(e){
      console.error('Audit save error:',e);
      if(typeof _showSaveStatus==='function')_showSaveStatus('failed','Audit log write failed — this action was not recorded.');
    });
}
function renderAuditPane(){
  if(!activeProfileName)return;
  var pane=document.getElementById('dpane-audit');pane.innerHTML='';
  if(spToken){
    pane.innerHTML='<div style="color:#5c7590;font-size:13px;padding:8px 0;">Loading audit history…</div>';
    fetch(API_BASE+'/audit/search',{method:'POST',headers:apiHeaders(),body:JSON.stringify({client:activeProfileName,limit:100})})
      .then(function(r){return r.ok?r.json():Promise.reject(r.status);})
      .then(function(rows){
        pane.innerHTML='';
        if(!rows.length){pane.innerHTML='<div style="color:#5c7590;font-size:13px;padding:16px 0;">No audit entries for this client yet.</div>';return;}
        var wrap=document.createElement('div');wrap.style.cssText='max-width:600px;';
        rows.forEach(function(e){
          var row=document.createElement('div');row.className='audit-row';
          var ts=e.created_at?new Date(e.created_at).toLocaleString():e.ts||'';
          row.innerHTML='<span class="audit-icon">—</span><div><div class="audit-text">'+esc(e.action)+'</div><div class="audit-who">'+esc(e.who)+' · '+esc(ts)+'</div></div>';
          wrap.appendChild(row);
        });
        pane.appendChild(wrap);
      })
      .catch(function(){
        // Fallback to local cache
        _renderAuditPaneLocal(pane);
      });
  } else {
    _renderAuditPaneLocal(pane);
  }
}
function _renderAuditPaneLocal(pane){
  var log=getAuditLog().filter(function(e){return e.client===activeProfileName;});
  if(!log.length){pane.innerHTML='<div style="color:#5c7590;font-size:13px;padding:16px 0;">No audit entries for this client yet.</div>';return;}
  var wrap=document.createElement('div');wrap.style.cssText='max-width:600px;';
  log.forEach(function(e){
    var row=document.createElement('div');row.className='audit-row';
    row.innerHTML='<span class="audit-icon">—</span><div><div class="audit-text">'+esc(e.action)+'</div><div class="audit-who">'+esc(e.who)+' · '+esc(e.ts)+'</div></div>';
    wrap.appendChild(row);
  });
  pane.appendChild(wrap);
}

// ============================================================
//  TASKS / TODOS
// ============================================================
function getTodos(){try{return JSON.parse(localStorage.getItem('lhca_todos')||'[]');}catch(e){return[];}}
function saveTodos(t){localStorage.setItem('lhca_todos',JSON.stringify(t));}
function todoId(){return 'td_'+Date.now()+'_'+Math.random().toString(36).slice(2,6);}
function addTodo(){
  var text=document.getElementById('todoText').value.trim();if(!text)return;
  var client=document.getElementById('todoClient').value;
  var due=document.getElementById('todoDue').value;
  var note=document.getElementById('todoNote')?document.getElementById('todoNote').value.trim():'';
  var todos=getTodos();
  todos.unshift({id:todoId(),text:text,client:client,due:due,note:note,done:false,created:new Date().toLocaleString()});
  saveTodos(todos);saveTaskAPI(todos[0]);
  document.getElementById('todoText').value='';
  document.getElementById('todoDue').value='';
  if(document.getElementById('todoNote'))document.getElementById('todoNote').value='';
  renderTodos();updateTaskBadge();
}
// Add a follow-up task that comes AFTER an existing task (linked via parentId)
// Generic show task edit/create modal — handles name, due date, and note
function showTaskEditModal(opts){
  opts=opts||{};
  var modal=document.getElementById('taskEditModal');
  document.getElementById('taskEditTitle').textContent=opts.title||'Edit Task';
  document.getElementById('taskEditSubtitle').textContent=opts.subtitle||'';
  var nameInp=document.getElementById('taskEditName');
  var dueInp=document.getElementById('taskEditDue');
  var noteInp=document.getElementById('taskEditNote');
  var clientSel=document.getElementById('taskEditClient');
  nameInp.value=opts.name||'';
  dueInp.value=opts.due||'';
  noteInp.value=opts.note||'';
  if(clientSel){
    var profs=getProfiles();
    // Any status — tasks get attached to In Progress / onboarding clients too, not just active
    var names=Object.keys(profs).sort(function(a,b){return a.localeCompare(b);});
    clientSel.innerHTML='<option value="">— No client —</option>'+names.map(function(n){return '<option value="'+esc(n)+'"'+(n===(opts.client||'')?' selected':'')+'>'+esc(n)+'</option>';}).join('');
  }
  var saveBtn=document.getElementById('taskEditSaveBtn');
  var cancelBtn=document.getElementById('taskEditCancelBtn');
  saveBtn.textContent=opts.okText||'Save';
  // Replace handlers freshly each open
  var newSave=saveBtn.cloneNode(true),newCancel=cancelBtn.cloneNode(true);
  saveBtn.parentNode.replaceChild(newSave,saveBtn);
  cancelBtn.parentNode.replaceChild(newCancel,cancelBtn);
  newSave.addEventListener('click',function(){
    var name=(nameInp.value||'').trim();if(!name){nameInp.focus();return;}
    modal.classList.remove('open');
    if(typeof opts.onSave==='function')opts.onSave({name:name,due:dueInp.value||'',note:(noteInp.value||'').trim(),client:clientSel?clientSel.value:''});
  });
  newCancel.addEventListener('click',function(){modal.classList.remove('open');});
  modal.classList.add('open');
  setTimeout(function(){nameInp.focus();nameInp.select();},50);
}

// Edit an existing task — name, due date, and note all in one modal
function editTodoTask(id){
  var todos=getTodos();var t=todos.find(function(x){return String(x.id)===String(id);});if(!t)return;
  showTaskEditModal({
    title:'Edit Task',
    subtitle:t.parentId?'(follow-up task)':'',
    name:t.text||'',
    due:t.due||'',
    note:t.note||'',
    client:t.client||'',
    okText:'Save Changes',
    onSave:function(vals){
      var todos2=getTodos();var t2=todos2.find(function(x){return String(x.id)===String(id);});if(!t2)return;
      t2.text=vals.name;t2.due=vals.due||'';t2.note=vals.note||'';t2.client=vals.client||'';
      saveTodos(todos2);saveTaskAPI(t2);renderTodos();updateTaskBadge();
    }
  });
}

function addFollowUpTask(parentId){
  var todos=getTodos();var parent=todos.find(function(x){return String(x.id)===String(parentId);});if(!parent)return;
  showTaskEditModal({
    title:'Add Follow-up Task',
    subtitle:'Comes after: '+parent.text,
    name:'',
    due:'',
    note:'',
    okText:'Add Follow-up',
    onSave:function(vals){
      var todos2=getTodos();
      var idx=todos2.findIndex(function(x){return String(x.id)===String(parentId);});
      var newTask={id:todoId(),text:vals.name,client:parent.client||'',due:vals.due||'',note:vals.note||'',parentId:String(parentId),done:false,created:new Date().toLocaleString()};
      if(idx>=0)todos2.splice(idx+1,0,newTask);else todos2.unshift(newTask);
      saveTodos(todos2);saveTaskAPI(newTask);renderTodos();updateTaskBadge();
    }
  });
}
function toggleTodo(id){
  // Use loose equality so legacy numeric IDs from DB still match string IDs from inline onclick
  var todos=getTodos(),t=todos.find(function(x){return String(x.id)===String(id);});
  if(t){t.done=!t.done;t.doneAt=t.done?new Date().toLocaleString():null;}
  saveTodos(todos);if(t) saveTaskAPI(t);renderTodos();updateTaskBadge();
}
function deleteTodo(id){
  // Use string-coerced comparison so legacy DB tasks (numeric IDs) match too
  var toDelete=getTodos().find(function(x){return String(x.id)===String(id);});
  if(!toDelete)return;
  var followUps=getTodos().filter(function(x){return String(x.parentId||'')===String(id);});
  var msg='Remove this task?';
  if(followUps.length)msg='Remove this task AND its '+followUps.length+' follow-up'+(followUps.length>1?'s':'')+'?';
  showConfirm(msg,function(){
    var idsToRemove={};idsToRemove[String(id)]=true;followUps.forEach(function(t){idsToRemove[String(t.id)]=true;});
    var remaining=getTodos().filter(function(x){return !idsToRemove[String(x.id)];});
    saveTodos(remaining);
    if(toDelete.dbId)deleteTaskAPI(toDelete.dbId);
    followUps.forEach(function(t){if(t.dbId)deleteTaskAPI(t.dbId);});
    renderTodos();updateTaskBadge();
  },{title:'Remove Task',okText:'Remove'});
}
function clearDoneTasks(){
  showConfirm('Remove all completed tasks?',function(){
    var all=getTodos();
    var done=all.filter(function(t){return t.done;});
    var keep=all.filter(function(t){return !t.done;});
    saveTodos(keep);
    // Also delete each completed task from the backend so they don't reappear on next load
    done.forEach(function(t){if(t.dbId)deleteTaskAPI(t.dbId);});
    renderTodos();updateTaskBadge();
  },{title:'Clear Completed Tasks',okText:'Remove All'});
}
function populateTodoClientSelect(){
  var sel=document.getElementById('todoClient');if(!sel)return;
  var cur=sel.value;sel.innerHTML='<option value="">— Link to client —</option>';
  Object.keys(getProfiles()).forEach(function(name){
    var o=document.createElement('option');o.value=name;o.textContent=name;if(name===cur)o.selected=true;sel.appendChild(o);
  });
}
function addTaskForClient(clientName){
  showTaskEditModal({
    title:'Add Task',
    subtitle:'For client: '+clientName,
    name:'',
    due:'',
    note:'',
    okText:'Add Task',
    onSave:function(vals){
      var todos=getTodos();
      todos.unshift({id:todoId(),text:vals.name,client:clientName,due:vals.due||'',note:vals.note||'',done:false,created:new Date().toLocaleString()});
      saveTodos(todos);saveTaskAPI(todos[0]);updateTaskBadge();
      renderOverviewPane();
      logActivity('client','Task added for '+clientName+': '+vals.name);
    }
  });
}
function renderTodos(filterClient){
  var todos=getTodos();
  if(filterClient)todos=todos.filter(function(t){return t.client===filterClient;});
  var open=todos.filter(function(t){return !t.done;});
  var done=todos.filter(function(t){return t.done;});
  var now=new Date();
  var oc=document.getElementById('openTaskCount');if(oc)oc.textContent='('+open.length+')';
  var dc=document.getElementById('doneTaskCount');if(dc)dc.textContent='('+done.length+')';

  // OPEN TASKS — render parents flat, with collapsible follow-ups under each
  var openC=document.getElementById('todoList');
  if(openC){
    openC.innerHTML='';
    if(!open.length){openC.innerHTML='<div style="color:#5c7590;font-size:13px;padding:8px 0;">No open tasks.</div>';}
    else {
      // Group: parents (no parentId) + follow-ups attached to them
      var parents=open.filter(function(t){return !t.parentId;});
      var followUpsByParent={};
      open.filter(function(t){return t.parentId;}).forEach(function(t){
        if(!followUpsByParent[t.parentId])followUpsByParent[t.parentId]=[];
        followUpsByParent[t.parentId].push(t);
      });
      // Orphaned follow-ups (parent already done/removed) — promote to top level
      var allParentIds={};parents.forEach(function(t){allParentIds[t.id]=true;});
      Object.keys(followUpsByParent).forEach(function(pid){
        if(!allParentIds[pid]){followUpsByParent[pid].forEach(function(t){t.parentId=null;parents.push(t);});delete followUpsByParent[pid];}
      });

      function renderTaskRow(t,opts){
        opts=opts||{};
        var overdue=t.due&&new Date(t.due)<now;
        var isFollowUp=!!opts.followUp;
        var item=document.createElement('div');item.className='todo-item';
        var indentStyle=isFollowUp?'margin-left:32px;':'';
        item.style.cssText='background:'+(isFollowUp?'#fafbfc':'#f5f8fc')+';border:1px solid #e1e5ea;border-left:3px solid '+(isFollowUp?'#b0c8e8':'#185FA5')+';border-radius:6px;padding:10px;margin-bottom:6px;'+indentStyle;
        item.innerHTML=
          '<input type="checkbox" class="todo-cb" onchange="toggleTodo(\''+t.id+'\')" style="margin-top:3px;">'+
          '<div class="todo-body">'+
            (isFollowUp?'<div style="font-size:10px;color:#5c7590;margin-bottom:2px;">↳ Follow-up</div>':'')+
            '<div class="todo-text" style="font-weight:600;">'+esc(t.text)+'</div>'+
            (t.note?'<div style="font-size:11px;color:#4a6a8a;margin-top:4px;padding:4px 6px;background:#e8f0fb;border-radius:3px;white-space:pre-wrap;">'+esc(t.note)+'</div>':'')+
            '<div class="todo-meta">'+
              (t.due?'<span class="'+(overdue?'todo-due-overdue':'')+'">Due: '+t.due+'</span>':'')+
              (!filterClient&&t.client?'<span class="todo-link" onclick="navDetail(\''+escJsAttr(t.client)+'\')">'+esc(t.client)+'</span>':'')+
            '</div>'+
          '</div>'+
          '<button class="btn btn-secondary btn-sm" onclick="editTodoTask(\''+t.id+'\')" style="padding:3px 7px;font-size:11px;" title="Edit task name, due date, or note">Edit</button>'+
          '<button class="btn btn-secondary btn-sm" onclick="addFollowUpTask(\''+(isFollowUp?(t.parentId||t.id):t.id)+'\')" style="padding:3px 7px;font-size:11px;" title="Add a follow-up that comes after this one">+ Follow-up</button>'+
          '<button class="btn btn-danger btn-sm" onclick="deleteTodo(\''+t.id+'\')" style="padding:3px 7px;font-size:11px;">Remove</button>';
        return item;
      }

      parents.forEach(function(parent){
        var fups=followUpsByParent[parent.id]||[];
        var parentRow=renderTaskRow(parent);
        // If has follow-ups, append a small toggle bar inside the parent row
        if(fups.length){
          var toggleBar=document.createElement('div');
          toggleBar.style.cssText='margin-top:8px;padding:8px 12px;background:#eef3fb;border:1px solid #d0dbe9;border-radius:5px;cursor:pointer;font-size:12px;color:#185FA5;font-weight:600;display:flex;justify-content:space-between;align-items:center;user-select:none;width:100%;transition:background 0.15s;';
          var toggleLabel=document.createElement('span');
          toggleLabel.style.cssText='display:flex;align-items:center;gap:8px;';
          toggleLabel.innerHTML='<span class="fup-arrow" style="display:inline-block;font-size:14px;transition:transform 0.18s;">▶</span><span>'+fups.length+' follow-up'+(fups.length>1?'s':'')+'</span>';
          var clearBtn=document.createElement('span');clearBtn.style.cssText='font-size:11px;color:#5c7590;font-weight:400;';clearBtn.textContent='click to expand';
          toggleBar.appendChild(toggleLabel);toggleBar.appendChild(clearBtn);
          toggleBar.addEventListener('mouseenter',function(){this.style.background='#dde8f5';});
          toggleBar.addEventListener('mouseleave',function(){this.style.background='#eef3fb';});
          // Insert toggle bar at end of parentRow body
          parentRow.querySelector('.todo-body').appendChild(toggleBar);
          openC.appendChild(parentRow);
          // Container for follow-ups (collapsed by default)
          var fupContainer=document.createElement('div');fupContainer.style.cssText='display:none;';
          fups.forEach(function(fu){fupContainer.appendChild(renderTaskRow(fu,{followUp:true}));});
          openC.appendChild(fupContainer);
          toggleBar.addEventListener('click',function(){
            var isOpen=fupContainer.style.display!=='none';
            fupContainer.style.display=isOpen?'none':'block';
            var arrow=toggleLabel.querySelector('.fup-arrow');
            if(arrow)arrow.style.transform=isOpen?'rotate(0deg)':'rotate(90deg)';
            clearBtn.textContent=isOpen?'click to expand':'click to collapse';
          });
        } else {
          openC.appendChild(parentRow);
        }
      });
    }
  }
  // DONE TASKS — compact
  var doneC=document.getElementById('todoListDone');
  if(doneC){
    doneC.innerHTML='';
    if(!done.length){doneC.innerHTML='<div style="color:#5c7590;font-size:13px;padding:8px 0;">Nothing completed yet.</div>';return;}
    done.forEach(function(t){
      var row=document.createElement('div');row.className='todo-item';row.style.opacity='0.6';
      row.innerHTML='<input type="checkbox" class="todo-cb" checked onchange="toggleTodo(\''+t.id+'\')">'+
        '<div class="todo-body"><div class="todo-text done">'+esc(t.text)+'</div>'+
        (t.doneAt?'<div class="todo-meta"><span>Completed: '+esc(t.doneAt)+'</span></div>':'')+
        '</div>'+
        '<button class="btn btn-danger btn-sm" onclick="deleteTodo(\''+t.id+'\')" style="padding:3px 6px;font-size:10px;">Remove</button>';
      doneC.appendChild(row);
    });
  }
}
function updateTaskBadge(){
  var overdue=getTodos().filter(function(t){
    return !t.done&&t.due&&new Date(t.due)<new Date();
  }).length;
  var badge=document.getElementById('taskBadge');
  if(badge){badge.style.display=overdue?'inline':'none';badge.textContent=overdue;}
}

// ============================================================
//  WORKFLOW
// ============================================================
var wfSteps=[];
function openWorkflowModal(){
  wfSteps=[];
  document.getElementById('wfStepsList').innerHTML='';
  document.getElementById('wfTemplate').value='';
  var sel=document.getElementById('wfClientSelect');
  sel.innerHTML='<option value="">— Link workflow to client (optional) —</option>';
  Object.keys(getProfiles()).forEach(function(n){var o=document.createElement('option');o.value=n;o.textContent=n;sel.appendChild(o);});
  if(activeProfileName){sel.value=activeProfileName;}
  document.getElementById('workflowModal').classList.add('open');
}
function loadWfTemplate(){
  var tpl=document.getElementById('wfTemplate').value;
  if(!tpl)return;
  var templates={
    invoice:[
      {text:'Prepare invoice',daysFromNow:0},
      {text:'Fill out service grid',daysFromNow:0},
      {text:'Review and sign invoice',daysFromNow:1},
      {text:'Submit invoice to MDHHS',daysFromNow:2},
      {text:'Confirm invoice received by worker',daysFromNow:5},
      {text:'Follow up if no payment by end of month',daysFromNow:28}
    ],
    newclient:[
      {text:'Collect intake form',daysFromNow:0},
      {text:'Verify Medicaid eligibility',daysFromNow:1},
      {text:'Assign caregiver',daysFromNow:2},
      {text:'Schedule first service visit',daysFromNow:3},
      {text:'Confirm MDHHS worker assigned',daysFromNow:3},
      {text:'Enter client into CRM',daysFromNow:0},
      {text:'File signed intake form',daysFromNow:7}
    ],
    custom:[]
  };
  wfSteps=(templates[tpl]||[]).map(function(s,i){
    var d=new Date();d.setDate(d.getDate()+s.daysFromNow);
    return {text:s.text,due:d.toISOString().slice(0,10),id:'wf_'+i};
  });
  renderWfSteps();
}
function renderWfSteps(){
  var list=document.getElementById('wfStepsList');list.innerHTML='';
  wfSteps.forEach(function(s,i){
    var row=document.createElement('div');row.className='wf-step';
    row.style.cssText=i===0?'':'opacity:0.7;';
    row.innerHTML='<div class="wf-step-num">'+(i+1)+'</div>'+
      '<input type="text" value="'+esc(s.text)+'" placeholder="Step description…" onchange="wfSteps['+i+'].text=this.value" style="'+(i===0?'':'background:#f8f9fa;')+'">' +
      '<input type="date" value="'+esc(s.due)+'" onchange="wfSteps['+i+'].due=this.value">'+
      '<button class="wf-step-del" onclick="wfSteps.splice('+i+',1);renderWfSteps()">✕</button>';
    list.appendChild(row);
  });
}
function addWfStep(){
  var d=new Date();d.setDate(d.getDate()+1);
  wfSteps.push({text:'',due:d.toISOString().slice(0,10),id:'wf_'+Date.now()});
  renderWfSteps();
}
function saveWorkflow(){
  var client=document.getElementById('wfClientSelect').value;
  // Filter out empty step rows
  var validSteps=wfSteps.filter(function(s){return s.text&&s.text.trim();});
  if(!validSteps.length){
    showConfirm('Add at least one step before saving the workflow.',function(){},{title:'No Steps',okText:'OK',danger:false});
    return;
  }
  // First step = parent task; remaining steps = follow-ups linked to it via parentId
  var todos=getTodos();
  var newTasks=[];
  var parentId=null;
  validSteps.forEach(function(s,i){
    var tId=todoId();
    var newTask={
      id:tId,
      text:s.text.trim(),
      client:client,
      due:s.due,
      note:'',
      done:false,
      created:new Date().toLocaleString(),
      parentId:i===0?null:parentId
    };
    if(i===0)parentId=tId;
    newTasks.push(newTask);
  });
  // Insert in reverse so the parent ends up at the top, follow-ups directly below
  todos.unshift.apply(todos,newTasks);
  saveTodos(todos);
  newTasks.forEach(function(t){saveTaskAPI(t);});
  updateTaskBadge();
  document.getElementById('workflowModal').classList.remove('open');
  if(document.getElementById('page-tasks').classList.contains('active'))renderTodos();
  if(activeProfileName&&document.getElementById('page-detail').classList.contains('active'))renderOverviewPane();
  logActivity('client','Workflow created'+(client?' for '+client:''));
  var msg=newTasks.length+' task'+(newTasks.length>1?'s':'')+' added'+(client?' for '+client:'')+'.';
  if(newTasks.length>1)msg+='\n\nThe first step is the parent task; the rest are follow-ups under it (collapsed by default).';
  showConfirm(msg,function(){},{title:'Workflow Created',okText:'OK',danger:false});
}

// ============================================================
//  EMAIL COMPOSER (Microsoft Graph)
// ============================================================
function sendGraphEmail(){
  if(!spToken){showAlert('Please sign in with your Microsoft account to send email.');return;}
  var to=document.getElementById('emailTo').value.trim();
  var subj=document.getElementById('emailSubject').value.trim();
  var body=document.getElementById('emailBody').value.trim();
  if(!to||!subj||!body){showAlert('Please fill in To, Subject, and Message.');return;}
  var btn=document.getElementById('emailSendBtn');btn.textContent='Sending…';btn.disabled=true;
  var payload={message:{subject:subj,body:{contentType:'Text',content:body},toRecipients:[{emailAddress:{address:to}}]},saveToSentItems:true};
  fetch('https://graph.microsoft.com/v1.0/me/sendMail',{
    method:'POST',headers:{'Authorization':'Bearer '+spToken,'Content-Type':'application/json'},
    body:JSON.stringify(payload)
  }).then(function(r){
    if(r.ok||r.status===202){
      addAuditEntry(activeProfileName||'—','Email sent to '+to+': '+subj);
      logActivity('edit','Email sent to '+to);
      document.getElementById('emailComposerModal').classList.remove('open');
      btn.textContent='Send Email';btn.disabled=false;
      showAlert('Email sent successfully.');
    } else {
      btn.textContent='Send Email';btn.disabled=false;
      r.text().then(function(t){showAlert('Error sending email: '+r.status+'\n'+t);});
    }
  }).catch(function(e){btn.textContent='Send Email';btn.disabled=false;showAlert('Network error: '+e.message);});
}
function renderReports(){
  var c=document.getElementById('reportsContent');if(!c)return;c.innerHTML='';
  var profiles=getProfiles(),allInvoices=[];
  Object.keys(profiles).forEach(function(name){
    (profiles[name].invoices||[]).forEach(function(inv){
      allInvoices.push({client:name,inv:inv});
    });
  });
  var statusCounts={draft:0,submitted:0,paid:0};
  allInvoices.forEach(function(r){statusCounts[r.inv.status||'draft']++;});
  var s1=document.createElement('div');s1.className='report-section';
  s1.innerHTML='<h3>Invoice Status Summary</h3>'+
    '<table class="report-table"><thead><tr><th>Status</th><th>Count</th><th>Visual</th></tr></thead><tbody>'+
    ['draft','submitted','paid'].map(function(st){
      var n=statusCounts[st],total=allInvoices.length||1;
      var pct=Math.round(n/total*100);
      var color=st==='paid'?'#2a9a5a':st==='submitted'?'#185FA5':'#aaa';
      return '<tr><td style="text-transform:capitalize;">'+st+'</td><td style="font-weight:600;">'+n+'</td>'+
        '<td style="width:200px;"><div class="report-bar"><div class="report-bar-fill" style="width:'+pct+'%;background:'+color+';"></div></div><span style="font-size:10px;color:#5c7590;">'+pct+'%</span></td></tr>';
    }).join('')+
    '<tr class="total-row"><td>Total</td><td>'+allInvoices.length+'</td><td></td></tr></tbody></table>';
  c.appendChild(s1);
  var s2=document.createElement('div');s2.className='report-section';
  var clientRows=Object.keys(profiles).map(function(name){
    var invs=profiles[name].invoices||[];
    var paid=invs.filter(function(i){return i.status==='paid';}).length;
    var open=invs.filter(function(i){return !i.status||i.status!=='paid';}).length;
    return {name:name,total:invs.length,paid:paid,open:open};
  }).sort(function(a,b){return b.total-a.total;});
  s2.innerHTML='<h3>Per-Client Invoice Breakdown</h3>'+
    (clientRows.length?'<table class="report-table"><thead><tr><th>Client</th><th>Total</th><th>Paid</th><th>Open</th></tr></thead><tbody>'+
    clientRows.map(function(r){
      return '<tr><td style="cursor:pointer;color:#185FA5;" onclick="navDetail(\''+escJsAttr(r.name)+'\')">'+esc(r.name)+'</td><td>'+r.total+'</td>'+
        '<td style="color:#1e7e34;font-weight:600;">'+r.paid+'</td><td style="color:'+(r.open?'#b07800':'#aaa')+';font-weight:600;">'+r.open+'</td></tr>';
    }).join('')+'</tbody></table>':'<div style="color:#5c7590;font-size:13px;">No client data yet.</div>');
  c.appendChild(s2);
  var byMonth={};
  allInvoices.forEach(function(r){var bp=r.inv.billingPeriod||'Unknown';byMonth[bp]=(byMonth[bp]||0)+1;});
  var months=Object.keys(byMonth).sort(function(a,b){
    var pa=a.split('/'),pb=b.split('/');
    if(pa.length<2||pb.length<2)return 0;
    return (parseInt(pb[1])*12+parseInt(pb[0]))-(parseInt(pa[1])*12+parseInt(pa[0]));
  }).slice(0,12);
  var s3=document.createElement('div');s3.className='report-section';
  var maxV=months.reduce(function(m,k){return Math.max(m,byMonth[k]);},1);
  s3.innerHTML='<h3>Monthly Invoice Volume <span style="font-size:11px;font-weight:normal;color:#5c7590;">(last 12 months)</span></h3>'+
    (months.length?'<table class="report-table"><thead><tr><th>Period</th><th>Count</th><th>Visual</th></tr></thead><tbody>'+
    months.map(function(m){
      var n=byMonth[m],pct=Math.round(n/maxV*100);
      return '<tr><td style="font-weight:500;">'+esc(m)+'</td><td>'+n+'</td>'+
        '<td style="width:200px;"><div class="report-bar"><div class="report-bar-fill" style="width:'+pct+'%;"></div></div></td></tr>';
    }).join('')+'</tbody></table>':'<div style="color:#5c7590;font-size:13px;">No invoices yet.</div>');
  c.appendChild(s3);
  var s4=document.createElement('div');s4.className='report-section';
  var allPeriods=Object.keys(byMonth).sort(function(a,b){
    var pa=a.split('/'),pb=b.split('/');
    if(pa.length<2||pb.length<2)return 0;
    return (parseInt(pb[1])*12+parseInt(pb[0]))-(parseInt(pa[1])*12+parseInt(pa[0]));
  });
  var allClients=Object.keys(profiles);
  // Default check period: most recent period that has at least 1 invoice, or current month
  var d=new Date();
  var defaultPeriod=String(d.getMonth()+1).padStart(2,'0')+'/'+d.getFullYear();
  if(allPeriods.length)defaultPeriod=allPeriods[0];
  var periodOpts=allPeriods.slice(0,12).map(function(p){return '<option value="'+esc(p)+'"'+(p===defaultPeriod?' selected':'')+'>'+esc(p)+'</option>';}).join('');
  // Also allow typing a period
  var missingNow=allClients.filter(function(name){
    if(!clientDueForInvoice(profiles[name],defaultPeriod))return false;   // #10: only clients actually due (start date on/before the period; excludes carriers/inactive)
    return !((profiles[name].invoices)||[]).some(function(i){return i.billingPeriod===defaultPeriod;});
  });
  s4.innerHTML='<h3>Clients Missing Invoice for Month</h3>'+
    '<div style="display:flex;align-items:center;gap:10px;margin-bottom:12px;flex-wrap:wrap;">'+
      '<label style="font-size:12px;color:#4d6c88;" for="missingPeriodSelect">Check period:</label>'+
      '<select id="missingPeriodSelect" onchange="updateMissingReport()" style="padding:5px 10px;border:1px solid #d0d8e4;border-radius:5px;font-size:12px;font-family:Arial,sans-serif;">'+periodOpts+'</select>'+
      '<input id="missingPeriodCustom" placeholder="or type MM/YYYY" maxlength="7" style="padding:5px 10px;border:1px solid #d0d8e4;border-radius:5px;font-size:12px;font-family:Arial,sans-serif;width:120px;" onblur="updateMissingReport(true)" onkeydown="if(event.key===\'Enter\')updateMissingReport(true)">'+
    '</div>'+
    '<div id="missingReportList"></div>';
  c.appendChild(s4);
  updateMissingReport();
}
function updateMissingReport(useCustom){
  var profiles=getProfiles(),allClients=Object.keys(profiles);
  var period=useCustom?document.getElementById('missingPeriodCustom').value.trim():document.getElementById('missingPeriodSelect').value;
  if(!period||period.length<7)return;
  var missing=allClients.filter(function(name){
    if(!clientDueForInvoice(profiles[name],period))return false;   // start date, carrier AND status (see clientDueForInvoice)
    return !((profiles[name].invoices)||[]).some(function(i){return i.billingPeriod===period;});
  });
  var list=document.getElementById('missingReportList');if(!list)return;
  if(!missing.length){list.innerHTML='<div style="color:#1e7e34;font-size:13px;font-weight:600;">✓ All clients have an invoice for '+esc(period)+'</div>';return;}
  list.innerHTML='<div style="font-size:12px;color:#b07800;font-weight:600;margin-bottom:8px;">'+missing.length+' client'+(missing.length>1?'s':'')+' missing invoice for '+esc(period)+':</div>'+
    missing.map(function(name){
      return '<div style="display:flex;align-items:center;gap:10px;padding:6px 0;border-bottom:1px solid #f0f3f7;">'+
        '<span style="cursor:pointer;color:#185FA5;font-size:13px;" onclick="navDetail(\''+escJsAttr(name)+'\')">'+esc(name)+'</span>'+
        '<button class="btn btn-primary btn-sm" style="font-size:11px;" onclick="activeProfileName=\''+escJsAttr(name)+'\';navInvoice()">+ Create Invoice</button>'+
      '</div>';
    }).join('');
}
function exportReportExcel(){
  var profiles=getProfiles(),rows=[['Client','Medicaid ID','Billing Period','Status','Saved At','Note']];
  Object.keys(profiles).forEach(function(name){
    (profiles[name].invoices||[]).forEach(function(inv){
      rows.push([name,profiles[name].medicaidId||'',inv.billingPeriod||'',inv.status||'draft',inv.savedAt||'',inv.invoiceNote||'']);
    });
  });
  var ws=XLSX.utils.aoa_to_sheet(rows);
  var wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Invoices');
  XLSX.writeFile(wb,'liberty_invoices_'+new Date().toISOString().slice(0,10)+'.xlsx');
}

// ============================================================
//  PROFILES (localStorage + SharePoint)
// ============================================================
// S8: SSN is kept in memory for this session only — never written to localStorage — so a
// lost/idle device or a future XSS can't read it off disk. It's repopulated from the
// authenticated API on load. saveProfilesLS strips ssn before persisting and caches it in
// memory; getProfiles overlays it back for display. Saves send ssn only when present, and
// the backend keeps the stored value when a save omits it (provided-guard).
// Windows this session opened that display PHI (caregiver task sheets). They are separate documents
// the idle/sign-out wipe cannot otherwise reach, so they are tracked and closed with it.
var _phiWindows = [];
var _ssnMem = Object.create(null);
// _clientSynced is the client dirty-tracking baseline, but its signature string EMBEDS the
// plaintext ssn (see _clientSig) — so it must be kept in memory ONLY, never written to disk,
// exactly like _ssnMem. Persisting it would leak SSN into localStorage (defeats S8).
var _clientSyncedMem = Object.create(null);
// Perf: getProfiles() JSON.parses the ENTIRE profiles store (incl. every invoice data blob)
// on each call, and a single navigation calls it ~4×. _profilesCache is a TICK-SCOPED cache:
// it collapses those repeated parses within one synchronous render pass, then self-clears on
// the next microtask — so it can never serve stale data to a later interaction. saveProfilesLS
// also clears it immediately, so a write is always reflected.
var _profilesCache=null;
function getProfiles(){
  if(_profilesCache) return _profilesCache;
  try{
    var p=JSON.parse(localStorage.getItem('lhca_profiles')||'{}');
    for(var name in p){ if(p[name]){
      if(_ssnMem[name]!==undefined) p[name].ssn=_ssnMem[name];
      if(_clientSyncedMem[name]!==undefined) p[name]._clientSynced=_clientSyncedMem[name];
    } }
    _profilesCache=p;
    Promise.resolve().then(function(){_profilesCache=null;});   // clear at end of this sync pass
    return p;
  }catch(e){return{};}
}
function saveProfilesLS(p){
  _profilesCache=null;   // a write invalidates the read cache immediately
  var out={};
  for(var name in p){
    var prof=p[name];
    if(prof && typeof prof==='object'){
      if(prof.ssn) _ssnMem[name]=prof.ssn;               // keep the in-memory copy current
      if(prof._clientSynced!==undefined) _clientSyncedMem[name]=prof._clientSynced; // memory-only (embeds ssn)
      // strip ssn AND _clientSynced from disk — both hold SSN and must never persist
      var clone={}; for(var k in prof){ if(k!=='ssn' && k!=='_clientSynced') clone[k]=prof[k]; }
      out[name]=clone;
    } else { out[name]=prof; }
  }
  try{
    localStorage.setItem('lhca_profiles',JSON.stringify(out));
  }catch(e){
    // Quota exceeded (invoice data blobs + signatures accumulate here). This used to throw straight
    // out of the save, so the caller's API write never ran and the user saw nothing at all.
    console.error('saveProfilesLS failed:',e);
    if(typeof _showSaveStatus==='function')_showSaveStatus('failed','Local storage is full — your change was NOT saved locally. Export a backup and use Settings → Clear Drafts, then try again.');
  }
}
// Perf: debounce the list filter inputs so typing doesn't re-parse the whole dataset and
// rebuild the entire table on every keystroke (~150ms after the last keystroke instead).
function _debounce(fn, ms){ var t; return function(){ var a=arguments, c=this; clearTimeout(t); t=setTimeout(function(){ fn.apply(c,a); }, ms||150); }; }
var filterClients     = _debounce(function(){ if(typeof renderClientTable==='function') renderClientTable(); }, 150);
var filterCaregivers  = _debounce(function(){ if(typeof renderCaregiverGrid==='function') renderCaregiverGrid(); }, 150);
var filterCaseworkers = _debounce(function(){ if(typeof renderCaseworkerList==='function') renderCaseworkerList(); }, 150);
var filterSupervisors = _debounce(function(){ if(typeof renderSupervisorList==='function') renderSupervisorList(); }, 150);
// F4: a note typed then the tab closed within the 600ms debounce was lost — the pending backend
// save never fired and pagehide's PHI wipe erased the LS copy. Register each debounced note save
// here so the pagehide/beforeunload handler can FLUSH it (fire the save while the note is still in
// LS, before the wipe). Best-effort on a hard close, but closes the guaranteed-loss window.
var _pendingNoteSaves = Object.create(null);
function _scheduleNoteSave(key, saveFn){
  var ex=_pendingNoteSaves[key]; if(ex) clearTimeout(ex._t);
  var entry={};
  entry._t=setTimeout(function(){ delete _pendingNoteSaves[key]; saveFn(); }, 600);
  entry.flush=function(){ clearTimeout(entry._t); delete _pendingNoteSaves[key]; saveFn(); };
  _pendingNoteSaves[key]=entry;
}
function _flushPendingNoteSaves(){ Object.keys(_pendingNoteSaves).forEach(function(k){ var e=_pendingNoteSaves[k]; if(e) e.flush(); }); }
function exportProfiles(){
  var p=getProfiles();if(!Object.keys(p).length){showConfirm('No clients yet.',function(){},{title:'Nothing to Export',okText:'OK',danger:false});return;}
  // Ask whether to include plaintext SSN (default: masked)
  showConfirm(
    'Export client data as JSON for backup or transfer between devices.\n\n'+
    'By default, SSN is masked (last 4 digits only). Choose "Include Full SSN" if this backup will be used to restore data — only do this if you will store the file in a secure location (encrypted drive, OneDrive, etc).',
    function(){doExportProfiles(false);},
    {
      title:'Export Clients (JSON)',
      okText:'Export with Masked SSN',
      danger:false,
      extraText:'Include Full SSN',
      extraDanger:true,
      onExtra:function(){doExportProfiles(true);}
    }
  );
}
// ──────────────────────────────────────────────────────────────────
//  AUTOMATIC WEEKLY ONEDRIVE BACKUP
//  Fires once per 7 days on first sign-in, silently uploads to
//  /Liberty Home Care Backups/. Keeps 26 most recent (auto-deletes older).
// ──────────────────────────────────────────────────────────────────
var ONEDRIVE_BACKUP_RETENTION=26;          // keep 26 weekly backups (~6 months)
var ONEDRIVE_BACKUP_INTERVAL_DAYS=7;
function _msSinceLastBackup(){
  var last=localStorage.getItem('lhca_last_onedrive_backup');
  if(!last)return Infinity;
  return Date.now()-new Date(last).getTime();
}
async function maybeAutoBackupOneDrive(){
  if(!spToken)return; // not signed in yet
  if(window._onedriveAutoRunning)return;
  // Wait for DB sync to finish — empty localStorage right after sign-in is normal
  // until loadProfilesAPI completes. Poll up to 60 seconds.
  var waitedMs=0;
  while((window._dbSyncPending>0||!Object.keys(getProfiles()).length)&&waitedMs<60000){
    await new Promise(function(r){setTimeout(r,1000);});
    waitedMs+=1000;
  }
  // If still no clients after 60s, there genuinely is nothing to back up — skip silently
  if(!Object.keys(getProfiles()).length){
    console.log('[OneDrive backup] Skipped — no clients to back up.');
    return;
  }
  var failed=localStorage.getItem('lhca_onedrive_backup_failed')==='1';
  var ms=_msSinceLastBackup();
  var dueByTime=ms>=ONEDRIVE_BACKUP_INTERVAL_DAYS*24*60*60*1000;
  if(!dueByTime&&!failed)return;
  window._onedriveAutoRunning=true;
  try{
    await doBackupToOneDrive(false,/*silent=*/true);
    localStorage.setItem('lhca_last_onedrive_backup',new Date().toISOString());
    localStorage.removeItem('lhca_onedrive_backup_failed');
    localStorage.removeItem('lhca_onedrive_backup_failed_msg');
    await pruneOldOneDriveBackups();
    renderOneDriveBackupBanner('done');
  }catch(e){
    console.error('Auto-backup failed:',e);
    // Soft-skip "No clients yet" — it's not really a failure
    if(/no clients yet/i.test(e.message||'')){
      console.log('[OneDrive backup] No clients to back up — skipping silently.');
      window._onedriveAutoRunning=false;
      return;
    }
    localStorage.setItem('lhca_onedrive_backup_failed','1');
    localStorage.setItem('lhca_onedrive_backup_failed_msg',e.message||'unknown error');
    renderOneDriveBackupBanner('failed');
  }finally{
    window._onedriveAutoRunning=false;
  }
}

// Show a small banner near the top of the page when auto-backup needs attention
function renderOneDriveBackupBanner(state){
  var existing=document.getElementById('odBackupBanner');
  if(state==='done'||state==='running'){
    if(existing)existing.remove();
    if(state==='done'){showToast('☁ Weekly backup to OneDrive complete',3500);}
    return;
  }
  if(state!=='failed')return;
  if(existing)return;
  var msg=localStorage.getItem('lhca_onedrive_backup_failed_msg')||'';
  var banner=document.createElement('div');
  banner.id='odBackupBanner';
  banner.style.cssText='position:fixed;top:0;left:0;right:0;z-index:9997;background:linear-gradient(90deg,#fff3cd,#ffe9a8);color:#856404;padding:8px 16px;border-bottom:1px solid #ffeaa7;box-shadow:0 2px 6px rgba(0,0,0,0.08);font-family:Arial,sans-serif;font-size:13px;display:flex;align-items:center;gap:12px;';
  banner.innerHTML='<span>⚠ Last weekly OneDrive backup failed'+(msg?': '+esc(msg.slice(0,80)):'')+'</span>'+
    '<button class="btn btn-primary btn-sm" onclick="retryOneDriveBackup()" style="margin-left:auto;">Retry Now</button>'+
    '<button class="btn btn-secondary btn-sm" onclick="document.getElementById(\'odBackupBanner\').remove();">Dismiss</button>';
  document.body.appendChild(banner);
}
async function retryOneDriveBackup(){
  var b=document.getElementById('odBackupBanner');if(b)b.remove();
  // Clear stale "no clients yet" failure (it's no longer a failure)
  var msg=(localStorage.getItem('lhca_onedrive_backup_failed_msg')||'');
  if(/no clients yet/i.test(msg)){
    localStorage.removeItem('lhca_onedrive_backup_failed');
    localStorage.removeItem('lhca_onedrive_backup_failed_msg');
  }
  await maybeAutoBackupOneDrive();
}

// Delete OneDrive backup files beyond the retention count (keeps newest N).
// ONLY prunes auto-backup files (date-only filename). Manual backups include
// '_manual' in the name and are kept forever.
async function pruneOldOneDriveBackups(){
  if(!spToken)return;
  try{
    var url='https://graph.microsoft.com/v1.0/me/drive/root:/Liberty Home Care Backups:/children?$select=id,name,lastModifiedDateTime&$orderby=lastModifiedDateTime desc&$top=200';
    var r=await fetch(url,{headers:{'Authorization':'Bearer '+spToken}});
    if(!r.ok)return;
    var data=await r.json();
    // Only auto-backup files (no '_manual' in the name) are eligible for prune
    var autoFiles=(data.value||[]).filter(function(f){
      return /^liberty_clients_\d{4}-\d{2}-\d{2}_(masked|FULL)\.json$/.test(f.name);
    });
    if(autoFiles.length<=ONEDRIVE_BACKUP_RETENTION)return;
    var toDelete=autoFiles.slice(ONEDRIVE_BACKUP_RETENTION);
    for(var i=0;i<toDelete.length;i++){
      try{
        await fetch('https://graph.microsoft.com/v1.0/me/drive/items/'+toDelete[i].id,{method:'DELETE',headers:{'Authorization':'Bearer '+spToken}});
      }catch(e){console.warn('Failed to prune',toDelete[i].name,e);}
    }
    console.log('[OneDrive] Pruned '+toDelete.length+' old auto-backup'+(toDelete.length>1?'s':'')+', keeping '+ONEDRIVE_BACKUP_RETENTION+' newest. Manual backups are not affected.');
  }catch(e){console.warn('Prune failed:',e);}
}

// One-click backup straight to user's OneDrive via Microsoft Graph.
// Asks whether to include full SSN, then uploads to /Liberty Home Care Backups/.
async function backupToOneDrive(){
  if(!spToken){
    showConfirm('Sign in with your Microsoft account to back up to OneDrive.',function(){signIn();},{title:'Sign In Required',okText:'Sign In',danger:false});
    return;
  }
  showConfirm(
    'Back up all client data to your OneDrive?\n\n'+
    'File goes to: OneDrive / Liberty Home Care Backups /\n\n'+
    'SSN is masked by default. Choose "Include Full SSN" if this backup needs to restore real SSNs (file stays on YOUR OneDrive — covered by your Microsoft BAA).',
    function(){doBackupToOneDrive(false);},
    {
      title:'Backup to OneDrive',
      okText:'Backup with Masked SSN',
      danger:false,
      extraText:'Include Full SSN',
      extraDanger:true,
      onExtra:function(){doBackupToOneDrive(true);}
    }
  );
}

async function doBackupToOneDrive(includeFullSSN,silent){
  var btn=document.getElementById('oneDriveBackupBtn');
  var oldText=btn?btn.textContent:'';
  if(btn&&!silent){btn.disabled=true;btn.textContent='Refreshing from DB…';}
  try{
    // Pull canonical state from DB before serializing — backups must reflect the
    // source of truth, not whatever happens to be in this device's LS cache.
    // Waits for the fetch to complete before building the payload.
    if(spToken && typeof loadProfilesAPI==='function'){
      try{ await loadProfilesAPI(); }catch(e){ console.warn('DB refresh failed before backup, using current LS:',e); }
    }
    if(btn&&!silent){btn.textContent='Uploading…';}
    // Build the backup payload (same logic as JSON export)
    var p=getProfiles();
    if(!Object.keys(p).length){throw new Error('No clients yet.');}
    var payload=_buildBackupPayload(includeFullSSN);
    var content=JSON.stringify(payload,null,2);
    var bytes=new Blob([content]).size;
    if(bytes>4*1024*1024){throw new Error('Backup too large for direct upload ('+(bytes/1024/1024).toFixed(1)+' MB). Contact support to add chunked upload.');}

    // Filename: auto-backups use date-only (eligible for retention prune),
    // manual backups include timestamp + 'manual' marker (kept forever).
    var ts=new Date().toISOString();
    var dateOnly=ts.slice(0,10);
    var manualMarker='';
    if(!silent){
      // HH-MM-SS for uniqueness across multiple manual saves on the same day
      var hms=ts.slice(11,19).replace(/:/g,'-');
      manualMarker='_'+hms+'_manual';
    }
    var fname='liberty_clients_'+dateOnly+manualMarker+(includeFullSSN?'_FULL':'_masked')+'.json';
    var path='Liberty Home Care Backups/'+fname;
    var url='https://graph.microsoft.com/v1.0/me/drive/root:/'+encodeURIComponent(path).replace(/%2F/g,'/')+':/content';
    var resp=await fetch(url,{
      method:'PUT',
      headers:{'Authorization':'Bearer '+spToken,'Content-Type':'application/json'},
      body:content
    });
    if(!resp.ok){
      var errTxt=await resp.text();
      // 403 likely means we don't have Files.ReadWrite scope yet — prompt re-sign-in
      if(resp.status===403){throw new Error('Permission denied. Sign out and sign back in, then accept OneDrive permission when prompted.');}
      throw new Error('Upload failed ('+resp.status+'): '+errTxt.slice(0,200));
    }
    var info=await resp.json();
    aiTrack('OneDriveBackup',{file:fname,size:bytes,fullSSN:!!includeFullSSN,auto:!!silent});
    // After a successful manual backup, also prune old files
    if(!silent)pruneOldOneDriveBackups();
    if(!silent){
      showConfirm(
        'Manual backup uploaded to OneDrive — this file is kept forever.\n\n'+
        'File: '+fname+'\n'+
        'Folder: Liberty Home Care Backups\n'+
        'Size: '+(bytes/1024).toFixed(1)+' KB\n'+
        (includeFullSSN?'\n⚠️ This backup contains full SSN — keep your OneDrive access controlled.':'')+
        '\n\nClick Open to view in OneDrive, or OK to dismiss.',
        function(){if(info.webUrl)window.open(info.webUrl,'_blank');},
        {title:'Manual Backup Complete',okText:'Open in OneDrive',danger:false,extraText:'OK',onExtra:function(){}}
      );
    }
    return info;
  }catch(e){
    console.error('OneDrive backup failed:',e);
    if(!silent)showConfirm('Backup failed: '+e.message,function(){},{title:'Backup Failed',okText:'OK',danger:false});
    throw e;
  }finally{
    if(btn&&!silent){btn.disabled=false;btn.textContent=oldText||'☁ Backup to OneDrive';}
  }
}

// Strip the internal dirty-tracking baseline from anything that leaves the app. _clientSig embeds
// the raw SSN, so a "masked" export/backup that keeps _clientSynced still ships the full number —
// masking prof.ssn alone is not enough. It is meaningless outside this device anyway.
function _stripInternalPHI(copy){
  Object.keys(copy).forEach(function(name){ if(copy[name]) delete copy[name]._clientSynced; });
  return copy;
}
function _maskSsnValue(v){
  var digits=String(v||'').replace(/\D/g,'');
  return digits.length>=4?'***-**-'+digits.slice(-4):'***-**-****';
}
// The whole backup payload, built in one testable place. Everything that leaves this device for
// OneDrive goes through here, so the SSN rules are enforced once and can be pinned by a test.
function _buildBackupPayload(includeFullSSN){
  var copy=_stripInternalPHI(JSON.parse(JSON.stringify(getProfiles())));
  if(!includeFullSSN){
    Object.keys(copy).forEach(function(name){
      var prof=copy[name]; if(prof && prof.ssn) prof.ssn=_maskSsnValue(prof.ssn);
    });
  }
  // CAREGIVERS were never masked at all — getCaregivers() re-attaches the SSN from _cgSsnMem, so
  // every backup (including the automatic weekly one, and the "masked" mode) uploaded every
  // caregiver's full SSN to OneDrive. Same rule as clients, and never ship the MI Login password.
  var cgCopy=JSON.parse(JSON.stringify(getCaregivers()));
  Object.keys(cgCopy).forEach(function(id){
    var c=cgCopy[id]; if(!c)return;
    delete c.miloginPassword; delete c.milogin_password;
    if(!includeFullSSN && c.ssn) c.ssn=_maskSsnValue(c.ssn);
  });
  return {
    _exportedAt:new Date().toISOString(),
    _exportedBy:currentUserEmail(),   // window.signedInEmail was never assigned — every backup said "unknown"
    _includesFullSSN:!!includeFullSSN,
    _appVersion:'liberty-homecare-v13',
    caregivers:cgCopy,
    caseworkers:getCaseworkers(),
    signatures:getSigs(),
    clients:copy
  };
}
function doExportProfiles(includeFullSSN){
  var p=getProfiles();
  // Deep clone so we don't mutate live data
  var copy=_stripInternalPHI(JSON.parse(JSON.stringify(p)));
  if(!includeFullSSN){
    Object.keys(copy).forEach(function(name){
      var prof=copy[name];if(prof.ssn){ prof.ssn=_maskSsnValue(prof.ssn); }
    });
  }
  // HIPAA §164.528: a bulk extract of the roster is a disclosure. It used to leave no record at all.
  try{
    var _n=Object.keys(copy).length;
    logActivity('export',(includeFullSSN?'FULL-SSN':'Masked')+' client export downloaded ('+_n+' clients) by '+currentUserEmail());
    if(includeFullSSN&&typeof addAuditEntry==='function')
      Object.keys(copy).forEach(function(nm){ addAuditEntry(nm,'Included in a FULL-SSN roster export'); });
  }catch(e){ console.error('export audit failed',e); }
  var fname='liberty_clients_'+new Date().toISOString().slice(0,10)+(includeFullSSN?'_FULL':'_masked')+'.json';
  var b=new Blob([JSON.stringify(copy,null,2)],{type:'application/json'}),u=URL.createObjectURL(b);
  var a=document.createElement('a');a.href=u;a.download=fname;a.click();URL.revokeObjectURL(u);
  if(includeFullSSN){
    showConfirm('Backup downloaded with FULL SSN.\n\nFile: '+fname+'\n\nStore this in a secure encrypted location. Do not email or share it.',function(){},{title:'Sensitive Backup',okText:'OK',danger:false});
  }
}

// Export all client+invoice data as multi-sheet Excel workbook
function exportClientsXLSX(){
  if(!window.XLSX){showAlert('Excel library not loaded — please reload the page.');return;}
  var profiles=getProfiles();
  var caregivers=getCaregivers();
  var caseworkers=getCaseworkers();
  if(!Object.keys(profiles).length){showAlert('No clients yet.');return;}

  // Sheet 1: Clients
  // Mask SSN to last 4 digits (PII safety in exports)
  function maskSSN(ssn){
    if(!ssn)return '';
    var digits=String(ssn).replace(/\D/g,'');
    if(digits.length<4)return '***-**-****';
    return '***-**-'+digits.slice(-4);
  }
  var clientsRows=Object.keys(profiles).map(function(name){
    var p=profiles[name];
    var cg=caregivers[p.caregiverId]||{};
    var cw=caseworkers.find(function(c){return c.id===p.caseworkerId||c.name===p.worker;})||{};
    return {
      'Client Name':name,
      'First Name':p.firstName||'','Last Name':p.lastName||'','Middle':p.middleName||'','Nickname':p.nickname||'',
      'Medicaid ID':p.medicaidId||'',
      'Status':clientStatusLabel(p.clientStatus||'active'),   // was p.status — never set, so every row exported as "active"
      'Hourly Pay Rate':p.hourlyRate||'',
      "Driver's License":p.driversLicense||'',
      'SSN (masked)':maskSSN(p.ssn),
      'Service Start':p.startDate||'',
      'Live-In':p.liveIn?'Yes':'',
      'Street':p.street||'','City':p.city||'','State':p.state||'','Zip':p.zip||'','County':p.county||'',
      'Caregiver':cg.name||'',
      'Caseworker':cw.name||'',
      'Caseworker Email':cw.email||'',
      'Bill To (Agency)':cw.agency||cw.county||'',
      'Invoice Count':(p.invoices||[]).length,
      'Notes':(p.clientNotes||'').replace(/\s+/g,' ').slice(0,500)
    };
  });
  // Sheet 2: Invoices (one row per invoice)
  var invoicesRows=[];
  Object.keys(profiles).forEach(function(name){
    var p=profiles[name];
    (p.invoices||[]).forEach(function(inv){
      invoicesRows.push({
        'Client Name':name,
        'Medicaid ID':p.medicaidId||'',
        'Billing Period':inv.billingPeriod||'',
        'Status':inv.status||'draft',
        'Saved At':inv.savedAt||'',
        'Hours (HH.MM)':((inv.data&&inv.data.svcHH)||'')+'.'+_padMin((inv.data&&inv.data.svcMM)||''),
        'Hourly Rate':(inv.data&&inv.data.hourlyRate)||'27.00',
        'Caseworker':p.worker||'',
        'Note':inv.invoiceNote||''
      });
    });
  });
  // Sheet 3: Caregivers (SSN masked to last 4)
  var caregiversRows=Object.keys(caregivers).map(function(id){
    var c=caregivers[id];
    return {'ID':id,'Name':c.name||'','First':c.firstName||'','Last':c.lastName||'','Status':c.status||'','Role':c.emptype||'','Phone':c.phone||'','Email':c.email||'',"Driver's License":c.driversLicense||'','SSN (masked)':(c.ssn?maskSSN(c.ssn):(c.ssnLast4?'***-**-'+c.ssnLast4:'')),'Hire Date':c.hireDate||'','Notes':(c.notes||'').slice(0,500)};
  });
  // Sheet 4: Caseworkers
  var caseworkersRows=caseworkers.map(function(c){
    return {'ID':c.id,'Name':c.name||'','Agency':c.agency||'','Phone':c.phone||'','Email':c.email||'','County':c.county||'','Notes':(c.notes||'').slice(0,500)};
  });

  var wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(clientsRows),'Clients');
  XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(invoicesRows),'Invoices');
  if(caregiversRows.length)XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(caregiversRows),'Caregivers');
  if(caseworkersRows.length)XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(caseworkersRows),'Caseworkers');
  var fname='liberty_clients_'+new Date().toISOString().slice(0,10)+'.xlsx';
  XLSX.writeFile(wb,fname);
}

// Export each client as a folder with all their invoice PDFs — all packaged into one ZIP
async function exportClientsAsPDFFolders(){
  if(!window.JSZip){showAlert('JSZip library not loaded — please reload the page.');return;}
  var profiles=getProfiles();
  var clientNames=Object.keys(profiles);
  if(!clientNames.length){showAlert('No clients yet.');return;}
  var clientsWithInvoices=clientNames.filter(function(n){return (profiles[n].invoices||[]).length>0;});
  if(!clientsWithInvoices.length){showAlert('No saved invoices to export.');return;}
  showConfirm(
    'Generate PDFs for '+clientsWithInvoices.length+' client'+(clientsWithInvoices.length>1?'s':'')+' and download as ZIP?\n\nThis may take a minute or two.',
    function(){_doExportClientsAsPDFFolders(clientsWithInvoices,profiles);},
    {title:'Export PDFs',okText:'Generate ZIP',danger:false}
  );
}
async function _doExportClientsAsPDFFolders(clientsWithInvoices,profiles){
  var btn=document.getElementById('exportPDFFoldersBtn');
  var oldText=btn?btn.textContent:'';
  if(btn){btn.disabled=true;btn.textContent='Preparing…';}

  // Stage invoice page for capture (must be active to render correctly)
  var savedActive=document.querySelector('.page.active')&&document.querySelector('.page.active').id;
  var invPage=document.getElementById('page-invoice');
  invPage.style.position='fixed';invPage.style.left='-9999px';invPage.style.top='0';invPage.style.zIndex='-1';
  invPage.classList.add('active');

  try{
    var zip=new JSZip();
    var totalInvoices=clientsWithInvoices.reduce(function(sum,n){return sum+(profiles[n].invoices||[]).length;},0);
    var done=0;
    for(var i=0;i<clientsWithInvoices.length;i++){
      var name=clientsWithInvoices[i];
      var prof=profiles[name];
      // Sanitize client name for folder safety
      var folder=name.replace(/[\/\\:*?"<>|]/g,'_');
      var clientFolder=zip.folder(folder);
      var invoices=(prof.invoices||[]).slice().sort(function(a,b){return (b.billingPeriod||'').localeCompare(a.billingPeriod||'');});
      for(var j=0;j<invoices.length;j++){
        var inv=invoices[j];
        if(btn)btn.textContent='Generating '+(++done)+' of '+totalInvoices+'…';
        try{
          await loadInvoiceForCapture(name,inv,inv.billingPeriod||'');
          var base64=await captureInvoicePDF();
          var bin=atob(base64);
          var bytes=new Uint8Array(bin.length);
          for(var k=0;k<bin.length;k++)bytes[k]=bin.charCodeAt(k);
          var pdfName=(inv.billingPeriod||'invoice').replace(/\//g,'_')+'.pdf';
          clientFolder.file(pdfName,bytes);
        }catch(e){console.error('PDF gen failed for period '+inv.billingPeriod,e);}
      }
    }
    if(btn)btn.textContent='Building ZIP…';
    var zipBlob=await zip.generateAsync({type:'blob',compression:'DEFLATE',compressionOptions:{level:6}});
    var url=URL.createObjectURL(zipBlob);
    var a=document.createElement('a');a.href=url;a.download='liberty_invoices_'+new Date().toISOString().slice(0,10)+'.zip';
    document.body.appendChild(a);a.click();document.body.removeChild(a);
    setTimeout(function(){URL.revokeObjectURL(url);},2000);
  }catch(e){showAlert('Export failed: '+e.message);console.error(e);}
  finally{
    invPage.classList.remove('active');
    invPage.style.position='';invPage.style.left='';invPage.style.top='';invPage.style.zIndex='';
    document.querySelectorAll('.page').forEach(function(el){el.classList.remove('active');});
    if(savedActive){var p=document.getElementById(savedActive);if(p)p.classList.add('active');}
    if(btn){btn.disabled=false;btn.textContent=oldText||'Export Clients + Invoices (PDFs)';}
  }
}
// Union two invoice arrays by stable identity (dbId, else billing-period + savedAt).
// Imported invoices win on conflict (the user confirmed "overwrite"), but any existing
// invoice the import DOESN'T contain is KEPT — so a stale/older backup can't silently
// drop invoices that were added locally after that backup was exported.
function mergeInvoiceLists(existingInvs, importedInvs){
  existingInvs = existingInvs || []; importedInvs = importedInvs || [];
  var keyOf=function(inv){ return inv&&inv.dbId ? ('db:'+inv.dbId) : ('bp:'+((inv&&inv.billingPeriod)||'')+'|'+((inv&&inv.savedAt)||'')); };
  var seen={}, out=[];
  importedInvs.forEach(function(inv){ seen[keyOf(inv)]=true; out.push(inv); });
  existingInvs.forEach(function(inv){ if(!seen[keyOf(inv)]) out.push(inv); });
  return out;
}
// A restored record carries the sync baselines it had at export time: `_clientSynced` on the client and
// `_synced` on each invoice. syncNewInvoices SKIPS any invoice whose baseline still matches its content
// (app.js syncNewInvoices), so a restored invoice was never written back to the DB — and
// _profileHasUnsyncedChanges then read the client as clean and let the next background load revert it.
// Restoring a backup to undo bad invoice edits silently undid itself. Clearing the baselines forces the
// restored data to be re-sent. dbId is deliberately KEPT so the server row is updated, not duplicated.
function _clearSyncBaselines(store){
  Object.keys(store||{}).forEach(function(nm){
    var p=store[nm]; if(!p||typeof p!=='object')return;
    (p.invoices||[]).forEach(function(iv){ if(iv) delete iv._synced; });
    delete p._clientSynced;
  });
  return store;
}
// Two different files can be imported, and they are NOT the same shape:
//   • the JSON export  → { "<client name>": {...}, ... }
//   • the OneDrive backup → { _exportedAt, _exportedBy, _appVersion, clients:{...}, caregivers:{...},
//                             caseworkers:[...], signatures:[...] }
// The importer used to treat every top-level key as a client name, so importing a BACKUP restored no
// clients at all and instead created junk records called "_exportedAt", "clients", "caregivers"… —
// i.e. the backup was not restorable by the app that wrote it, and importing one polluted the roster.
function _parseBackupFile(raw){
  var imp=JSON.parse(raw);
  if(!imp||typeof imp!=='object')throw new Error('not an object');
  var wrapped=(imp.clients&&typeof imp.clients==='object'&&!Array.isArray(imp.clients));
  if(wrapped){
    return { clients:imp.clients, caregivers:imp.caregivers||null, caseworkers:imp.caseworkers||null,
             signatures:imp.signatures||null,
             meta:{ exportedAt:imp._exportedAt||'', exportedBy:imp._exportedBy||'', includesFullSSN:!!imp._includesFullSSN } };
  }
  // Flat export. Drop any underscore-prefixed metadata defensively so it can never become a client.
  var clients={}; Object.keys(imp).forEach(function(k){ if(k.charAt(0)!=='_')clients[k]=imp[k]; });
  return { clients:clients, caregivers:null, caseworkers:null, signatures:null, meta:{} };
}
function importProfiles(ev){
  var file=ev.target.files[0];if(!file)return;
  var r=new FileReader();r.onload=function(e){
    try{
      var parsed=_parseBackupFile(e.target.result);
      var imp=parsed.clients,ex=getProfiles(),keys=Object.keys(imp);
      if(!keys.length){showConfirm('No clients found in that file.',function(){},{title:'Nothing to Import',okText:'OK',danger:false});return;}
      // For each imported record: if SSN looks masked (***-**-****), preserve any existing real SSN
      keys.forEach(function(name){
        var incoming=imp[name],existing=ex[name];
        if(incoming&&incoming.ssn&&/^\*\*\*-\*\*-/.test(incoming.ssn)){
          // Masked SSN — keep existing real value if we have one, otherwise drop the masked placeholder
          if(existing&&existing.ssn&&!/^\*\*\*-\*\*-/.test(existing.ssn)){
            incoming.ssn=existing.ssn;
          } else {
            delete incoming.ssn;
          }
        }
      });
      var conf=keys.filter(function(k){return ex[k];});
      function doImport(){
        // Merge per-client: imported scalar fields overwrite existing, but UNION the
        // invoices arrays (see mergeInvoiceLists) so an older backup can't drop invoices
        // added locally since. Previously Object.assign replaced the whole record incl.
        // its invoices array.
        var merged=Object.assign({},ex);
        keys.forEach(function(name){
          var incoming=imp[name],existing=ex[name];
          if(existing){
            var rec=Object.assign({},existing,incoming);
            rec.invoices=mergeInvoiceLists(existing.invoices, incoming&&incoming.invoices);
            merged[name]=rec;
          } else {
            merged[name]=incoming;
          }
        });
        _clearSyncBaselines(merged);
        saveProfilesLS(merged);
        // A backup also carries the rosters. Restore any that are MISSING locally (a wiped device has
        // none of them) and push each to the server; existing records are left alone so a restore can
        // never silently overwrite work done since the backup was taken.
        var restored={caregivers:0,caseworkers:0,signatures:0};
        try{
          if(parsed.caregivers&&typeof parsed.caregivers==='object'){
            var cgs=getCaregivers(),cgAdd=[];
            Object.keys(parsed.caregivers).forEach(function(id){
              if(!cgs[id]&&parsed.caregivers[id]){ cgs[id]=parsed.caregivers[id]; cgAdd.push(id); }
            });
            if(cgAdd.length){ saveCaregiversLS(cgs); restored.caregivers=cgAdd.length;
              cgAdd.forEach(function(id){ if(typeof saveCaregiverAPI==='function')saveCaregiverAPI(id,cgs[id],true); }); }
          }
          if(Array.isArray(parsed.caseworkers)){
            var cws=getCaseworkers(),have={}; cws.forEach(function(c){ if(c&&c.id!=null)have[c.id]=1; });
            var cwAdd=parsed.caseworkers.filter(function(c){ return c&&c.id!=null&&!have[c.id]; });
            if(cwAdd.length){ saveCaseworkersLS(cws.concat(cwAdd)); restored.caseworkers=cwAdd.length;
              cwAdd.forEach(function(c){ if(typeof saveCaseworkerAPI==='function')saveCaseworkerAPI(c,true); }); }
          }
          if(Array.isArray(parsed.signatures)&&typeof saveSigsLS==='function'){
            var sigs=getSigs(),sHave={}; sigs.forEach(function(x){ if(x&&x.id)sHave[x.id]=1; });
            var sAdd=parsed.signatures.filter(function(x){ return x&&x.id&&!sHave[x.id]; });
            if(sAdd.length){ saveSigsLS(sigs.concat(sAdd)); restored.signatures=sAdd.length; }
          }
        }catch(e){ console.error('roster restore failed',e); }
        // Persist each imported client to the backend so the data isn't just local
        keys.forEach(function(name){
          try{saveProfileSP(name,merged[name]);}catch(e){console.error('Failed to sync an imported client to DB',e);}
        });
        renderSidebarClients();renderClientGrid();updateStats();
        var _extra=[];
        if(restored.caregivers)_extra.push(restored.caregivers+' caregiver'+(restored.caregivers>1?'s':''));
        if(restored.caseworkers)_extra.push(restored.caseworkers+' caseworker'+(restored.caseworkers>1?'s':''));
        if(restored.signatures)_extra.push(restored.signatures+' signature'+(restored.signatures>1?'s':''));
        var _meta=parsed.meta&&parsed.meta.exportedAt?('\n\nBackup taken '+new Date(parsed.meta.exportedAt).toLocaleString()+
          (parsed.meta.exportedBy?(' by '+parsed.meta.exportedBy):'')+
          (parsed.meta.includesFullSSN?'':'\n\nNote: this backup was MASKED, so SSNs were not restored — they remain as stored on the server.')):'';
        showConfirm('Imported '+keys.length+' client'+(keys.length>1?'s':'')+
          (_extra.length?(', plus '+_extra.join(', ')):'')+'.'+_meta,
          function(){},{title:'Import Complete',okText:'OK',danger:false});
      }
      if(conf.length){
        showConfirm('Overwrite existing client'+(conf.length>1?'s':'')+'?\n\n'+conf.join(', '),doImport,{title:'Overwrite Confirm',okText:'Overwrite'});
      } else {
        doImport();
      }
    }catch(err){showConfirm('Could not read that file. It may be corrupted or not a valid JSON export.',function(){},{title:'Import Failed',okText:'OK',danger:false});}
    ev.target.value='';
  };r.readAsText(file);
}
function esc(s){return String(s||'').replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/'/g,'&#39;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}
// Escape a value that goes inside a JS single-quoted string which itself sits inside an
// HTML attribute — e.g. onclick="fn('<HERE>')". esc() ALONE IS UNSAFE there: the HTML
// parser decodes &#39; back to ' before the JS runs, so a quote in the data breaks out of
// the string and injects code. Fix: backslash-escape the JS-significant chars FIRST, then
// HTML-escape. A "'" becomes "\&#39;", which the HTML parser decodes to "\'" — an escaped
// quote the JS engine sees safely. Use this for ALL data in on* handler arguments.
function escJsAttr(v){
  return esc(String(v==null?'':v)
    .replace(/\\/g,'\\\\')
    .replace(/'/g,"\\'")
    .replace(/[\r\n\u2028\u2029]/g,' '));
}

// ============================================================
//  INVOICE FORM
// ============================================================
function loadProfileIntoForm(prof){
  document.getElementById('clientName').value=prof.clientName||'';document.getElementById('clientName2').value=prof.clientName||'';
  document.getElementById('medicaidId').value=prof.medicaidId||'';
  // Bill To: use caseworker's county code (e.g. "50-MACOMB"); fall back to county name
  var cwRec=getCaseworkers().find(function(c){return c.id===prof.caseworkerId||c.name===prof.worker;})||{};
  document.getElementById('billTo').value=(cwRec.agency||cwRec.county||'');
  document.getElementById('worker').value=prof.worker||'';document.getElementById('worker2').value=prof.worker||'';
  // Invoice bills at the client's own rate (set from their DHS-1210 authorization or manually),
  // falling back to the Settings state rate. NOT the caregiver's pay rate.
  document.getElementById('hourlyRate').value=clientInvoiceRate(prof);
  document.getElementById('showComplex').checked=prof.hasComplex||false;toggleComplex();
  if(prof.tasks){applyStates(prof.tasks);lastLoadedStates=prof.tasks;}
  // Agent email from caseworker record
  var ef=document.getElementById('activeAgentEmail');if(ef)ef.value=cwRec.email||'';
}
function captureFullInvoice(){
  var f=['clientName','medicaidId','billTo','worker','billingPeriod','hourlyRate','svcHH','svcMM','cplxHH','cplxMM','p1HH','p1MM','grandHH','grandMM','dateSubmitted','sigDate1','sigDate2'],d={};
  f.forEach(function(id){var el=document.getElementById(id);if(el)d[id]=el.value;});
  d.hasComplex=document.getElementById('showComplex').checked;d.tasks=captureStates();return d;
}
function applyFullInvoice(data){
  // A signature belongs to the invoice it was placed on. Loading another invoice (from history, or
  // via "copy from last") used to leave the previous stamp in the DOM, and both the PDF and the email
  // path read it straight from there — so a later month could go out CERTIFIED with a signature the
  // user never placed for it.
  resetSigArea(1); resetSigArea(2);
  var f=['clientName','medicaidId','worker','billingPeriod','svcHH','svcMM','cplxHH','cplxMM','p1HH','p1MM','grandHH','grandMM','dateSubmitted','sigDate1','sigDate2'];
  f.forEach(function(id){var el=document.getElementById(id);if(el&&data[id]!==undefined)el.value=data[id];});
  // Hourly rate = the client's own rate (from DHS-1210 or manual), else the Settings state rate
  var profCur=(activeProfileName&&getProfiles()[activeProfileName])||{};
  document.getElementById('hourlyRate').value=clientInvoiceRate(profCur);
  // Bill To: always from caseworker billing code, not whatever was saved on the invoice
  var cwApply=getCaseworkers().find(function(c){return c.id===profCur.caseworkerId||c.name===(data.worker||profCur.worker);})||{};
  document.getElementById('billTo').value=(cwApply.agency||cwApply.county||data.billTo||'');
  document.getElementById('clientName2').value=data.clientName||'';document.getElementById('worker2').value=data.worker||'';document.getElementById('billingPeriod2').value=data.billingPeriod||'';
  var bp=data.billingPeriod||'',p=bp.split('/');rebuild(p.length===2&&p[1].length===4?daysIn(p[0],p[1]):31);
  if(data.tasks){applyStates(data.tasks);lastLoadedStates=data.tasks;}
  document.getElementById('showComplex').checked=data.hasComplex||false;toggleComplex();
}
function saveInvoiceToClient(){
  if(!activeProfileName){showAlert('No client selected.');return;}
  var bp=document.getElementById('billingPeriod').value.trim();if(!bp){showAlert('Enter a billing period first.');return;}
  var p=getProfiles();if(!p[activeProfileName])return;if(!p[activeProfileName].invoices)p[activeProfileName].invoices=[];
  var ex=p[activeProfileName].invoices.findIndex(function(i){return i.billingPeriod===bp;});
  if(ex>=0){
    var existingStatus=p[activeProfileName].invoices[ex].status||'draft';
    if(existingStatus==='paid'){
      showAlert('Invoice '+bp+' is marked Paid and cannot be overwritten. Change the status to Draft first if you need to edit it.',{title:'Invoice Locked'});
      return;
    }
    var msg=existingStatus==='submitted'
      ? 'Invoice '+bp+' has already been submitted. Are you sure you want to overwrite it?'
      : 'Invoice for '+bp+' already exists. Overwrite?';
    showConfirm(msg,function(){_doSaveInvoiceToClient(bp,ex,existingStatus);},{title:'Overwrite Invoice',okText:'Overwrite'});
    return;
  }
  _doSaveInvoiceToClient(bp,-1,null);
}
function _doSaveInvoiceToClient(bp,ex,existingStatus){
  aiTrack('InvoiceSaved',{client:activeProfileName,period:bp});
  var p=getProfiles();if(!p[activeProfileName])return;if(!p[activeProfileName].invoices)p[activeProfileName].invoices=[];
  if(ex>=0){
    var prevInv=p[activeProfileName].invoices[ex];
    p[activeProfileName].invoices[ex]=Object.assign({},prevInv,{
      billingPeriod:bp,
      savedAt:new Date().toLocaleString(),
      status:existingStatus,
      data:captureFullInvoice()
    });
    addAuditEntry(activeProfileName,'Invoice '+bp+' overwritten');
  } else {
    p[activeProfileName].invoices.unshift({billingPeriod:bp,savedAt:new Date().toLocaleString(),status:'draft',invoiceNote:'',data:captureFullInvoice()});
    addAuditEntry(activeProfileName,'Invoice '+bp+' created');
  }
  saveProfilesLS(p);saveProfileSP(activeProfileName,p[activeProfileName]);
  logActivity('invoice','Invoice '+bp+' saved for '+activeProfileName);
  try{localStorage.removeItem('lhca_draft_'+activeProfileName);}catch(e){}
  var btn=document.getElementById('saveInvoiceBtn');btn.textContent='Saved';setTimeout(function(){btn.textContent='Save Invoice';},1800);
  document.getElementById('dupWarning').style.display='none';
  updateStats();
}
function clearInvoiceForm(){
  showConfirm('Clear the invoice form? Unsaved changes will be lost.',function(){
    ['clientName','medicaidId','billTo','worker','billingPeriod','billingPeriod2','hourlyRate','clientName2','worker2','svcHH','svcMM','cplxHH','cplxMM','p1HH','p1MM','grandHH','grandMM'].forEach(function(id){var e=document.getElementById(id);if(e)e.value='';});
    var bpp=document.getElementById('billingPeriodPicker');if(bpp)bpp.value='';
    var T=today();document.getElementById('dateSubmitted').value=T;document.getElementById('sigDate1').value=T;document.getElementById('sigDate2').value=T;
    document.getElementById('showComplex').checked=false;document.getElementById('complexSection').style.display='none';rebuild(31);resetSigArea(1);resetSigArea(2);lastLoadedStates=null;
    document.getElementById('dupWarning').style.display='none';
  },{title:'Clear Invoice Form',okText:'Clear'});
}

// ============================================================
//  AUTOSAVE DRAFT
// ============================================================
// (Removed) Draft-autosave was write-only: it wrote lhca_draft_* every 30s and flashed a
// "Draft autosaved" badge, but no code ever read the key back — drafts could never be
// recovered, so the timer + badge only misled the user. The removeItem('lhca_draft_'+…)
// cleanups elsewhere are kept so any legacy draft keys still in a browser get purged.

// ============================================================
//  INVOICE TABLE HELPERS (unchanged from original)
// ============================================================
function today(){var d=new Date();return String(d.getMonth()+1).padStart(2,'0')+'/'+String(d.getDate()).padStart(2,'0')+'/'+d.getFullYear();}
function daysIn(mm,yyyy){var m=parseInt(mm),y=parseInt(yyyy);if(isNaN(m)||isNaN(y)||m<1||m>12)return 31;return new Date(y,m,0).getDate();}
function buildRows(tbodyId,cols){
  var tb=document.getElementById(tbodyId);tb.innerHTML='';
  var isSvc=(tbodyId==='svcBody');
  var hospIdx=isSvc?(cols-1):-1; // Hospital column is the LAST column in SVC table
  for(var d=1;d<=31;d++){
    var tr=document.createElement('tr');if(d>active)tr.classList.add('inactive');
    // First cell: per-row "All" button (own column, hidden in print)
    var tdAll=document.createElement('td');tdAll.className='all-col';
    if(d<=active){
      var allBtn=document.createElement('span');
      allBtn.className='row-all-btn';allBtn.textContent='All';
      allBtn.title='Toggle entire row';
      allBtn.style.cssText='cursor:pointer;font-size:8pt;color:#185FA5;padding:1px 6px;border:1px solid #aac4e0;border-radius:3px;background:#eaf4ff;font-weight:bold;display:inline-block;line-height:1.4;font-family:Arial,sans-serif;';
      allBtn.addEventListener('mouseenter',function(){this.style.background='#bbddff';});
      allBtn.addEventListener('mouseleave',function(){this.style.background='#eaf4ff';});
      allBtn.addEventListener('click',function(e){
        e.stopPropagation();
        var rowCells=this.closest('tr').querySelectorAll('td.mc');
        var nonHospOn=Array.from(rowCells).slice(0,hospIdx>=0?hospIdx:rowCells.length).some(function(c){return c.classList.contains('on');});
        if(nonHospOn){
          rowCells.forEach(function(c){c.classList.remove('on');});
        } else {
          rowCells.forEach(function(c,i){if(hospIdx>=0&&i===hospIdx)c.classList.remove('on');else c.classList.add('on');});
        }
      });
      tdAll.appendChild(allBtn);
    }
    tr.appendChild(tdAll);
    var td0=document.createElement('td');td0.className='dc';td0.textContent=d;
    tr.appendChild(td0);
    for(var c=0;c<cols;c++){
      var td=document.createElement('td');td.className='mc';
      var box=document.createElement('span');box.className='box';td.appendChild(box);
      if(d<=active){
        (function(colIdx){
          td.addEventListener('click',function(){
            this.classList.toggle('on');
            // Hospital exclusivity (SVC table only)
            if(hospIdx<0)return;
            var rowCells=this.parentNode.querySelectorAll('td.mc');
            if(colIdx===hospIdx){
              // Just turned hospital on/off — if on, clear all other cells in row
              if(this.classList.contains('on')){
                rowCells.forEach(function(c,i){if(i!==hospIdx)c.classList.remove('on');});
              }
            } else {
              // Just toggled a service column — if turned on, clear hospital
              if(this.classList.contains('on')){
                var hosp=rowCells[hospIdx];if(hosp)hosp.classList.remove('on');
              }
            }
          });
        })(c);
      }
      tr.appendChild(td);
    }
    tb.appendChild(tr);
  }
}
function buildAllRow(rowId,tbodyId,cols){
  // Keep first 2 children (the all-col placeholder + day-col placeholder); rebuild service-col cells
  var row=document.getElementById(rowId);while(row.children.length>2)row.removeChild(row.lastChild);
  // Helper: get start-of-month day-of-week (Mon=0..Sun=6)
  function getStartDow(){
    var bp=document.getElementById('billingPeriod').value.trim(),startDow=0;
    if(bp&&bp.length===7){var pts=bp.split('/'),m=parseInt(pts[0]),y=parseInt(pts[1]);if(!isNaN(m)&&!isNaN(y))startDow=(new Date(y,m-1,1).getDay()+6)%7;}
    return startDow;
  }
  // Apply a day-of-week pattern (array of dow indices 0-6 where 0=Mon) to a column
  function applyDowPattern(bid,col,dowSet){
    var startDow=getStartDow();
    var rows=Array.from(document.getElementById(bid).querySelectorAll('tr:not(.inactive)'));
    rows.forEach(function(tr){var cell=tr.querySelectorAll('td')[col+2];if(cell)cell.classList.remove('on');});
    rows.forEach(function(tr,i){if(dowSet.indexOf((startDow+i)%7)!==-1){var cell=tr.querySelectorAll('td')[col+2];if(cell)cell.classList.add('on');}});
  }
  for(var c=0;c<cols;c++){
    var td=document.createElement('td');td.style.cssText='padding:0;vertical-align:middle;';
    (function(col,bid){
      function mk(txt,css,fn){var d=document.createElement('div');d.textContent=txt;d.style.cssText=css;d.addEventListener('click',fn);return d;}
      var btnCss='cursor:pointer;font-size:7pt;font-weight:bold;padding:1px 0;border-bottom:1px solid #aac4e0;';
      var allBtn=mk('All','cursor:pointer;font-size:7.5pt;color:#333;padding:1px 0;border-bottom:1px solid #aac4e0;',function(){
        allBtn._on=!allBtn._on;document.getElementById(bid).querySelectorAll('tr:not(.inactive)').forEach(function(tr){var cell=tr.querySelectorAll('td')[col+2];if(cell)cell.classList.toggle('on',allBtn._on);});
      });
      allBtn.addEventListener('mouseenter',function(){this.style.background='#bbddff';});allBtn.addEventListener('mouseleave',function(){this.style.background='';});
      var wkdyBtn=mk('Wkdy',btnCss+'color:#0c5460;',function(){applyDowPattern(bid,col,[0,1,2,3,4]);});
      wkdyBtn.title='Weekdays (Mon-Fri)';
      wkdyBtn.addEventListener('mouseenter',function(){this.style.background='#bee5eb';});wkdyBtn.addEventListener('mouseleave',function(){this.style.background='';});
      var oneBtn=mk('1x/wk',btnCss+'color:#856404;',function(){
        var cycle=parseInt(oneBtn.dataset.cycle)||0;applyDowPattern(bid,col,[cycle]);oneBtn.dataset.cycle=String((cycle+1)%7);
      });
      oneBtn.title='Once a week — click to cycle through days';
      oneBtn.addEventListener('mouseenter',function(){this.style.background='#fff3cd';});oneBtn.addEventListener('mouseleave',function(){this.style.background='';});
      var rndBtn=mk('2x/wk',btnCss+'color:#155724;',function(){
        var patterns=[[1,3],[2,4],[3,5]],cycle=parseInt(rndBtn.dataset.cycle)||0,pair=patterns[cycle];rndBtn.dataset.cycle=String((cycle+1)%patterns.length);
        applyDowPattern(bid,col,pair);
      });
      rndBtn.title='2x a week — click to cycle patterns';
      rndBtn.addEventListener('mouseenter',function(){this.style.background='#b8dfc4';});rndBtn.addEventListener('mouseleave',function(){this.style.background='';});
      var threeBtn=mk('3x/wk',btnCss+'color:#1a7740;',function(){
        var patterns=[[0,2,4],[1,3,5],[0,2,5],[1,3,4]],cycle=parseInt(threeBtn.dataset.cycle)||0;applyDowPattern(bid,col,patterns[cycle]);threeBtn.dataset.cycle=String((cycle+1)%patterns.length);
      });
      threeBtn.title='3x a week — click to cycle patterns';
      threeBtn.addEventListener('mouseenter',function(){this.style.background='#a8e6c4';});threeBtn.addEventListener('mouseleave',function(){this.style.background='';});
      var cpyBtn=mk('Copy','cursor:pointer;font-size:7pt;color:#721c24;font-weight:bold;padding:1px 0;border-bottom:1px solid #aac4e0;',function(){
        clipboard[bid]=Array.from(document.getElementById(bid).querySelectorAll('tr')).map(function(tr){var cell=tr.querySelectorAll('td')[col+2];return cell?cell.classList.contains('on'):false;});
        cpyBtn.style.background='#f8d7da';setTimeout(function(){cpyBtn.style.background='';},600);
      });
      cpyBtn.addEventListener('mouseenter',function(){this.style.background='#f8d7da';});cpyBtn.addEventListener('mouseleave',function(){this.style.background='';});
      var pstBtn=mk('Paste','cursor:pointer;font-size:7pt;color:#004085;font-weight:bold;padding:1px 0;',function(){
        if(!clipboard[bid])return;
        Array.from(document.getElementById(bid).querySelectorAll('tr')).forEach(function(tr,i){var cell=tr.querySelectorAll('td')[col+2];if(cell&&!tr.classList.contains('inactive'))cell.classList.toggle('on',clipboard[bid][i]||false);});
        pstBtn.style.background='#cce5ff';setTimeout(function(){pstBtn.style.background='';},600);
      });
      pstBtn.addEventListener('mouseenter',function(){this.style.background='#cce5ff';});pstBtn.addEventListener('mouseleave',function(){this.style.background='';});
      td.appendChild(allBtn);td.appendChild(wkdyBtn);td.appendChild(oneBtn);td.appendChild(rndBtn);td.appendChild(threeBtn);td.appendChild(cpyBtn);td.appendChild(pstBtn);
    })(c,tbodyId);
    row.appendChild(td);
  }
}
function rebuild(days){active=days;buildRows('svcBody',SVC);buildRows('cplxBody',CPLX);buildAllRow('svcAllRow','svcBody',SVC);buildAllRow('cplxAllRow','cplxBody',CPLX);}
function onBillingTextInput(el){
  var raw=el.value.replace(/\D/g,'');
  var formatted=raw;
  if(raw.length>2)formatted=raw.slice(0,2)+'/'+raw.slice(2);
  // Cap at 7 chars (MM/YYYY)
  if(formatted.length>7)formatted=formatted.slice(0,7);
  el.value=formatted;
  document.getElementById('billingPeriod2').value=formatted;
  // If complete, rebuild grid — preserve current checkbox state
  var parts=formatted.split('/');
  if(parts.length===2&&parts[1].length===4){
    var savedStates=captureStates();
    rebuild(daysIn(parts[0],parts[1]));
    applyStates(savedStates);
    checkDuplicatePeriod(formatted);
  }
}
function onBillingBlur(el){
  var val=el.value.trim();
  // Handle mmyy shorthand → MM/YYYY (e.g. 0425 → 04/2025)
  var raw=val.replace(/\D/g,'');
  if(raw.length===4){
    var mm=raw.slice(0,2),yy=raw.slice(2,4);
    var fullYear=parseInt(yy)<50?'20'+yy:'19'+yy;
    val=mm+'/'+fullYear;
    el.value=val;
    document.getElementById('billingPeriod2').value=val;
    var parts=val.split('/');
    var savedStates=captureStates();
    rebuild(daysIn(parts[0],parts[1]));
    applyStates(savedStates);
    checkDuplicatePeriod(val);
  } else if(raw.length===5){
    // "9/2025" types as 92025 — the month digit needs padding. Left unrepaired, the value stays
    // malformed ("92/025") and every period comparison downstream (duplicate check, missing-invoice
    // report, monthly preview) silently fails to match it.
    var mm5='0'+raw.slice(0,1), yyyy5=raw.slice(1,5);
    val=mm5+'/'+yyyy5;
    el.value=val;
    document.getElementById('billingPeriod2').value=val;
    var parts5=val.split('/');
    var savedStates5=captureStates();
    rebuild(daysIn(parts5[0],parts5[1]));
    applyStates(savedStates5);
    checkDuplicatePeriod(val);
  } else if(raw.length===6){
    // mmyyyy shorthand e.g. 042025
    var mm=raw.slice(0,2),yyyy=raw.slice(2,6);
    val=mm+'/'+yyyy;
    el.value=val;
    document.getElementById('billingPeriod2').value=val;
    var parts=val.split('/');
    var savedStates=captureStates();
    rebuild(daysIn(parts[0],parts[1]));
    applyStates(savedStates);
    checkDuplicatePeriod(val);
  }
}
function syncBillingPeriodFields(){
  var bp=document.getElementById('billingPeriod').value;
  document.getElementById('billingPeriod2').value=bp||'';
}
function checkDuplicatePeriod(bp){
  var warn=document.getElementById('dupWarning');if(!warn||!activeProfileName)return;
  var p=getProfiles();if(!p[activeProfileName])return;
  var ex=((p[activeProfileName].invoices)||[]).find(function(i){return i.billingPeriod===bp;});
  if(ex){
    var st=ex.status||'draft';
    var msg=st==='paid'?
      '[Locked] A <strong>Paid</strong> invoice already exists for '+esc(bp)+'. It is locked and cannot be overwritten.':
      'Warning: An invoice for '+esc(bp)+' already exists (status: '+esc(st)+'). Saving will overwrite it.';
    warn.innerHTML=msg;warn.style.display='block';
  } else {warn.style.display='none';}
}
function syncFields(){document.getElementById('clientName2').value=document.getElementById('clientName').value;document.getElementById('worker2').value=document.getElementById('worker').value;}
function toggleComplex(){document.getElementById('complexSection').style.display=document.getElementById('showComplex').checked?'block':'none';}
function captureStates(){
  var svc=Array.from(document.getElementById('svcBody').querySelectorAll('tr')).map(function(tr){return Array.from(tr.querySelectorAll('td.mc')).map(function(td){return td.classList.contains('on');});});
  var cplx=Array.from(document.getElementById('cplxBody').querySelectorAll('tr')).map(function(tr){return Array.from(tr.querySelectorAll('td.mc')).map(function(td){return td.classList.contains('on');});});
  return{svc:svc,cplx:cplx};
}
function applyStates(states){
  ['svc','cplx'].forEach(function(key){
    var tbodyId=key==='svc'?'svcBody':'cplxBody',rows=document.getElementById(tbodyId).querySelectorAll('tr');
    states[key].forEach(function(rowState,i){
      if(!rows[i])return;
      // Never re-apply a mark to a day the CURRENT month doesn't have. Copying a 31-day month into a
      // 30-day one used to carry day 31's checkmarks across; the print CSS re-shows .inactive rows and
      // buildInvoiceHTML strips the class entirely, so the phantom day printed on the certified form.
      if(rows[i].classList.contains('inactive')){
        rows[i].querySelectorAll('td.mc').forEach(function(c){c.classList.remove('on');});
        return;
      }
      var cells=rows[i].querySelectorAll('td.mc');
      rowState.forEach(function(on,j){if(cells[j])cells[j].classList.toggle('on',on);});
    });
  });
}

// ── Toast ─────────────────────────────────────────────────────
// a11y: announce a message to screen readers via a visually-hidden aria-live region, so
// toasts / "Saved ✓" / "Save failed" (visual-only otherwise) are spoken. WCAG 4.1.3.
function _ariaAnnounce(msg){
  var r=document.getElementById('ariaLive');
  if(!r){
    r=document.createElement('div');r.id='ariaLive';
    r.setAttribute('aria-live','polite');r.setAttribute('role','status');
    r.style.cssText='position:absolute;width:1px;height:1px;overflow:hidden;clip:rect(0 0 0 0);clip-path:inset(50%);white-space:nowrap;';
    document.body.appendChild(r);
  }
  r.textContent='';setTimeout(function(){r.textContent=msg;},50); // clear+set so repeats re-announce
}
function showToast(msg,ms){
  var t=document.getElementById('lhcaToast');if(!t){t=document.createElement('div');t.id='lhcaToast';t.className='lhca-toast';document.body.appendChild(t);}
  t.textContent=msg;t.classList.add('show');
  _ariaAnnounce(msg);
  clearTimeout(t._tid);t._tid=setTimeout(function(){t.classList.remove('show');},ms||3500);
}
// ── Save-status toast — used by every API save path so failures are never silent ──
// Usage: trackSave('client', () => fetch(...))  — returns the promise so callers can await/chain
// On failure: shows a red toast with a Retry button that re-invokes the save.
var _saveStatusEl=null;
function _showSaveStatus(state,label,onRetry){
  if(!_saveStatusEl){
    _saveStatusEl=document.createElement('div');
    _saveStatusEl.id='saveStatusToast';
    _saveStatusEl.style.cssText='position:fixed;bottom:18px;right:18px;background:#fff;border:1px solid #d0d8e4;border-radius:8px;padding:10px 14px;box-shadow:0 4px 12px rgba(0,0,0,0.15);font-size:12px;z-index:10000;display:none;align-items:center;gap:10px;font-family:Arial,sans-serif;';
    document.body.appendChild(_saveStatusEl);
  }
  _saveStatusEl.style.display='flex';
  if(state==='saving'){
    _saveStatusEl.style.borderColor='#d0d8e4';_saveStatusEl.style.background='#fff';
    _saveStatusEl.innerHTML='<span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:#cbb26b;animation:sfPulse 1.4s ease-in-out infinite;"></span><span style="color:#1a2b45;">Saving '+esc(label)+'…</span>';
    clearTimeout(_saveStatusEl._t);_saveStatusEl._t=setTimeout(function(){_saveStatusEl.style.display='none';},5000);
  } else if(state==='saved'){
    _saveStatusEl.style.borderColor='#b9e4c9';_saveStatusEl.style.background='#eef9f1';
    _saveStatusEl.innerHTML='<span style="color:#1a7740;font-weight:700;">✓</span><span style="color:#1a7740;">Saved '+esc(label)+'</span>';
    _ariaAnnounce('Saved '+label);
    clearTimeout(_saveStatusEl._t);_saveStatusEl._t=setTimeout(function(){_saveStatusEl.style.display='none';},2800);
  } else if(state==='failed'){
    _saveStatusEl.style.borderColor='#f0c0c0';_saveStatusEl.style.background='#fdecec';
    _saveStatusEl.innerHTML='<span style="color:#a00;font-weight:700;">✗</span><span style="color:#a00;flex:1;">Save failed: '+esc(label)+'</span>'+
      (onRetry?'<button class="btn btn-secondary btn-sm" style="padding:4px 10px;" id="_sstRetryBtn">Retry</button>':'')+
      '<button style="background:none;border:none;color:#a00;cursor:pointer;font-size:14px;padding:0 4px;" onclick="document.getElementById(\'saveStatusToast\').style.display=\'none\';" aria-label="Dismiss">✕</button>';
    _ariaAnnounce('Save failed: '+label);
    if(onRetry){
      var btn=document.getElementById('_sstRetryBtn');
      if(btn)btn.addEventListener('click',function(){onRetry();});
    }
    // Don't auto-dismiss failures — user must acknowledge or retry
  }
}
function trackSave(label,doSave){
  _showSaveStatus('saving',label);
  return doSave()
    .then(function(r){_showSaveStatus('saved',label);return r;})
    .catch(function(e){
      _showSaveStatus('failed',label+' ('+(e&&e.message?e.message:'unknown')+')',function(){trackSave(label,doSave);});
      throw e;
    });
}
// ── Silent-save guards (audit Batch 2: D8/D10/D13) ──────────────────
// Optimistic UI updates and "quiet" auto-saves must never report success when
// the API call actually failed. Both helpers expect a promise that REJECTS on a
// non-2xx response (the save fns already do `if(!r.ok)throw`). On failure they
// raise the persistent red save-status toast with a Retry button; success stays
// quiet (surfaceSaveFailure) or shows the inline "· Saved ✓" tick (flashQuietSave).
function surfaceSaveFailure(promise,label,onRetry){
  // Fire-and-forget callers don't chain, so swallow after surfacing (retry is wired
  // through the toast button) — rethrowing here would just be an unhandled rejection.
  return promise.catch(function(e){
    _showSaveStatus('failed',label+' ('+(e&&e.message?e.message:'network error')+')',onRetry||null);
  });
}
function flashQuietSave(promise,flashElId,label,onRetry){
  return promise.then(function(){
    var f=flashElId&&document.getElementById(flashElId);
    if(f){f.textContent='· Saved ✓';f.style.color='#1a7740';f.style.display='inline';setTimeout(function(){if(f)f.style.display='none';},2000);}
  }).catch(function(e){
    var f=flashElId&&document.getElementById(flashElId);
    if(f)f.style.display='none';
    _showSaveStatus('failed',label+' — not saved ('+(e&&e.message?e.message:'network error')+')',onRetry||null);
  });
}

// ── PDF / Print helpers ────────────────────────────────────────
function buildInvoiceHTML(){
  var p1=document.getElementById('page1').cloneNode(true);
  var cs=document.getElementById('complexSection').cloneNode(true);
  var showCplx=document.getElementById('showComplex').checked;
  var hasCplxChecked=document.querySelectorAll('#cplxTable td.mc.on').length>0;
  var includeCplx=showCplx&&hasCplxChecked;

  [p1,cs].forEach(function(root){
    // Remove UI quick-fill rows
    root.querySelectorAll('.allrow').forEach(function(r){if(r.parentNode)r.parentNode.removeChild(r);});
    // Inactive rows (beyond days in month): keep them but un-gray — they appear blank on the form
    root.querySelectorAll('tr.inactive').forEach(function(r){r.classList.remove('inactive');});
    // Strip "Click to place signature" placeholder text — keep just the underline
    root.querySelectorAll('.sig-placeholder').forEach(function(el){el.textContent='';});
  });

  // Column headers: replace writing-mode (not supported by html2canvas) with
  // transform:rotate on a width-constrained box so long text wraps correctly.
  // Pre-rotation width = 82px ≈ header height → becomes the visual column height after rotation.
  // Pre-rotation height = text height (1-3 lines) → becomes visual column width after rotation.
  [p1,cs].forEach(function(root){
    root.querySelectorAll('th.th').forEach(function(th){
      var text=th.textContent.trim();
      var safeText=text;
      th.removeAttribute('style');
      th.className='th';
      // The inner span is centered then rotated. width:82px constrains wrapping.
      th.innerHTML=
        '<div style="position:relative;width:100%;height:88px;overflow:hidden;">'+
          '<span style="position:absolute;left:50%;top:50%;display:inline-block;'+
          'width:82px;text-align:center;white-space:normal;word-break:break-word;line-height:1.25;'+
          'font-size:8pt;font-weight:normal;font-family:\'Times New Roman\',Times,serif;'+
          'transform:translate(-50%,-50%) rotate(-90deg);">'+
          safeText+
          '</span>'+
        '</div>';
    });
  });

  function syncInputs(root){
    root.querySelectorAll('input:not([type="checkbox"]):not([type="hidden"])').forEach(function(inp){
      inp.setAttribute('value',inp.value||'');
    });
    root.querySelectorAll('input[type="checkbox"]').forEach(function(cb){
      if(cb.checked)cb.setAttribute('checked',''); else cb.removeAttribute('checked');
    });
  }
  syncInputs(p1); syncInputs(cs);

  [p1,cs].forEach(function(root){
    root.querySelectorAll('[onclick],[onchange],[oninput],[onblur]').forEach(function(el){
      ['onclick','onchange','oninput','onblur'].forEach(function(a){el.removeAttribute(a);});
    });
  });

  var css=[
    // @page margin:0 so the browser doesn't add its own margins on top of body margins
    '@page{margin:0;size:letter;}',
    'body{margin:0.44in 0.44in;font-family:"Times New Roman",Times,serif;background:#fff;color:#000;}',
    '.pg-title{text-align:center;margin-bottom:2px;}',
    '.pg-title .t1{font-size:13pt;font-weight:bold;font-family:"Times New Roman",Times,serif;}',
    '.pg-title .t2{font-size:9.5pt;font-family:"Times New Roman",Times,serif;}',
    '.bi{width:100%;border-collapse:collapse;}',
    '.bi td{border:1px solid #000;padding:1px 3px;vertical-align:top;font-size:8.5pt;}',
    '.bi .lbl{font-size:7.5pt;display:block;font-family:"Times New Roman",Times,serif;}',
    '.bi input{border:none;width:100%;font-size:8.5pt;font-family:"Times New Roman",Times,serif;background:transparent;outline:none;color:#000;}',
    '.sec-hdr{border:1px solid #000;border-top:none;background:#fff;text-align:center;font-weight:bold;font-size:9pt;padding:2px 4px;font-family:"Times New Roman",Times,serif;}',
    '.tt{width:100%;border-collapse:collapse;border-left:1px solid #000;border-right:1px solid #000;border-bottom:1px solid #000;table-layout:fixed;}',
    '.tt th,.tt td{border:1px solid #000;text-align:center;vertical-align:middle;padding:0;}',
    '.tt th.dh{width:52px;font-size:8pt;font-weight:normal;padding:2px;font-family:"Times New Roman",Times,serif;line-height:1.2;}',
    '.tt th.th{height:88px;padding:0;overflow:hidden;vertical-align:middle;position:relative;}',
    '.tt td.dc{font-size:8pt;text-align:center;width:52px;padding:0;font-family:"Times New Roman",Times,serif;line-height:1;}',
    '.tt td.mc{padding:0;height:14px;vertical-align:middle;text-align:center;}',
    '.tt td.mc .box{display:inline-block;width:11px;height:11px;border:1px solid #000;position:relative;vertical-align:middle;}',
    '.tt td.mc.on .box::before,.tt td.mc.on .box::after{content:"";position:absolute;left:50%;top:-1px;width:0;height:calc(100% + 2px);border-left:1.4px solid #000;}',
    '.tt td.mc.on .box::before{transform:rotate(45deg);}',
    '.tt td.mc.on .box::after{transform:rotate(-45deg);}',
    '.bottom-block{border:1px solid #000;border-top:none;font-size:8.5pt;font-family:"Times New Roman",Times,serif;}',
    '.bottom-block .tot-row{padding:2px 4px;font-weight:bold;}',
    '.bottom-block .tot-row input{border:none;border-bottom:1px solid #000;font-size:8.5pt;font-family:"Times New Roman",Times,serif;background:transparent;outline:none;text-align:center;width:28px;}',
    '.bottom-block .cert-row{padding:2px 4px;}',
    '.bottom-block .label-row{display:flex;justify-content:space-between;padding:2px 4px;align-items:flex-end;}',
    '.sig-label{display:flex;flex-direction:column;gap:0;width:60%;}',
    '.sig-label span{font-size:7.5pt;font-family:"Times New Roman",Times,serif;}',
    '.sig-placeholder{border-bottom:1px solid #000;height:28px;display:flex;align-items:center;}',
    '.sig-stamp{max-height:28px;display:block;}',
    '.pg-footer{display:flex;justify-content:space-between;font-size:7.5pt;margin-top:2px;font-family:"Times New Roman",Times,serif;}',
    '.p2info{width:100%;border-collapse:collapse;}',
    '.p2info td{border:1px solid #000;padding:1px 3px;font-size:8.5pt;vertical-align:top;}',
    '.p2info .lbl{font-size:7.5pt;display:block;font-family:"Times New Roman",Times,serif;}',
    '.p2info input{border:none;width:100%;font-size:8.5pt;font-family:"Times New Roman",Times,serif;background:transparent;outline:none;color:#000;}',
    '.pagebreak{margin-top:24px;}'
  ].join('\n');

  return '<!DOCTYPE html><html><head><meta charset="utf-8"><style>'+css+'</style></head><body>'+
    p1.outerHTML+(includeCplx?'\n'+cs.outerHTML:'')+'</body></html>';
}

async function printInvoiceAsPDF(){
  var cn=document.getElementById('clientName')?document.getElementById('clientName').value.trim():'';
  var bp=document.getElementById('billingPeriod')?document.getElementById('billingPeriod').value.trim():'';
  aiTrack('InvoicePrinted',{clientName:cn,billingPeriod:bp});
  var btn=document.querySelector('button[onclick="printInvoiceAsPDF()"]');
  var oldText=btn?btn.textContent:'';
  if(btn){btn.disabled=true;btn.textContent='Generating…';}
  try{
    var base64=await captureInvoicePDF();
    // Convert base64 to blob and download
    var bin=atob(base64);
    var bytes=new Uint8Array(bin.length);
    for(var i=0;i<bin.length;i++)bytes[i]=bin.charCodeAt(i);
    var blob=new Blob([bytes],{type:'application/pdf'});
    var url=URL.createObjectURL(blob);
    var fname=((cn||'Invoice').replace(/[^a-z0-9]/gi,'_'))+'_'+((bp||'').replace('/','_'))+'.pdf';
    var a=document.createElement('a');a.href=url;a.download=fname;
    document.body.appendChild(a);a.click();document.body.removeChild(a);
    setTimeout(function(){URL.revokeObjectURL(url);},2000);
  }catch(e){
    console.error('PDF generation failed:',e);
    showAlert('PDF generation failed: '+e.message);
  }finally{
    if(btn){btn.disabled=false;btn.textContent=oldText||'Print / PDF';}
  }
}

// ============================================================
//  VECTOR PDF — direct jsPDF drawing (small file, government-form fidelity)
// ============================================================
// Letter = 612x792 pt; we use 18pt (0.25") margins matching @page setting.
// Output PDFs are 80-150 KB vs 1.6 MB for the raster path.
// Signature is the only raster element (small embedded PNG).
async function captureInvoicePDFVector(){
  var jsPDF=window.jspdf.jsPDF;
  var pdf=new jsPDF({unit:'pt',format:'letter',orientation:'p'});
  var data=captureFullInvoice();

  var SVC_COLS=['Bathing','Dressing','Eating','Grooming','Mobility','Toileting',
    'Transferring','Housework','Laundry','Travel Time for Laundry','Medication',
    'Meal Preparation','Shopping','Travel Time for Shopping','Hospital/Nursing Facility Stay'];
  var CPLX_COLS=['Bowel Program','Catheters for Leg Bags','Colostomy Care',
    'Eating or Feeding Assistance','Peritoneal Dialysis','Range of Motion Exercises',
    'Specialized Skin Care','Suctioning','Wound Care'];

  // Pre-crop any signature images so legacy padded sigs render at full effective size
  var sig1El=document.getElementById('sigArea1');
  var sig2El=document.getElementById('sigArea2');
  var sig1Cropped=(sig1El&&sig1El.tagName==='IMG'&&sig1El.src)?await cropDataURL(sig1El.src):null;
  var sig2Cropped=(sig2El&&sig2El.tagName==='IMG'&&sig2El.src)?await cropDataURL(sig2El.src):null;
  data._sig1Cropped=sig1Cropped;
  data._sig2Cropped=sig2Cropped;

  drawInvoicePageVector(pdf,data,false,SVC_COLS);
  var hasCplx=data.hasComplex && (data.tasks.cplx||[]).some(function(r){return r.some(function(c){return c;});});
  if(hasCplx){pdf.addPage();drawInvoicePageVector(pdf,data,true,CPLX_COLS);}

  return pdf.output('datauristring').split(',')[1];
}

function drawInvoicePageVector(pdf,data,isPage2,cols){
  // ── Layout constants (points, 612x792 letter) ──
  var ML=18,MT=18,W=576;        // content area
  var x0=ML,y=MT;

  // ── Title block ──
  pdf.setFont('helvetica','bold');pdf.setFontSize(14);
  pdf.text('HOME HELP AGENCY INVOICE',ML+W/2,y+14,{align:'center'});
  pdf.setFont('helvetica','normal');pdf.setFontSize(13);
  pdf.text('Michigan Department of Health and Human Services',ML+W/2,y+30,{align:'center'});
  y+=38;

  // ── Info table — sizes bumped to match Arial 12pt visually ──
  var GRID_INFO=0.5;
  var rowH=26,labelSize=11,valSize=12;
  function cell(x,y,w,h,label,value,boldVal){
    pdf.setLineWidth(GRID_INFO);pdf.rect(x,y,w,h);
    pdf.setFont('helvetica','normal');pdf.setFontSize(labelSize);
    pdf.text(label,x+3,y+11);
    pdf.setFont('helvetica',boldVal?'bold':'normal');pdf.setFontSize(valSize);
    pdf.text(String(value||''),x+3,y+22);
  }
  // Inline cell: label + value on same line, ONE cell (no internal divider)
  function inlineCell(x,y,w,h,label,value){
    pdf.setLineWidth(GRID_INFO);pdf.rect(x,y,w,h);
    pdf.setFont('helvetica','normal');pdf.setFontSize(11);
    pdf.text(label,x+3,y+13);
    var lblW=pdf.getTextWidth(label);
    pdf.text(String(value||''),x+3+lblW+5,y+13);
  }
  // Inline with vertical DIVIDER (used only for Bill To per March reference)
  function inlineDividerCell(x,y,labelW,valueW,h,label,value){
    pdf.setLineWidth(GRID_INFO);
    pdf.rect(x,y,labelW,h);
    pdf.rect(x+labelW,y,valueW,h);
    pdf.setFont('helvetica','normal');pdf.setFontSize(11);
    pdf.text(label,x+3,y+13);
    pdf.text(String(value||''),x+labelW+3,y+13);
  }
  var inlineH=18; // tighter rows for Client/Medicaid/BillTo/Attention

  if(!isPage2){
    // Row 1: Agency Name | Provider Number | Phone (stacked label/value)
    var c1w=288,c2w=144,c3w=144;
    cell(x0,y,c1w,rowH,'Agency Name','Liberty Home Care Assistance');
    cell(x0+c1w,y,c2w,rowH,'Agency Provider Number','6221933');
    cell(x0+c1w+c2w,y,c3w,rowH,'Agency Phone Number','248-291-4106');
    y+=rowH;
    // Row 2: Contact wider | Period | Date narrower | Rate narrower (less internal whitespace)
    var b1=235,b2=155,b3=100,b4=W-b1-b2-b3; // 235+155+100 = 490; b4 = 86 (Hourly Rate snug to "27.00")
    cell(x0,y,b1,rowH,'Contact Person','Thomas Jaboro');
    cell(x0+b1,y,b2,rowH,'Billing Period',data.billingPeriod);
    cell(x0+b1+b2,y,b3,rowH,'Date Submitted',data.dateSubmitted);
    cell(x0+b1+b2+b3,y,b4,rowH,'Hourly Rate',data.hourlyRate||'27.00');
    y+=rowH;
    // Row 3: 'Client Name: VAL' (one cell)  |  'Client Medicaid ID Number: VAL' (one cell)
    var halfW=W/2;
    inlineCell(x0,y,halfW,inlineH,'Client Name:',data.clientName);
    inlineCell(x0+halfW,y,halfW,inlineH,'Client Medicaid ID Number:',data.medicaidId);
    y+=inlineH;
    // Row 4: 'Bill To:' (cell) | VAL (cell) WITH vertical divider | 'Attention: VAL' (one cell, no divider)
    inlineDividerCell(x0,y,52,halfW-52,inlineH,'Bill To:',data.billTo);
    inlineCell(x0+halfW,y,halfW,inlineH,'Attention:',data.worker);
    y+=inlineH;
  } else {
    // Page 2 has a 3-col info row + the title repeated below it
    var p1=W*0.34,p2=W*0.33,p3=W-p1-p2;
    cell(x0,y,p1,rowH,'Client Name',data.clientName);
    cell(x0+p1,y,p2,rowH,'Caseworker Name',data.worker);
    cell(x0+p1+p2,y,p3,rowH,'Billing Period',data.billingPeriod);
    y+=rowH+4;
    pdf.setFont('helvetica','normal');pdf.setFontSize(9.5);
    pdf.text('Michigan Department of Health and Human Services',ML+W/2,y+9,{align:'center'});
    pdf.setFont('helvetica','bold');pdf.setFontSize(13);
    pdf.text('HOME HELP AGENCY INVOICE',ML+W/2,y+22,{align:'center'});
    y+=28;
  }

  // ── Section header bar — 12pt per March ──
  var secH=18;
  pdf.setLineWidth(0.6);pdf.rect(x0,y,W,secH);
  pdf.setFont('helvetica','bold');pdf.setFontSize(12);
  var secText=isPage2?'VERIFICATION OF COMPLEX CARE TASKS':'VERIFICATION OF SERVICES AND HOSPITAL/NURSING FACILITY STAYS';
  pdf.text(secText,ML+W/2,y+13,{align:'center'});
  y+=secH;

  // ── Day grid ──
  var dayColW=44,headerH=92,dayRowH=13,nCols=cols.length;
  var colW=(W-dayColW)/nCols;

  // ── Grid weight: 0.5pt — visible but not heavy, matches reference ──
  var GRID_W=0.5;

  // Header row
  pdf.setLineWidth(GRID_W);pdf.rect(x0,y,W,headerH);
  // Day-of-month label
  pdf.setFont('helvetica','normal');pdf.setFontSize(9);
  var dayLabel=['Days','of','Billing','Month'];
  dayLabel.forEach(function(line,i){
    pdf.text(line,x0+dayColW/2,y+headerH/2-15+i*10,{align:'center'});
  });
  // Vertical separator after day col
  pdf.line(x0+dayColW,y,x0+dayColW,y+headerH);
  // Column headers (rotated 90°) — Helvetica 11pt ≈ Arial 12pt visually (March reference size)
  pdf.setFontSize(11);
  for(var c=0;c<nCols;c++){
    var cx=x0+dayColW+c*colW;
    if(c>0)pdf.line(cx,y,cx,y+headerH);
    // Wrap long header text into 2 lines (max ~15 chars per line)
    var label=cols[c],lines=wrapHeaderText(label,15);
    var lineSpace=12;
    var totalW=lines.length*lineSpace;
    var startOffset=(totalW)/2-lineSpace/2;
    lines.forEach(function(line,li){
      // +3pt offset shifts the rotated text slightly right of center (matches March)
      var lineX=cx+colW/2-startOffset+li*lineSpace+3;
      pdf.text(line,lineX,y+headerH-4,{angle:90});
    });
  }
  y+=headerH;

  // Day rows (1..31)
  pdf.setFontSize(8.5);
  var dataRows=isPage2?(data.tasks.cplx||[]):(data.tasks.svc||[]);
  for(var d=1;d<=31;d++){
    pdf.setLineWidth(GRID_W);
    pdf.rect(x0,y,W,dayRowH);
    pdf.line(x0+dayColW,y,x0+dayColW,y+dayRowH);
    pdf.setFont('helvetica','normal');pdf.text(String(d),x0+dayColW/2,y+9,{align:'center'});
    var rowState=dataRows[d-1]||[];
    for(var c=0;c<nCols;c++){
      var cx=x0+dayColW+c*colW;
      pdf.setLineWidth(GRID_W);
      if(c>0)pdf.line(cx,y,cx,y+dayRowH);
      // Checkbox slightly taller than wide
      var bxW=10,bxH=11,bxX=cx+colW/2-bxW/2,bxY=y+dayRowH/2-bxH/2;
      pdf.rect(bxX,bxY,bxW,bxH);
      if(rowState[c]){
        // X stroke — lighter than before (was 1.1pt — too dark per user feedback)
        pdf.setLineWidth(0.8);
        pdf.line(bxX,bxY,bxX+bxW,bxY+bxH);
        pdf.line(bxX+bxW,bxY,bxX,bxY+bxH);
        pdf.setLineWidth(GRID_W);
      }
    }
    y+=dayRowH;
  }
  pdf.setLineWidth(GRID_W);

  // ── Bottom block — ONE outer rect, NO internal dividers, tight per March ──
  var totH=15,certH=14,sigH=70;
  var blockH=totH+certH+sigH;
  pdf.setLineWidth(GRID_W);
  pdf.rect(x0,y,W,blockH);

  // Total time row (no divider below) — 12pt, snug spacing per March
  if(!isPage2){
    pdf.setFont('helvetica','bold');pdf.setFontSize(12);
    var totLabel='Total Time for Services Above:';
    pdf.text(totLabel,x0+5,y+11);
    var lblWidth=pdf.getTextWidth(totLabel);
    pdf.setFont('helvetica','normal');pdf.setFontSize(12);
    pdf.text((data.svcHH||'')+'.'+_padMin(data.svcMM||''),x0+5+lblWidth+10,y+11);
  } else {
    pdf.setFontSize(9);
    var lblY=y+14;
    function dualText(label,val,xPos){
      pdf.setFont('helvetica','bold');pdf.text(label,xPos,lblY);
      var w=pdf.getTextWidth(label);
      pdf.setFont('helvetica','normal');pdf.text(val,xPos+w+4,lblY);
    }
    dualText('Total Time for Services Above:',(data.cplxHH||'')+'.'+_padMin(data.cplxMM||''),x0+5);
    dualText('Total Time from Previous Page:',(data.p1HH||'')+'.'+_padMin(data.p1MM||''),x0+200);
    dualText('Total Time for Billing Period:',(data.grandHH||'')+'.'+_padMin(data.grandMM||''),x0+395);
  }

  // Cert row (no divider above or below) — close to total time row
  pdf.setFont('helvetica','normal');pdf.setFontSize(12);
  pdf.text('I certify that Liberty Home Care Assistance has provided all the services as checked above.',x0+5,y+totH+10);

  // Signature row — NO vertical divider between sig and date.
  // Underline at sigY+44 leaves generous room for a much larger signature image above.
  var sigY=y+totH+certH;
  var sigBoxW=W*0.62;
  var sigUnderlineY=sigY+44;
  // Place signature image — uses pre-cropped (tight) version with known dimensions
  var sigInfo=isPage2?data._sig2Cropped:data._sig1Cropped;
  if(sigInfo&&sigInfo.dataUrl){
    try{
      var aspectRatio=sigInfo.w/sigInfo.h;
      // Generous height up to 40pt, width capped at 85% of sig box for wider rendering
      var sigImgH=40,sigImgW=sigImgH*aspectRatio;
      var maxW=sigBoxW*0.85;
      if(sigImgW>maxW){sigImgW=maxW;sigImgH=sigImgW/aspectRatio;}
      pdf.addImage(sigInfo.dataUrl,'PNG',x0+8,sigUnderlineY-sigImgH-1,sigImgW,sigImgH);
    }catch(e){console.warn('Signature embed failed:',e);}
  }
  // Underlines (the only lines inside the bottom block other than the outer rect)
  pdf.line(x0+5,sigUnderlineY,x0+sigBoxW-5,sigUnderlineY);
  pdf.line(x0+sigBoxW+5,sigUnderlineY,x0+W-5,sigUnderlineY);
  // Date value above the date underline
  var sigDate=isPage2?(data.sigDate2||''):(data.sigDate1||'');
  pdf.setFontSize(12);pdf.text(sigDate,x0+sigBoxW+5,sigUnderlineY-3);
  // Labels a bit further below the underlines; ~13pt whitespace below labels before bottom border
  pdf.setFontSize(12);
  pdf.text('Signature of Authorized Representative',x0+5,sigUnderlineY+13);
  pdf.text('Date',x0+sigBoxW+5,sigUnderlineY+13);

  y+=blockH;

  // ── Footer — MSA form id on left, page number CENTERED (per March) ──
  pdf.setFont('helvetica','normal');pdf.setFontSize(10);
  pdf.text('MSA-1904 (1/2022)',x0,y+13);
  pdf.text(isPage2?'2':'1',x0+W/2,y+13,{align:'center'});
}

// Wrap header label text by greedy word break at maxLen chars
function wrapHeaderText(s,maxLen){
  if(s.length<=maxLen)return [s];
  var words=s.split(' '),lines=[],cur='';
  words.forEach(function(w){
    if((cur+' '+w).trim().length<=maxLen)cur=(cur?cur+' ':'')+w;
    else{if(cur)lines.push(cur);cur=w;}
  });
  if(cur)lines.push(cur);
  return lines;
}

// ============================================================
//  RASTER FALLBACK — html2canvas → JPEG (kept for emergency rollback)
// ============================================================
async function captureInvoicePDFRaster(){
  var jsPDF=window.jspdf.jsPDF;
  var pdf=new jsPDF('p','mm','letter');
  var pageW=195,pageH=257,marginL=10,marginT=11;

  var html=buildInvoiceHTML();
  var blob=new Blob([html],{type:'text/html;charset=utf-8'});
  var blobUrl=URL.createObjectURL(blob);

  var iframe=document.createElement('iframe');
  iframe.style.cssText='position:fixed;left:-9999px;top:0;width:816px;height:4000px;border:none;overflow:hidden;';
  document.body.appendChild(iframe);
  await new Promise(function(resolve,reject){iframe.onload=resolve;iframe.onerror=reject;iframe.src=blobUrl;});
  await new Promise(function(r){setTimeout(r,500);});

  var iDoc=iframe.contentDocument||iframe.contentWindow.document;

  function addPageImage(canvas,pdf,mL,mT,pW,pH){
    var ratio=canvas.height/canvas.width;
    var iW=pW,iH=iW*ratio;
    if(iH>pH){iH=pH;iW=pH/ratio;}
    pdf.addImage(canvas.toDataURL('image/jpeg',0.78),'JPEG',mL,mT,iW,iH);
  }

  var el1=iDoc.getElementById('page1');
  var c1=await html2canvas(el1,{scale:1.5,useCORS:true,allowTaint:true,backgroundColor:'#fff',logging:false});
  addPageImage(c1,pdf,marginL,marginT,pageW,pageH);

  var showCplx=document.getElementById('showComplex').checked;
  var hasCplxChecked=document.querySelectorAll('#cplxTable td.mc.on').length>0;
  if(showCplx&&hasCplxChecked){
    var el2=iDoc.getElementById('complexSection');
    if(el2&&el2.style.display!=='none'){
      pdf.addPage('letter','p');
      var c2=await html2canvas(el2,{scale:1.5,useCORS:true,allowTaint:true,backgroundColor:'#fff',logging:false});
      addPageImage(c2,pdf,marginL,marginT,pageW,pageH);
    }
  }

  document.body.removeChild(iframe);
  URL.revokeObjectURL(blobUrl);
  return pdf.output('datauristring').split(',')[1];
}

// ── Unified entry point — vector by default, raster fallback ──
// Set localStorage 'lhca_pdf_mode' to 'raster' to force the old path.
async function captureInvoicePDF(){
  var mode=(localStorage.getItem('lhca_pdf_mode')||'vector').toLowerCase();
  if(mode==='raster')return captureInvoicePDFRaster();
  try{
    return await captureInvoicePDFVector();
  }catch(e){
    console.error('[PDF] Vector render failed, falling back to raster:',e);
    return captureInvoicePDFRaster();
  }
}

// ── Graph API Email with PDF Attachments ──────────────────────
// ──────────────────────────────────────────────────────────────────
//  Graph email send — auto picks fast path or upload-session path
//  based on total attachment size:
//   - Total ≤ 3.5 MB: single POST /me/sendMail (fast, ~1 sec)
//   - Total > 3.5 MB: createDraft → uploadSession per attachment → send
//                      (handles up to ~150 MB; ~3-10 sec for 30 PDFs)
//  attachments: [{name, base64}], onProgress?: (done,total,label) => void
// ──────────────────────────────────────────────────────────────────
async function sendMailWithPDF(toEmail,subject,bodyHtml,attachments,onProgress,ccEmails){
  if(!spToken)return {ok:false,err:'Not signed in'};
  // Calculate total attachment size in bytes (base64 inflates ~33%, so actual=base64Length*0.75)
  var totalBytes=0;
  attachments.forEach(function(a){totalBytes+=Math.floor(a.base64.length*0.75);});
  var THRESHOLD_BYTES=3.5*1024*1024; // stay safely under Graph's 4 MB sendMail cap
  if(totalBytes<=THRESHOLD_BYTES){
    return _sendMailDirect(toEmail,subject,bodyHtml,attachments,ccEmails);
  }
  console.log('[Graph] Total attachments '+(totalBytes/1024/1024).toFixed(2)+' MB exceeds 3.5 MB — using upload-session flow');
  return _sendMailUploadSession(toEmail,subject,bodyHtml,attachments,onProgress,ccEmails);
}

// Fast path — direct sendMail with attachments inline (≤4 MB total)
async function _sendMailDirect(toEmail,subject,bodyHtml,attachments,ccEmails){
  var ccList=(ccEmails||[]).filter(Boolean).map(function(e){return{emailAddress:{address:e}};});
  var msg={subject:subject,body:{contentType:'HTML',content:bodyHtml},toRecipients:[{emailAddress:{address:toEmail}}],attachments:attachments.map(function(a){return{'@odata.type':'#microsoft.graph.fileAttachment',name:a.name,contentType:'application/pdf',contentBytes:a.base64};})};
  if(ccList.length)msg.ccRecipients=ccList;
  try{
    var resp=await fetch('https://graph.microsoft.com/v1.0/me/sendMail',{method:'POST',headers:{'Authorization':'Bearer '+spToken,'Content-Type':'application/json'},body:JSON.stringify({message:msg,saveToSentItems:true})});
    if(resp.ok||resp.status===202||resp.status===204)return {ok:true};
    var errText=await resp.text();
    console.error('Graph sendMail failed',resp.status,errText);
    return {ok:false,status:resp.status,err:errText};
  }catch(e){console.error('Graph email error:',e);return {ok:false,err:e.message};}
}

// Upload-session path — for emails > 4 MB. Creates draft, uploads each attachment via session, then sends.
async function _sendMailUploadSession(toEmail,subject,bodyHtml,attachments,onProgress,ccEmails){
  function progress(done,total,label){if(typeof onProgress==='function')onProgress(done,total,label);}
  try{
    // 1) Create draft message (no attachments yet)
    progress(0,attachments.length,'Creating draft…');
    var ccList=(ccEmails||[]).filter(Boolean).map(function(e){return{emailAddress:{address:e}};});
    var draftBody={subject:subject,body:{contentType:'HTML',content:bodyHtml},toRecipients:[{emailAddress:{address:toEmail}}]};
    if(ccList.length)draftBody.ccRecipients=ccList;
    var draftResp=await fetch('https://graph.microsoft.com/v1.0/me/messages',{method:'POST',headers:{'Authorization':'Bearer '+spToken,'Content-Type':'application/json'},body:JSON.stringify(draftBody)});
    if(!draftResp.ok){var t=await draftResp.text();return {ok:false,status:draftResp.status,err:'Draft creation failed: '+t};}
    var draft=await draftResp.json();
    var msgId=draft.id;

    // 2) Upload each attachment via createUploadSession
    for(var i=0;i<attachments.length;i++){
      var att=attachments[i];
      progress(i,attachments.length,'Uploading '+att.name+'…');
      // Decode base64 to byte length
      var byteLen=Math.floor(att.base64.length*0.75);
      // a) Create upload session
      var sessResp=await fetch('https://graph.microsoft.com/v1.0/me/messages/'+msgId+'/attachments/createUploadSession',{
        method:'POST',
        headers:{'Authorization':'Bearer '+spToken,'Content-Type':'application/json'},
        body:JSON.stringify({AttachmentItem:{attachmentType:'file',name:att.name,size:byteLen,contentType:'application/pdf'}})
      });
      if(!sessResp.ok){var et=await sessResp.text();return {ok:false,status:sessResp.status,err:'Upload session failed for '+att.name+': '+et};}
      var sess=await sessResp.json();
      var uploadUrl=sess.uploadUrl;
      // b) Decode base64 → bytes (Uint8Array)
      var bin=atob(att.base64);
      var bytes=new Uint8Array(bin.length);
      for(var j=0;j<bin.length;j++)bytes[j]=bin.charCodeAt(j);
      // c) Upload — for files <4 MB, single PUT works; for >4 MB, chunk in 4 MB pieces
      var CHUNK=4*1024*1024;
      if(bytes.length<=CHUNK){
        var putResp=await fetch(uploadUrl,{method:'PUT',headers:{'Content-Type':'application/octet-stream','Content-Range':'bytes 0-'+(bytes.length-1)+'/'+bytes.length},body:bytes});
        if(!putResp.ok&&putResp.status!==201){var pt=await putResp.text();return {ok:false,status:putResp.status,err:'Upload failed for '+att.name+': '+pt};}
      } else {
        // Chunked upload (rare — only triggered for very large signed PDFs)
        for(var off=0;off<bytes.length;off+=CHUNK){
          var end=Math.min(off+CHUNK,bytes.length);
          var slice=bytes.slice(off,end);
          var pr=await fetch(uploadUrl,{method:'PUT',headers:{'Content-Type':'application/octet-stream','Content-Range':'bytes '+off+'-'+(end-1)+'/'+bytes.length},body:slice});
          if(!pr.ok&&pr.status!==201&&pr.status!==202){var prt=await pr.text();return {ok:false,status:pr.status,err:'Chunked upload failed: '+prt};}
        }
      }
    }

    // 3) Send the draft
    progress(attachments.length,attachments.length,'Sending email…');
    var sendResp=await fetch('https://graph.microsoft.com/v1.0/me/messages/'+msgId+'/send',{method:'POST',headers:{'Authorization':'Bearer '+spToken}});
    if(!sendResp.ok&&sendResp.status!==202){var st=await sendResp.text();return {ok:false,status:sendResp.status,err:'Send failed: '+st};}
    return {ok:true};
  }catch(e){console.error('Graph upload-session error:',e);return {ok:false,err:e.message};}
}

// ── Load invoice into page for capture ────────────────────────
async function loadInvoiceForCapture(clientName,inv,period){
  var prof=getProfiles()[clientName]||{};
  var parts=(period||'').split('/');
  rebuild(parts.length===2?daysIn(parts[0],parts[1]):31);
  document.getElementById('clientName').value=clientName;
  document.getElementById('clientName2').value=clientName;
  document.getElementById('billingPeriod').value=period;
  document.getElementById('billingPeriod2').value=period;
  document.getElementById('medicaidId').value=prof.medicaidId||'';
  document.getElementById('hourlyRate').value=stateRate();
  var cwRecCapture=getCaseworkers().find(function(c){return c.id===prof.caseworkerId||c.name===prof.worker;})||{};
  document.getElementById('billTo').value=(cwRecCapture.agency||cwRecCapture.county||'');
  document.getElementById('worker').value=prof.worker||'';
  document.getElementById('worker2').value=prof.worker||'';
  document.getElementById('dateSubmitted').value=today();
  document.getElementById('sigDate1').value=today();
  document.getElementById('sigDate2').value=today();
  if(inv.data&&inv.data.tasks)applyStates(inv.data.tasks);
  var fields=['svcHH','svcMM','cplxHH','cplxMM','p1HH','p1MM','grandHH','grandMM'];
  fields.forEach(function(id){var el=document.getElementById(id);if(el)el.value=(inv.data&&inv.data[id])||'';});
  var hc=inv.data&&inv.data.hasComplex;
  document.getElementById('showComplex').checked=!!hc;
  document.getElementById('complexSection').style.display=hc?'block':'none';
  // Auto-place the user's primary signature (so every captured PDF has a sig)
  // Reset first to avoid stale stamps from prior client; then stamp if any sig exists.
  resetSigArea(1);resetSigArea(2);
  var sigs=getSigs();
  if(sigs.length){
    stampSignatureData(1,sigs[0].data);
    if(hc)stampSignatureData(2,sigs[0].data);
  }
  await new Promise(function(r){setTimeout(r,150);});
}

// ── Mark invoice submitted (local + DB) ───────────────────────
// ──────────────────────────────────────────────────────────────────
//  EMAIL AUDIT LOG — HIPAA "accounting of disclosures"
//  Each PHI-bearing email send writes one log entry. Stored in
//  localStorage and mirrored to App Insights. Downloadable as CSV
//  from Settings > Email Audit Log.
// ──────────────────────────────────────────────────────────────────
function getEmailAuditLog(){try{return JSON.parse(localStorage.getItem('lhca_email_audit')||'[]');}catch(e){return[];}}
function saveEmailAuditLog(arr){
  // Cap at 1000 entries to avoid unbounded growth
  if(arr.length>1000)arr=arr.slice(-1000);
  try{localStorage.setItem('lhca_email_audit',JSON.stringify(arr));}catch(e){console.error('Audit log save failed:',e);}
}
function renderEmailAuditTable(){
  var wrap=document.getElementById('emailAuditTableWrap');if(!wrap)return;
  // Instant render from LS cache, then refresh from DB (source of truth)
  _renderEmailAuditTableFromArray(getEmailAuditLog());
  fetch(API_BASE+'/audit-events?type=email_send&limit=1000',{headers:apiHeaders()})
    .then(function(r){if(!r.ok)throw new Error('HTTP '+r.status);return r.json();})
    .then(function(rows){
      // Unpack metadata blob back into the rendering shape
      var unpacked=rows.map(function(ev){
        try{var m=JSON.parse(ev.metadata||'{}');m.timestamp=m.timestamp||ev.created_at;return m;}
        catch(e){return {timestamp:ev.created_at,recipient:'(parse error)',success:false};}
      });
      _renderEmailAuditTableFromArray(unpacked);
    })
    .catch(function(e){console.error('Email audit DB fetch failed (using LS cache):',e);});
}
function _renderEmailAuditTableFromArray(arr){
  var wrap=document.getElementById('emailAuditTableWrap');if(!wrap)return;
  if(!arr.length){wrap.innerHTML='<div style="padding:14px;color:#5c7590;text-align:center;">No log entries yet.</div>';return;}
  var rows=arr.slice().reverse().map(function(r){
    var dt=new Date(r.timestamp);var when=dt.toLocaleDateString()+' '+dt.toLocaleTimeString([],{hour:'2-digit',minute:'2-digit'});
    var clientList=(r.clientNames||[]).join(', ');
    var statusIcon=r.success?'<span style="color:#1a7740;">✓</span>':'<span style="color:#a00;">✗</span>';
    var detail=r.success?'Sent':('Failed: '+esc((r.errorMsg||'').toString().slice(0,80)));
    return '<tr>'+
      '<td style="padding:6px 8px;">'+statusIcon+'</td>'+
      '<td style="padding:6px 8px;white-space:nowrap;">'+when+'</td>'+
      '<td style="padding:6px 8px;">'+esc(r.recipient||'')+'</td>'+
      '<td style="padding:6px 8px;">'+esc(r.caseworkerName||'')+'</td>'+
      '<td style="padding:6px 8px;white-space:nowrap;">'+esc(r.billingPeriod||'')+'</td>'+
      '<td style="padding:6px 8px;text-align:center;">'+(r.attachmentCount||0)+'</td>'+
      '<td style="padding:6px 8px;">'+esc(clientList)+'</td>'+
      '<td style="padding:6px 8px;color:'+(r.success?'#1a7740':'#a00')+';">'+detail+'</td>'+
    '</tr>';
  }).join('');
  wrap.innerHTML='<table style="width:100%;border-collapse:collapse;font-size:11px;">'+
    '<thead style="background:#f4f8fc;position:sticky;top:0;"><tr>'+
      '<th style="padding:6px 8px;text-align:left;border-bottom:1px solid #d0d8e4;"></th>'+
      '<th style="padding:6px 8px;text-align:left;border-bottom:1px solid #d0d8e4;">When</th>'+
      '<th style="padding:6px 8px;text-align:left;border-bottom:1px solid #d0d8e4;">Recipient</th>'+
      '<th style="padding:6px 8px;text-align:left;border-bottom:1px solid #d0d8e4;">Caseworker</th>'+
      '<th style="padding:6px 8px;text-align:left;border-bottom:1px solid #d0d8e4;">Period</th>'+
      '<th style="padding:6px 8px;text-align:center;border-bottom:1px solid #d0d8e4;">#</th>'+
      '<th style="padding:6px 8px;text-align:left;border-bottom:1px solid #d0d8e4;">Clients</th>'+
      '<th style="padding:6px 8px;text-align:left;border-bottom:1px solid #d0d8e4;">Status</th>'+
    '</tr></thead><tbody>'+rows+'</tbody></table>';
}

function downloadEmailAuditCSV(){
  // Pull canonical history from DB (HIPAA requirement: complete and authoritative)
  fetch(API_BASE+'/audit-events?type=email_send&limit=1000',{headers:apiHeaders()})
    .then(function(r){if(!r.ok)throw new Error('HTTP '+r.status);return r.json();})
    .then(function(rows){
      var arr=rows.map(function(ev){
        try{var m=JSON.parse(ev.metadata||'{}');m.timestamp=m.timestamp||ev.created_at;return m;}
        catch(e){return {timestamp:ev.created_at,recipient:'(parse error)',success:false};}
      });
      if(!arr.length){showAlert('No audit log entries to download.');return;}
      _doEmailAuditCSVDownload(arr);
    })
    .catch(function(e){
      // Fall back to LS cache if DB unreachable
      console.error('Email audit DB fetch failed, using LS cache:',e);
      var arr=getEmailAuditLog();
      if(!arr.length){showAlert('No audit log entries to download.');return;}
      _doEmailAuditCSVDownload(arr);
    });
}
function _doEmailAuditCSVDownload(arr){
  // Excel/Sheets evaluate a cell beginning = + - @ or tab. This file is a HIPAA disclosure register
  // whose cells carry operator-typed names and Graph error text, so neutralise the lead character.
  function csvEscape(v){
    var s=String(v==null?'':v);
    if(/^[=+\-@\t\r]/.test(s))s="'"+s;
    return /[",\n]/.test(s)?'"'+s.replace(/"/g,'""')+'"':s;
  }
  var header=['Timestamp','SentBy','Type','Recipient','CaseworkerName','BillingPeriod','AttachmentCount','ClientNames','Success','ErrorMsg'].join(',');
  var lines=arr.map(function(r){return [
    r.timestamp,r.sentBy,r.type,r.recipient,r.caseworkerName,r.billingPeriod,r.attachmentCount,(r.clientNames||[]).join('; '),r.success?'YES':'NO',r.errorMsg||''
  ].map(csvEscape).join(',');});
  // The backend caps /audit-events at 1000 rows and offers no paging, so this export can silently
  // omit the OLDEST disclosures. Say so on the face of the file rather than shipping a truncated
  // record that looks complete.
  var truncated=(arr.length>=1000);
  var csv=header+'\n'+lines.join('\n')+
    (truncated?'\n\n"WARNING: this export hit the 1000-record limit and is INCOMPLETE — disclosures older than the newest 1000 are NOT included."':'');
  var blob=new Blob([csv],{type:'text/csv;charset=utf-8;'});
  var url=URL.createObjectURL(blob);
  var a=document.createElement('a');a.href=url;
  a.download='email_audit_log_'+new Date().toISOString().slice(0,10)+'.csv';
  document.body.appendChild(a);a.click();document.body.removeChild(a);
  setTimeout(function(){URL.revokeObjectURL(url);},2000);
}

function logEmailSend(entry){
  // entry: { recipient, caseworkerName, billingPeriod, clientNames[], attachmentCount, sentBy, success, errorMsg?, cc? }
  var record=Object.assign({
    timestamp:new Date().toISOString(),
    sentBy:currentUserEmail()   // window.signedInEmail is never assigned anywhere — this logged "unknown" for every disclosure
  },entry);
  // HIPAA-critical — persist to DB FIRST, then mirror to LS as display cache
  var summary='Email to '+(record.recipient||'?')+' · '+((record.clientNames||[]).length)+' client'+((record.clientNames||[]).length===1?'':'s')+' · '+(record.success?'sent':'failed');
  // D13: this DB write is the authoritative HIPAA record (the LS copy is only a
  // display cache). A swallowed failure let an operator believe the trail was
  // complete when the server never got the row — check r.ok and surface it loudly.
  var doAudit=function(){
    return fetch(API_BASE+'/audit-events',{
      method:'POST',
      headers:apiHeaders(),
      body:JSON.stringify({
        event_type:'email_send',
        actor:record.sentBy,
        target_type:record.caseworkerName?'caseworker':'',
        target_id:record.caseworkerName||'',
        summary:summary,
        metadata:JSON.stringify(record)
      })
    }).then(function(r){if(!r.ok)throw new Error('HTTP '+r.status);return r;});
  };
  doAudit().catch(function(e){
    console.error('Email audit DB write failed:',e);
    _showSaveStatus('failed','HIPAA email-audit record not saved to server ('+(e&&e.message?e.message:'network error')+')',doAudit);
  });
  // Mirror to LS for instant display (keep as cache only — DB is source of truth)
  var arr=getEmailAuditLog();arr.push(record);saveEmailAuditLog(arr);
  // App Insights for centralized observability
  try{aiTrack('PHIEmailSend',{
    recipient:record.recipient,
    caseworker:record.caseworkerName,
    period:record.billingPeriod,
    clientCount:(record.clientNames||[]).length,
    success:String(record.success),
    errorMsg:record.errorMsg||''
  });}catch(e){}
  return record;
}

function markInvoiceSubmitted(clientName,period){
  var p=getProfiles();
  if(!p[clientName])return;
  var inv=(p[clientName].invoices||[]).find(function(i){return i.billingPeriod===period;});
  if(!inv||inv.status==='paid')return;
  inv.status='submitted';
  saveProfilesLS(p);
  // saveProfileSP persists the status via the invoice upsert (the changed invoice is now
  // dirty and gets sent). The separate /status PATCH that used to run here double-wrote the
  // row AND discarded the fresh row_version it returned — leaving a stale token that could
  // spuriously 409 the invoice's next save — so it was removed.
  saveProfileSP(clientName,p[clientName]);
}

// ── Email tone helpers (personalized but SAFE — never asserts a fact that could be wrong) ──
// Date-aware sign-off: only fires on unambiguous, always-correct windows; otherwise "Thanks,".
function _emailClose(){
  var d=new Date();
  if(d.getMonth()===0 && d.getDate()<=3) return 'Happy New Year,';
  if(d.getDay()===4 || d.getDay()===5) return 'Have a great weekend,';
  return 'Thanks,';
}
// Deterministic index from a string — STABLE for a given caseworker+month (so a resend
// matches), but different across caseworkers so a batch doesn't read identically to all.
function _emailSeed(s,n){var h=5381;for(var i=0;i<(s||'').length;i++)h=((h*33)^s.charCodeAt(i))>>>0;return h%n;}
// Occasional "here and there" appreciation line (~1 in 3), keyed off the same seed.
function _apprecLine(seed,plural){return _emailSeed(seed+'|a',3)===0?(' I appreciate your help with '+(plural?'these':'this')+', as always.'):'';}
// Standard sign-off block with the date-aware close.
function _emailSig(){return '<p>'+_emailClose()+'<br><b>Thomas Jaboro</b><br>Liberty Home Care Assistance<br>(248) 291-4106</p>';}

// ── Send single invoice email ─────────────────────────────────
async function sendEmail(){
  var cn=document.getElementById('clientName').value.trim();
  if(cn&&isCarrierClient(getProfiles()[cn])){showAlert('This is a managed-care (carrier) client — invoices are handled in the carrier’s software, so there is nothing to email here.');return;}
  var bp=document.getElementById('billingPeriod').value.trim();
  var ae=document.getElementById('activeAgentEmail').value.trim();
  var w=document.getElementById('worker').value.trim();
  // Validate before sending so we don't email blank/incomplete invoices
  if(cn&&bp){
    var profSE=getProfiles()[cn]||{};
    var invSE=(profSE.invoices||[]).find(function(i){return i.billingPeriod===bp;});
    var cwRecSE=getCaseworkers().find(function(c){return c.id===profSE.caseworkerId||c.name===profSE.worker;})||{};
    var issuesSE=validateInvoiceForSend(cn,profSE,invSE,cwRecSE);
    if(issuesSE.length){
      var proceed=false;
      await new Promise(function(res){
        showConfirm(
          'This invoice has issues:\n\n• '+issuesSE.join('\n• ')+'\n\nSend anyway?',
          function(){proceed=true;res();},
          {title:'Invoice Has Issues',okText:'Send Anyway',danger:true,onCancel:function(){res();}}
        );
      });
      if(!proceed)return;
    }
  }
  // If no email cached, try to look up from caseworker record
  if(!ae&&activeProfileName){
    var prof2=getProfiles()[activeProfileName]||{};
    var cwRec2=getCaseworkers().find(function(c){return c.id===prof2.caseworkerId||c.name===prof2.worker;})||{};
    ae=cwRec2.email||'';
    if(ae){var ef=document.getElementById('activeAgentEmail');if(ef)ef.value=ae;}
  }
  if(!ae){showAlert('No caseworker email on file. Add an email to the caseworker record from the Caseworkers page.');return;}
  if(!spToken){
    showConfirm('Sign in with your Microsoft account to send email via Outlook?\n\nClicking Sign In will redirect you to authenticate.',function(){signIn();},{title:'Sign In Required',okText:'Sign In',danger:false});
    return;
  }
  var btn=document.getElementById('sendEmailInvBtn');
  if(btn){btn.disabled=true;btn.textContent='Generating PDF…';}
  try{
    // Persist whatever's currently in the form BEFORE sending — so the saved record matches
    // what the caseworker received. Otherwise the PDF can be emailed with values the user
    // typed but never clicked Save on, and the stored invoice ends up out of sync.
    if(activeProfileName&&bp){
      var pSE=getProfiles();
      if(pSE[activeProfileName]){
        if(!pSE[activeProfileName].invoices)pSE[activeProfileName].invoices=[];
        var idxSE=pSE[activeProfileName].invoices.findIndex(function(i){return i.billingPeriod===bp;});
        var snapshot=captureFullInvoice();
        if(idxSE>=0){
          var existing=pSE[activeProfileName].invoices[idxSE];
          if(existing.status!=='paid'){
            pSE[activeProfileName].invoices[idxSE]=Object.assign({},existing,{savedAt:new Date().toLocaleString(),data:snapshot});
          }
        } else {
          pSE[activeProfileName].invoices.unshift({billingPeriod:bp,savedAt:new Date().toLocaleString(),status:'draft',invoiceNote:'',data:snapshot});
        }
        saveProfilesLS(pSE);saveProfileSP(activeProfileName,pSE[activeProfileName]);
      }
    }
    var base64=await captureInvoicePDF();
    var fname=(cn||'Invoice').replace(/[^a-z0-9]/gi,'_')+'_'+(bp||'').replace('/','_')+'.pdf';
    var _mo=['January','February','March','April','May','June','July','August','September','October','November','December'];
    var _bpP=(bp||'').split('/');
    var bpLabel=(_bpP.length===2&&parseInt(_bpP[0],10)>=1&&parseInt(_bpP[0],10)<=12)?_mo[parseInt(_bpP[0],10)-1]+' '+_bpP[1]:bp;
    var wFirst=(w||'').split(/\s+/)[0];
    // Subject always leads with "Invoice" + the month/year (agency direction), then client.
    var subj='Invoice'+(bpLabel?' '+bpLabel:'')+(cn?' – '+cn:'');
    var body='<p>Hi'+(wFirst?' '+esc(wFirst):'')+',</p>'+
      '<p>Attached is the invoice for <b>'+esc(cn)+'</b> for <b>'+esc(bpLabel)+'</b>. Please review it at your convenience, and let me know if you have any questions or if anything needs adjusting.'+_apprecLine((ae||'')+'|'+(bp||''),false)+'</p>'+
      _emailSig();
    if(btn)btn.textContent='Sending…';
    var result=await sendMailWithPDF(ae,subj,body,[{name:fname,base64:base64}]);
    // HIPAA audit log entry
    logEmailSend({
      type:'single',
      recipient:ae,
      caseworkerName:w||'(none)',
      billingPeriod:bp,
      clientNames:[cn],
      attachmentCount:1,
      success:!!result.ok,
      errorMsg:result.ok?null:(result.err||result.status||'unknown')
    });
    if(result.ok){
      aiTrack('InvoiceEmailed',{clientName:cn,billingPeriod:bp,recipient:ae});
      markInvoiceSubmitted(cn,bp);
      renderOverviewPane();updateStats();
      showToast('✓ Email sent to '+ae+' — invoice marked Submitted');
    }else{
      var msg='Email failed to send.';
      if(result.status===401)msg='Authentication error (401) — please sign out and sign back in.';
      else if(result.status===403)msg='Permission denied (403) — sign out, sign back in, and accept the Mail.Send permission when prompted.';
      else if(result.err)msg='Error ('+( result.status||'?')+'):\n'+result.err.slice(0,300);
      showAlert(msg);
    }
  }catch(e){showAlert('Error generating PDF: '+e.message);}
  if(btn){btn.disabled=false;btn.textContent='✉ Email Worker';}
}
// The government/state hourly rate billed on every invoice (NOT the caregiver pay rate).
// Configurable in Settings so the once-a-year rate change is a setting, not a code edit.
function stateRate(){ var v=(localStorage.getItem('lhca_state_rate')||'').trim(); return v||'27.00'; }
// The rate to bill a client's invoices at: ALWAYS the state rate ($27). The state pays a flat
// provider rate, so every invoice uses it — the client "Hourly Rate" field is NOT a billing rate
// (it does not affect invoices). prof kept for call-site compatibility.
function clientInvoiceRate(prof){
  return stateRate();
}
function saveStateRate(){
  var el=document.getElementById('stateRateInput'); if(!el)return;
  var raw=(el.value||'').replace(/[^0-9.]/g,'').trim();
  var n=parseFloat(raw);
  var st=document.getElementById('stateRateStatus');
  if(!raw||isNaN(n)||n<=0){ if(st){st.style.color='#c0392b';st.textContent='Enter a valid rate';} return; }
  var val=n.toFixed(2);
  localStorage.setItem('lhca_state_rate',val); el.value=val;
  if(st){st.style.color='#5c7590';st.textContent='Saving…';}
  // Report the real result. loadSettingsAPI makes the SERVER authoritative on the next load, so a
  // silent failure here means the rate reverts and every invoice bills at the old rate.
  _pushSettings({state_rate:val}).then(function(){
    if(st){st.style.color='#1a7740';st.textContent='✓ Saved — $'+val+'/hr on all invoices';}
  },function(e){
    if(!st)return;
    if(e&&e.notSignedIn){ st.style.color='#1a7740'; st.textContent='✓ Saved — $'+val+'/hr on all invoices (sign in to sync it to your other devices)'; return; }
    st.style.color='#c0392b';
    st.textContent='⚠ Saved on this device only — the server rejected it ('+((e&&e.message)||'error')+'). It will revert on reload.';
  });
}
// ── App settings sync (state rate + agency) across the owner's devices ──
// These two were localStorage-only, so they didn't follow the owner device-to-device. Now mirrored
// to a backend AppSettings store: pushed on save, pulled on load (server is the source of truth).
// Returns a promise so callers can report the REAL outcome. This used to swallow every failure with a
// console.warn while the UI said "Saved ✓" — and because loadSettingsAPI treats the server as the source
// of truth, a failed push is silently reverted on the next load. For the state rate that means every
// invoice quietly goes back to billing at the old hourly rate.
function _pushSettings(obj){
  if(typeof _apiToken==='undefined' || !_apiToken || typeof API_BASE!=='string'){
    // Not signed in is the normal local-only state, not a sync failure — flag it so callers stay quiet.
    var _e=new Error('not signed in'); _e.notSignedIn=true; return Promise.reject(_e);
  }
  return fetch(API_BASE+'/settings',{method:'POST',headers:apiHeaders(),body:JSON.stringify(obj)})
    .then(function(r){ if(!r.ok) throw new Error('HTTP '+r.status); return r; });
}
function loadSettingsAPI(){
  return fetch(API_BASE+'/settings',{headers:apiHeaders()})
    .then(function(r){ if(!r.ok)throw new Error('HTTP '+r.status); return r.json(); })
    .then(function(s){
      if(!s||typeof s!=='object') return;
      if(s.state_rate!=null && s.state_rate!=='') localStorage.setItem('lhca_state_rate', String(s.state_rate));
      if(s.agency && typeof s.agency==='object') localStorage.setItem('lhca_agency', JSON.stringify(s.agency));
      if(typeof renderAgencySettings==='function' && document.getElementById('ag-name')) renderAgencySettings();
      var sr=document.getElementById('stateRateInput'); if(sr && s.state_rate) sr.value=String(s.state_rate);
    })
    .catch(function(e){ console.warn('Load settings failed (using local):', e); });
}
function copyMonth(){
  var bp=document.getElementById('billingPeriod').value.trim();if(!bp||bp.length<7){showAlert('Enter a billing period first (MM/YYYY).');return;}
  var pts=bp.split('/'),m=parseInt(pts[0]),y=parseInt(pts[1]);m++;if(m>12){m=1;y++;}
  var newBP=String(m).padStart(2,'0')+'/'+y;
  // Smart fill: capture current task state to carry forward
  var states=captureStates();
  var cn=document.getElementById('clientName').value;
  var mid=document.getElementById('medicaidId').value;
  var hr=document.getElementById('hourlyRate').value;
  var bt=document.getElementById('billTo').value; // carries forward the address string
  var wk=document.getElementById('worker').value;
  var hc=document.getElementById('showComplex').checked;
  // If current form has no billing period set yet, try pulling from most recent saved invoice
  if(!bp&&activeProfileName){
    var prof=getProfiles()[activeProfileName];
    if(prof&&prof.invoices&&prof.invoices.length){
      var latest=prof.invoices[0];
      if(latest.data&&latest.data.tasks)states=latest.data.tasks;
    }
  }
  ['svcHH','svcMM','cplxHH','cplxMM','p1HH','p1MM','grandHH','grandMM'].forEach(function(id){var e=document.getElementById(id);if(e)e.value='';});
  var T2=today();
  document.getElementById('billingPeriod').value=newBP;document.getElementById('billingPeriod2').value=newBP;
  document.getElementById('clientName').value=cn;document.getElementById('clientName2').value=cn;
  document.getElementById('medicaidId').value=mid;
  document.getElementById('hourlyRate').value=stateRate();document.getElementById('billTo').value=bt;
  document.getElementById('worker').value=wk;document.getElementById('worker2').value=wk;
  document.getElementById('dateSubmitted').value=T2;document.getElementById('sigDate1').value=T2;document.getElementById('sigDate2').value=T2;
  document.getElementById('showComplex').checked=hc;toggleComplex();
  rebuild(daysIn(String(m).padStart(2,'0'),String(y)));applyStates(states);resetSigArea(1);resetSigArea(2);
}

// ============================================================
//  SIGNATURE (multi-sig from Settings)
// ============================================================
function initSigCanvas(){
  sigCanvas=document.getElementById('sigCanvas');sigCtx=sigCanvas.getContext('2d');sigCtx.strokeStyle='#000';sigCtx.lineWidth=2;sigCtx.lineCap='round';sigCtx.lineJoin='round';
  function gp(e){var r=sigCanvas.getBoundingClientRect(),src=e.touches?e.touches[0]:e;return{x:src.clientX-r.left,y:src.clientY-r.top};}
  sigCanvas.addEventListener('mousedown',function(e){sigDrawing=true;var p=gp(e);sigCtx.beginPath();sigCtx.moveTo(p.x,p.y);});
  sigCanvas.addEventListener('mousemove',function(e){if(!sigDrawing)return;var p=gp(e);sigCtx.lineTo(p.x,p.y);sigCtx.stroke();});
  sigCanvas.addEventListener('mouseup',function(){sigDrawing=false;});sigCanvas.addEventListener('mouseleave',function(){sigDrawing=false;});
  sigCanvas.addEventListener('touchstart',function(e){e.preventDefault();sigDrawing=true;var p=gp(e);sigCtx.beginPath();sigCtx.moveTo(p.x,p.y);},{passive:false});
  sigCanvas.addEventListener('touchmove',function(e){e.preventDefault();if(!sigDrawing)return;var p=gp(e);sigCtx.lineTo(p.x,p.y);sigCtx.stroke();},{passive:false});
  sigCanvas.addEventListener('touchend',function(){sigDrawing=false;});
}
function placeSignature(target){
  pendingSigTarget=target;
  var sigs=getSigs();
  if(!sigs.length){openAddSigModal();return;}
  if(sigs.length===1){stampSignatureData(target,sigs[0].data);return;}
  var list=document.getElementById('pickSigList');list.innerHTML='';
  sigs.forEach(function(s,i){
    var item=document.createElement('div');item.style.cssText='display:flex;align-items:center;gap:10px;padding:8px;border:1px solid #e1e5ea;border-radius:6px;cursor:pointer;';
    item.innerHTML='<img src="'+s.data+'" style="max-height:26px;max-width:140px;"><span style="font-size:12px;color:#1a2b45;">'+esc(s.label||'Signature '+(i+1))+'</span>';
    item.addEventListener('click',function(){stampSignatureData(target,s.data);document.getElementById('pickSigModal').classList.remove('open');});
    list.appendChild(item);
  });
  document.getElementById('pickSigModal').classList.add('open');
}
function stampSignatureData(target,dataUrl){
  var area=document.getElementById('sigArea'+target);if(!area)return;
  var img=document.createElement('img');img.src=dataUrl;img.className='sig-stamp';img.title='Click to clear';img.id='sigArea'+target;
  img.addEventListener('click',function(){resetSigArea(target);});area.parentNode.replaceChild(img,area);
}
function resetSigArea(t){var a=document.getElementById('sigArea'+t);if(!a)return;var ph=document.createElement('div');ph.id='sigArea'+t;ph.className='sig-placeholder';ph.textContent='Click to place signature';ph.addEventListener('click',function(){placeSignature(t);});a.parentNode.replaceChild(ph,a);}
// Trim a canvas to its non-transparent content (tight bounding box + small padding).
// Returns a data URL, or null if canvas is fully transparent.
function trimCanvasToContent(canvas,padding){
  padding=padding==null?4:padding;
  var w=canvas.width,h=canvas.height;
  var ctx=canvas.getContext('2d');
  var imgData;try{imgData=ctx.getImageData(0,0,w,h).data;}catch(e){return null;}
  var minX=w,minY=h,maxX=-1,maxY=-1;
  for(var y=0;y<h;y++){
    for(var x=0;x<w;x++){
      var idx=(y*w+x)*4;
      // alpha threshold OR non-white pixel (covers white-background canvases too)
      if(imgData[idx+3]>10&&!(imgData[idx]>240&&imgData[idx+1]>240&&imgData[idx+2]>240)){
        if(x<minX)minX=x;if(x>maxX)maxX=x;
        if(y<minY)minY=y;if(y>maxY)maxY=y;
      }
    }
  }
  if(maxX<0)return null;
  minX=Math.max(0,minX-padding);minY=Math.max(0,minY-padding);
  maxX=Math.min(w-1,maxX+padding);maxY=Math.min(h-1,maxY+padding);
  var nw=maxX-minX+1,nh=maxY-minY+1;
  var crop=document.createElement('canvas');crop.width=nw;crop.height=nh;
  // Fill white background so PDFs don't show transparency artifacts
  var cctx=crop.getContext('2d');cctx.fillStyle='#fff';cctx.fillRect(0,0,nw,nh);
  cctx.drawImage(canvas,minX,minY,nw,nh,0,0,nw,nh);
  return crop.toDataURL('image/png');
}

// Async version for stored data URLs (used at PDF embed time for legacy padded sigs)
// Returns { dataUrl, w, h } — natural dimensions of the cropped image
function cropDataURL(dataUrl){
  return new Promise(function(resolve){
    if(!dataUrl){resolve(null);return;}
    var img=new Image();
    img.onload=function(){
      var c=document.createElement('canvas');c.width=img.width;c.height=img.height;
      c.getContext('2d').drawImage(img,0,0);
      var trimmed=trimCanvasToContent(c,4);
      if(!trimmed){resolve({dataUrl:dataUrl,w:img.width,h:img.height});return;}
      // Load the trimmed image to get its dimensions
      var t=new Image();
      t.onload=function(){resolve({dataUrl:trimmed,w:t.width,h:t.height});};
      t.onerror=function(){resolve({dataUrl:trimmed,w:img.width,h:img.height});};
      t.src=trimmed;
    };
    img.onerror=function(){resolve(null);};
    img.src=dataUrl;
  });
}

function confirmSig(){
  var data=sigCanvas.toDataURL('image/png'),blank=document.createElement('canvas');blank.width=sigCanvas.width;blank.height=sigCanvas.height;
  if(blank.toDataURL()===data){showAlert('Please draw your signature first.');return;}
  var trimmed=trimCanvasToContent(sigCanvas,4);if(trimmed)data=trimmed;
  var label=document.getElementById('sigLabel').value.trim()||'Signature';
  var id=sigId();
  var sigs=getSigs();sigs.push({id:id,label:label,data:data});saveSigsLS(sigs);
  // D6/D12: identity comes from Easy Auth (apiHeaders), NOT the Graph token —
  // persist unconditionally and surface a failure/retry instead of saving to LS only.
  trackSave('signature',function(){
    return fetch(API_BASE+'/signatures',{method:'POST',headers:apiHeaders(),body:JSON.stringify({id:id,label:label,data_url:data})})
      .then(function(r){if(!r.ok)throw new Error('HTTP '+r.status);return r;});
  });
  closeSigModal();
  if(pendingSigTarget!==null){stampSignatureData(pendingSigTarget,data);}
  if(document.getElementById('page-settings').classList.contains('active'))renderSigSettings();
}
function clearSigPad(){if(sigCtx)sigCtx.clearRect(0,0,sigCanvas.width,sigCanvas.height);}
function closeSigModal(){document.getElementById('sigModal').classList.remove('open');}

// ── Signature modal tab switching ────────────────────────────────
function switchSigTab(tab){
  var panes={draw:'sigPaneDraw',type:'sigPaneType',upload:'sigPaneUpload'};
  var tabs={draw:'sigTabDraw',type:'sigTabType',upload:'sigTabUpload'};
  Object.keys(panes).forEach(function(k){
    var p=document.getElementById(panes[k]);if(p)p.style.display=(k===tab?'':'none');
    var t=document.getElementById(tabs[k]);if(t){
      var active=(k===tab);
      t.style.cssText='flex:1;padding:7px;border:none;cursor:pointer;font-size:13px;font-family:Arial,sans-serif;background:'+(active?'#1a2b45':'#f0f3f7')+';color:'+(active?'#fff':'#1a2b45')+';';
    }
  });
  if(tab==='type'){previewCursiveSig();}
}

// ── Upload signature handling ───────────────────────────────────
window._sigUploadOriginal=null; // raw uploaded image (HTMLImageElement)
function handleSigUpload(ev){
  var file=ev.target.files&&ev.target.files[0];
  if(!file)return;
  if(!/^image\//.test(file.type)){showAlert('Please choose an image file (PNG or JPG).');return;}
  var hint=document.getElementById('sigUploadHint');if(hint)hint.textContent=file.name;
  var reader=new FileReader();
  reader.onload=function(e){
    var img=new Image();
    img.onload=function(){window._sigUploadOriginal=img;rerenderSigUpload();};
    img.onerror=function(){showAlert('Could not load that image. Try a different file.');};
    img.src=e.target.result;
  };
  reader.readAsDataURL(file);
}

// Re-render preview using current settings (line removal toggle)
function rerenderSigUpload(){
  var img=window._sigUploadOriginal;if(!img)return;
  var removeLine=document.getElementById('sigUploadRemoveLine').checked;
  var c=document.createElement('canvas');c.width=img.width;c.height=img.height;
  var ctx=c.getContext('2d');
  ctx.fillStyle='#fff';ctx.fillRect(0,0,c.width,c.height); // flatten transparency to white bg
  ctx.drawImage(img,0,0);
  if(removeLine)removeHorizontalLinesFromCanvas(c);
  var trimmed=trimCanvasToContent(c,8);
  // Draw trimmed result into preview canvas
  var preview=document.getElementById('sigUploadCanvas');
  var pctx=preview.getContext('2d');
  pctx.clearRect(0,0,preview.width,preview.height);
  pctx.fillStyle='#fff';pctx.fillRect(0,0,preview.width,preview.height);
  if(trimmed){
    var pImg=new Image();
    pImg.onload=function(){
      // Fit preserving aspect ratio inside 390x80 preview
      var sx=preview.width/pImg.width,sy=preview.height/pImg.height;
      var s=Math.min(sx,sy,1);
      var dw=pImg.width*s,dh=pImg.height*s;
      pctx.drawImage(pImg,(preview.width-dw)/2,(preview.height-dh)/2,dw,dh);
      window._sigUploadProcessed=trimmed;
    };
    pImg.src=trimmed;
  } else {
    pctx.fillStyle='#aaa';pctx.font='12px Arial';
    pctx.fillText('Image looks empty after processing.',10,40);
    window._sigUploadProcessed=null;
  }
}

// Detect rows that are mostly dark (a horizontal line) and clear them.
// Used to strip signature lines from scanned signatures.
function removeHorizontalLinesFromCanvas(c){
  var ctx=c.getContext('2d'),w=c.width,h=c.height;
  var data;try{data=ctx.getImageData(0,0,w,h);}catch(e){return;}
  var px=data.data;
  // For each row, count dark pixels (R+G+B < 300, alpha > 100)
  for(var y=0;y<h;y++){
    var dark=0;
    for(var x=0;x<w;x++){
      var i=(y*w+x)*4;
      if(px[i+3]>100&&(px[i]+px[i+1]+px[i+2])<300)dark++;
    }
    // Heuristic: row is a "line" if >=50% of its pixels are dark
    if(dark>=w*0.5){
      // Erase this row + 2px above/below (typical line is 1-3px tall)
      for(var dy=-2;dy<=2;dy++){
        var yy=y+dy;if(yy<0||yy>=h)continue;
        for(var x2=0;x2<w;x2++){
          var i2=(yy*w+x2)*4;
          px[i2]=255;px[i2+1]=255;px[i2+2]=255;px[i2+3]=255;
        }
      }
    }
  }
  ctx.putImageData(data,0,0);
}

function confirmUploadedSig(){
  var data=window._sigUploadProcessed;
  if(!data){showAlert('Please upload an image first.');return;}
  var label=document.getElementById('sigLabel').value.trim()||'Uploaded Signature';
  var id=sigId();
  var sigs=getSigs();sigs.push({id:id,label:label,data:data});saveSigsLS(sigs);
  // D6/D12: identity comes from Easy Auth (apiHeaders), NOT the Graph token —
  // persist unconditionally and surface a failure/retry instead of saving to LS only.
  trackSave('signature',function(){
    return fetch(API_BASE+'/signatures',{method:'POST',headers:apiHeaders(),body:JSON.stringify({id:id,label:label,data_url:data})})
      .then(function(r){if(!r.ok)throw new Error('HTTP '+r.status);return r;});
  });
  closeSigModal();
  if(pendingSigTarget!==null){stampSignatureData(pendingSigTarget,data);}
  if(document.getElementById('page-settings').classList.contains('active'))renderSigSettings();
}

// ── Cursive typed signature preview ──────────────────────────────
function previewCursiveSig(){
  var name=(document.getElementById('sigTypeName')||{}).value||'';
  var fontSel=document.getElementById('sigTypeFont');
  var fontFamily=fontSel?fontSel.value:'Caveat';
  var tc=document.getElementById('sigTypeCanvas');if(!tc)return;
  var ctx=tc.getContext('2d');
  ctx.clearRect(0,0,tc.width,tc.height);
  if(!name)return;
  // Real-signature fonts (Allura, Italianno, Great Vibes) need larger sizes to read well
  var sizeBoost=(fontFamily==='Allura'||fontFamily==='Italianno'||fontFamily==='Great Vibes'||fontFamily==='Mr Dafoe')?1.25:1.0;
  // Cedarville Cursive renders smaller naturally — boost it for legibility
  if(fontFamily==='Cedarville Cursive')sizeBoost=1.4;
  var fontSize=Math.min(56,Math.max(32,Math.floor(tc.width*0.11*sizeBoost)));
  ctx.font=fontSize+'px "'+fontFamily+'", cursive';
  ctx.fillStyle='#000';
  ctx.textBaseline='middle';
  var metrics=ctx.measureText(name);
  if(metrics.width>tc.width-16){
    fontSize=Math.floor(fontSize*(tc.width-16)/metrics.width);
    ctx.font=fontSize+'px "'+fontFamily+'", cursive';
    metrics=ctx.measureText(name);
  }
  var x=Math.max(8,(tc.width-metrics.width)/2);
  ctx.fillText(name,x,tc.height/2);
}

// ── Save cursive typed signature ──────────────────────────────────
function confirmTypedSig(){
  var name=(document.getElementById('sigTypeName')||{}).value.trim();
  if(!name){showAlert('Please type your name first.');return;}
  var tc=document.getElementById('sigTypeCanvas');
  var blank=document.createElement('canvas');blank.width=tc.width;blank.height=tc.height;
  if(blank.toDataURL()===tc.toDataURL()){showAlert('Please type your name first.');return;}
  // Trim padding around the typed text — keeps the signature tight when embedded in PDF
  var data=trimCanvasToContent(tc,4)||tc.toDataURL('image/png');
  var label=document.getElementById('sigLabel').value.trim()||name;
  var id=sigId();
  var sigs=getSigs();sigs.push({id:id,label:label,data:data});saveSigsLS(sigs);
  // D6/D12: identity comes from Easy Auth (apiHeaders), NOT the Graph token —
  // persist unconditionally and surface a failure/retry instead of saving to LS only.
  trackSave('signature',function(){
    return fetch(API_BASE+'/signatures',{method:'POST',headers:apiHeaders(),body:JSON.stringify({id:id,label:label,data_url:data})})
      .then(function(r){if(!r.ok)throw new Error('HTTP '+r.status);return r;});
  });
  closeSigModal();
  if(pendingSigTarget!==null){stampSignatureData(pendingSigTarget,data);}
  if(document.getElementById('page-settings').classList.contains('active'))renderSigSettings();
}

// ============================================================
//  PRINT GUARDS — fix page 3, page 2 only if filled
// ============================================================
window.addEventListener('beforeprint',function(){
  // Hide complex section if no complex tasks filled in
  var cplxRows=document.getElementById('cplxBody').querySelectorAll('td.mc.on');
  var hasComplex=cplxRows.length>0;
  var cs=document.getElementById('complexSection');
  if(!hasComplex){cs.setAttribute('data-print-hidden','1');cs.style.display='none';}

  // Swap column headers to rotated-span technique so print matches email PDF exactly
  document.querySelectorAll('.tt th.th').forEach(function(th){
    th.dataset.printOrigHtml=th.innerHTML;
    th.dataset.printOrigStyle=th.getAttribute('style')||'';
    var text=th.textContent.trim();
    th.setAttribute('style','height:110px;padding:0;overflow:hidden;vertical-align:middle;writing-mode:initial;transform:none;');
    th.innerHTML=
      '<div style="position:relative;width:100%;height:110px;overflow:hidden;">'+
        '<span style="position:absolute;left:50%;top:50%;display:inline-block;'+
        'width:105px;text-align:center;white-space:normal;word-break:break-word;line-height:1.25;'+
        'font-size:8pt;font-weight:normal;font-family:\'Times New Roman\',Times,serif;'+
        'transform:translate(-50%,-50%) rotate(-90deg);">'+
        text+
        '</span>'+
      '</div>';
  });
});
window.addEventListener('afterprint',function(){
  var cs=document.getElementById('complexSection');
  if(cs.getAttribute('data-print-hidden')==='1'){
    cs.removeAttribute('data-print-hidden');
    if(document.getElementById('showComplex').checked)cs.style.display='block';
  }
  document.querySelectorAll('.tt th.th').forEach(function(th){
    if(th.dataset.printOrigHtml!==undefined){
      th.innerHTML=th.dataset.printOrigHtml;
      if(th.dataset.printOrigStyle){th.setAttribute('style',th.dataset.printOrigStyle);}
      else{th.removeAttribute('style');}
      delete th.dataset.printOrigHtml;
      delete th.dataset.printOrigStyle;
    }
  });
});

// ============================================================
//  AUTH & API (Azure Functions replaces SharePoint for data)
// ============================================================
function initMSAL() {
  var cfg = {
    auth: {
      clientId: SP_CLIENT_ID,
      authority: 'https://login.microsoftonline.com/' + SP_TENANT_ID,
      redirectUri: REDIRECT_URI,
    },
    cache: {
      cacheLocation: 'localStorage',
      // iOS Safari Intelligent Tracking Prevention wipes sessionStorage mid-redirect
      // during the OAuth dance. Storing auth state in a cookie survives the trip.
      storeAuthStateInCookie: true,
    },
  };
  msalInstance = new msal.PublicClientApplication(cfg);
  msalInstance.initialize().then(function () {
    msalInstance.handleRedirectPromise().then(function (r) {
      if (r && r.accessToken) {
        var email = (r.account && r.account.username || '').toLowerCase();
        if (!ALLOWED_USERS.map(function(u){return u.toLowerCase();}).includes(email)) {
          msalInstance.logoutRedirect(); return;
        }
        window._aiUser = email;
        if(window.appInsights&&window.appInsights.setAuthenticatedUserContext)window.appInsights.setAuthenticatedUserContext(email);
        aiTrack('UserSignIn',{method:'redirect'});
        // Detect whether this redirect returned a Graph token or an API token
        var isApiToken = (r.scopes||[]).some(function(s){return String(s).indexOf('user_impersonation')>=0;});
        if (isApiToken) {
          _apiToken = r.accessToken;
          var ttl = r.expiresOn ? (r.expiresOn.getTime() - Date.now() - 600000) : 3000000;
          _apiTokenTimer = setTimeout(refreshApiToken, Math.max(ttl, 60000));
          // Now also need Graph token for email — try silent
          msalInstance.acquireTokenSilent({ scopes: GRAPH_SCOPES, account: r.account })
            .then(function(g){ spToken = g.accessToken; setTimeout(maybeAutoBackupOneDrive,8000); }).catch(function(){});
          updateAuthUI(true); loadProfilesAPI();
        } else {
          spToken = r.accessToken;
          refreshApiToken().then(function(){ updateAuthUI(true); loadProfilesAPI(); setTimeout(maybeAutoBackupOneDrive,8000); });
        }
        return;
      }
      var acc = msalInstance.getAllAccounts();
      if (acc.length) {
        var email2 = (acc[0].username || '').toLowerCase();
        if (!ALLOWED_USERS.map(function(u){return u.toLowerCase();}).includes(email2)) {
          msalInstance.logoutRedirect(); return;
        }
        window._aiUser = email2;
        if(window.appInsights&&window.appInsights.setAuthenticatedUserContext)window.appInsights.setAuthenticatedUserContext(email2);
        msalInstance.acquireTokenSilent({
          scopes: GRAPH_SCOPES, // Graph only — API token fetched separately in refreshApiToken()
          account: acc[0], redirectUri: REDIRECT_URI,
        }).then(function (r2) { spToken = r2.accessToken; aiTrack('UserSignIn',{method:'silent'}); refreshApiToken().then(function(){ updateAuthUI(true); loadProfilesAPI(); setTimeout(maybeAutoBackupOneDrive,8000); }); })
          .catch(function () {
            msalInstance.loginRedirect({ scopes: GRAPH_SCOPES, redirectUri: REDIRECT_URI });
          });
      } else { updateAuthUI(false); }
    }).catch(function () { updateAuthUI(false); });
  });
}
function signIn() {
  // Only Graph scopes here — mixing API_SCOPE causes 400 on token endpoint (different resource).
  // refreshApiToken() acquires API token silently after login (admin consent already granted).
  msalInstance.loginRedirect({
    scopes: GRAPH_SCOPES,
    redirectUri: REDIRECT_URI,
  });
}
function clearPHIFromStorage() {
  // F4: before wiping, fire any pending note save (it reads the note from LS synchronously),
  // so a note typed just before an idle-wipe / sign-out / tab-close still reaches the backend.
  try{ if(typeof _flushPendingNoteSaves==='function') _flushPendingNoteSaves(); }catch(_){}
  // Remove ALL PHI/PII from localStorage (HIPAA idle-wipe / sign-out). Wipe every lhca_*
  // key EXCEPT an explicit whitelist of non-PHI *settings*, so a PHI key added later is
  // covered automatically instead of silently surviving. (The old fixed list missed
  // lhca_email_audit, lhca_supervisors, and lhca_autogen_undo — all client-identifying.)
  // KEEP = column widths, page sizes, the state billing rate, PDF-mode preference, and the
  // weekly-backup timestamp (a plain date, not PHI — wiping it made the "weekly" OneDrive
  // backup re-fire on every session, since the wipe erased its memory of the last run).
  var KEEP = /(_col_widths|_page_size)$|^lhca_state_rate$|^lhca_pdf_mode$|^lhca_last_onedrive_backup$|^lhca_agency$/;
  Object.keys(localStorage)
    .filter(function(k){ return k.indexOf('lhca_') === 0 && !KEEP.test(k); })
    .forEach(function(k){ try{ localStorage.removeItem(k); }catch(e){} });
  // S8: also wipe the in-memory SSN caches so nothing lingers after sign-out/idle.
  // _clientSyncedMem holds the dirty-tracking signatures, which embed SSN — wipe it too.
  _ssnMem = Object.create(null);
  _clientSyncedMem = Object.create(null);
  _profilesCache = null;   // the profiles store was just wiped — drop the read cache too
  if (typeof _cgSsnMem !== 'undefined') _cgSsnMem = Object.create(null);
  // PHI also sits in the OPEN FORM. Wiping storage while an SSN / MiLogin password / client detail
  // stayed readable in an input behind the login wall left PHI on screen after "sign out".
  try{
    // Close any PHI-bearing popup this session opened (caregiver task sheets). They are separate
    // documents, so wiping storage and inputs here never touched them.
    try{
      for(var _w=0;_w<_phiWindows.length;_w++){
        try{ if(_phiWindows[_w] && !_phiWindows[_w].closed) _phiWindows[_w].close(); }catch(_e){}
      }
      _phiWindows=[];
    }catch(e){}
    var _els=document.querySelectorAll('input, textarea');
    for(var _i=0;_i<_els.length;_i++){
      var _el=_els[_i], _t=(_el.type||'').toLowerCase();
      if(_t==='checkbox'||_t==='radio'||_t==='button'||_t==='submit'||_t==='hidden')continue;
      try{ _el.value=''; }catch(_e){}
    }
  }catch(e){}
}
function signOut() {
  aiTrack('UserSignOut');
  clearPHIFromStorage();
  spToken = null;
  _apiToken = null;
  clearTimeout(_apiTokenTimer); _apiTokenTimer = null;
  msalInstance.logoutPopup({ redirectUri: REDIRECT_URI });
}
function updateAuthUI(on) {
  applyInvoiceAdminVisibility();
  var wall = document.getElementById('loginWall');
  var wallMsg = document.getElementById('loginWallMsg');
  var wallBtn = document.getElementById('loginWallBtn');
  if (on) {
    // Authenticated — hide wall, show app, start inactivity timer
    if (wall) wall.style.display = 'none';
    resetSessionTimer();
  } else {
    // Not authenticated — stop timer, clear PHI, show wall
    clearTimeout(_sessionTimer); clearTimeout(_sessionWarnTimer);
    clearPHIFromStorage();
    if (wall) wall.style.display = 'flex';
    if (wallMsg) wallMsg.textContent = 'Sign in with your Microsoft account to access client data.';
    if (wallBtn) wallBtn.style.display = 'block';
  }
  var b = document.getElementById('authBtn'), s = document.getElementById('authStatus');
  if (b) { b.textContent = on ? 'Sign Out' : 'Sign In'; b.onclick = on ? signOut : signIn; }
  if (s) { s.textContent = on ? '✓ Signed in — Azure SQL sync active' : 'Not signed in'; s.style.color = on ? '#6dcf95' : '#435f7a'; }
  var lsl = document.getElementById('lastSyncedLabel');
  if (lsl) { var ls = localStorage.getItem('lhca_last_synced'); lsl.textContent = on && ls ? 'Last synced: ' + ls : (on ? 'Syncing…' : ''); }
  updateSettingsAuth();
}

// ── DB sync-state tracker — disables Add buttons during initial load ──
window._dbSyncPending=0;
function syncStart(){
  window._dbSyncPending++;
  // Only show the full-page banner on initial cold load — when LS is genuinely empty
  // and Add buttons would be unsafe. After the first sync completes, LS has data and
  // background revalidation on nav should be silent (Phase-5 SWR).
  if(!window._initialLoadDone)document.body.classList.add('db-syncing');
}
function syncEnd(){
  window._dbSyncPending=Math.max(0,window._dbSyncPending-1);
  if(window._dbSyncPending===0){
    document.body.classList.remove('db-syncing');
    // Mark initial load complete — every subsequent sync is background revalidation
    window._initialLoadDone=true;
  }
}
// ── DB load-error banner: surface a failed load instead of leaving a silent empty page ──
function showDbError(what){
  var el=document.getElementById('dbErrorBanner');
  if(!el){
    el=document.createElement('div');
    el.id='dbErrorBanner';
    el.style.cssText='position:fixed;top:0;left:0;right:0;z-index:9999;background:#fdecec;color:#a00;padding:9px 16px;font-family:Arial,sans-serif;font-size:13px;border-bottom:1px solid #f0c0c0;box-shadow:0 2px 6px rgba(0,0,0,0.08);display:flex;align-items:center;justify-content:center;gap:12px;';
    document.body.appendChild(el);
  }
  el.innerHTML='<span>⚠️ Couldn’t reach the database'+(what?' ('+esc(what)+')':'')+'. Your data is safe — this is usually a brief connection hiccup.</span>'+
    '<button class="btn btn-secondary btn-sm" id="_dbRetryBtn" style="padding:3px 12px;">Retry</button>'+
    '<button id="_dbErrDismiss" style="background:none;border:none;color:#a00;cursor:pointer;font-size:16px;line-height:1;padding:0 4px;">✕</button>';
  el.style.display='flex';
  var rb=document.getElementById('_dbRetryBtn');
  if(rb)rb.onclick=function(){hideDbError();if(typeof loadProfilesAPI==='function')loadProfilesAPI();};
  var db=document.getElementById('_dbErrDismiss');
  if(db)db.onclick=hideDbError;
}
function hideDbError(){var el=document.getElementById('dbErrorBanner');if(el)el.style.display='none';}
// ── SWR: revalidate DB → LS on nav, throttled to avoid hammering on rapid clicks ──
// Skipped if any unsaved-changes flag is active so an in-flight edit doesn't get clobbered.
var _lastRevalidate=0;
var REVALIDATE_THROTTLE_MS=30000; // 30 sec — bandwidth vs freshness trade-off for a 3-user CRM
function revalidate(force){
  // Guard 1: avoid clobbering in-progress edits
  if(typeof unsavedChanges!=='undefined'&&unsavedChanges)return;
  if(typeof cgUnsavedChanges!=='undefined'&&cgUnsavedChanges)return;
  if(typeof cwUnsavedChanges!=='undefined'&&cwUnsavedChanges)return;
  // Guard 1b: never reload while a save is still in flight. unsavedChanges is cleared
  // synchronously (before the async save resolves) and invoice-only saves don't set it at
  // all — so this counter is what actually prevents a background reload from reverting a
  // pending save's writes (dbId/rowVersion writebacks + saved content).
  if(_savesInFlight>0)return;
  // Guard 2: throttle to once per 30 sec unless forced
  var now=Date.now();
  if(!force&&(now-_lastRevalidate)<REVALIDATE_THROTTLE_MS)return;
  _lastRevalidate=now;
  // Need an auth token to actually fetch — silently skip on cold start
  if(!spToken)return;
  // loadProfilesAPI already fans out to caregivers/caseworkers/supervisors/tasks/signatures
  // internally, so calling those here too double-fetched each of them every revalidate. Just
  // call loadProfilesAPI — it re-renders all the entity UIs.
  if(typeof loadProfilesAPI==='function')loadProfilesAPI();
}

// ── LOAD all clients + invoices from Azure SQL (bulk fetch) ───────────────
// One round-trip instead of N+1. Backend returns clients with invoices nested.
function loadProfilesAPI() {
  syncStart();
  // Load the supporting datasets in parallel. These calls were previously placed AFTER
  // the return below, making them unreachable dead code — so caregivers/caseworkers/etc.
  // never loaded on sign-in, only when their tab was first opened.
  if (typeof loadCaregiversAPI === 'function') loadCaregiversAPI();
  if (typeof loadCaseworkersAPI === 'function') loadCaseworkersAPI();
  if (typeof loadSupervisorsAPI === 'function') loadSupervisorsAPI();
  if (typeof loadSettingsAPI === 'function') loadSettingsAPI();   // state rate + agency, synced across devices
  if (typeof loadTasksAPI === 'function') loadTasksAPI();
  if (typeof loadSignaturesAPI === 'function') loadSignaturesAPI();
  // Return the promise so callers (e.g. OneDrive backup) can await freshness
  return fetch(API_BASE + '/homecare-clients-with-invoices', { headers: apiHeaders() })
    .then(function (r) {
      if (!r.ok) throw new Error('HTTP ' + r.status);
      return r.json();
    })
    .then(function (clients) {
      if (!Array.isArray(clients)) throw new Error('Bulk fetch returned non-array');
      var idMap = getIdMap();
      var profiles = {};
      clients.forEach(function (c) {
        var name = c.client_name;
        idMap[name] = c.id;
        // Dedupe invoices by billing_period — keep the newest, schedule old duplicates for DB delete
        var invs = (c.invoices || []).slice().sort(function (a, b) {
          return new Date(b.saved_at || 0) - new Date(a.saved_at || 0);
        });
        // Dedupe by billing period for DISPLAY only. NEVER delete from the DB on a read
        // path — a same-period row may be a legitimate corrected re-issue. Hide only an
        // EXACT duplicate (identical invoice_data); if the content differs, show both.
        var seenByPeriod = {};
        var dedupedInvs = [];
        invs.forEach(function (inv) {
          var key = (inv.billing_period || '').trim();
          if (!key) { dedupedInvs.push(inv); return; }
          var prior = seenByPeriod[key];
          if (prior) {
            if ((prior.invoice_data || '') === (inv.invoice_data || '')) return; // exact repeat — hide, keep DB intact
            dedupedInvs.push(inv); // different content, same period — keep both visible
            return;
          }
          seenByPeriod[key] = inv;
          dedupedInvs.push(inv);
        });
        var mappedInvs = dedupedInvs.map(function (inv) {
          var data = null; try { data = JSON.parse(inv.invoice_data || 'null'); } catch (e) {}
          var m = {
            dbId: inv.id, billingPeriod: inv.billing_period || '',
            status: inv.status || 'draft', invoiceNote: inv.invoice_note || '',
            savedAt: inv.saved_at ? new Date(inv.saved_at).toLocaleString() : '', data: data,
            // Optimistic-concurrency token: sent back on save so a stale write is rejected
            // (409) instead of clobbering another user's edit. The bundled load endpoint
            // returns it as row_version_hex; the /invoices endpoint as a hex row_version
            // string. A raw binary row_version (SELECT *) is NOT a usable token, so ignore
            // that shape.
            rowVersion: inv.row_version_hex || (typeof inv.row_version === 'string' ? inv.row_version : null),
          };
          // Dirty-tracking baseline: this is the server-confirmed state, so mark it synced.
          m._synced = _invoiceSig(m);
          return m;
        });
        profiles[name] = {
          clientName: name, firstName: c.first_name || '', lastName: c.last_name || '',
          middleName: c.middle_name || '', nickname: c.nickname || '',
          medicaidId: c.medicaid_id || '', medicare: c.medicare || '', hourlyRate: c.hourly_rate || '',
          program: c.program || '', carrier: c.carrier || '', memberId: c.member_id || '',
          worker: c.worker || '', caseworkerId: c.caseworker_id || '',
          street: c.street || '', city: c.city || '', state: c.state || '',
          zip: c.zip || '', county: c.county || '',
          phone: c.phone || '', homePhone: c.home_phone || '', clientEmail: c.client_email || '', caregiverId: c.caregiver_id || '',
          dob: c.dob || '', gender: c.gender || '',
          driversLicense: c.drivers_license || '', ssn: c.ssn || '',
          startDate: c.start_date || '', liveIn: !!c.live_in,
          clientStatus: c.client_status || 'active', hasComplex: !!c.has_complex,
          clientNotes: c.client_notes || '', _dbId: c.id,
          authorization: _parseAuth(c.dhs_authorization),
          // Optimistic-concurrency token (see saveProfileSP). Tolerant of both the bundled
          // (row_version_hex) and plain (/homecare-clients) endpoints.
          _rowVersion: c.row_version_hex || (typeof c.row_version === 'string' ? c.row_version : null),
          invoices: mappedInvs,
        };
        // Client-record dirty-tracking baseline (server-confirmed state) — mirror of the
        // per-invoice _synced above. saveProfileSP skips the client POST while this matches.
        profiles[name]._clientSynced = _clientSig(profiles[name]);
      });
      // A save may have started DURING this fetch (the round-trip is async). If so, abort
      // the destructive full-store replace — otherwise it reverts the pending save's local
      // writes (content + dbId/rowVersion writebacks). A later revalidate will refresh.
      if (_savesInFlight > 0) { syncEnd(); return; }
      // Merge (not blind-replace) so a client save that FAILED — its data left only in LS after
      // _savesInFlight fell back to 0 — isn't reverted by this background refresh. Server wins for
      // clean synced clients; unsynced local adds/edits are preserved; deletes still propagate.
      saveProfilesLS(_mergeProfilesLoad(profiles, getProfiles()));
      localStorage.setItem('lhca_id_map', JSON.stringify(idMap));
      hideDbError(); // fresh data arrived — clear any stale connection-error banner
      var now = new Date().toLocaleString(); localStorage.setItem('lhca_last_synced', now);
      renderSidebarClients(); renderClientTable(); updateStats();
      var lsl = document.getElementById('lastSyncedLabel');
      if (lsl) lsl.textContent = 'Last synced: ' + now;
      _restoreRouteAfterLoad();
      syncEnd();
    })
    .catch(function (e) {
      console.error('Bulk load error:', e);
      showDbError('clients');
      renderSidebarClients(); renderClientTable(); updateStats();
      syncEnd();
    });
}

// ── SAVE client profile to Azure SQL ────────────────────────
// Returns a Promise. Shows save-status toast (Saving → Saved ✓ / Save failed [Retry]).
// On failure the LS write done by the caller is preserved so the user can keep working
// against the optimistic state, and the Retry button re-invokes this exact save.
function saveProfileSP(name, data, quiet) {
  var idMap = getIdMap();
  var dbId = data._dbId || idMap[name];
  var body = {
    id: dbId || undefined, client_name: name,
    first_name: data.firstName || '', last_name: data.lastName || '',
    middle_name: data.middleName || '', nickname: data.nickname || '',
    medicaid_id: data.medicaidId || '', medicare: data.medicare || '', hourly_rate: data.hourlyRate || '',
    program: data.program || '', carrier: data.carrier || '', member_id: data.memberId || '',
    worker: data.worker || '', caseworker_id: data.caseworkerId || '',
    street: data.street || '', city: data.city || '', state: data.state || '',
    zip: data.zip || '', county: data.county || '',
    phone: data.phone || '', home_phone: data.homePhone || '', client_email: data.clientEmail || '', caregiver_id: data.caregiverId || '',
    client_status: data.clientStatus || 'active', has_complex: data.hasComplex ? 1 : 0,
    dob: data.dob || '', gender: data.gender || '',
    drivers_license: data.driversLicense || '', ssn: data.ssn || '',
    start_date: data.startDate || '', live_in: data.liveIn ? 1 : 0,
    client_notes: data.clientNotes || '',
    // DHS-1210 authorization (hours/dates/tasks). Sent as an object; the backend
    // JSON.stringify's it. null when the client has no imported authorization yet.
    dhs_authorization: data.authorization || null,
  };
  // S8: only send ssn when we actually have it in memory. Omitting it makes the backend
  // keep the stored (encrypted) value rather than blanking it — so a save that happens
  // before ssn has loaded (e.g. right after a reload) can never wipe it.
  if (!data.ssn) delete body.ssn;
  // Optimistic concurrency: send the row_version we last read for an EXISTING client so a
  // stale save is rejected (409) instead of clobbering another user's edit. Omitted for a
  // new client, or one loaded before this feature (backend then updates unconditionally).
  if (dbId && data._rowVersion) body.expected_version = data._rowVersion;
  // Client-record dirty-tracking: if this is an EXISTING client whose own fields are
  // unchanged since the last confirmed save, skip the client POST entirely. That POST
  // otherwise re-writes the row and bumps its row_version on EVERY save (even an invoice-
  // only one), which made a concurrent invoice edit falsely 409 at the client level. The
  // invoices still sync (they have their own dirty-tracking + version check).
  var _clientSigNow = _clientSig(data);
  var _clientUnchanged = dbId && data._clientSynced === _clientSigNow;
  var _doSave = function(){
    if (_clientUnchanged) {
      var now0 = new Date().toLocaleString(); localStorage.setItem('lhca_last_synced', now0);
      var lsl0 = document.getElementById('lastSyncedLabel'); if (lsl0) lsl0.textContent = 'Last synced: ' + now0;
      return syncNewInvoices(name, data).then(function(){ return { skipped: true }; });
    }
    // keepalive: a note save flushed on pagehide/tab-close (F4 flush via clearPHIFromStorage)
    // must survive the page dying — a normal fetch is aborted on unload and the note is lost
    // (LS is wiped right after). The body here is one client record (no invoice blobs), so it's
    // well under keepalive's 64KB cap. Harmless for ordinary saves.
    return fetch(API_BASE + '/homecare-clients', { method: 'POST', headers: apiHeaders(), body: JSON.stringify(body), keepalive: true })
      .then(function (r) {
        if (r.status === 409) {
          // Someone else changed this client since we loaded it. The stale write was
          // REJECTED (not clobbered) — surface it so the user reloads + re-applies rather
          // than silently overwriting the other person's change.
          var ce = new Error("This client's info was changed by someone else. Reload to get the latest, then re-apply your edit.");
          ce.isConflict = true;
          throw ce;
        }
        if (!r.ok) throw new Error('HTTP ' + r.status);
        return r.json();
      })
      .then(function (result) {
        if (!dbId && result.id) {
          idMap[name] = result.id; localStorage.setItem('lhca_id_map', JSON.stringify(idMap));
          var p = getProfiles(); if (p[name]) { p[name]._dbId = result.id; saveProfilesLS(p); }
        }
        // Refresh the concurrency token so the NEXT save carries the current one (both in
        // memory and in the stored profile), otherwise the next edit would falsely 409.
        if (result.row_version) {
          data._rowVersion = result.row_version;
          var pv = getProfiles(); if (pv[name]) { pv[name]._rowVersion = result.row_version; saveProfilesLS(pv); }
        }
        // Client write confirmed — advance the dirty-tracking baseline to the state we just
        // sent, so the NEXT save skips the client POST if nothing changes.
        data._clientSynced = _clientSigNow;
        var pcs = getProfiles(); if (pcs[name]) { pcs[name]._clientSynced = _clientSigNow; saveProfilesLS(pcs); }
        aiTrack('ClientInfoUpdated',{clientName:name,clientStatus:body.client_status});
        var now = new Date().toLocaleString(); localStorage.setItem('lhca_last_synced', now);
        var lsl = document.getElementById('lastSyncedLabel'); if (lsl) lsl.textContent = 'Last synced: ' + now;
        // Sync any invoices not yet in DB. CHAIN it (was fire-and-forget) so a failed
        // invoice write rejects the whole save → trackSave shows "Save failed [Retry]"
        // instead of a false "Saved ✓" while the invoice is silently lost.
        return syncNewInvoices(name, data).then(function () { return result; });
      });
  };
  // quiet = auto-save (notes typing): persist without the corner toast — the inline
  // "Saved" indicator already covers it. Deliberate saves keep the toast.
  // Track in-flight so a background reload can't clobber this save (see revalidate/
  // loadProfilesAPI guards). Decrement on settle (success OR failure).
  _savesInFlight++;
  var _p;
  // try/finally so a SYNCHRONOUS throw (before _p is assigned) still schedules the
  // decrement — otherwise the counter would stick >0 and block every future background
  // reload permanently.
  try { _p = quiet ? _doSave() : trackSave(name, _doSave); }
  finally { Promise.resolve(_p).then(function(){ _savesInFlight--; }, function(){ _savesInFlight--; }); }
  return _p;
}
// Signature of an invoice's PERSISTED fields (exactly what syncNewInvoices sends). Used
// for dirty-tracking so an invoice is re-sent only when it actually changed since its last
// server-confirmed save (or it has no dbId yet). Deliberately covers billing period,
// status, note, AND the data blob — a status-only or note-only edit must count as dirty.
function _invoiceSig(inv) {
  return JSON.stringify([
    inv.billingPeriod || '', inv.status || 'draft', inv.invoiceNote || '',
    inv.data ? JSON.stringify(inv.data) : '',
  ]);
}
// Signature of a client's OWN persisted fields (everything saveProfileSP sends EXCEPT the
// invoices, which have their own dirty-tracking). Lets an invoice-only save skip re-POSTing
// the client record — which otherwise bumps the client row_version on every save and made a
// concurrent invoice edit falsely 409 at the CLIENT level. ssn IS included so an ssn-only
// edit is never skipped (never lost); the cost is only a harmless re-send if ssn lazy-loads
// into memory after the load-time baseline.
function _clientSig(d) {
  return JSON.stringify([
    d.firstName||'', d.lastName||'', d.middleName||'', d.nickname||'', d.medicaidId||'', d.medicare||'',
    d.hourlyRate||'', d.worker||'', d.caseworkerId||'', d.street||'', d.city||'', d.state||'',
    d.zip||'', d.county||'', d.phone||'', d.homePhone||'', d.clientEmail||'', d.caregiverId||'',
    d.clientStatus||'active', d.hasComplex?1:0, d.dob||'', d.gender||'', d.driversLicense||'',
    d.ssn||'', d.startDate||'', d.liveIn?1:0, d.clientNotes||'',
    d.program||'', d.carrier||'', d.memberId||'',
    d.authorization?JSON.stringify(d.authorization):'',
  ]);
}
// Does a locally-cached client profile hold changes not yet confirmed by the server? True if its
// own fields differ from the last-synced baseline, or ANY invoice is new (no dbId) or edited since
// its baseline. Used to protect a pending/failed save from being reverted by a background load's
// store refresh. Errs toward TRUE (keep local) on any uncertainty — the safe direction is to never
// discard local work.
function _profileHasUnsyncedChanges(loc){
  if(!loc) return false;
  try{
    // Only judge client-field dirtiness when there's an actual saved baseline to compare against.
    // _clientSynced is MEMORY-ONLY (stripped from localStorage, like SSN), so on the FIRST load of a
    // session it's absent for every client — without this guard the merge treated every client as
    // "unsynced" and kept the local (SSN-stripped) copy instead of the fresh server copy, making the
    // SSN field show blank. With no baseline, the server is authoritative.
    if(typeof _clientSig==='function' && loc._clientSynced!=null && loc._clientSynced !== _clientSig(loc)) return true;
    var invs = loc.invoices || [];
    for(var i=0;i<invs.length;i++){
      var inv = invs[i]; if(!inv) continue;
      if(!inv.dbId) return true;                                                    // never-synced new invoice
      if(typeof _invoiceSig==='function' && inv._synced !== _invoiceSig(inv)) return true; // edited since baseline
    }
  }catch(e){ return true; }
  return false;
}
// Merge a freshly-loaded client store over the local one WITHOUT reverting unsynced local work.
// loadProfilesAPI used to blind-replace, so a client save that FAILED (data left only in LS, with
// _savesInFlight already back to 0) was silently wiped by the next background revalidate. Rules,
// mirroring the roster/task merges: server wins for clean synced clients; a local client with
// unsynced changes is kept; a local-only unsynced ADD (no _dbId) is kept; a local client that WAS
// synced (_dbId) but is gone from the server was deleted elsewhere → dropped.
function _mergeProfilesLoad(serverProfiles, localProfiles){
  var out = {}, name;
  for(name in serverProfiles) if(_rosterHas(serverProfiles,name)) out[name]=serverProfiles[name];
  for(name in localProfiles){
    if(!_rosterHas(localProfiles,name)) continue;
    var loc = localProfiles[name]; if(!loc) continue;
    if(_rosterHas(out,name)){
      if(_profileHasUnsyncedChanges(loc)){
        // Keep the pending local copy — but SSN is stripped from localStorage (PHI), so on a COLD
        // load loc has no SSN. Backfill it from the fresh server copy so keeping-local never blanks
        // the SSN. Only when loc has none: an in-session SSN edit lives in _ssnMem and wins.
        var srv=out[name];
        if(srv && !loc.ssn && srv.ssn) loc.ssn=srv.ssn;
        out[name]=loc;                                                  // protect a pending/failed edit
      }
    } else if(!loc._dbId || _profileHasUnsyncedChanges(loc)){
      out[name]=loc;                                                    // unsynced add (or edited row deleted elsewhere)
    }
    // else: previously-synced, clean, and gone from the server → deleted elsewhere → drop.
  }
  return out;
}
// DHS-1210 authorization JSON may arrive as a string (from SQL) or already-parsed object.
function _parseAuth(v){
  if(!v)return null;
  if(typeof v==='object')return v;
  try{return JSON.parse(v);}catch(e){return null;}
}
// Persist a client's invoices to SQL. New invoices (no dbId) INSERT; existing ones (with
// dbId) send their id so the backend UPDATE branch fires — this is what makes EDITS to a
// saved invoice (amount/status/note) actually reach the DB. Resolves only when EVERY write
// succeeds and REJECTS otherwise, so saveProfileSP reports a real "Save failed" instead of a
// false "Saved ✓". Existing invoices send the row_version last seen; the backend rejects a
// stale write with 409 (optimistic concurrency). Fresh row_versions + new dbIds are written
// back to LS in one pass. Every save path routes through here via saveProfileSP.
function syncNewInvoices(name, data) {
  var idMap = getIdMap(); var clientDbId = idMap[name];
  if (!clientDbId || !data.invoices) return Promise.resolve();
  var writeBacks = [];   // {isNew, dbId, period, savedAt, newId, newVersion, sig}
  var conflicts = [];    // billing periods the server rejected as stale (409)
  var hardError = null;  // first non-conflict failure (network / 5xx)
  return Promise.all(data.invoices.map(function (inv) {
    // Dirty-tracking: an existing invoice (has dbId) whose persisted fields are unchanged
    // since its last confirmed save is SKIPPED — so editing a client field or a note no
    // longer re-pushes every invoice (which caused write amplification and false conflicts
    // on rows the user never touched). A brand-new invoice (no dbId) is always sent.
    var sig = _invoiceSig(inv);
    if (inv.dbId && inv._synced === sig) return Promise.resolve();
    var payload = {
      homecare_client_id: clientDbId,
      billing_period: inv.billingPeriod || '',
      status: inv.status || 'draft',
      invoice_note: inv.invoiceNote || '',
      invoice_data: inv.data ? JSON.stringify(inv.data) : null,
    };
    var isNew = !inv.dbId;
    // HIGH#2 defense: an invoice with a server rowVersion but NO dbId lost its id (e.g. a
    // dropped writeback). Sending it as "new" would hit the backend's no-id MERGE, which
    // overwrites the matching (client, period) row UNCONDITIONALLY — bypassing the 409
    // concurrency guard that only the id-UPDATE branch enforces. Refuse to send it as new;
    // surface a conflict so the user reloads (which restores the dbId) instead of silently
    // clobbering whatever is on the server for that period.
    if (isNew && inv.rowVersion) { conflicts.push(inv.billingPeriod || '(no period)'); return Promise.resolve(); }
    if (!isNew) {
      payload.id = inv.dbId; // existing -> UPDATE; else INSERT
      // Send the concurrency token when we have one; omit for invoices loaded before this
      // feature (the backend then does an unconditional update and returns a fresh token).
      if (inv.rowVersion) payload.expected_version = inv.rowVersion;
    }
    var period = inv.billingPeriod || '', savedAt = inv.savedAt;
    return fetch(API_BASE + '/invoices', {
      method: 'POST', headers: apiHeaders(), body: JSON.stringify(payload),
    }).then(function (r) {
      if (r.status === 409) { conflicts.push(period || '(no period)'); return null; }
      if (!r.ok) throw new Error('HTTP ' + r.status);
      return r.json();
    }).then(function (result) {
      if (result) writeBacks.push({ isNew: isNew, dbId: inv.dbId, period: period, savedAt: savedAt, newId: result.id, newVersion: result.row_version, sig: sig });
    }).catch(function (e) {
      console.error('Sync invoice error (' + period + '):', e);
      if (!hardError) hardError = e;
    });
  })).then(function () {
    // One LS pass: persist new dbIds + fresh row_versions by invoice IDENTITY (dbId for
    // existing; period + savedAt for brand-new) so the next save carries the right token.
    if (writeBacks.length) {
      var p2 = getProfiles();
      if (p2[name] && p2[name].invoices) {
        writeBacks.forEach(function (w) {
          var cand;
          if (w.isNew) {
            cand = p2[name].invoices.filter(function (x) { return !x.dbId && (x.billingPeriod || '') === w.period && x.savedAt === w.savedAt; });
            if (cand.length !== 1) cand = p2[name].invoices.filter(function (x) { return !x.dbId && (x.billingPeriod || '') === w.period; });
            if (cand.length === 1 && w.newId) { cand[0].dbId = w.newId; if (w.newVersion) cand[0].rowVersion = w.newVersion; cand[0]._synced = w.sig; }
          } else {
            cand = p2[name].invoices.filter(function (x) { return x.dbId === w.dbId; });
            if (cand.length === 1) { if (w.newVersion) cand[0].rowVersion = w.newVersion; cand[0]._synced = w.sig; }
          }
        });
        saveProfilesLS(p2);
      }
    }
    if (conflicts.length) {
      // Another user changed these invoices since we loaded them. The stale write was
      // REJECTED (not clobbered) — surface it so the user reloads + re-applies rather than
      // losing their edit or silently overwriting the other person's.
      var ce = new Error('Invoice ' + conflicts.join(', ') + ' was changed by someone else. Reload to get the latest, then re-apply your edit.');
      ce.isConflict = true;
      throw ce;
    }
    if (hardError) throw hardError;
  });
}

// ── DELETE client from Azure SQL ─────────────────────────────
function deleteProfileSP(name) {
  var dbId = getIdMap()[name]; if (!dbId) return;
  // D10: only drop the id-map entry after a confirmed OK; surface failure otherwise.
  // Re-READ the map inside the callback instead of closing over a snapshot: a bulk delete fires N of
  // these in one synchronous pass, and each used to write back its own stale copy — so the last
  // resolver restored the N-1 entries the others had removed. A client later re-created under a
  // resurrected name then picked up a dead dbId and every save from that device 404'd for good.
  var doDel=function(){return surfaceSaveFailure(
    fetch(API_BASE + '/homecare-clients/' + dbId, { method: 'DELETE', headers: apiHeaders() })
      .then(function (r) {
        if(!r.ok)throw new Error('HTTP '+r.status);
        var m = getIdMap(); delete m[name]; localStorage.setItem('lhca_id_map', JSON.stringify(m));
        return r;
      }),
    'Delete client',doDel);};
  doDel();
}

// ── INVOICE status update via API ───────────────────────────
// (updateInvoiceStatusAPI removed — status now persists through the invoice upsert in
// saveProfileSP/syncNewInvoices; a separate status PATCH double-wrote the row and
// self-conflicted under optimistic concurrency. The PATCH /invoices/{id}/status backend
// route still exists but is no longer called from the client.)
function deleteInvoiceAPI(dbId, clientName, billingPeriod, onDeleted) {
  if (!dbId) { if (onDeleted) onDeleted(); return; }
  aiTrack('InvoiceDeleted',{invoiceDbId:dbId,clientName:clientName||'',billingPeriod:billingPeriod||''});
  // D10: surface a failure/retry so a "deleted" invoice can't reappear on next sync.
  // onDeleted runs ONLY on a confirmed 2xx — so the local removal happens after the server
  // delete succeeds, never before (a failed delete must not leave the row gone locally but
  // alive on the server, where the next load would resurrect it).
  var doDel=function(){return surfaceSaveFailure(
    fetch(API_BASE + '/invoices/' + dbId, { method: 'DELETE', headers: apiHeaders() })
      .then(function(r){if(!r.ok)throw new Error('HTTP '+r.status); if(onDeleted) onDeleted(); return r;}),
    'Delete invoice',doDel);};
  doDel();
}

// ── CAREGIVERS API ───────────────────────────────────────────
// ── Roster load merge ────────────────────────────────────────────────────────
// A fresh server roster used to REPLACE the local cache wholesale. On a cold device a worker can
// add a row while the initial load is still in flight; the load then lands and wipes the just-added
// row. These merges keep server data authoritative for rows the server knows about, but preserve a
// LOCAL-ONLY row *only when it's an unsynced addition* (no _rowVersion yet). A local row that once
// had a _rowVersion but is now absent from the server was deleted on another device → we drop it,
// so cross-device deletes still propagate.
// A row whose save FAILED (or never finished) is flagged `_unsaved` by _rosterMarkUnsaved. That row
// exists ONLY in localStorage, so the server copy is STALE — the merge must keep the local one, the
// same protection _mergeProfilesLoad gives clients. Without it, the next background load quietly
// reverted the edit the user thought they'd made.
function _rosterHas(o, k){ return Object.prototype.hasOwnProperty.call(o, k); }
function _rosterKeepLocal(row){ return !!row && (!row._rowVersion || row._unsaved === true); }
function _mergeRosterMap(serverMap, localMap){   // caregivers, supervisors (id-keyed objects)
  var out = {}, k;
  for (k in serverMap) if (_rosterHas(serverMap, k)) out[k] = serverMap[k];
  for (k in localMap){
    if (!_rosterHas(localMap, k) || !localMap[k]) continue;
    if (_rosterHas(out, k)) { if (localMap[k]._unsaved === true) out[k] = localMap[k]; }  // failed save wins
    else if (_rosterKeepLocal(localMap[k])) out[k] = localMap[k];                          // unsynced add
  }
  return out;
}
function _mergeRosterArr(serverArr, localArr){   // caseworkers (array of {id,…})
  var idx = {}; (serverArr || []).forEach(function(x, i){ if (x && x.id != null) idx[x.id] = i; });
  var out = (serverArr || []).slice();
  (localArr || []).forEach(function(x){
    if (!x || x.id == null) return;
    if (idx[x.id] != null) { if (x._unsaved === true) out[idx[x.id]] = x; }   // failed save wins
    else if (_rosterKeepLocal(x)) out.push(x);                                 // unsynced add
  });
  return out;
}
// A 409 is NOT a failed save in the sense the merge cares about: it means the SERVER holds a NEWER
// row than this device did. Marking it `_unsaved` pinned the stale local copy (and its stale
// _rowVersion) over the server's, so every retry re-sent the same stale version and 409'd again —
// the record could never be saved on that device. Conflicts must let the server copy through; only
// genuine failures (network, 5xx) mean "the local copy is the newer one, keep it".
function _isConflict(e){ return !!(e && e.isConflict); }
// Flag/clear the "this row's save failed" marker used by the merges above.
function _rosterMarkUnsaved(kind, id, failed){
  try{
    if (kind === 'caseworker'){
      var cws = getCaseworkers(); var c = cws.find(function(x){ return x && String(x.id) === String(id); });
      if (!c) return; if (failed) c._unsaved = true; else delete c._unsaved; saveCaseworkersLS(cws);
    } else if (kind === 'supervisor'){
      var sups = getSupervisors(); if (!sups[id]) return;
      if (failed) sups[id]._unsaved = true; else delete sups[id]._unsaved; saveSupervisorsLS(sups);
    } else {
      var cgs = getCaregivers(); if (!cgs[id]) return;
      if (failed) cgs[id]._unsaved = true; else delete cgs[id]._unsaved; saveCaregiversLS(cgs);
    }
  }catch(e){ /* storage full / parse error — the merge just falls back to server-wins */ }
}
function loadCaregiversAPI() {
  syncStart();
  fetch(API_BASE + '/caregivers', { headers: apiHeaders() })
    .then(function (r) {
      if (!r.ok) throw new Error('HTTP ' + r.status);
      return r.json();
    })
    .then(function (cgs) {
      if (!Array.isArray(cgs)) { console.warn('Caregivers API returned non-array:', cgs); syncEnd(); return; }
      var obj = {};
      cgs.forEach(function (cg) {
        obj[cg.id] = {
          name: cg.name, firstName: cg.first_name || '', lastName: cg.last_name || '',
          middleName: cg.middle_name || '', nickname: cg.nickname || '',
          status: cg.status, phone: cg.phone, email: cg.email,
          hireDate: cg.hire_date || cg.start_date || '',
          emptype: cg.emptype || cg.role || '',
          payRate: cg.pay_rate || '', maxHours: cg.max_hours || '',
          certs: cg.certifications || '',
          ecName: cg.ec_name || '', ecPhone: cg.ec_phone || '',
          champsId: cg.champs_id || '', gender: cg.gender || '',
          miloginUsername: cg.milogin_username || '', miloginPassword: cg.milogin_password || '',
          street: cg.street || '', city: cg.city || '',
          state: cg.state || '', zip: cg.zip || '', county: cg.county || '',
          dob: cg.dob || '', driversLicense: cg.drivers_license || '',
          // The list endpoint no longer returns the full SSN (minimum necessary) — only the last 4,
          // which is what the masked export needs. The full value is fetched when the field is focused.
          ssn: cg.ssn || '', ssnLast4: cg.ssn_last4 || '',
          notes: cg.notes || '',
          _rowVersion: cg.row_version_hex || null   // optimistic-concurrency token
        };
      });
      // A save may have started DURING this fetch — its writes aren't in this response, so merging
      // now would show pre-save data (mirrors the guard in loadProfilesAPI).
      if (_savesInFlight > 0) { syncEnd(); return; }
      // Merge (not replace) so an unsynced local addition isn't wiped by the load; the merge also
      // keeps local rows when the server response is transiently empty.
      saveCaregiversLS(_mergeRosterMap(obj, getCaregivers()));
      // Repaint the grid now that fresh data is in — matches loadProfilesAPI/loadCaseworkersAPI.
      // Without this, a cold cache (new device/URL, or after the idle-timeout clears storage)
      // renders an empty grid and never refreshes when the fetch lands.
      if (typeof renderCaregiverGrid === 'function' && document.getElementById('cgTableBody')) renderCaregiverGrid();
      // Also refresh the dashboard count + client table caregiver column, so a cold load
      // never shows a false "0 caregivers" or blank caregiver names once the data lands.
      if (typeof updateStats === 'function') updateStats();
      if (typeof renderClientTable === 'function' && document.getElementById('clientTableBody')) renderClientTable();
      _restoreRouteAfterLoad();
      syncEnd();
    }).catch(function (e) { console.error('Load caregivers error:', e); showDbError('caregivers'); syncEnd(); });
}
// Roster saves (caregiver/caseworker/supervisor) join the SAME in-flight counter that client
// saves use (saveProfileSP), so a background revalidate / loadProfilesAPI can't reload the roster
// mid-save and briefly show pre-save data. Increment synchronously as the save starts; decrement
// when it settles (success OR failure). 409/data-loss is already handled by row_version; this only
// closes the transient-staleness window.
function _trackRosterSave(p){
  _savesInFlight++;
  Promise.resolve(p).then(function(){_savesInFlight--;},function(){_savesInFlight--;});
  return p;
}
function saveCaregiverAPI(id, cg, quiet) {
  var _doSave = function(){
    var body = {
        id: id, name: cg.name || '',
        first_name: cg.firstName || '', last_name: cg.lastName || '',
        middle_name: cg.middleName || '', nickname: cg.nickname || '',
        role: cg.emptype || '', emptype: cg.emptype || '',
        phone: cg.phone || '', email: cg.email || '',
        start_date: cg.hireDate || '', hire_date: cg.hireDate || '',
        status: cg.status || 'active', notes: cg.notes || '',
        street: cg.street || cg.address || '', city: cg.city || '',
        state: cg.state || '', zip: cg.zip || '', county: cg.county || '',
        pay_rate: cg.payRate || '', max_hours: cg.maxHours || '',
        certifications: cg.certs || cg.certifications || '',
        ec_name: cg.ecName || '', ec_phone: cg.ecPhone || '',
        champs_id: cg.champsId || '', gender: cg.gender || '',
        milogin_username: cg.miloginUsername || '', milogin_password: cg.miloginPassword || '',
        // Identity (sensitive — these fields are masked in exports too)
        dob: cg.dob || '', drivers_license: cg.driversLicense || '', ssn: cg.ssn || ''
    };
    // S8: only send ssn when present; omitting it makes the backend keep the stored value.
    if (!cg.ssn) delete body.ssn;
    // Optimistic concurrency: send the row_version we last read so a stale save is rejected
    // (409) instead of silently overwriting another user's edit.
    if (cg._rowVersion) body.expected_version = cg._rowVersion;
    return fetch(API_BASE + '/caregivers', {
      method: 'POST', headers: apiHeaders(),
      body: JSON.stringify(body),
    }).then(function(r){
      if (r.status === 409) {
        var ce = new Error("This caregiver's info was changed by someone else. Reload to get the latest, then re-apply your edit.");
        ce.isConflict = true; throw ce;
      }
      if(!r.ok)throw new Error('HTTP '+r.status);
      return r.json();
    }).then(function(result){
      // Refresh the concurrency token so the NEXT save carries the current one.
      if (result && result.row_version) {
        cg._rowVersion = result.row_version;
        var cgs = getCaregivers(); if (cgs[id]) { cgs[id]._rowVersion = result.row_version; saveCaregiversLS(cgs); }
      }
      return result;
    });
  };
  var _pr = quiet ? _doSave() : trackSave(cg.name||id, _doSave);
  _pr.then(function(){ _rosterMarkUnsaved('caregiver', id, false); },
           function(e){ _rosterMarkUnsaved('caregiver', id, !_isConflict(e)); });  // handled here; caller keeps _pr
  return _trackRosterSave(_pr);
}
function deleteCaregiverAPI(id) {
  // D10: surface a failure/retry so a "deleted" caregiver can't reappear on next sync.
  var doDel=function(){return surfaceSaveFailure(
    fetch(API_BASE + '/caregivers/' + id, { method: 'DELETE', headers: apiHeaders() })
      .then(function(r){if(!r.ok)throw new Error('HTTP '+r.status);return r;}),
    'Delete caregiver',doDel);};
  doDel();
}

// ── TASKS API ────────────────────────────────────────────────
function loadTasksAPI() {
  fetch(API_BASE + '/tasks?source=homecare', { headers: apiHeaders() })
    .then(function (r) { if (!r.ok) throw new Error('HTTP ' + r.status); return r.json(); })
    .then(function (tasks) {
      var todos = (Array.isArray(tasks) ? tasks : []).map(function (t) {
        // CRITICAL: id MUST be a string to match the format used by todoId() (e.g. 'td_xxx_yyy').
        // If id is a number, all task button click handlers (deleteTodo / toggleTodo / etc.)
        // fail their strict-equality lookup and silently do nothing.
        return {
          id: String(t.id), dbId: t.id, text: t.task_text, done: !!t.done, doneAt: t.done_at || null,
          due: t.due_date ? t.due_date.split('T')[0] : '', client: t.client_name || '',
          priority: t.priority || 'normal', note: t.note || '',
          parentId: t.parent_id ? String(t.parent_id) : null,
        };
      });
      // Merge so a task created on this device while the load was in flight isn't wiped; a task
      // that had a dbId but is gone from the server was deleted elsewhere, so it's dropped.
      // Same guard every other loader has: a save that started during this fetch isn't in the
      // response, so merging it now reverts the just-ticked task.
      if(_savesInFlight>0)return;
      saveTodos(_mergeByIdKeepUnsynced(todos, getTodos(), 'dbId'));
      // Re-render if user is currently on the tasks page so DB-loaded tasks become clickable
      if(document.getElementById('page-tasks')&&document.getElementById('page-tasks').classList.contains('active')){
        renderTodos();updateTaskBadge();
      }
    }).catch(function (e) { console.error('Load tasks error:', e); showDbError('tasks'); });
}
function saveTaskAPI(todo) {
  // Join the in-flight counter the loaders gate on (rosters already do this via _trackRosterSave).
  return _trackRosterSave(trackSave('task: '+(todo.text||'').slice(0,32), function(){
    return fetch(API_BASE + '/tasks', {
      method: 'POST', headers: apiHeaders(),
      body: JSON.stringify({
        id: todo.dbId || undefined, text: todo.text, done: todo.done ? 1 : 0,
        due: todo.due || null, client: todo.client || '', priority: todo.priority || 'normal', source: 'homecare',
        parent_id: todo.parentId || null, note: todo.note || null, done_at: todo.doneAt || null,
      }),
    }).then(function (r) {
      if (!r.ok) throw new Error('HTTP ' + r.status);
      return r.json();
    }).then(function (result) {
      if (!todo.dbId && result.id) {
        todo.dbId = result.id;
        // Persist it: every caller runs saveTodos() BEFORE this promise resolves, so without this
        // write-back the id never reaches localStorage and the next load treats the task as unsynced
        // — keeping the local copy AND accepting the server's, i.e. a duplicate task.
        try{ var _t=getTodos(); var _i=_t.findIndex(function(x){return x && x.id===todo.id;});
             if(_i>=0 && !_t[_i].dbId){ _t[_i].dbId=result.id; saveTodos(_t); } }catch(e){}
      }
      return result;
    });
  }));
}
function deleteTaskAPI(dbId) {
  if (!dbId) return;
  // D10: surface a failure/retry so a "deleted" task can't reappear on next sync.
  var doDel=function(){return surfaceSaveFailure(
    fetch(API_BASE + '/tasks/' + dbId, { method: 'DELETE', headers: apiHeaders() })
      .then(function(r){if(!r.ok)throw new Error('HTTP '+r.status);return r;}),
    'Delete task',doDel);};
  doDel();
}

// ============================================================
//  CASEWORKERS
// ============================================================
// Caseworkers: source of truth is Azure SQL via /api/caseworkers.
// localStorage only used as a UI cache (refilled by loadCaseworkersAPI on sign-in).
function getCaseworkers(){try{return JSON.parse(localStorage.getItem('lhca_caseworkers')||'[]');}catch(e){return[];}}
function saveCaseworkersLS(arr){localStorage.setItem('lhca_caseworkers',JSON.stringify(arr));}
function loadCaseworkersAPI(){
  syncStart();
  return fetch(API_BASE + '/caseworkers', { headers: apiHeaders() })
    .then(function(r){ if(!r.ok)throw new Error('HTTP '+r.status); return r.json(); })
    .then(function(rows){
      var arr = (rows||[]).map(function(c){
        // D7: map middle_name/nickname back so they survive a reload (was dropping both).
        return { id:c.id, name:c.name||'', title:c.title||'', first_name:c.first_name||'', middle_name:c.middle_name||'', last_name:c.last_name||'', nickname:c.nickname||'',
                 agency:c.agency||'', org:c.org||'', phone:c.phone||'', email:c.email||'', fax:c.fax||'',
                 street:c.street||'', city:c.city||'', state:c.state||'', zip:c.zip||'', county:c.county||'',
                 notes:c.notes||'', supervisor_id:c.supervisor_id||'',
                 _rowVersion:c.row_version_hex||null };   // optimistic-concurrency token
      });
      if (_savesInFlight > 0) { syncEnd(); return; }   // a save raced this fetch — don't merge stale data
      // D11: don't let a transient empty/partial response wipe the good caseworker cache — and
      // merge so an unsynced local addition isn't dropped by the load (cross-device deletes still
      // propagate: a row that had a _rowVersion but is gone from the server is dropped).
      saveCaseworkersLS(_mergeRosterArr(arr, getCaseworkers()));
      if (typeof renderCaseworkerList === 'function' && document.getElementById('cwList')) renderCaseworkerList();
      _restoreRouteAfterLoad();
      syncEnd();
    })
    .catch(function(e){ console.error('Load caseworkers error:', e); showDbError('caseworkers'); syncEnd(); });
}
function saveCaseworkerAPI(cw, quiet){
  var _doSave = function(){
    var body = Object.assign({}, cw);
    // Optimistic concurrency: send the row_version we last read (as expected_version) so a
    // stale save is rejected (409) instead of silently overwriting another user's edit.
    if (cw._rowVersion) body.expected_version = cw._rowVersion;
    delete body._rowVersion;   // internal token, not a DB column
    delete body._unsaved;      // internal failed-save marker, not a DB column
    return fetch(API_BASE + '/caseworkers', {
      method: 'POST', headers: apiHeaders(), body: JSON.stringify(body),
    }).then(function(r){
      if (r.status === 409) {
        var ce = new Error("This caseworker's info was changed by someone else. Reload to get the latest, then re-apply your edit.");
        ce.isConflict = true; throw ce;
      }
      if(!r.ok)throw new Error('HTTP '+r.status);
      return r.json();
    }).then(function(result){
      if (result && result.row_version) {
        cw._rowVersion = result.row_version;
        var cws = getCaseworkers(); var c = cws.find(function(x){return x.id===cw.id;});
        if (c) { c._rowVersion = result.row_version; saveCaseworkersLS(cws); }
      }
      return result;
    });
  };
  var _pr = quiet ? _doSave() : trackSave(cw.name||cw.id||'caseworker', _doSave);
  _pr.then(function(){ _rosterMarkUnsaved('caseworker', cw.id, false); },
           function(e){ _rosterMarkUnsaved('caseworker', cw.id, !_isConflict(e)); });
  return _trackRosterSave(_pr);
}
function deleteCaseworkerAPI(id){
  // Surface a failure/retry (parity with deleteCaregiverAPI) so a "deleted" caseworker
  // can't silently reappear on the next sync while the UI showed it gone.
  var doDel=function(){return surfaceSaveFailure(
    fetch(API_BASE + '/caseworkers/' + encodeURIComponent(id), { method: 'DELETE', headers: apiHeaders() })
      .then(function(r){if(!r.ok)throw new Error('HTTP '+r.status);return r;}),
    'Delete caseworker',doDel);};
  return doDel();
}
function cwId(){return 'cw_'+Date.now()+'_'+Math.random().toString(36).slice(2,7);}

// ============================================================
//  SUPERVISORS — managed inline via caseworker form dropdown
//  (no dedicated page; CC'd on monthly invoice emails when assigned)
// ============================================================
function getSupervisors(){try{return JSON.parse(localStorage.getItem('lhca_supervisors')||'{}');}catch(e){return{};}}
function saveSupervisorsLS(map){localStorage.setItem('lhca_supervisors',JSON.stringify(map));}
function supId(){return 'sup_'+Date.now()+'_'+Math.random().toString(36).slice(2,7);}
function loadSupervisorsAPI(){
  syncStart();
  return fetch(API_BASE+'/supervisors',{headers:apiHeaders()})
    .then(function(r){if(!r.ok)throw new Error('HTTP '+r.status);return r.json();})
    .then(function(rows){
      var map={};
      (rows||[]).forEach(function(s){
        map[s.id]={id:s.id,name:s.name||'',title:s.title||'',phone:s.phone||'',email:s.email||'',
          _rowVersion:s.row_version_hex||null};   // optimistic-concurrency token
      });
      if (_savesInFlight > 0) { syncEnd(); return; }   // a save raced this fetch — don't merge stale data
      // Merge so an unsynced local addition survives the load and a transient empty response
      // doesn't wipe the cache (a synced supervisor now carries _rowVersion, so cross-device
      // deletes propagate just like caregivers/caseworkers).
      saveSupervisorsLS(_mergeRosterMap(map, getSupervisors()));
      // If a caseworker form is open, refresh its dropdown
      refreshSupervisorDropdowns();
      syncEnd();
    })
    .catch(function(e){console.error('Load supervisors error:',e);showDbError('supervisors');syncEnd();});
}
function saveSupervisorAPI(sup){
  return _trackRosterSave(_rosterSaveMarked('supervisor', sup.id, trackSave('supervisor: '+(sup.name||sup.id||''), function(){
    // Optimistic concurrency (mirror of caseworkers): send the row_version we last read so a stale
    // save is rejected (409) instead of silently overwriting another device's edit.
    var body=Object.assign({},sup);
    if(sup._rowVersion) body.expected_version=sup._rowVersion;
    delete body._rowVersion;   // internal token, not a DB column
    delete body._unsaved;      // internal failed-save marker, not a DB column
    return fetch(API_BASE+'/supervisors',{
      method:'POST',headers:apiHeaders(),body:JSON.stringify(body)
    }).then(function(r){
      if(r.status===409){ var ce=new Error("This supervisor was changed by someone else. Reload to get the latest, then re-apply your edit."); ce.isConflict=true; throw ce; }
      if(!r.ok)throw new Error('HTTP '+r.status);
      return r.json();
    }).then(function(result){
      if(result&&result.row_version){
        sup._rowVersion=result.row_version;
        var sups=getSupervisors(); if(sups[sup.id]){ sups[sup.id]._rowVersion=result.row_version; saveSupervisorsLS(sups); }
      }
      return result;
    });
  })));
}
// Attach the failed/succeeded marker to a save promise without changing what the caller receives.
function _rosterSaveMarked(kind, id, pr){
  pr.then(function(){ _rosterMarkUnsaved(kind, id, false); },
          function(e){ _rosterMarkUnsaved(kind, id, !_isConflict(e)); });
  return pr;
}
function deleteSupervisorAPI(id){
  // Surface a failure/retry (parity with deleteCaregiverAPI) so a "deleted" supervisor
  // can't silently reappear on the next sync while the UI showed it gone.
  var doDel=function(){return surfaceSaveFailure(
    fetch(API_BASE+'/supervisors/'+encodeURIComponent(id),{ method:'DELETE', headers:apiHeaders() })
      .then(function(r){if(!r.ok)throw new Error('HTTP '+r.status);return r;}),
    'Delete supervisor',doDel);};
  return doDel();
}
function populateSupervisorDropdown(selectEl,currentId){
  if(!selectEl)return;
  var sups=getSupervisors();
  var ids=Object.keys(sups).sort(function(a,b){return (sups[a].name||'').localeCompare(sups[b].name||'');});
  selectEl.innerHTML='<option value="">— None —</option>'+ids.map(function(id){
    var s=sups[id];
    return '<option value="'+esc(id)+'"'+(id===currentId?' selected':'')+'>'+esc(s.name)+(s.email?' ('+esc(s.email)+')':'')+'</option>';
  }).join('');
  // Show/hide the Edit button next to the dropdown based on selection
  var editBtn=document.getElementById(selectEl.id+'-edit-btn');
  if(editBtn)editBtn.style.display=selectEl.value?'inline-block':'none';
  selectEl.onchange=function(){
    var eb=document.getElementById(selectEl.id+'-edit-btn');
    if(eb)eb.style.display=selectEl.value?'inline-block':'none';
  };
}
function refreshSupervisorDropdowns(){
  var sel1=document.getElementById('cw-supervisor');
  if(sel1)populateSupervisorDropdown(sel1,sel1.value);
  var sel2=document.getElementById('cwi-supervisor');
  if(sel2)populateSupervisorDropdown(sel2,sel2.value);
}
// ── Supervisor profile (detail page) ─────────────────────────
// Caseworker + supervisor emails are all @michigan.gov. Type just the username and this
// fills in the domain on blur; a value that already has "@" is left untouched.
function _autoGovEmail(input){
  if(!input)return;
  var v=(input.value||'').trim();
  if(v && v.indexOf('@')===-1){ input.value=v+'@michigan.gov'; }
}
function openSupDetail(id){
  var sups=getSupervisors(); var s=sups[id]; if(!s)return;
  // supDetailView is a sibling of the two list views inside cwGridView — hide the lists, not
  // the wrapper (hiding cwGridView would hide the profile itself).
  var cwList=document.getElementById('cwViewCaseworkers'); if(cwList)cwList.style.display='none';
  var supView=document.getElementById('cwViewSupervisors'); if(supView)supView.style.display='none';
  var det=document.getElementById('supDetailView'); if(det)det.style.display='block';
  document.getElementById('supd-id').value=id;
  document.getElementById('supd-title').value=s.title||'';
  document.getElementById('supd-name').value=s.name||'';
  document.getElementById('supd-phone').value=s.phone||'';
  document.getElementById('supd-email').value=s.email||'';
  var ini=(s.name||'?').split(/\s+/).map(function(w){return w[0]||'';}).slice(0,2).join('').toUpperCase();
  var av=document.getElementById('supDetailAvatar'); if(av)av.textContent=ini||'?';
  var nm=document.getElementById('supDetailName'); if(nm)nm.textContent=(s.title?s.title+' ':'')+(s.name||'');
  renderSupCaseworkers(id);
}
function renderSupCaseworkers(id){
  var host=document.getElementById('supd-caseworkers'); if(!host)return;
  var cws=getCaseworkers().filter(function(c){return String(c.supervisor_id||'')===String(id);})
    .sort(function(a,b){return (a.name||'').localeCompare(b.name||'');});
  if(!cws.length){ host.innerHTML='<div style="color:#5c7590;font-size:13px;padding:4px 0;">No caseworkers assigned to this supervisor yet.</div>'; return; }
  host.innerHTML=cws.map(function(c){
    var nm=(c.title?c.title+' ':'')+(c.name||'');
    var sub=[c.agency||'',c.email||''].filter(Boolean).map(esc).join(' · ');
    return '<div style="display:flex;align-items:center;justify-content:space-between;gap:10px;padding:8px 10px;border:1px solid #e8ecf0;border-radius:6px;margin-bottom:6px;">'+
      '<div style="min-width:0;"><div style="font-size:13px;font-weight:600;color:#1a2b45;">'+esc(nm)+'</div>'+
      '<div style="font-size:11px;color:#5c7590;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">'+sub+'</div></div>'+
      '<button class="btn btn-secondary btn-sm" onclick="openCwDetail(\''+escJsAttr(c.id)+'\')" style="white-space:nowrap;">Open ↗</button>'+
    '</div>';
  }).join('');
}
function saveSupDetail(){
  var id=document.getElementById('supd-id').value;
  var name=(document.getElementById('supd-name').value||'').trim();
  if(!name){ showAlert('Name is required.'); return; }
  var sup={ id:id, name:name, title:document.getElementById('supd-title').value||'',
    phone:(document.getElementById('supd-phone').value||'').trim(),
    email:(document.getElementById('supd-email').value||'').trim() };
  var sups=getSupervisors(); sups[id]=sup; saveSupervisorsLS(sups);
  saveSupervisorAPI(sup);
  var nm=document.getElementById('supDetailName'); if(nm)nm.textContent=(sup.title?sup.title+' ':'')+sup.name;
  if(typeof refreshSupervisorDropdowns==='function') refreshSupervisorDropdowns();
}
function backToSupList(){
  var det=document.getElementById('supDetailView'); if(det)det.style.display='none';
  var supView=document.getElementById('cwViewSupervisors'); if(supView)supView.style.display='block';
  if(typeof renderSupervisorList==='function') renderSupervisorList();
}
function deleteSupervisorFromDetail(){
  var id=document.getElementById('supd-id').value;
  var sups=getSupervisors(); var name=(sups[id]&&sups[id].name)||'this supervisor';
  showConfirm('Delete supervisor "'+name+'"? Caseworkers assigned to them will show no supervisor.',function(){
    var s=getSupervisors(); delete s[id]; saveSupervisorsLS(s);
    if(typeof deleteSupervisorAPI==='function') deleteSupervisorAPI(id);
    backToSupList();
  },{title:'Delete Supervisor',okText:'Delete'});
}
function openSupervisorModal(){_openSupervisorModal(null);}
function editSelectedSupervisor(){
  var sel=document.getElementById('cw-supervisor')||document.getElementById('cwi-supervisor');
  if(sel && sel.value)_openSupervisorModal(sel.value);
}
function _openSupervisorModal(editId){
  var sups=getSupervisors();
  var sup=editId?(sups[editId]||{}):{};
  var existing=document.getElementById('supModal');
  if(existing)existing.remove();
  var overlay=document.createElement('div');
  overlay.id='supModal';
  overlay.className='modal-overlay open';
  overlay.innerHTML=
    '<div class="modal-box" style="max-width:420px;">'+
      '<h3>'+(editId?'Edit Supervisor':'New Supervisor')+'</h3>'+
      '<div style="display:flex;gap:8px;margin-top:8px;">'+
        '<div class="ff" style="flex:0 0 110px;"><label for="sup-title">Title</label>'+
          '<select id="sup-title">'+['','Mr.','Mrs.','Ms.','Miss','Dr.','Mx.'].map(function(t){return '<option'+(String(sup.title||'')===t?' selected':'')+'>'+t+'</option>';}).join('')+'</select></div>'+
        '<div class="ff" style="flex:1;"><label for="sup-name">Name *</label><input id="sup-name" value="'+esc(sup.name||'')+'" placeholder="Full name" maxlength="120"></div>'+
      '</div>'+
      '<div class="ff" style="margin-top:8px;"><label for="sup-phone">Phone</label><input id="sup-phone" value="'+esc(sup.phone||'')+'" placeholder="(555) 555-5555" maxlength="30"></div>'+
      '<div class="ff" style="margin-top:8px;"><label for="sup-email">Email <span style="font-weight:400;font-size:10px;color:#5c7590;">(used for CC on monthly invoice emails)</span></label><input id="sup-email" type="email" value="'+esc(sup.email||'')+'" placeholder="supervisor@michigan.gov" maxlength="120" onblur="_autoGovEmail(this)"></div>'+
      '<div class="modal-row" style="justify-content:space-between;">'+
        (editId?'<button class="btn btn-danger btn-sm" onclick="_deleteSupervisorFromModal(\''+escJsAttr(editId)+'\')">Delete</button>':'<span></span>')+
        '<div style="display:flex;gap:8px;">'+
          '<button class="btn btn-secondary" onclick="document.getElementById(\'supModal\').remove()">Cancel</button>'+
          '<button class="btn btn-primary" onclick="_saveSupervisorFromModal(\''+(editId||'')+'\')">Save</button>'+
        '</div>'+
      '</div>'+
    '</div>';
  document.body.appendChild(overlay);
  setTimeout(function(){var n=document.getElementById('sup-name');if(n)n.focus();},80);
}
function _saveSupervisorFromModal(editId){
  var name=(document.getElementById('sup-name').value||'').trim();
  if(!name){showAlert('Supervisor name is required.');return;}
  var phone=(document.getElementById('sup-phone').value||'').trim();
  var email=(document.getElementById('sup-email').value||'').trim();
  var titleEl=document.getElementById('sup-title');
  var title=titleEl?(titleEl.value||''):'';
  var sups=getSupervisors();
  var id=editId||supId();
  sups[id]={id:id,name:name,title:title,phone:phone,email:email};
  saveSupervisorsLS(sups);
  saveSupervisorAPI(sups[id]);
  document.getElementById('supModal').remove();
  // Refresh both dropdowns + select the newly added/edited one if we were in a form context
  ['cw-supervisor','cwi-supervisor'].forEach(function(selId){
    var sel=document.getElementById(selId);
    if(sel){
      populateSupervisorDropdown(sel,id);
      sel.value=id;
      var eb=document.getElementById(selId+'-edit-btn');
      if(eb)eb.style.display='inline-block';
    }
  });
  // If on the Supervisors view, refresh that table too
  if(typeof renderSupervisorList==='function' && document.getElementById('cwViewSupervisors') && document.getElementById('cwViewSupervisors').style.display!=='none')renderSupervisorList();
  // No manual showToast here — saveSupervisorAPI (wrapped in trackSave since Phase 4) already
  // shows a "✓ Saved supervisor: …" toast on success. A second one caused the overlap.
}
function _deleteSupervisorFromModal(id){
  var sups=getSupervisors();
  var name=sups[id]?sups[id].name:'this supervisor';
  // Close the edit modal FIRST so the confirm isn't hidden behind it — both share
  // z-index:9000 so the later-appended supModal was stacking on top of showConfirm.
  var m=document.getElementById('supModal');if(m)m.remove();
  // Warn about caseworker assignments that would be orphaned
  var assigned=getCaseworkers().filter(function(c){return c.supervisor_id===id;});
  var warn=assigned.length?'\n\n'+assigned.length+' caseworker'+(assigned.length>1?'s':'')+' assigned to this supervisor will be unassigned.':'';
  showConfirm(
    'Delete '+name+'?'+warn+'\n\nThis cannot be undone.',
    function(){
      delete sups[id];
      saveSupervisorsLS(sups);
      deleteSupervisorAPI(id);
      // Unassign locally on any caseworkers that had this supervisor (server is already stateless about it)
      var arr=getCaseworkers();
      arr.forEach(function(c){if(c.supervisor_id===id){c.supervisor_id='';saveCaseworkerAPI(c);}});
      saveCaseworkersLS(arr);
      refreshSupervisorDropdowns();
      if(typeof renderCaseworkerList==='function')renderCaseworkerList();
      if(typeof renderSupervisorList==='function' && document.getElementById('cwViewSupervisors') && document.getElementById('cwViewSupervisors').style.display!=='none')renderSupervisorList();
      showToast('Supervisor deleted',3000);
    },
    {title:'Delete Supervisor',okText:'Delete',danger:true}
  );
}
function navCaseworkers(){
  showPage('caseworkers');bc([{l:'Caseworkers'}]);document.getElementById('topbarActions').innerHTML='';
  showCwGrid();
  // Default back to Caseworkers view when arriving at the page
  showCwView('caseworkers');
  // Cold cache: fetch directly, bypassing the 30s revalidate throttle.
  if(typeof spToken!=='undefined'&&spToken&&getCaseworkers().length===0&&typeof loadCaseworkersAPI==='function')loadCaseworkersAPI();
  if(typeof revalidate==='function')revalidate();
}

// Pill toggle on the Caseworkers page — flips between Caseworkers list and Supervisors list
function showCwView(view){
  var isSup = view === 'supervisors';
  var cwView=document.getElementById('cwViewCaseworkers');
  var supView=document.getElementById('cwViewSupervisors');
  // Always leave the supervisor profile page when switching pills.
  var supDet=document.getElementById('supDetailView'); if(supDet)supDet.style.display='none';
  var pillCw=document.getElementById('pillCaseworkers');
  var pillSup=document.getElementById('pillSupervisors');
  var addBtn=document.getElementById('cwAddBtn');
  var title=document.getElementById('cwPageTitle');
  var sub=document.getElementById('cwPageSub');
  if(cwView)cwView.style.display=isSup?'none':'';
  if(supView)supView.style.display=isSup?'':'none';
  if(pillCw)pillCw.classList.toggle('active',!isSup);
  if(pillSup)pillSup.classList.toggle('active',isSup);
  if(addBtn){
    addBtn.textContent=isSup?'+ Add Supervisor':'+ Add Caseworker';
    addBtn.onclick=isSup?function(){openSupervisorModal();}:function(){showCaseworkerForm();};
  }
  if(title)title.textContent=isSup?'Supervisors':'Caseworkers';
  if(sub)sub.textContent=isSup?'Manage supervisors — assigned to caseworkers as the CC on monthly invoice emails.':'Manage caseworkers and their client assignments.';
  if(isSup)renderSupervisorList();
}

// ── Supervisors list (within Caseworkers page) ──────────────
var supBulkSelected={};
// Sort / pagination / column-width state for Supervisors table — mirrors Clients pattern
var _supSort={key:'name',dir:'asc'};
// Widths sum to ~90; Supervisor column trimmed from 26 → 22.
var _supColDefaults={name:22,phone:20,email:32,caseworkers:16};
var _supColumnWidths=null;
var _supPage=1;
var _supPageSize=25;
(function loadSupPrefs(){
  try{
    var w=localStorage.getItem('lhca_sup_col_widths');if(w)_supColumnWidths=JSON.parse(w);
    var ps=localStorage.getItem('lhca_sup_page_size');
    if(ps==='all')_supPageSize=Infinity;else if(ps){var n=parseInt(ps);if(n>0)_supPageSize=n;}
  }catch(e){}
})();
function sortSupBy(key){if(_supSort.key===key)_supSort.dir=_supSort.dir==='asc'?'desc':'asc';else{_supSort.key=key;_supSort.dir='asc';}renderSupervisorList();}
function _supSortCompare(a,b,sups,countBySup){
  var sa=sups[a],sb=sups[b];
  var key=_supSort.key,dir=_supSort.dir==='asc'?1:-1;
  function val(k,s,id){
    if(k==='name')return (s.name||'').toLowerCase();
    if(k==='phone')return (s.phone||'').toLowerCase();
    if(k==='email')return (s.email||'').toLowerCase();
    if(k==='caseworkers')return countBySup[id]||0;
    return 0;
  }
  var va=val(key,sa,a),vb=val(key,sb,b);
  if(typeof va==='string')return va.localeCompare(vb)*dir;
  return (va<vb?-1:va>vb?1:0)*dir;
}
function applySupColWidths(){var headers=document.querySelectorAll('#supTable thead th[data-col]');headers.forEach(function(th){var col=th.dataset.col;var w=(_supColumnWidths&&_supColumnWidths[col])||_supColDefaults[col];if(w)th.style.width=w+'%';});}
function startSupColResize(e,col){
  e.preventDefault();e.stopPropagation();
  var th=e.target.closest('th');if(!th)return;
  var startX=e.pageX,startWidth=th.offsetWidth,tableWidth=th.closest('table').offsetWidth;
  document.body.style.cursor='col-resize';document.body.style.userSelect='none';
  function onMove(ev){var d=ev.pageX-startX;var nx=Math.max(60,startWidth+d);var pct=(nx/tableWidth)*100;th.style.width=pct+'%';if(!_supColumnWidths)_supColumnWidths={};_supColumnWidths[col]=pct;}
  function onUp(){document.removeEventListener('mousemove',onMove);document.removeEventListener('mouseup',onUp);document.body.style.cursor='';document.body.style.userSelect='';try{localStorage.setItem('lhca_sup_col_widths',JSON.stringify(_supColumnWidths||{}));}catch(e){}}
  document.addEventListener('mousemove',onMove);document.addEventListener('mouseup',onUp);
}
function supPageSizeChange(){var s=document.getElementById('supPageSize');var v=s.value;if(v==='all')_supPageSize=Infinity;else _supPageSize=parseInt(v)||25;_supPage=1;try{localStorage.setItem('lhca_sup_page_size',v);}catch(e){}renderSupervisorList();}
function goToSupPage(n){_supPage=n;renderSupervisorList();}

function renderSupervisorList(){
  var sups=getSupervisors();
  var q=(document.getElementById('supSearch')?document.getElementById('supSearch').value:'').toLowerCase();
  var cws=getCaseworkers();
  var countBySup={};
  cws.forEach(function(c){if(c.supervisor_id){countBySup[c.supervisor_id]=(countBySup[c.supervisor_id]||0)+1;}});

  var allIds=Object.keys(sups);
  var ids=allIds.filter(function(id){
    if(!q)return true;
    var s=sups[id];
    return (s.name||'').toLowerCase().includes(q)||(s.email||'').toLowerCase().includes(q)||(s.phone||'').includes(q);
  });
  ids.sort(function(a,b){return _supSortCompare(a,b,sups,countBySup);});

  var headers=document.querySelectorAll('#supTable thead th.sortable');
  headers.forEach(function(h){h.classList.remove('sort-asc','sort-desc');if(h.dataset.sortkey===_supSort.key)h.classList.add(_supSort.dir==='asc'?'sort-asc':'sort-desc');});
  applySupColWidths();

  var totalMatched=ids.length;
  var totalPages=Math.max(1,Math.ceil(totalMatched/_supPageSize));
  if(_supPage>totalPages)_supPage=totalPages;
  var startIdx=(_supPage-1)*_supPageSize;
  var endIdx=Math.min(startIdx+_supPageSize,totalMatched);
  var visibleIds=ids.slice(startIdx,endIdx);
  var countEl=document.getElementById('supTableCount');
  if(countEl){
    if(!totalMatched)countEl.textContent='';
    else if(_supPageSize===Infinity||totalMatched<=_supPageSize)countEl.textContent=totalMatched+' supervisor'+(totalMatched===1?'':'s');
    else countEl.textContent='Showing '+(startIdx+1)+'–'+endIdx+' of '+totalMatched;
  }
  var pageControls=document.getElementById('supPageControls');
  if(pageControls){
    if(totalPages<=1)pageControls.innerHTML='';
    else {
      var html='<button class="pg-btn" onclick="goToSupPage('+(_supPage-1)+')"'+(_supPage===1?' disabled':'')+'>‹</button>';
      var startPg=Math.max(1,_supPage-2),endPg=Math.min(totalPages,startPg+4);
      if(endPg-startPg<4)startPg=Math.max(1,endPg-4);
      for(var p=startPg;p<=endPg;p++)html+='<button class="pg-btn'+(p===_supPage?' active':'')+'" onclick="goToSupPage('+p+')">'+p+'</button>';
      html+='<button class="pg-btn" onclick="goToSupPage('+(_supPage+1)+')"'+(_supPage===totalPages?' disabled':'')+'>›</button>';
      pageControls.innerHTML=html;
    }
  }

  var tbody=document.getElementById('supTableBody');if(!tbody)return;tbody.innerHTML='';
  var empty=document.getElementById('supTableEmpty');
  if(!ids.length){if(empty)empty.style.display='block';return;}
  if(empty)empty.style.display='none';
  visibleIds.forEach(function(id){
    var s=sups[id];
    var cwCount=countBySup[id]||0;
    var tr=document.createElement('tr');
    var checked=supBulkSelected[id]?'checked':'';
    tr.innerHTML=
      '<td style="width:26px;" onclick="event.stopPropagation()"><input type="checkbox" class="sup-select" data-id="'+esc(id)+'" '+checked+' onchange="toggleBulkSupervisor(\''+escJsAttr(id)+'\',this)" style="width:12px;height:12px;cursor:pointer;"></td>'+
      '<td><div class="ct-name">'+esc((s.title?s.title+' ':'')+(s.name||''))+'</div></td>'+
      '<td style="color:var(--text-muted);font-size:12px;">'+esc(s.phone||'—')+'</td>'+
      '<td style="color:var(--text-muted);font-size:12px;">'+esc(s.email||'—')+'</td>'+
      '<td style="color:var(--text);font-size:12px;">'+cwCount+'</td>';
    tr.addEventListener('click',function(e){
      if(e.target.closest('a')||e.target.tagName==='BUTTON'||e.target.tagName==='INPUT')return;
      openSupDetail(id);
    });
    tbody.appendChild(tr);
  });

  var sel=document.getElementById('supPageSize');
  if(sel){var savedPs=_supPageSize===Infinity?'all':String(_supPageSize);if(sel.value!==savedPs)sel.value=savedPs;}
}
function toggleBulkSupervisor(id,cb){
  if(cb.checked)supBulkSelected[id]=true;else delete supBulkSelected[id];
  var count=Object.keys(supBulkSelected).length;
  var bar=document.getElementById('supBulkBar');
  if(bar){bar.classList.toggle('visible',count>0);var lbl=document.getElementById('supBulkCount');if(lbl)lbl.textContent=count+' selected';}
}
function clearSupBulkSelect(){supBulkSelected={};var bar=document.getElementById('supBulkBar');if(bar)bar.classList.remove('visible');var sa=document.getElementById('supSelectAll');if(sa)sa.checked=false;renderSupervisorList();}
function toggleAllSupervisors(cb){
  document.querySelectorAll('.sup-select').forEach(function(c){
    c.checked=cb.checked;
    var id=c.dataset.id;
    if(cb.checked)supBulkSelected[id]=true;else delete supBulkSelected[id];
  });
  var count=Object.keys(supBulkSelected).length;
  var bar=document.getElementById('supBulkBar');
  if(bar){bar.classList.toggle('visible',count>0);var lbl=document.getElementById('supBulkCount');if(lbl)lbl.textContent=count+' selected';}
}
function bulkDeleteSupervisors(){
  var ids=Object.keys(supBulkSelected);if(!ids.length){showAlert('No supervisors selected.');return;}
  var sups=getSupervisors();
  var preview=ids.slice(0,5).map(function(id){return '• '+(sups[id]?sups[id].name||id:id);}).join('\n')+(ids.length>5?'\n• …and '+(ids.length-5)+' more':'');
  // Count caseworker re-assignments that would orphan
  var cws=getCaseworkers();
  var affected=cws.filter(function(c){return ids.indexOf(c.supervisor_id)>=0;}).length;
  var warn=affected?'\n\n'+affected+' caseworker'+(affected>1?'s':'')+' assigned to these supervisors will be unassigned.':'';
  showConfirm(
    'Delete '+ids.length+' supervisor'+(ids.length>1?'s':'')+'?\n\n'+preview+warn+'\n\nThis cannot be undone.',
    function(){
      ids.forEach(function(id){
        delete sups[id];
        try{deleteSupervisorAPI(id);}catch(e){}
      });
      saveSupervisorsLS(sups);
      cws.forEach(function(c){if(ids.indexOf(c.supervisor_id)>=0){c.supervisor_id='';saveCaseworkerAPI(c);}});
      saveCaseworkersLS(cws);
      logActivity('delete','Bulk supervisor delete: '+ids.length+' removed');
      clearSupBulkSelect();
      showAlert(ids.length+' supervisor'+(ids.length!==1?'s':'')+' deleted.');
    },
    {title:'Delete '+ids.length+' Supervisors?',okText:'Delete '+ids.length+(ids.length>1?' Supervisors':' Supervisor'),danger:true}
  );
}
// Agency is an MDHHS-only detail (which office). Show it only when Org = MDHHS; a carrier coordinator
// has no MDHHS agency. Used on the caseworker grid form (cw-*) and detail-inline pane (cwi-*).
function cwOrgToggle(){
  var org=(document.getElementById('cw-org')||{}).value||'';
  var w=document.getElementById('cw-agency-wrap'); if(w)w.style.display=(org==='MDHHS')?'':'none';
}
function cwiOrgToggle(){
  var org=(document.getElementById('cwi-org')||{}).value||'';
  var ag=document.getElementById('cwi-agency'); if(!ag)return;
  var wrap=ag.closest('.info-field'); if(wrap)wrap.style.display=(org==='MDHHS')?'':'none';
}
function showCaseworkerForm(id){
  // Hide detail view always; for new caseworker keep grid visible (scroll to form), for edit from detail hide grid
  var dv=document.getElementById('cwDetailView');if(dv)dv.style.display='none';
  if(id&&activeCwId){
    // Editing from detail view — hide grid too
    var gv=document.getElementById('cwGridView');if(gv)gv.style.display='none';
  } else if(!id){
    // New caseworker — keep grid visible so they can see the list
    var gv2=document.getElementById('cwGridView');if(gv2)gv2.style.display='';
  }
  document.getElementById('cwFormWrap').style.display='block';
  document.getElementById('cwFormTitle').textContent=id?'Edit Caseworker':'New Caseworker';
  document.getElementById('cw-editing-id').value=id||'';
  if(id){
    var cw=getCaseworkers().find(function(c){return c.id===id;});
    if(cw){
      var nameParts=(cw.name||'').trim().split(' ');
      var firstName=nameParts[0]||'';
      var lastName=nameParts.slice(1).join(' ')||'';
      document.getElementById('cw-title').value=cw.title||'';
      document.getElementById('cw-first-name').value=cw.first_name||firstName;
      document.getElementById('cw-middle-name').value=cw.middle_name||'';
      document.getElementById('cw-last-name').value=cw.last_name||lastName;
      document.getElementById('cw-nickname').value=cw.nickname||'';
      document.getElementById('cw-agency').value=cw.agency||'';
      var _cwOrgEl=document.getElementById('cw-org');if(_cwOrgEl)_cwOrgEl.value=cw.org||'';
      document.getElementById('cw-phone').value=cw.phone||'';
      document.getElementById('cw-fax').value=cw.fax||'';
      document.getElementById('cw-email').value=cw.email||'';
      // Supervisor dropdown is populated separately by populateSupervisorDropdown() after the form opens
      document.getElementById('cw-street').value=cw.street||'';
      document.getElementById('cw-city').value=cw.city||'';
      document.getElementById('cw-state').value=cw.state||'';
      document.getElementById('cw-zip').value=cw.zip||'';
      document.getElementById('cw-county').value=cw.county||'';
      document.getElementById('cw-notes').value=cw.notes||'';
      // Populate supervisor dropdown with the currently assigned one selected
      populateSupervisorDropdown(document.getElementById('cw-supervisor'),cw.supervisor_id||'');
    }
    document.getElementById('cwDeleteBtn').style.display='inline-block';
  } else {
    ['cw-title','cw-first-name','cw-middle-name','cw-last-name','cw-nickname','cw-agency','cw-org','cw-phone','cw-fax','cw-email','cw-street','cw-city','cw-state','cw-zip','cw-county','cw-notes'].forEach(function(fid){var e=document.getElementById(fid);if(e)e.value='';});
    // Fresh form — populate dropdown with no selection
    populateSupervisorDropdown(document.getElementById('cw-supervisor'),'');
    document.getElementById('cwDeleteBtn').style.display='none';
  }
  cwOrgToggle();   // show the Agency field only when Org = MDHHS
  wireCopyableFields('cwFormWrap');
  document.getElementById('cwFormWrap').scrollIntoView({behavior:'smooth'});
}
function hideCaseworkerForm(){
  document.getElementById('cwFormWrap').style.display='none';
  // If no detail view active and grid is hidden, show grid
  var dv=document.getElementById('cwDetailView');
  var gv=document.getElementById('cwGridView');
  if(gv&&gv.style.display==='none'&&(!dv||dv.style.display==='none')){
    gv.style.display='';
  }
}
function saveCaseworker(){
  var firstName=document.getElementById('cw-first-name').value.trim();
  var middleName=document.getElementById('cw-middle-name').value.trim();
  var lastName=document.getElementById('cw-last-name').value.trim();
  var nickname=document.getElementById('cw-nickname').value.trim();
  var name=(firstName+(middleName?' '+middleName:'')+' '+lastName).trim();
  if(!name){showAlert('Name is required.');return;}
  var cws=getCaseworkers();
  var editingId=document.getElementById('cw-editing-id').value;
  var rec={
    id:editingId||cwId(),name:name,title:document.getElementById('cw-title').value,first_name:firstName,middle_name:middleName,last_name:lastName,nickname:nickname,
    // Agency is only shown when Org = MDHHS (see cwOrgToggle) — clear it otherwise so a coordinator
    // moved to a carrier org doesn't keep a stale MDHHS agency on the record.
    agency:(((document.getElementById('cw-org')||{}).value||'')==='MDHHS')?document.getElementById('cw-agency').value:'',
    org:(document.getElementById('cw-org')||{}).value||'',
    phone:document.getElementById('cw-phone').value,
    fax:document.getElementById('cw-fax').value,
    email:document.getElementById('cw-email').value,
    street:document.getElementById('cw-street').value,
    city:document.getElementById('cw-city').value,
    state:document.getElementById('cw-state').value,
    zip:document.getElementById('cw-zip').value,
    county:document.getElementById('cw-county').value,
    notes:document.getElementById('cw-notes').value,
    supervisor_id:(document.getElementById('cw-supervisor')||{}).value||''
  };
  if(editingId){
    var idx=cws.findIndex(function(c){return c.id===editingId;});
    if(idx>=0){cws[idx]=rec;}
  } else {
    cws.push(rec);
  }
  saveCaseworkersLS(cws);
  saveCaseworkerAPI(rec);
  hideCaseworkerForm();
  if(editingId&&activeCwId===editingId){
    // Return to detail view if we were editing from the detail view
    document.getElementById('cwGridView').style.display='none';
    openCwDetail(editingId);
  } else {
    showCwGrid();
  }
}
function deleteCaseworker(){
  var id=document.getElementById('cw-editing-id').value;if(!id)return;
  var cws=getCaseworkers().find(function(c){return c.id===id;});if(!cws)return;
  showConfirm('Delete caseworker "'+cws.name+'"? This cannot be undone.',function(){
    addAuditEntry('','Caseworker DELETED: '+cws.name+' by '+currentUserEmail());
    saveCaseworkersLS(getCaseworkers().filter(function(c){return c.id!==id;}));
    deleteCaseworkerAPI(id);
    activeCwId=null;
    hideCaseworkerForm();
    showCwGrid();
  },{title:'Delete Caseworker',okText:'Delete'});
}
function deleteCaseworkerFromDetail(){
  if(!activeCwId)return;
  var cw=getCaseworkers().find(function(c){return c.id===activeCwId;});if(!cw)return;
  var idToDelete=activeCwId;
  showConfirm('Delete caseworker "'+cw.name+'"? This cannot be undone.',function(){
    saveCaseworkersLS(getCaseworkers().filter(function(c){return c.id!==idToDelete;}));
    deleteCaseworkerAPI(idToDelete);
    aiTrack('CaseworkerDeleted',{caseworkerId:idToDelete,caseworkerName:cw.name||idToDelete});
    showCwGrid();
  },{title:'Delete Caseworker',okText:'Delete'});
}
var cwBulkSelected={};
// Sort / pagination / column-width state for Caseworkers table — mirrors Clients pattern
var _cwSort={key:'name',dir:'asc'};
// Widths sum to ~92; Caseworker column trimmed from 20 → 17.
var _cwColDefaults={name:17,phone:14,email:22,county:12,supervisor:16,clients:12};
var _cwColumnWidths=null;
var _cwPage=1;
var _cwPageSize=25;
(function loadCwPrefs(){
  try{
    var w=localStorage.getItem('lhca_cw_col_widths');if(w)_cwColumnWidths=JSON.parse(w);
    var ps=localStorage.getItem('lhca_cw_page_size');
    if(ps==='all')_cwPageSize=Infinity;else if(ps){var n=parseInt(ps);if(n>0)_cwPageSize=n;}
  }catch(e){}
})();
function sortCwBy(key){if(_cwSort.key===key)_cwSort.dir=_cwSort.dir==='asc'?'desc':'asc';else{_cwSort.key=key;_cwSort.dir='asc';}renderCaseworkerList();}
function _cwSortCompare(a,b,profiles,sups){
  var key=_cwSort.key,dir=_cwSort.dir==='asc'?1:-1;
  function clientCount(cw){return Object.keys(profiles).filter(function(k){return (profiles[k].caseworkerId===cw.id||profiles[k].worker===cw.name) && (profiles[k].clientStatus||'active')==='active';}).length;}
  function val(k,c){
    if(k==='name')return (c.name||'').toLowerCase();
    if(k==='phone')return (c.phone||'').toLowerCase();
    if(k==='email')return (c.email||'').toLowerCase();
    if(k==='county')return (c.county||'').toLowerCase();
    if(k==='supervisor'){var s=c.supervisor_id&&sups[c.supervisor_id]?sups[c.supervisor_id].name:'';return s.toLowerCase();}
    if(k==='clients')return clientCount(c);
    return 0;
  }
  var va=val(key,a),vb=val(key,b);
  if(typeof va==='string')return va.localeCompare(vb)*dir;
  return (va<vb?-1:va>vb?1:0)*dir;
}
function applyCwColWidths(){var headers=document.querySelectorAll('#cwTable thead th[data-col]');headers.forEach(function(th){var col=th.dataset.col;var w=(_cwColumnWidths&&_cwColumnWidths[col])||_cwColDefaults[col];if(w)th.style.width=w+'%';});}
function startCwColResize(e,col){
  e.preventDefault();e.stopPropagation();
  var th=e.target.closest('th');if(!th)return;
  var startX=e.pageX,startWidth=th.offsetWidth,tableWidth=th.closest('table').offsetWidth;
  document.body.style.cursor='col-resize';document.body.style.userSelect='none';
  function onMove(ev){var d=ev.pageX-startX;var nx=Math.max(60,startWidth+d);var pct=(nx/tableWidth)*100;th.style.width=pct+'%';if(!_cwColumnWidths)_cwColumnWidths={};_cwColumnWidths[col]=pct;}
  function onUp(){document.removeEventListener('mousemove',onMove);document.removeEventListener('mouseup',onUp);document.body.style.cursor='';document.body.style.userSelect='';try{localStorage.setItem('lhca_cw_col_widths',JSON.stringify(_cwColumnWidths||{}));}catch(e){}}
  document.addEventListener('mousemove',onMove);document.addEventListener('mouseup',onUp);
}
function cwPageSizeChange(){var s=document.getElementById('cwPageSize');var v=s.value;if(v==='all')_cwPageSize=Infinity;else _cwPageSize=parseInt(v)||25;_cwPage=1;try{localStorage.setItem('lhca_cw_page_size',v);}catch(e){}renderCaseworkerList();}
function goToCwPage(n){_cwPage=n;renderCaseworkerList();}

function renderCaseworkerList(){
  var cws=getCaseworkers();
  var q=(document.getElementById('cwSearch')?document.getElementById('cwSearch').value:'').toLowerCase();
  var profiles=getProfiles();
  var sups=getSupervisors();

  var filtered=cws.filter(function(cw){return !q||(cw.name||'').toLowerCase().includes(q)||(cw.agency||'').toLowerCase().includes(q)||(cw.email||'').toLowerCase().includes(q);});
  filtered.sort(function(a,b){return _cwSortCompare(a,b,profiles,sups);});

  var headers=document.querySelectorAll('#cwTable thead th.sortable');
  headers.forEach(function(h){h.classList.remove('sort-asc','sort-desc');if(h.dataset.sortkey===_cwSort.key)h.classList.add(_cwSort.dir==='asc'?'sort-asc':'sort-desc');});
  applyCwColWidths();

  var totalMatched=filtered.length;
  var totalPages=Math.max(1,Math.ceil(totalMatched/_cwPageSize));
  if(_cwPage>totalPages)_cwPage=totalPages;
  var startIdx=(_cwPage-1)*_cwPageSize;
  var endIdx=Math.min(startIdx+_cwPageSize,totalMatched);
  var visible=filtered.slice(startIdx,endIdx);
  var countEl=document.getElementById('cwTableCount');
  if(countEl){
    if(!totalMatched)countEl.textContent='';
    else if(_cwPageSize===Infinity||totalMatched<=_cwPageSize)countEl.textContent=totalMatched+' caseworker'+(totalMatched===1?'':'s');
    else countEl.textContent='Showing '+(startIdx+1)+'–'+endIdx+' of '+totalMatched;
  }
  var pageControls=document.getElementById('cwPageControls');
  if(pageControls){
    if(totalPages<=1)pageControls.innerHTML='';
    else {
      var html='<button class="pg-btn" onclick="goToCwPage('+(_cwPage-1)+')"'+(_cwPage===1?' disabled':'')+'>‹</button>';
      var startPg=Math.max(1,_cwPage-2),endPg=Math.min(totalPages,startPg+4);
      if(endPg-startPg<4)startPg=Math.max(1,endPg-4);
      for(var p=startPg;p<=endPg;p++)html+='<button class="pg-btn'+(p===_cwPage?' active':'')+'" onclick="goToCwPage('+p+')">'+p+'</button>';
      html+='<button class="pg-btn" onclick="goToCwPage('+(_cwPage+1)+')"'+(_cwPage===totalPages?' disabled':'')+'>›</button>';
      pageControls.innerHTML=html;
    }
  }

  var tbody=document.getElementById('cwTableBody');if(!tbody)return;tbody.innerHTML='';
  var empty=document.getElementById('cwTableEmpty');
  if(!filtered.length){if(empty)empty.style.display='block';return;}
  if(empty)empty.style.display='none';
  visible.forEach(function(cw){
    var clientCount=Object.keys(profiles).filter(function(k){return (profiles[k].caseworkerId===cw.id||profiles[k].worker===cw.name) && (profiles[k].clientStatus||'active')==='active';}).length;
    var hrefCw=buildCaseworkerUrl(cw.id);
    var supName=cw.supervisor_id && sups[cw.supervisor_id] ? sups[cw.supervisor_id].name : '';
    var tr=document.createElement('tr');
    var checked=cwBulkSelected[cw.id]?'checked':'';
    tr.innerHTML=
      '<td style="width:26px;" onclick="event.stopPropagation()"><input type="checkbox" class="cw-select" data-id="'+esc(cw.id)+'" '+checked+' onchange="toggleBulkCaseworker(\''+escJsAttr(cw.id)+'\',this)" style="width:12px;height:12px;cursor:pointer;"></td>'+
      '<td><a href="'+hrefCw+'" class="link-plain" style="display:block;" onclick="return navClick(event,this.getAttribute(\'href\'))"><div class="ct-name">'+esc(cw.name||'')+'</div><div class="ct-id">'+esc(cw.agency||'No agency')+' '+_cwOrgBadge(cw)+'</div></a></td>'+
      '<td style="color:var(--text-muted);font-size:12px;">'+esc(cw.phone||'—')+'</td>'+
      '<td style="color:var(--text-muted);font-size:12px;">'+esc(cw.email||'—')+'</td>'+
      '<td style="color:var(--text-muted);font-size:12px;">'+esc(cw.county||'—')+'</td>'+
      '<td style="color:var(--text-muted);font-size:12px;">'+esc(supName||'—')+'</td>'+
      '<td style="color:var(--text);font-size:12px;">'+clientCount+'</td>';
    tr.addEventListener('click',function(e){
      if(e.target.closest('a')||e.target.closest('button')||e.target.closest('input'))return;
      openCwDetail(cw.id);
    });
    tbody.appendChild(tr);
  });

  var sel=document.getElementById('cwPageSize');
  if(sel){var savedPs=_cwPageSize===Infinity?'all':String(_cwPageSize);if(sel.value!==savedPs)sel.value=savedPs;}
}
function toggleBulkCaseworker(id,cb){
  if(cb.checked)cwBulkSelected[id]=true;else delete cwBulkSelected[id];
  var count=Object.keys(cwBulkSelected).length;
  var bar=document.getElementById('cwBulkBar');
  if(bar){bar.classList.toggle('visible',count>0);var lbl=document.getElementById('cwBulkCount');if(lbl)lbl.textContent=count+' selected';}
}
function clearCwBulkSelect(){cwBulkSelected={};var bar=document.getElementById('cwBulkBar');if(bar)bar.classList.remove('visible');var sa=document.getElementById('cwSelectAll');if(sa)sa.checked=false;renderCaseworkerList();}
function toggleAllCaseworkers(cb){
  document.querySelectorAll('.cw-select').forEach(function(c){
    c.checked=cb.checked;
    var id=c.dataset.id;
    if(cb.checked)cwBulkSelected[id]=true;else delete cwBulkSelected[id];
  });
  var count=Object.keys(cwBulkSelected).length;
  var bar=document.getElementById('cwBulkBar');
  if(bar){bar.classList.toggle('visible',count>0);var lbl=document.getElementById('cwBulkCount');if(lbl)lbl.textContent=count+' selected';}
}
function bulkDeleteCaseworkers(){
  var ids=Object.keys(cwBulkSelected);if(!ids.length){showAlert('No caseworkers selected.');return;}
  var cws=getCaseworkers();
  var byId={};cws.forEach(function(c){byId[c.id]=c;});
  var preview=ids.slice(0,5).map(function(id){return '• '+(byId[id]?byId[id].name||id:id);}).join('\n')+(ids.length>5?'\n• …and '+(ids.length-5)+' more':'');
  showConfirm(
    'Delete '+ids.length+' caseworker'+(ids.length>1?'s':'')+'?\n\n'+preview+'\n\nClients assigned to these caseworkers will be left without a caseworker. This cannot be undone.',
    function(){
      var keep=cws.filter(function(c){return ids.indexOf(c.id)===-1;});
      saveCaseworkersLS(keep);
      ids.forEach(function(id){try{deleteCaseworkerAPI(id);}catch(e){}});
      logActivity('delete','Bulk caseworker delete: '+ids.length+' removed');
      clearCwBulkSelect();
      showAlert(ids.length+' caseworker'+(ids.length!==1?'s':'')+' deleted.');
    },
    {title:'Delete '+ids.length+' Caseworkers?',okText:'Delete '+ids.length+(ids.length>1?' Caseworkers':' Caseworker'),danger:true}
  );
}

// ============================================================
//  CASEWORKER DETAIL VIEW
// ============================================================
var activeCwId=null;
function openCwDetail(id){
  var cw=getCaseworkers().find(function(c){return c.id===id;});
  if(!cw)return;
  activeCwId=id;
  // Same topbar breadcrumb pattern as Client + Caregiver detail: "Caseworkers > [name]".
  bc([{l:'Caseworkers',fn:navCaseworkers},{l:cw.name||''}]);
  document.getElementById('cwGridView').style.display='none';
  document.getElementById('cwFormWrap').style.display='none';
  document.getElementById('cwDetailView').style.display='block';
  var ini=(cw.name||'?').split(' ').map(function(w){return w[0]||'';}).slice(0,2).join('').toUpperCase();
  document.getElementById('cwDetailAvatar').textContent=ini;
  document.getElementById('cwDetailName').textContent=(cw.title?cw.title+' ':'')+(cw.name||'');
  document.getElementById('cwDetailMeta').innerHTML=esc(cw.agency||'')+(cw.phone?' · '+esc(cw.phone):'')+(cw.org?' &nbsp;'+_cwOrgBadge(cw):'');
  switchCwTab('overview');
}
function showCwGrid(){
  activeCwId=null;
  var dv=document.getElementById('cwDetailView');if(dv)dv.style.display='none';
  var gv=document.getElementById('cwGridView');if(gv)gv.style.display='';
  hideCaseworkerForm();
  renderCaseworkerList();
}
var cwUnsavedChanges=false;
function switchCwTab(tab){
  if(cwUnsavedChanges && tab!=='info'){
    showConfirm('You have unsaved changes on the caseworker Info tab. Leave anyway?',function(){
      cwUnsavedChanges=false;
      _doSwitchCwTab(tab);
    },{title:'Unsaved Changes',okText:'Discard & Leave',danger:true});
    return;
  }
  _doSwitchCwTab(tab);
}
function _doSwitchCwTab(tab){
  ['overview','info','clients','notes','docs','audit'].forEach(function(t){
    var tb=document.getElementById('cwtab-'+t);
    var pn=document.getElementById('cwpane-'+t);
    if(tb)tb.classList.toggle('active',t===tab);
    if(pn)pn.classList.toggle('active',t===tab);
  });
  if(tab==='overview')renderCwOverviewPane();
  if(tab==='info')renderCwInfoPane();
  if(tab==='clients')renderCwClientsPane();
  if(tab==='notes')renderCwNotesPane();
  if(tab==='docs')renderCwDocsPane();
  if(tab==='audit')renderCwAuditPane();
}
function renderCwOverviewPane(){
  if(!activeCwId)return;
  var cw=getCaseworkers().find(function(c){return c.id===activeCwId;});
  var pane=document.getElementById('cwpane-overview');
  if(!cw||!pane)return;
  var profiles=getProfiles();
  var assignedNames=Object.keys(profiles).filter(function(k){return profiles[k].worker===cw.name||profiles[k].caseworkerId===cw.id;}).sort();
  var addrParts=[cw.street,cw.city?(cw.city+(cw.state?' '+cw.state:'')+(cw.zip?' '+cw.zip:'')):''].filter(Boolean);
  var addrStr=addrParts.join(', ');
  var clientListHtml='';
  if(assignedNames.length){
    clientListHtml='<div style="display:flex;flex-direction:column;gap:4px;margin-top:6px;">'+
      assignedNames.map(function(name){
        var p=profiles[name]||{};
        var st=p.clientStatus||'active';
        var stColor=st==='active'?'#1e7e34':st==='inactive'?'#888':'#a83232';
        return '<div onclick="navDetail(\''+escJsAttr(name)+'\')" style="display:flex;align-items:center;gap:8px;padding:6px 8px;background:#f7faff;border:1px solid #e1e5ea;border-radius:5px;cursor:pointer;font-size:12px;" onmouseover="this.style.borderColor=\'#b0c8e8\'" onmouseout="this.style.borderColor=\'#e1e5ea\'">'+
          '<span style="flex:1;color:#185FA5;font-weight:500;">'+esc(name)+'</span>'+
          (p.medicaidId?'<span style="color:#5c7590;">'+esc(p.medicaidId)+'</span>':'')+
          '<span style="color:'+stColor+';font-size:10px;font-weight:600;text-transform:uppercase;">'+(st==='inactive'?'In Progress':st)+'</span>'+
        '</div>';
      }).join('')+
    '</div>';
  } else {
    clientListHtml='<div style="font-size:12px;color:#5c7590;margin-top:6px;">No clients assigned.</div>';
  }
  pane.innerHTML='<div class="overview-grid">'+
    '<div class="ov-card"><h4>Contact Info</h4>'+
      (cw.phone?'<div class="ov-row"><span class="ov-label">Phone</span><span class="ov-value">'+esc(cw.phone)+'</span></div>':'')+
      (cw.fax?'<div class="ov-row"><span class="ov-label">Fax</span><span class="ov-value">'+esc(cw.fax)+'</span></div>':'')+
      (cw.email?'<div class="ov-row"><span class="ov-label">Email</span><span class="ov-value">'+esc(cw.email)+'</span></div>':'')+
      (cw.agency?'<div class="ov-row"><span class="ov-label">Agency</span><span class="ov-value">'+esc(cw.agency)+'</span></div>':'')+
      (addrStr?'<div class="ov-row"><span class="ov-label">Address</span><span class="ov-value">'+esc(addrStr)+'</span></div>':'')+
    '</div>'+
    '<div class="ov-card"><h4>Assigned Clients ('+assignedNames.length+')</h4>'+
      clientListHtml+
    '</div>'+
  '</div>';
}
function renderCwInfoPane(){
  if(!activeCwId)return;
  var cw=getCaseworkers().find(function(c){return c.id===activeCwId;});
  var c=document.getElementById('cwInfoContent');c.innerHTML='';
  if(!cw)return;
  cwUnsavedChanges=false;
  c.oninput=function(){cwUnsavedChanges=true;};
  c.onchange=function(){cwUnsavedChanges=true;};

  var g=document.createElement('div');g.className='info-grid';c.appendChild(g);

  function mkF(id,label,val,full,attrs){
    var d=document.createElement('div');d.className='info-field'+(full?' full':'');
    d.innerHTML='<label for="'+id+'">'+label+'</label><input id="'+id+'" value="'+esc(val||'')+'"'+(attrs||'')+'>';g.appendChild(d);
  }
  function mkDiv(label){var d=document.createElement('div');d.className='form-section-divider full';d.innerHTML='<span>'+label+'</span>';g.appendChild(d);}
  function mkRow(html){var d=document.createElement('div');d.className='info-field-row full';d.innerHTML=html;g.appendChild(d);}

  var firstName=cw.first_name||(cw.name||'').split(' ')[0]||'';
  var lastName=cw.last_name||(cw.name||'').split(' ').slice(1).join(' ')||'';
  var dName=document.createElement('div');dName.className='info-field-row full';dName.style.gridTemplateColumns='84px 1fr 1fr 1fr';
  dName.innerHTML='<div class="info-field"><label for="cwi-title">Title</label><select id="cwi-title">'+['','Mr.','Mrs.','Ms.','Miss','Dr.','Mx.'].map(function(t){return '<option'+(String(cw.title||'')===t?' selected':'')+'>'+t+'</option>';}).join('')+'</select></div>'+
    '<div class="info-field"><label for="cwi-first">First Name *</label><input id="cwi-first" value="'+esc(firstName)+'"></div>'+
    '<div class="info-field"><label for="cwi-middle">Middle Name</label><input id="cwi-middle" value="'+esc(cw.middle_name||'')+'"></div>'+
    '<div class="info-field"><label for="cwi-last">Last Name *</label><input id="cwi-last" value="'+esc(lastName)+'"></div>';
  g.appendChild(dName);

  mkF('cwi-agency','Agency',cw.agency,true);
  var _cwOrgOpts=['','MDHHS','Humana','Priority Health','Aetna','UnitedHealthcare','HAP','Molina','Meridian','Other'];
  var dCwOrg=document.createElement('div');dCwOrg.className='info-field full';
  dCwOrg.innerHTML='<label for="cwi-org">Organization <span style="font-weight:400;font-size:10px;color:#5c7590;">(MDHHS, or the carrier for a managed-care coordinator)</span></label>'+
    '<select id="cwi-org" onchange="cwiOrgToggle()">'+_cwOrgOpts.map(function(o){return '<option value="'+esc(o)+'"'+((cw.org||'')===o?' selected':'')+'>'+(o||'— Select —')+'</option>';}).join('')+'</select>';
  g.appendChild(dCwOrg);
  cwiOrgToggle();   // Agency (MDHHS office) shows only when Org = MDHHS
  mkRow('<div class="info-field"><label for="cwi-phone">Phone</label><input id="cwi-phone" value="'+esc(cw.phone||'')+'"></div>'+
    '<div class="info-field"><label for="cwi-fax">Fax</label><input id="cwi-fax" value="'+esc(cw.fax||'')+'"></div>');
  mkF('cwi-email','Email',cw.email,true,' onblur="_autoGovEmail(this)"');
  // Supervisor dropdown + Add/Edit buttons (CC'd on monthly invoice emails when set)
  var supDiv=document.createElement('div');supDiv.className='info-field full';
  supDiv.innerHTML='<label for="cwi-supervisor">Supervisor <span style="font-weight:400;font-size:10px;color:#5c7590;">(CC\'d on monthly invoice emails when set)</span></label>'+
    '<div style="display:flex;gap:6px;align-items:center;">'+
      '<select id="cwi-supervisor" style="flex:1;"></select>'+
      '<button type="button" class="btn btn-secondary btn-sm" onclick="openSupervisorModal()" style="white-space:nowrap;">+ Add</button>'+
      '<button type="button" id="cwi-supervisor-edit-btn" class="btn btn-secondary btn-sm" onclick="editSelectedSupervisor()" style="white-space:nowrap;display:none;">Edit</button>'+
    '</div>';
  g.appendChild(supDiv);
  // Populate after element is attached so the populate fn can find #cwi-supervisor
  setTimeout(function(){populateSupervisorDropdown(document.getElementById('cwi-supervisor'),cw.supervisor_id||'');},0);

  mkDiv('Address');
  mkF('cwi-street','Street',cw.street,true);
  mkRow('<div class="info-field"><label for="cwi-city">City</label><input id="cwi-city" value="'+esc(cw.city||'')+'"></div>'+
    '<div class="info-field"><label for="cwi-state">State</label><input id="cwi-state" value="'+esc(cw.state||'')+'"></div>');
  mkRow('<div class="info-field"><label for="cwi-zip">ZIP</label><input id="cwi-zip" value="'+esc(cw.zip||'')+'" oninput="lookupZip(\'cwi-zip\',\'cwi-city\',\'cwi-state\',\'cwi-county\')"></div>'+
    '<div class="info-field"><label for="cwi-county">County</label><input id="cwi-county" value="'+esc(cw.county||'')+'"></div>');

  var actions=document.createElement('div');actions.style.cssText='margin-top:16px;display:flex;gap:8px;';
  actions.innerHTML='<button class="btn btn-primary" id="cwSaveInfoBtn" onclick="saveCwInfoPane()">Save Changes</button>'+
    '<button class="btn btn-danger btn-sm" onclick="deleteCaseworkerFromDetail()" style="padding:6px 14px;">Delete Caseworker</button>';
  c.appendChild(actions);
}
function saveCwInfoPane(){
  if(!activeCwId)return;
  var arr=getCaseworkers();var cw=arr.find(function(c){return c.id===activeCwId;});if(!cw)return;
  var first=(document.getElementById('cwi-first').value||'').trim();
  var last=(document.getElementById('cwi-last').value||'').trim();
  if(!first||!last){showAlert('First and last name are required.');return;}
  var titleEl=document.getElementById('cwi-title');cw.title=titleEl?(titleEl.value||''):(cw.title||'');
  cw.first_name=first;cw.middle_name=document.getElementById('cwi-middle').value;cw.last_name=last;
  cw.name=(first+(document.getElementById('cwi-middle').value?' '+document.getElementById('cwi-middle').value:'')+' '+last).trim();
  cw.agency=document.getElementById('cwi-agency').value;
  var _cwiOrg=document.getElementById('cwi-org');if(_cwiOrg)cw.org=_cwiOrg.value;
  if((cw.org||'')!=='MDHHS')cw.agency='';   // Agency is MDHHS-only (hidden otherwise) — don't keep a stale one
  cw.phone=document.getElementById('cwi-phone').value;cw.fax=document.getElementById('cwi-fax').value;
  cw.email=document.getElementById('cwi-email').value;
  cw.street=document.getElementById('cwi-street').value;cw.city=document.getElementById('cwi-city').value;
  cw.state=document.getElementById('cwi-state').value;cw.zip=document.getElementById('cwi-zip').value;cw.county=document.getElementById('cwi-county').value;
  var cwiSup=document.getElementById('cwi-supervisor');if(cwiSup)cw.supervisor_id=cwiSup.value||'';
  saveCaseworkersLS(arr);saveCaseworkerAPI(cw);
  document.getElementById('cwDetailName').textContent=(cw.title?cw.title+' ':'')+cw.name;
  document.getElementById('cwDetailMeta').innerHTML=esc(cw.agency||'')+(cw.phone?' · '+esc(cw.phone):'')+(cw.org?' &nbsp;'+_cwOrgBadge(cw):'');
  var btn=document.getElementById('cwSaveInfoBtn');if(btn){btn.textContent='Saved ✓';setTimeout(function(){btn.textContent='Save Changes';},1800);}
  addAuditEntry(cw.name,'Caseworker profile updated');
  cwUnsavedChanges=false;
  renderCaseworkerList();
}
function renderCwClientsPane(){
  if(!activeCwId)return;
  var cw=getCaseworkers().find(function(c){return c.id===activeCwId;});
  var c=document.getElementById('cwClientsContent');
  if(!c||!cw)return;
  var profiles=getProfiles();
  var assigned=Object.keys(profiles).filter(function(k){return profiles[k].worker===cw.name||profiles[k].caseworkerId===cw.id;});
  if(!assigned.length){c.innerHTML='<div class="empty-state"><h3>No clients assigned</h3><p style="font-size:13px;">Assign this caseworker from a client\'s Profile tab.</p></div>';return;}
  c.innerHTML='';
  assigned.forEach(function(name){
    var prof=profiles[name];
    var st=prof.clientStatus||'active';
    var ini=name.split(' ').map(function(w){return w[0]||'';}).slice(0,2).join('').toUpperCase();
    var row=document.createElement('div');
    row.style.cssText='display:flex;align-items:center;gap:12px;padding:11px 14px;background:#fff;border:1px solid #e1e5ea;border-radius:8px;margin-bottom:8px;cursor:pointer;transition:border-color 0.15s;max-width:500px;';
    row.innerHTML='<div class="cc-avatar" style="width:36px;height:36px;font-size:13px;">'+ini+'</div>'+
      '<div style="flex:1;min-width:0;">'+
        '<div style="font-size:13px;font-weight:600;color:#1a2b45;">'+esc(name)+'</div>'+
        '<div style="font-size:11px;color:#4d6c88;">'+esc(prof.medicaidId||'No Medicaid ID')+(prof.phone?' · '+esc(prof.phone):'')+'</div>'+
      '</div>'+
      '<span class="cs-badge cs-'+st+'">'+clientStatusLabel(st)+'</span>'+
      '<span style="font-size:11px;color:#185FA5;font-weight:500;">Open →</span>';
    row.addEventListener('mouseenter',function(){this.style.borderColor='#b0c8e8';});
    row.addEventListener('mouseleave',function(){this.style.borderColor='#e1e5ea';});
    row.addEventListener('click',function(){navDetail(name);});
    c.appendChild(row);
  });
}
function renderCwNotesPane(){
  if(!activeCwId)return;
  var cw=getCaseworkers().find(function(c){return c.id===activeCwId;});
  var c=document.getElementById('cwNotesContent');
  if(!c||!cw)return;
  c.innerHTML='<div style="margin-bottom:6px;font-size:11px;color:#5c7590;">Auto-saves as you type <span id="cwNotesSavedFlash" style="display:none;color:#1a7740;font-weight:600;">· Saved ✓</span></div>'+
    '<textarea id="cwNotesArea" style="width:100%;min-height:200px;padding:12px;border:1px solid #d0d8e4;border-radius:6px;font-size:13px;font-family:Arial,sans-serif;outline:none;resize:vertical;max-width:620px;">'+esc(cw.notes||'')+'</textarea>';
  var ta=document.getElementById('cwNotesArea');
  var t=null;
  ta.addEventListener('input',function(){
    clearTimeout(t);
    t=setTimeout(function(){
      var arr=getCaseworkers();
      var cwRec=arr.find(function(c){return c.id===activeCwId;});
      if(!cwRec)return;
      cwRec.notes=ta.value;
      saveCaseworkersLS(arr);
      // D8: only claim "Saved ✓" after the API resolves; surface failure otherwise.
      var doSave=function(){return flashQuietSave(saveCaseworkerAPI(cwRec,true),'cwNotesSavedFlash','Caseworker note',doSave);};
      doSave();
    },600);
  });
}
function renderCwDocsPane(){
  if(!activeCwId)return;
  var cw=getCaseworkers().find(function(c){return c.id===activeCwId;});
  var c=document.getElementById('cwDocsContent');
  if(!c||!cw)return;
  c.innerHTML=docUploaderHtml({title:'Documents for '+esc(cw.name||'Caseworker'),subtitle:'Letters, authorization docs, etc.',hasCategory:true,catId:'cwDocCategory',fileId:'cwDocFileInput',scanId:'cwDocScanInput',scanFn:'handleCwDocScan(this)',uploadFn:'uploadCwDoc()',statusId:'cwDocUploadStatus',listId:'cwDocListAzure'});
  fetch(API_BASE+'/documents?clientType=caseworker&clientId='+activeCwId,{headers:apiHeaders()})
    .then(function(r){ if(!r.ok)throw new Error('HTTP '+r.status); return r.json(); })
    .then(function(docs){ renderCwDocListAzure(activeCwId, Array.isArray(docs)?docs:[]); })
    .catch(function(err){ _renderDocLoadError('cwDocListAzure', 'renderCwDocsPane()', err); });
}
function renderCwDocListAzure(cwId,docs){
  renderDocGrid(document.getElementById('cwDocListAzure'),docs,{clientType:'caseworker',clientId:cwId,deleteExpr:function(name){return 'deleteCwDoc(\''+cwId+'\',\''+encodeURIComponent(name)+'\')';},refresh:function(){renderCwDocsPane();}});
}
function uploadCwDoc(){
  var input=document.getElementById('cwDocFileInput');
  if(!input||!input.files||!input.files.length){showAlert('Please select a file first.');return;}
  var status=document.getElementById('cwDocUploadStatus');status.textContent='Uploading...';
  var fd=new FormData();fd.append('clientType','caseworker');fd.append('clientId',activeCwId);fd.append('category',(document.getElementById('cwDocCategory')||{}).value||'Other');
  Array.from(input.files).forEach(function(f){fd.append('file',f);});
  fetch(API_BASE+'/documents',{method:'POST',headers:authUploadHeaders(),body:fd})
    .then(function(r){if(!r.ok)throw new Error('Upload failed ('+r.status+')');return r;})
    .then(function(){status.textContent='';input.value='';renderCwDocsPane();})
    .catch(function(e){status.textContent='Upload failed: '+e;});
}
function deleteCwDoc(cwId,encodedName){
  showConfirm('Delete this document?',function(){
    fetch(API_BASE+'/documents?clientType=caseworker&clientId='+cwId,{method:'DELETE',headers:apiHeaders(),body:JSON.stringify({name:decodeURIComponent(encodedName)})})
      .then(function(r){if(!r.ok)throw new Error('HTTP '+r.status);renderCwDocsPane();}).catch(function(e){showAlert('Delete failed: '+e);});
  },{title:'Delete Document',okText:'Delete'});
}
function handleCwDocScan(input){
  if(!activeCwId){showAlert('Open a caseworker first.');return;}
  if(!input||!input.files||!input.files.length)return;
  var status=document.getElementById('cwDocUploadStatus');
  if(status)status.textContent='Uploading scanned image…';
  var fd=new FormData();fd.append('clientType','caseworker');fd.append('clientId',activeCwId);fd.append('category',(document.getElementById('cwDocCategory')||{}).value||'Other');
  fd.append('file',input.files[0]);
  fetch(API_BASE+'/documents',{method:'POST',headers:authUploadHeaders(),body:fd})
    .then(function(r){if(!r.ok)throw new Error('Upload failed ('+r.status+')');return r;})
    .then(function(){if(status)status.textContent='';input.value='';renderCwDocsPane();})
    .catch(function(e){if(status)status.textContent='Upload failed: '+e;});
}
function renderCwAuditPane(){
  if(!activeCwId)return;
  var pane=document.getElementById('cwpane-audit');
  var cw=getCaseworkers().find(function(c){return c.id===activeCwId;});
  if(!pane||!cw)return;
  pane.innerHTML='<div style="color:#5c7590;font-size:13px;padding:8px 0;">Loading…</div>';
  if(spToken){
    fetch(API_BASE+'/audit/search',{method:'POST',headers:apiHeaders(),body:JSON.stringify({client:cw.name||activeCwId,limit:100})})
      .then(function(r){return r.ok?r.json():Promise.reject();})
      .then(function(rows){
        pane.innerHTML='';
        if(!rows.length){pane.innerHTML='<div style="color:#5c7590;font-size:13px;padding:16px 0;">No audit entries yet.</div>';return;}
        var wrap=document.createElement('div');wrap.style.cssText='max-width:600px;';
        rows.forEach(function(e){
          var row=document.createElement('div');row.className='audit-row';
          var ts=e.created_at?new Date(e.created_at).toLocaleString():e.ts||'';
          row.innerHTML='<span class="audit-icon">—</span><div><div class="audit-text">'+esc(e.action)+'</div><div class="audit-who">'+esc(e.who)+' · '+esc(ts)+'</div></div>';
          wrap.appendChild(row);
        });
        pane.appendChild(wrap);
      })
      .catch(function(){pane.innerHTML='<div style="color:#5c7590;font-size:13px;padding:16px 0;">No audit entries yet.</div>';});
  } else {
    pane.innerHTML='<div style="color:#5c7590;font-size:13px;padding:16px 0;">Sign in to view audit log.</div>';
  }
}

// ============================================================
//  KEYBOARD SHORTCUTS & NAVIGATION GUARD
// ============================================================
document.addEventListener('keydown',function(e){
  // a11y: Escape dismisses the topmost open modal (keyboard users had no way to close one).
  if(e.key==='Escape'){
    var open=document.querySelectorAll('.modal-overlay.open');
    if(open.length){ open[open.length-1].classList.remove('open'); e.preventDefault(); return; }
  }
  if((e.ctrlKey||e.metaKey)&&e.key==='s'){
    e.preventDefault();
    // Save whichever context is active
    var activePage=document.querySelector('.page.active');
    if(!activePage)return;
    var id=activePage.id;
    if(id==='page-invoice'){saveInvoiceToClient();}
    else if(id==='page-detail'){
      var activeTab=document.querySelector('.dtab.active');
      if(activeTab&&activeTab.id==='dtab-info')saveClientInfo();
    }
  }
});
window.addEventListener('beforeunload',function(e){
  // Warn on any in-progress edit — client, caregiver, OR caseworker panes.
  var dirty=unsavedChanges||
    (typeof cgUnsavedChanges!=='undefined'&&cgUnsavedChanges)||
    (typeof cwUnsavedChanges!=='undefined'&&cwUnsavedChanges);
  if(dirty){e.preventDefault();e.returnValue='';}
});
// HIPAA: wipe cached PHI (lhca_* + in-memory SSN/signature caches) when the app is actually
// being left or the tab closed — so the client roster doesn't sit at rest in the browser
// profile after the session ends. Internal SPA navigation is hash-based and does NOT fire
// pagehide, so this only triggers on real unload. On the next visit the roster reloads from
// the server. Skip when the page is going into bfcache (persisted) since it may be restored.
window.addEventListener('pagehide',function(e){
  if(e&&e.persisted)return;
  try{clearPHIFromStorage();}catch(_){}
});

// ============================================================
//  AUTO SESSION TIMEOUT (HIPAA — 45 min inactivity)
// ============================================================
var SESSION_TIMEOUT_MS  = 45 * 60 * 1000; // sign out after 45 min idle
var SESSION_WARN_MS     = 43 * 60 * 1000; // warn at 43 min
var _sessionTimer, _sessionWarnTimer;

function resetSessionTimer(){
  clearTimeout(_sessionTimer);
  clearTimeout(_sessionWarnTimer);
  if(!spToken) return; // only enforce when signed in
  _sessionWarnTimer = setTimeout(function(){
    showToast('Warning: Session expiring in 2 minutes due to inactivity. Click anywhere to stay signed in.', 9000);
  }, SESSION_WARN_MS);
  _sessionTimer = setTimeout(function(){
    aiTrack('SessionTimeout',{reason:'inactivity'});
    clearPHIFromStorage();
    spToken = null;
    // Kill the API bearer token AND stop its auto-refresh, otherwise the pending
    // refresh timer silently re-acquires it and the "sign-out" isn't real.
    _apiToken = null;
    clearTimeout(_apiTokenTimer); _apiTokenTimer = null;
    updateAuthUI(false);
    showToast('Signed out automatically due to inactivity.', 5000);
  }, SESSION_TIMEOUT_MS);
}
// Restart the timer on GENUINE interaction only. mousemove and scroll used to be in
// this list, which meant any cursor drift — or an animated/auto-scrolling element —
// kept an unattended session alive forever, defeating the purpose.
['mousedown','keydown','touchstart','click'].forEach(function(ev){
  document.addEventListener(ev, resetSessionTimer, {passive:true});
});

// ============================================================
//  INIT
// ============================================================
var T=today();
document.getElementById('dateSubmitted').value=T;document.getElementById('sigDate1').value=T;document.getElementById('sigDate2').value=T;
rebuild(31);
// Legacy sig migration: if old single sig exists, port to new array
try{
  var oldSig=localStorage.getItem('lhca_sig');
  if(oldSig){var sigs=getSigs();if(!sigs.length){sigs.push({id:sigId(),label:'Thomas Jaboro',data:oldSig});saveSigsLS(sigs);localStorage.removeItem('lhca_sig');}}
}catch(e){}
try{initMSAL();}catch(e){console.log('MSAL init:',e);}
renderActivityFeed();
updateTaskBadge();
// If the URL has a hash route (e.g. #/client/Adnan), navigate there; else default to home.
if(window.location.hash){routeFromHash();}else{navHome();}

// ============================================================
//  SESSION 2 — FORMS TAB
// ============================================================
var activeFormType='';
var activeFormClientName='';

function navForms(){
  showPage('forms');
  bc([{l:'Forms'}]);
  document.getElementById('topbarActions').innerHTML='';
  renderFormsClientDropdown();
}
function renderFormsClientDropdown(){
  var profiles=getProfiles();
  var sel=document.getElementById('formClientSelect');if(!sel)return;
  var prev=activeFormClientName||sel.value;
  sel.innerHTML='<option value="">— Select a client to pre-fill —</option>';
  Object.keys(profiles).sort().forEach(function(name){
    var o=document.createElement('option');o.value=name;o.textContent=name;sel.appendChild(o);
  });
  if(prev)sel.value=prev;
  activeFormClientName=sel.value;
}
function openStateForm(type){
  activeFormType=type;
  var selEl=document.getElementById('formClientSelect');
  activeFormClientName=selEl?selEl.value:'';
  var titles={msa4676:'MSA-4676 — Home Help Services Agreement',
              dhs390:'DHS-390 — Application for Home Help / Adult Services'};
  showPage('form-fill');
  bc([{l:'Forms',fn:navForms},{l:titles[type]||type}]);
  ['sb-home','sb-caregivers','sb-settings','sb-tasks','sb-reports','sb-caseworkers','sb-forms'].forEach(function(id){var el=document.getElementById(id);if(el)el.classList.remove('active');});
  var fb=document.getElementById('sb-forms');if(fb)fb.classList.add('active');
  document.getElementById('formFillTitle').textContent=titles[type]||'';
  // Single-pane PDF view — auto-fill from CRM on load, user can switch clients here
  var profsForPicker=getProfiles();
  // Show every client regardless of status — In Progress / onboarding clients are
  // exactly the ones who need these state forms, so never filter by status here
  // (matches the Forms-page picker, which already lists all clients).
  var clientNames=Object.keys(profsForPicker).sort();
  var clientOpts='<option value=""'+(!activeFormClientName?' selected':'')+'>— Pick a client —</option>'+
    clientNames.map(function(n){return '<option value="'+esc(n)+'"'+(n===activeFormClientName?' selected':'')+'>'+esc(n)+'</option>';}).join('');
  document.getElementById('formFillContent').innerHTML=
    '<div style="max-width:1100px;margin:0 auto;">'+
      '<div style="font-size:12px;color:#5a7296;margin-bottom:8px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px;">'+
        '<div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;">'+
          '<label style="font-size:12px;font-weight:600;color:#1a2b45;white-space:nowrap;" for="sfClientPicker">Pre-fill for:</label>'+
          '<select id="sfClientPicker" onchange="changeStateFormClient(this.value)" style="padding:5px 8px;border:1px solid #d0d8e4;border-radius:5px;font-size:12px;background:#fff;min-width:220px;">'+clientOpts+'</select>'+
          (activeFormClientName?'<span style="color:#1a7740;font-size:11px;">✓ pre-filled</span>':'<span style="color:#a05a00;font-size:11px;">type directly, or pick a client to auto-fill</span>')+
        '</div>'+
        '<span id="sfPreviewStatusText" style="font-size:11px;color:#4d6c88;">Loading…</span>'+
      '</div>'+
      '<iframe id="sfPreviewFrame" style="width:100%;height:calc(100vh - 130px);border:1px solid #d0d8e4;border-radius:8px;background:#f0f3f7;" title="Form"></iframe>'+
    '</div>';
  scheduleSfPreview(0);
  var topbar=document.querySelector('.form-topbar');
  if(topbar){
    var titleSpan=document.getElementById('formFillTitle');
    topbar.innerHTML='';
    var back=document.createElement('button');back.className='btn btn-secondary btn-sm';back.textContent='← Back to Forms';back.onclick=navForms;
    topbar.appendChild(back);
    if(titleSpan)topbar.appendChild(titleSpan);
    var spacer=document.createElement('div');spacer.style.flex='1';topbar.appendChild(spacer);
    // "Fillable" is offered for templates that really have form fields (the DHS-390 does; the
    // MSA-4676's template has none, so an unflattened copy of it would be no more editable).
    if(_formHasAcroFields(activeFormType)){
      var dlf=document.createElement('button');dlf.className='btn btn-secondary';dlf.textContent='⬇ Download fillable';
      dlf.title='Download with the PDF form fields still editable — for the sections the client fills in themselves.';
      dlf.onclick=function(){downloadStateFormPdf(true);};
      topbar.appendChild(dlf);
    }
    var dl=document.createElement('button');dl.className='btn btn-primary';dl.textContent='⬇ Download PDF';dl.title='Download the auto-filled PDF, flattened for printing and signature. Upload the signed scan via the Documents tab.';dl.onclick=function(){downloadStateFormPdf(false);};
    topbar.appendChild(dl);
  }
}

// True when the form's TEMPLATE is a real AcroForm (fillable). Seeded for the DHS-390 and corrected
// from the last render, so it stays right if a template is swapped for a fillable revision.
var _sfAcroForms={dhs390:true};
function _formHasAcroFields(type){ return _sfAcroForms[type]===true; }
// Re-render with a different client without leaving the form view
function changeStateFormClient(name){
  activeFormClientName=name||'';
  var statusEl=document.querySelector('#sfClientPicker').nextElementSibling;
  if(statusEl){
    statusEl.style.color=name?'#1a7740':'#a05a00';
    statusEl.style.fontSize='11px';
    statusEl.textContent=name?'✓ pre-filled':'type directly, or pick a client to auto-fill';
  }
  // Keep the Forms-page picker in sync
  var formsPagePicker=document.getElementById('formClientSelect');
  if(formsPagePicker)formsPagePicker.value=name||'';
  scheduleSfPreview(0);
}

// Re-render the PDF in the iframe whenever an input changes (debounced)
var _sfPreviewTimer=null,_sfPreviewBlobUrl=null,_sfPreviewBytes=null;
function scheduleSfPreview(delay){
  clearTimeout(_sfPreviewTimer);
  var ms=(delay==null||typeof delay==='object')?350:delay;
  var s=document.querySelector('.sf-preview-status');var t=document.getElementById('sfPreviewStatusText');
  if(s){s.classList.remove('idle');}if(t)t.textContent='Updating preview…';
  _sfPreviewTimer=setTimeout(renderSfPreview,ms);
}
// Each state form's left-pane input IDs → dict keys. Lets us read live-edited
// values from the form on the left and flow them into the PDF preview/save.
// Only the MSA-4676 is offered in the Forms UI now — the DHS-390 / DHS-4771 / MDHHS-6200 / BPHASA-2421
// definitions were removed as dead code when their form cards were retired.
var STATE_FORM_INPUT_MAPS={
  msa4676:{
    msa_cname:'client_name', msa_mid:'medicaid_id',
    msa_addr:'client_address', msa_city:'client_city', msa_st:'client_state', msa_zip:'client_zip',
    msa_phone:'client_phone', msa_county:'client_county',
    msa_cgname:'caregiver_full_name', msa_cgphone:'caregiver_phone',
    msa_start:'start_date', msa_date:'today_date'
  }
};
// Checkboxes to tick on an AcroForm template, by exact field name. Only ever used for facts the
// FORM ITSELF establishes — the DHS-390 is the Home Help application, so "Home Help" is the service
// being applied for. Anything the client attests to (which benefits they receive, who they live
// with) is left for them to check, since they sign it.
var STATE_FORM_CHECKS={
  dhs390:{ 'Home Help':true }
};
// Convert YYYY-MM-DD (HTML date-input format) → MM/DD/YYYY (state-form format)
function _normalizeDate(v){
  if(!v)return '';
  var m=String(v).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if(m)return m[2]+'/'+m[3]+'/'+m[1];
  return v;
}

// Agency (provider) info — auto-filled onto state forms (Section 2 of MSA-4676, etc.).
// These are the DEFAULTS; the live values are editable in Settings and stored in localStorage
// (lhca_agency), so the owner can update them without a code change. Not PHI — it's the agency's
// own business info — so lhca_agency is in the clearPHIFromStorage KEEP whitelist (survives wipe).
// Defaults as a hoisted function (not a top-level var) so it's available everywhere regardless of
// file position — including test harnesses that don't execute the whole file top-to-bottom.
function _agencyDefaults(){
  return {
    agency_provider_name:'Thomas Jaboro',
    agency_provider_id:'6221933',
    agency_address:'2741 Balsam Way Dr',
    agency_city:'Sterling Heights',
    agency_state:'MI',
    agency_zip:'48314',
    agency_phone:'(248) 291-4106',
    agency_relationship:'N/A - Agency'
  };
}
function getAgencyInfo(){
  var D=_agencyDefaults();
  var saved={}; try{ saved=JSON.parse(localStorage.getItem('lhca_agency')||'{}')||{}; }catch(e){ saved={}; }
  // Merge over defaults so a partially-filled/older saved object still yields every field.
  var out={}; for(var k in D){ out[k]=(saved[k]!==undefined&&saved[k]!=='')?saved[k]:D[k]; }
  return out;
}
function saveAgencyInfo(obj){
  var D=_agencyDefaults();
  var clean={}; for(var k in D){ clean[k]=(obj&&obj[k]!=null)?String(obj[k]):D[k]; }
  localStorage.setItem('lhca_agency', JSON.stringify(clean));
  // _pushSettings swallowed failures, and loadSettingsAPI treats the SERVER as the source of truth —
  // so a failed push silently reverted the agency details (which are stamped onto the MSA-4676) on the
  // next reload, on this device too. Report the real outcome.
  var _agencyPush=(typeof _pushSettings==='function')?_pushSettings({agency:clean}):null;
  if(_agencyPush&&_agencyPush.catch)_agencyPush.catch(function(e){
    if(e&&e.notSignedIn)return;   // local-only until sign-in — expected, not a failure
    showAlert('The agency details were saved on this device, but the server did not accept them ('+((e&&e.message)||'error')+').\n\nThey will revert on the next reload and forms will use the old details. Try saving again.',{title:'Agency details not synced'});
  });   // sync across devices
  return clean;
}
// Settings > Agency Information — populate the inputs from the stored/default values.
function renderAgencySettings(){
  var a=getAgencyInfo();
  var set=function(id,v){var e=document.getElementById(id);if(e)e.value=v||'';};
  set('ag-name',a.agency_provider_name); set('ag-id',a.agency_provider_id); set('ag-phone',a.agency_phone);
  set('ag-address',a.agency_address); set('ag-city',a.agency_city); set('ag-state',a.agency_state); set('ag-zip',a.agency_zip);
}
function saveAgencyFromSettings(){
  var g=function(id){var e=document.getElementById(id);return e?e.value.trim():'';};
  saveAgencyInfo({
    agency_provider_name:g('ag-name'), agency_provider_id:g('ag-id'), agency_phone:g('ag-phone'),
    agency_address:g('ag-address'), agency_city:g('ag-city'), agency_state:g('ag-state'), agency_zip:g('ag-zip'),
    agency_relationship:getAgencyInfo().agency_relationship
  });
  var s=document.getElementById('agencyStatus'); if(s){s.textContent='Saved ✓';setTimeout(function(){s.textContent='';},1800);}
}
// Set of dict keys that hold dates and should be normalized to MM/DD/YYYY
var _DATE_DICT_KEYS={client_dob:1,caregiver_dob:1,signature_date:1,today_date:1,last_seen:1,resolved_date:1,start_date:1};

// Build the data dictionary the renderer fills AcroForm fields from.
// CRM data is the default; any non-empty value typed into the left-pane form
// overrides it (so live edits flow into the PDF preview).
// clientName defaults to the Forms-tab selection (activeFormClientName) so existing callers are
// unchanged; the Assistant passes a name explicitly to build a form for any client without the UI.
function _buildFormDataDict(clientName){
  if(clientName==null)clientName=activeFormClientName;
  var prof=clientName?(getProfiles()[clientName]||{}):{};
  var cw=getCaseworkers().find(function(c){return (c.name||'').toLowerCase()===(prof.worker||'').toLowerCase();})||{};
  var cgs=getCaregivers();
  var assignedCg=(prof.caregiverId&&cgs[prof.caregiverId])||{};
  var td=today();
  var _agency=getAgencyInfo();
  var fullName=clientName||'';
  var fParts=fullName.split(' ');
  var firstN=fParts[0]||'',lastN=fParts.slice(1).join(' ')||'';
  var cgFull=assignedCg.name||'';
  var cgParts=cgFull.split(' ');
  var cgFirst=assignedCg.firstName||cgParts[0]||'';
  var cgLast=assignedCg.lastName||cgParts.slice(1).join(' ')||'';
  var dict={
    client_name:fullName, client_first_name:firstN, client_last_name:lastN,
    client_dob:_normalizeDate(prof.dob||''), medicaid_id:prof.medicaidId||'', case_number:'', recipient_id:prof.medicaidId||'',
    client_address:prof.street||prof.address||'', client_city:prof.city||'', client_state:prof.state||'MI', client_zip:prof.zip||'',
    client_phone:prof.phone||'', client_email:prof.clientEmail||prof.cemail||'', client_county:prof.county||'',
    worker_name:cw.name||prof.worker||'', worker_phone:cw.phone||'', worker_email:cw.email||'', worker_fax:'',
    caregiver_first_name:cgFirst, caregiver_last_name:cgLast, caregiver_full_name:cgFull,
    caregiver_dob:_normalizeDate(assignedCg.dob||''), caregiver_address:assignedCg.street||assignedCg.address||'',
    caregiver_city:assignedCg.city||'', caregiver_state:assignedCg.state||'MI', caregiver_zip:assignedCg.zip||'',
    caregiver_phone:assignedCg.phone||'', caregiver_email:assignedCg.email||'',
    caregiver_champs_id:assignedCg.champsId||assignedCg.champs_id||'',
    // Agency (editable in Settings — used as Section 2 on MSA-4676, etc.)
    agency_provider_name:_agency.agency_provider_name,
    agency_provider_id:_agency.agency_provider_id,
    agency_address:_agency.agency_address,
    agency_city:_agency.agency_city,
    agency_state:_agency.agency_state,
    agency_zip:_agency.agency_zip,
    agency_phone:_agency.agency_phone,
    agency_relationship:_agency.agency_relationship,
    today_date:td, signature_date:td, log_number:''
  };
  // Final pass: normalize any date-key in the dict that's still YYYY-MM-DD (defensive)
  Object.keys(_DATE_DICT_KEYS).forEach(function(k){if(dict[k])dict[k]=_normalizeDate(dict[k]);});
  return dict;
}
// Per-form override maps — exact field-name → data-key. Used when the
// fuzzy matcher can't read the field name (e.g. MSA-4676's garbled names
// caused by the PDF's custom font encoding).
// Per-form explicit field maps (Adobe's auto-detected field name → our data-dict key). Only the
// MSA-4676 remains — the DHS-4771 / BPHASA-2421 maps were removed as dead code with their retired cards.
var STATE_FORM_FIELD_MAPS={
  // DHS-390 — Adobe's auto-detected names carry the printed item numbers ("9 Date of Birth …").
  // Mapped explicitly so a template revision that renumbers items fails loudly here instead of
  // quietly filling the wrong box. Everything the CLIENT must answer (household composition,
  // benefits, interpreter needs, signature) is deliberately left blank.
  dhs390:{
    'Full Name':'client_name',
    '9 Date of Birth MMDDYYYY':'client_dob',
    '10 MedicaidRecipient ID':'medicaid_id',
    '11 Street Address':'client_address',
    'City':'client_city',
    'State':'client_state',
    'Zip Code':'client_zip',
    '12 Phone Number':'client_phone',
    '14 Email Address':'client_email',
    'Date':'today_date'
    // 'Client Signature X' — left blank on purpose; the client signs it.
  },
  msa4676:{
    // Caseworker / MDHHS office block (top of form)
    'Caseworker Name':'worker_name',
    'Caseworker County':'client_county',
    'Caseworker Phone #':'worker_phone',
    // Client / Beneficiary block
    'Client Name':'client_name',
    'Medicaid #':'medicaid_id',
    'Client Street':'client_address',
    'Client Date of Birth':'client_dob',
    'Client City':'client_city',
    'Client State':'client_state',
    'Client Zip Code':'client_zip',
    'Client Phone # with area code':'client_phone',
    // Section 2 — AGENCY PROVIDER (hardcoded — Liberty Home Care, Thomas Jaboro)
    'Agnecy Provider Name':'agency_provider_name',
    'Agency Provider Name':'agency_provider_name',
    'Provider ID #':'agency_provider_id',
    'Agency Street':'agency_address',
    'Agency City':'agency_city',
    'Agency State':'agency_state',
    'Agency Zip':'agency_zip',
    'Phone #':'agency_phone',
    'RelationShip':'agency_relationship',
    // Manual-fill: per-case start date
    'Start Date':''
  }
};
// Match Adobe's auto-detected field names (which use the form's printed labels)
// to a key in our data dict. Returns the matching value, or '' if no match.
function _matchAcroFormField(fieldName,dict,formType){
  // 1. Per-form explicit map wins
  var explicit=(STATE_FORM_FIELD_MAPS[formType]||{})[fieldName];
  if(explicit&&dict[explicit])return dict[explicit];
  var n=(fieldName||'').toLowerCase().replace(/[^a-z0-9]+/g,' ').trim();
  // direct exact match by our convention
  var exactKey=fieldName.toLowerCase().replace(/[^a-z0-9_]/g,'');
  if(dict.hasOwnProperty(exactKey))return dict[exactKey];
  // Fuzzy match by keywords. ORDER MATTERS — more specific rules come first
  // so "Email Address" doesn't get caught by the generic /address/ rule.
  var rules=[
    // Email FIRST (before address) so "Email Address" doesn't trip the address rule
    [/email/, 'client_email'],
    // Phone variants
    [/(phone number|telephone|^phone$|phone\s*#|cell)/, 'client_phone'],
    [/tty/, 'client_tty'],
    // ID-like
    [/case name/, 'client_name'],
    [/log number/, 'log_number'],
    [/recipient id/, 'recipient_id'],
    [/medicaid/, 'medicaid_id'],
    [/(date of birth|dob|birth date)/, 'client_dob'],
    // Names
    [/patient.*name/, 'client_name'],
    [/full name/, 'client_name'],
    [/client.*name/, 'client_name'],
    [/beneficiary.*name/, 'client_name'],
    // Caregiver / provider fields BEFORE generic address (so "Caregiver Address" doesn't grab client)
    [/(caregiver|provider).*first/, 'caregiver_first_name'],
    [/(caregiver|provider).*last/, 'caregiver_last_name'],
    [/champs/, 'caregiver_champs_id'],
    // Worker
    [/worker name/, 'worker_name'],
    [/worker phone/, 'worker_phone'],
    [/worker email/, 'worker_email'],
    [/(return fax|fax number|fax$)/, 'worker_fax'],
    // Address (after email/name catches)
    [/(street|^address)/, 'client_address'],
    [/^city$/, 'client_city'],
    [/^state$/, 'client_state'],
    [/(zip|postal)/, 'client_zip'],
    [/county/, 'client_county'],
    // Caregiver name (broad — last because it catches anything with "caregiver/provider")
    [/(caregiver|provider).*name/, 'caregiver_full_name'],
    // Generic dates
    [/sig.*date|date.*sig/, 'signature_date'],
    [/^date$/, 'today_date'],
    [/today/, 'today_date']
  ];
  for(var i=0;i<rules.length;i++){
    if(rules[i][0].test(n)){var k=rules[i][1];if(dict[k])return dict[k];}
  }
  return '';
}

// Shared render path. Used by live preview (flatten:false) and exports (flatten:true).
async function _renderStateFormBytes(opts){
  opts=opts||{};
  var def=STATE_FORM_OVERLAYS[activeFormType];
  if(!def)throw new Error('Form not configured: '+activeFormType);
  if(!window.PDFLib)throw new Error('PDF library still loading');
  var resp=await fetch(def.file);
  if(!resp.ok)throw new Error('Cannot load template ('+resp.status+')');
  var bytes=new Uint8Array(await resp.arrayBuffer());
  var pdfDoc=await PDFLib.PDFDocument.load(bytes);
  var helv=await pdfDoc.embedFont(PDFLib.StandardFonts.Helvetica);
  var pages=pdfDoc.getPages();
  var acroForm=pdfDoc.getForm();
  var acroFields=acroForm.getFields();
  var usedAcroForm=false;
  _sfAcroForms[activeFormType]=acroFields.length>0;   // drives the "Download fillable" button
  if(acroFields.length>0){
    usedAcroForm=true;
    var dict=_buildFormDataDict();
    acroFields.forEach(function(field){
      var fname=field.getName();
      try{
        if(typeof field.setText==='function'){
          var v=_matchAcroFormField(fname,dict,activeFormType);
          if(v)field.setText(String(v));
          // Force black text on every text field
          try{
            var da='0 0 0 rg /Helv 10 Tf';
            field.acroField.dict.set(PDFLib.PDFName.of('DA'),PDFLib.PDFString.of(da));
          }catch(e){}
        } else if(typeof field.check==='function'){
          // Checkboxes: only the per-form allowlist above is ticked; everything else stays as-is.
          if((STATE_FORM_CHECKS[activeFormType]||{})[fname]===true){try{field.check();}catch(e){}}
        } else if(typeof field.select==='function'){
          var v2=_matchAcroFormField(fname,dict,activeFormType);
          if(v2){try{field.select(String(v2));}catch(e){}}
        }
      }catch(e){console.warn('[StateForm] field "'+fname+'" set failed:',e.message);}
    });
    try{acroForm.updateFieldAppearances(helv);}
    catch(e){
      try{acroForm.acroForm.dict.set(PDFLib.PDFName.of('NeedAppearances'),PDFLib.PDFBool.True);}catch(e2){}
    }
    // Flatten on export — bakes field values into the page so recipient can't edit
    if(opts.flatten){try{acroForm.flatten();}catch(e){console.warn('[StateForm] flatten failed:',e.message);}}
  }
  // Coord-stamping fallback for flat PDFs
  if(!usedAcroForm)(def.fields||[]).forEach(function(f){
    var el=document.getElementById(f.inputId);if(!el)return;
    var v=(el.type==='checkbox')?(el.checked?(f.checkedText||'X'):''):((el.value||'').trim());
    if(!v)return;
    var page=pages[f.page||0];if(!page)return;
    if(f.coverRect){
      page.drawRectangle({x:f.coverRect.x,y:f.coverRect.y,width:f.coverRect.w,height:f.coverRect.h,color:PDFLib.rgb(1,1,1),borderColor:PDFLib.rgb(1,1,1),borderWidth:0});
    }
    var size=f.size||10;
    if(f.maxWidth){while(size>6&&helv.widthOfTextAtSize(v,size)>f.maxWidth)size-=0.5;}
    page.drawText(v,{x:f.x,y:f.y,size:size,font:helv,color:PDFLib.rgb(0,0,0)});
  });
  return await pdfDoc.save();
}

async function renderSfPreview(){
  try{
    var def=STATE_FORM_OVERLAYS[activeFormType];
    var iframe=document.getElementById('sfPreviewFrame');
    if(!def||!iframe)return;
    if(!window.PDFLib){setTimeout(renderSfPreview,400);return;}
    // Live preview keeps fields editable so user can also tweak in the iframe
    var out=await _renderStateFormBytes({flatten:false});
    _sfPreviewBytes=out;
    var blob=new Blob([out],{type:'application/pdf'});
    if(_sfPreviewBlobUrl)URL.revokeObjectURL(_sfPreviewBlobUrl);
    _sfPreviewBlobUrl=URL.createObjectURL(blob);
    iframe.src=_sfPreviewBlobUrl+'#toolbar=1&navpanes=0&view=FitH';
    var s=document.querySelector('.sf-preview-status');var t=document.getElementById('sfPreviewStatusText');
    if(s)s.classList.add('idle');if(t)t.textContent='Preview up to date';
  }catch(e){
    console.error('[StateForm] preview render failed',e);
    var t2=document.getElementById('sfPreviewStatusText');
    if(t2)t2.textContent='Preview error: '+(e.message||e);
  }
}

// keepFillable:true leaves the PDF's form fields editable, so a form the CLIENT still has to
// complete (the DHS-390's household/benefits/signature sections) can be finished on screen instead
// of by hand. The default export stays flattened — a signed agreement shouldn't be editable.
async function downloadStateFormPdf(keepFillable){
  try{
    showToast(keepFillable?'Generating fillable PDF…':'Generating final (flattened) PDF…',2000);
    var bytes=await _renderStateFormBytes({flatten:!keepFillable});
    var def=STATE_FORM_OVERLAYS[activeFormType]||{};
    var clientTag=(activeFormClientName||'client').replace(/[^a-z0-9]/gi,'_');
    var fname=(def.title||activeFormType)+'_'+clientTag+(keepFillable?'_fillable':'')+'_'+today().replace(/\//g,'-')+'.pdf';
    var blob=new Blob([bytes],{type:'application/pdf'});
    var url=URL.createObjectURL(blob);
    var a=document.createElement('a');a.href=url;a.download=fname;a.style.display='none';document.body.appendChild(a);a.click();
    setTimeout(function(){document.body.removeChild(a);URL.revokeObjectURL(url);},2000);
    showToast('✓ Downloaded '+fname,4000);
  }catch(e){
    console.error('[StateForm] download failed',e);
    showAlert('Could not generate the PDF: '+(e.message||e));
  }
}

/* ──────────────────────────────────────────────────────────────────
 *  STATE-FORM PDF OVERLAY
 *  Loads the official state PDF (kept in /forms/) and overlays typed
 *  text from the on-screen inputs at calibrated coordinates. Output
 *  is byte-for-byte the state's layout with our data on top — no
 *  HTML approximation, no risk of margin/font rejection.
 *
 *  Coordinate space: pdf-lib uses bottom-left origin in PDF points
 *  (1 pt = 1/72 inch). Letter page = 612 × 792 pts.
 *
 *  STATE_FORM_OVERLAYS[type] = {
 *    file:'/forms/X.pdf',
 *    fields: [{ inputId:'msa_cname', page:0, x:.., y:.., size:10, maxWidth?:.. }, ...],
 *    signature?: { inputId:'msa_sig_present', page, x, y, w, h }  // optional, draws stored sig PNG
 *  }
 * ────────────────────────────────────────────────────────────────── */
// State-form PDF overlays — calibrated coordinates for stamping data onto the official template.
// The MSA-4676 uses coordinate stamping (its template has no usable form fields). The DHS-390 is a
// real AcroForm — 38 named fields — so it needs no `fields` coordinates at all: _renderStateFormBytes
// detects the form fields and fills them by name (see STATE_FORM_FIELD_MAPS/CHECKS below).
var STATE_FORM_OVERLAYS={
  dhs390:{
    file:'/forms/DHS-390.pdf',
    title:'DHS-390_Application_for_Home_Help',
    fields:[],            // AcroForm template — filled by field name, not by coordinates
    signature:null        // the CLIENT signs the DHS-390; never auto-sign it
  },
  msa4676:{
    file:'/forms/MSA-4676.pdf',
    title:'MSA-4676_Home_Help_Agreement',
    fields:[
      {inputId:'msa_cname',       page:0, x:120, y:660, size:10, maxWidth:240},
      {inputId:'msa_mid',         page:0, x:380, y:660, size:10, maxWidth:140},
      {inputId:'msa_addr',        page:0, x:60,  y:625, size:10, maxWidth:480},
      {inputId:'msa_city',        page:0, x:60,  y:600, size:10, maxWidth:160},
      {inputId:'msa_st',          page:0, x:240, y:600, size:10, maxWidth:30},
      {inputId:'msa_zip',         page:0, x:280, y:600, size:10, maxWidth:80},
      {inputId:'msa_phone',       page:0, x:380, y:600, size:10, maxWidth:140},
      {inputId:'msa_cgname',      page:0, x:60,  y:560, size:10, maxWidth:240},
      {inputId:'msa_cgphone',     page:0, x:320, y:560, size:10, maxWidth:160},
      {inputId:'msa_start',       page:0, x:60,  y:530, size:10, maxWidth:140},
      {inputId:'msa_county',      page:0, x:240, y:530, size:10, maxWidth:140},
      {inputId:'msa_date',        page:0, x:430, y:160, size:10, maxWidth:120}
    ],
    // No auto-signature: beneficiary + caregiver sign MSA-4676
    signature:null
  }
};

// Build a state-form PDF for ANY client WITHOUT the Forms UI: compute the field values from CRM data
// (the same dict the Forms tab fills from) and overlay them onto the official template. Returns the
// PDF bytes (Uint8Array). This is what lets the Assistant generate + attach a form (e.g. the
// MSA-4676) on command, for a client the user names, without opening the Forms page.
async function buildStateFormBytes(formType, clientName){
  if(!window.PDFLib) throw new Error('PDF library still loading — try again in a moment.');
  var def=STATE_FORM_OVERLAYS[formType];
  if(!def) throw new Error('Form not configured: '+formType);
  var inputMap=STATE_FORM_INPUT_MAPS[formType]||{};
  var dict=_buildFormDataDict(clientName);
  var resp=await fetch(def.file);
  if(!resp.ok) throw new Error('HTTP '+resp.status+' fetching '+def.file);
  var bytes=new Uint8Array(await resp.arrayBuffer());
  var pdfDoc=await PDFLib.PDFDocument.load(bytes);
  var helv=await pdfDoc.embedFont(PDFLib.StandardFonts.Helvetica);
  var pages=pdfDoc.getPages();
  (def.fields||[]).forEach(function(f){
    var v=dict[inputMap[f.inputId]]; v=(v==null?'':String(v)).trim();
    if(!v)return;
    var page=pages[f.page||0]; if(!page)return;
    var size=f.size||10;
    if(f.maxWidth){ while(size>6&&helv.widthOfTextAtSize(v,size)>f.maxWidth)size-=0.5; }
    page.drawText(v,{x:f.x,y:f.y,size:size,font:helv,color:PDFLib.rgb(0,0,0)});
  });
  if(def.signature){
    var sigs=(typeof getSigs==='function')?getSigs():[];
    if(sigs.length){
      var sigBytes=_dataUrlToUint8(sigs[0].data);
      if(sigBytes){
        var img=sigs[0].data.indexOf('image/png')>=0?await pdfDoc.embedPng(sigBytes):await pdfDoc.embedJpg(sigBytes);
        var p=pages[def.signature.page||0];
        if(p)p.drawImage(img,{x:def.signature.x,y:def.signature.y,width:def.signature.w,height:def.signature.h});
      }
    }
  }
  return await pdfDoc.save();
}
async function generateStateFormPdf(){
  if(!window.PDFLib){showAlert('PDF library still loading. Please try again in a moment.');return;}
  var def=STATE_FORM_OVERLAYS[activeFormType];
  if(!def){
    showAlert('Official PDF overlay for this form is not configured yet. Use "Print HTML preview" for now — overlay coming soon.');
    return;
  }
  try{
    showToast('Loading '+def.title+'…',2000);
    var resp=await fetch(def.file);
    if(!resp.ok)throw new Error('HTTP '+resp.status+' fetching '+def.file);
    var bytes=new Uint8Array(await resp.arrayBuffer());
    var pdfDoc=await PDFLib.PDFDocument.load(bytes);
    var helv=await pdfDoc.embedFont(PDFLib.StandardFonts.Helvetica);
    var pages=pdfDoc.getPages();
    (def.fields||[]).forEach(function(f){
      var el=document.getElementById(f.inputId);
      if(!el)return;
      var v=(el.value||'').trim();
      if(!v)return;
      var page=pages[f.page||0];if(!page)return;
      // Auto-shrink size if value would overflow maxWidth
      var size=f.size||10;
      if(f.maxWidth){
        while(size>6&&helv.widthOfTextAtSize(v,size)>f.maxWidth)size-=0.5;
      }
      page.drawText(v,{x:f.x,y:f.y,size:size,font:helv,color:PDFLib.rgb(0,0,0)});
    });
    // Signature placement (if a sig is saved and the user opted to include)
    if(def.signature){
      var sigOptIn=true;
      if(def.signature.checkboxId){
        var cb=document.getElementById(def.signature.checkboxId);
        sigOptIn=cb?cb.checked:true;
      }
      if(sigOptIn){
        var sigs=(typeof getSigs==='function')?getSigs():[];
        if(sigs.length){
          var sigBytes=_dataUrlToUint8(sigs[0].data);
          if(sigBytes){
            var img=sigs[0].data.indexOf('image/png')>=0?
              await pdfDoc.embedPng(sigBytes):
              await pdfDoc.embedJpg(sigBytes);
            var p=pages[def.signature.page||0];
            p.drawImage(img,{x:def.signature.x,y:def.signature.y,width:def.signature.w,height:def.signature.h});
          }
        }
      }
    }
    var out=await pdfDoc.save();
    var blob=new Blob([out],{type:'application/pdf'});
    var url=URL.createObjectURL(blob);
    var clientTag=(activeFormClientName||'client').replace(/[^a-z0-9]/gi,'_');
    var fname=def.title+'_'+clientTag+'_'+today().replace(/\//g,'-')+'.pdf';
    // Open in a new tab so user can print or save
    var w=window.open(url,'_blank');
    var a=document.createElement('a');a.href=url;a.download=fname;a.style.display='none';document.body.appendChild(a);a.click();setTimeout(function(){document.body.removeChild(a);URL.revokeObjectURL(url);},2000);
    showToast('✓ Generated '+fname,5000);
  }catch(e){
    console.error('[StateForm] Generate failed:',e);
    showAlert('Could not generate the PDF: '+(e.message||e));
  }
}
function _dataUrlToUint8(dataUrl){
  if(!dataUrl||dataUrl.indexOf(',')<0)return null;
  var b64=dataUrl.split(',')[1];
  try{var bin=atob(b64);var arr=new Uint8Array(bin.length);for(var i=0;i<bin.length;i++)arr[i]=bin.charCodeAt(i);return arr;}catch(e){return null;}
}

// ============================================================
//  SESSION 3 — MONTHLY INVOICE EMAILS
// ============================================================
function openMonthlyInvModal(){
  if(!isInvoiceAdmin()){showAlert('Only the account owner can send monthly emails.');return;}
  var modal=document.getElementById('monthlyInvModal');if(!modal)return;
  modal.classList.add('open');
  // Force-refresh caseworkers + clients from DB so "No email on file" doesn't show against
  // a stale LS cache. Force=true bypasses the 30-sec SWR throttle since this is user-initiated.
  if(typeof revalidate==='function')revalidate(true);
  // Default to previous month. C2: pin the day to 1 FIRST — otherwise on the 31st,
  // setMonth overflows (e.g. Jul 31 → "Jun 31" → Jul 1) and defaults to the wrong month.
  var d=new Date();d.setDate(1);d.setMonth(d.getMonth()-1);
  var m=String(d.getMonth()+1).padStart(2,'0'),y=d.getFullYear();
  var inp=document.getElementById('monthlyInvPeriod');
  if(inp&&!inp.value)inp.value=m+'/'+y;
  document.getElementById('monthlyInvResults').innerHTML='<div style="color:#5c7590;font-size:13px;text-align:center;padding:24px 0;">Enter a billing period and click <b>Preview</b> to see caseworker email groups.</div>';
}
function closeMonthlyInvModal(){
  var modal=document.getElementById('monthlyInvModal');if(modal)modal.classList.remove('open');
}
// Generate the vector PDF for one client's invoice and open in new tab
async function previewClientInvoice(clientName,period){
  var prof=getProfiles()[clientName];
  if(!prof){showAlert('Client not found.');return;}
  var inv=(prof.invoices||[]).find(function(i){return i.billingPeriod===period;});
  if(!inv){showAlert('No invoice found for '+clientName+' in '+period+'.');return;}
  // Stage current state so we can restore (loadInvoiceForCapture mutates the form)
  var savedActive=document.querySelector('.page.active')&&document.querySelector('.page.active').id;
  var invPage=document.getElementById('page-invoice');
  // Briefly activate invoice page off-screen to populate data
  invPage.style.position='fixed';invPage.style.left='-9999px';invPage.style.top='0';invPage.style.zIndex='-1';
  invPage.classList.add('active');
  try{
    await loadInvoiceForCapture(clientName,inv,period);
    var base64=await captureInvoicePDF();
    var bin=atob(base64);
    var bytes=new Uint8Array(bin.length);
    for(var i=0;i<bin.length;i++)bytes[i]=bin.charCodeAt(i);
    var blob=new Blob([bytes],{type:'application/pdf'});
    var url=URL.createObjectURL(blob);
    window.open(url,'_blank');
    // Cleanup blob after a delay so the new tab can render
    setTimeout(function(){URL.revokeObjectURL(url);},10000);
  }catch(e){showAlert('Preview failed: '+e.message);console.error(e);}
  finally{
    invPage.classList.remove('active');
    invPage.style.position='';invPage.style.left='';invPage.style.top='';invPage.style.zIndex='';
    document.querySelectorAll('.page').forEach(function(el){el.classList.remove('active');});
    if(savedActive){var p=document.getElementById(savedActive);if(p)p.classList.add('active');}
  }
}

// Per-invoice validation — returns array of issue strings (empty = ready)
function validateInvoiceForSend(name,prof,inv,cwRec){
  var issues=[];
  if(!inv)return ['No invoice for this period — fill out the invoice first'];
  // Submitted/paid invoices are already done — no point flagging issues on them
  if(inv.status==='submitted'||inv.status==='paid')return [];
  var d=inv.data||{};
  if(!prof.medicaidId)issues.push('Missing Medicaid ID');
  if(!cwRec||!cwRec.agency)issues.push('Caseworker has no Agency set (Bill To will be empty)');
  if(!cwRec||!cwRec.email)issues.push('Caseworker has no email');
  if(!prof.worker&&!prof.caseworkerId)issues.push('No caseworker assigned');
  // Total Time HH must be present
  if(!d.svcHH||String(d.svcHH).trim()===''){
    var hasTask=d.tasks&&d.tasks.svc&&d.tasks.svc.some(function(r){return r&&r.some(function(c){return c;});});
    if(!hasTask)issues.push('Total Time empty AND no tasks checked');
    else issues.push('Total Time hours blank');
  }
  // Sanity-check the billed time itself. These fields are free text and were only ever checked for
  // PRESENCE — so a typo ("99", "-8", "abc", 75 minutes) printed straight onto the certified MSA-1904
  // and went to MDHHS as a claim. Nothing between the keystroke and the state caught it.
  var _hm=function(hhKey,mmKey,label){
    var hh=String(d[hhKey]==null?'':d[hhKey]).trim(), mm=String(d[mmKey]==null?'':d[mmKey]).trim();
    if(hh===''&&mm==='')return;
    if(hh!==''&&!/^\d{1,3}$/.test(hh))issues.push(label+' hours is not a plain number ("'+hh+'")');
    if(mm!==''&&!/^\d{1,2}$/.test(mm))issues.push(label+' minutes is not a plain number ("'+mm+'")');
    if(/^\d{1,2}$/.test(mm)&&parseInt(mm,10)>59)issues.push(label+' minutes is '+mm+' — minutes must be 0-59');
    if(/^\d{1,3}$/.test(hh)&&parseInt(hh,10)>744)issues.push(label+' hours is '+hh+' — more than a full month');
  };
  _hm('svcHH','svcMM','Total Time'); _hm('cplxHH','cplxMM','Complex care time'); _hm('grandHH','grandMM','Billing-period total');
  // Cross-check against the authorization: billing more than MDHHS approved is the error that costs the
  // agency a recoupment, and nothing else in the app compares these two numbers.
  try{
    var a=prof&&prof.authorization;
    if(a&&a.hours!=null&&String(a.hours).trim()!==''&&/^\d{1,3}$/.test(String(d.svcHH||'').trim())){
      var billed=parseInt(d.svcHH,10)*60+(parseInt(d.svcMM,10)||0);
      var appr=parseInt(a.hours,10)*60+(parseInt(a.minutes,10)||0);
      if(appr>0&&billed>appr)
        issues.push('Billed time ('+d.svcHH+':'+_padMin(d.svcMM||'00')+') exceeds the authorized '+a.hours+':'+_padMin(a.minutes||0));
    }
  }catch(e){}
  // Signature: at least one must exist locally so the PDF auto-places it
  if(!getSigs().length)issues.push('Missing signature — PDF will export with no signature on it');
  return issues;
}

function previewMonthlyInvoices(){
  var period=(document.getElementById('monthlyInvPeriod').value||'').trim();
  if(!period||period.length<7){showAlert('Enter a billing period in MM/YYYY format.');return;}
  // Show a "loading" hint and force a fresh fetch of caseworkers + clients from DB before rendering
  // so we never render "No email on file" against stale LS. If the fetch takes >200ms we swap the
  // hint for the real preview; if it fails we fall through to whatever LS has (offline tolerance).
  var resultsEl=document.getElementById('monthlyInvResults');
  if(resultsEl)resultsEl.innerHTML='<div style="color:#5c7590;font-size:13px;text-align:center;padding:24px 0;">Refreshing caseworker + client data…</div>';
  var freshFetch=Promise.all([
    (typeof loadCaseworkersAPI==='function'?loadCaseworkersAPI():Promise.resolve()),
    (typeof loadProfilesAPI==='function'?loadProfilesAPI():Promise.resolve()),
  ]).catch(function(e){console.warn('Fresh fetch failed, using LS cache:',e);});
  freshFetch.then(function(){_previewMonthlyInvoicesRender(period);});
}
function _previewMonthlyInvoicesRender(period){
  var profiles=getProfiles();
  var cws=getCaseworkers();
  // Group ALL active clients by worker name
  var groups={};
  Object.keys(profiles).forEach(function(name){
    var prof=profiles[name];
    if(!isInvoiceableStatus(prof))return;   // In Progress / Lost / Terminated are never invoiced
    if(isCarrierClient(prof))return;   // managed-care clients are billed by the carrier — never invoiced here
    // Skip clients whose service started AFTER this billing period
    if(!clientWasActiveInPeriod(prof,period))return;
    // Resolve the caseworker identity. Prefer caseworkerId; fall back to fuzzy name match.
    var rawWorker=(prof.worker||'').trim();
    var cwRec=cws.find(function(c){return c.id&&prof.caseworkerId&&String(c.id)===String(prof.caseworkerId);})||
              cws.find(function(c){return (c.name||'').trim().toLowerCase()===rawWorker.toLowerCase()&&rawWorker;})||
              {};
    // Group key: stable caseworker ID if known, else normalized name. Display name comes from the caseworker record so typo'd `prof.worker` strings still merge.
    var groupKey=cwRec.id?'cw:'+cwRec.id:(rawWorker?'nm:'+rawWorker.toLowerCase():'(No Worker Assigned)');
    var displayName=(cwRec.name||rawWorker||'(No Worker Assigned)').trim();
    if(!groups[groupKey])groups[groupKey]={clients:[],cwName:displayName,cwRec:cwRec,email:cwRec.email||''};
    var inv=(prof.invoices||[]).find(function(i){return i.billingPeriod===period;})||null;
    var issues=validateInvoiceForSend(name,prof,inv,cwRec);
    groups[groupKey].clients.push({name:name,prof:prof,inv:inv,issues:issues});
  });
  var groupKeys=Object.keys(groups).sort(function(a,b){return (groups[a].cwName||'').localeCompare(groups[b].cwName||'');});
  if(!groupKeys.length){
    document.getElementById('monthlyInvResults').innerHTML='<div style="color:#5c7590;font-size:13px;text-align:center;padding:20px;">No active clients found.</div>';
    return;
  }
  // Pre-compute global "send all" eligibility
  var totalReady=0,totalIssues=0,totalSent=0,totalEmpty=0,totalCaseworkersSendable=0;
  groupKeys.forEach(function(gk){
    var g=groups[gk];if(!g.email)return;
    var ready=g.clients.filter(function(c){return c.inv && c.inv.status!=='submitted' && c.inv.status!=='paid' && (!c.issues||c.issues.length===0);}).length;
    var iss=g.clients.filter(function(c){return c.inv && c.inv.status!=='submitted' && c.inv.status!=='paid' && c.issues && c.issues.length;}).length;
    var sent=g.clients.filter(function(c){return c.inv && (c.inv.status==='submitted'||c.inv.status==='paid');}).length;
    var none=g.clients.filter(function(c){return !c.inv;}).length;
    totalReady+=ready;totalIssues+=iss;totalSent+=sent;totalEmpty+=none;
    if(ready>0)totalCaseworkersSendable++;
  });
  // Count eligible auto-gen clients (active, missing invoice for period, have prior invoice)
  var eligibleAutoGen=findClientsEligibleForAutoGen(period).length;
  var html='<div style="font-size:12px;color:#4d6c88;margin-bottom:14px;display:flex;align-items:center;gap:14px;flex-wrap:wrap;">'+
    '<span><b>'+groupKeys.length+'</b> caseworker group'+(groupKeys.length!==1?'s':'')+' for billing period <b>'+esc(period)+'</b></span>'+
    '<span style="color:#1e7e34;">✓ '+totalReady+' ready</span>'+
    (totalIssues?'<span style="color:#c07000;">⚠ '+totalIssues+' issues</span>':'')+
    (totalSent?'<span style="color:#1565a0;">✓ '+totalSent+' already sent</span>':'')+
    (totalEmpty?'<span style="color:#888;">— '+totalEmpty+' missing</span>':'')+
    '<div style="margin-left:auto;display:flex;gap:8px;">'+
      (eligibleAutoGen>0?'<button class="btn btn-secondary btn-sm" onclick="autoGenerateMonthlyInvoices(\''+escJsAttr(period)+'\')" style="white-space:nowrap;" title="Copy each missing client\'s last invoice into '+esc(period)+' with day-shifted patterns">🔄 Auto-Generate '+eligibleAutoGen+'</button>':'')+
      (totalCaseworkersSendable>1?'<button class="btn btn-primary btn-sm" id="sendAllCwBtn" onclick="sendAllCaseworkerEmails(\''+escJsAttr(period)+'\')" style="white-space:nowrap;">Send All ('+totalCaseworkersSendable+' caseworkers)</button>':'')+
    '</div>'+
  '</div>';
  groupKeys.forEach(function(gk){
    var g=groups[gk];
    var wname=g.cwName||gk;
    var email=g.email||'';
    var hasInv=g.clients.filter(function(c){return c.inv;}).length;
    var missingInv=g.clients.length-hasInv;
    html+='<div class="cw-email-card">';
    html+='<div class="cw-email-hdr">';
    html+='<div style="width:34px;height:34px;border-radius:50%;background:#e8f0f9;color:#185FA5;font-size:12px;font-weight:700;display:flex;align-items:center;justify-content:center;flex-shrink:0;">'+
      esc((wname.split(' ').map(function(w){return w[0]||'';}).slice(0,2).join('').toUpperCase()||'?'))+
    '</div>';
    html+='<div><div class="cw-email-name">'+esc(wname)+'</div>';
    if(email)html+='<div style="font-size:11px;color:#4d6c88;">'+esc(email)+'</div>';
    html+='</div>';
    html+='<div style="flex:1;"></div>';
    if(missingInv>0)html+='<span style="font-size:11px;color:#c07000;font-weight:600;margin-right:8px;">'+missingInv+' missing invoice'+(missingInv>1?'s':'')+'</span>';
    if(email){
      var clientsJson=JSON.stringify(g.clients.map(function(c){return {name:c.name,medicaidId:c.prof.medicaidId||'',invStatus:c.inv?(c.inv.status||'draft'):'none',hasIssues:(c.issues||[]).length>0,issues:c.issues||[]};}));
      html+='<button class="btn btn-primary btn-sm" style="white-space:nowrap;" onclick=\'sendMonthlyEmail('+esc(JSON.stringify(email))+','+esc(JSON.stringify(wname))+','+esc(clientsJson)+','+esc(JSON.stringify(period))+')\'>Send Email</button>';
    } else {
      html+='<span style="font-size:11px;color:#b03030;font-style:italic;margin-right:8px;">No email on file</span>';
      html+='<button class="btn btn-secondary btn-sm" onclick="closeMonthlyInvModal();navCaseworkers()">Add Email</button>';
    }
    html+='</div>';
    html+='<div class="cw-client-list">';
    g.clients.forEach(function(c){
      var st=c.inv?(c.inv.status||'draft'):'none';
      var stColor=st==='paid'?'#1e7e34':st==='submitted'?'#1565a0':st==='none'?'#b03030':'#888';
      var stLabel=st==='none'?'No invoice':st.charAt(0).toUpperCase()+st.slice(1);
      var hasIssues=c.issues&&c.issues.length>0;
      var issueTitle=hasIssues?c.issues.join('\n• '):'';
      html+='<div class="cw-client-row" style="display:flex;align-items:center;gap:6px;">';
      // Warning indicator if validation issues
      if(hasIssues){
        html+='<span title="• '+esc(issueTitle)+'" style="color:#c07000;font-size:14px;cursor:help;flex-shrink:0;" aria-label="Has issues">⚠</span>';
      } else if(c.inv){
        html+='<span title="Looks complete" style="color:#1e7e34;font-size:13px;flex-shrink:0;">✓</span>';
      } else {
        html+='<span style="display:inline-block;width:14px;flex-shrink:0;"></span>';
      }
      html+='<span style="flex:1;font-weight:500;color:#1a2b45;">'+esc(c.name)+'</span>'+
        (c.prof.medicaidId?'<span style="color:#5c7590;font-size:11px;margin-right:6px;">ID: '+esc(c.prof.medicaidId)+'</span>':'');
      // Preview button (only if invoice exists)
      if(c.inv){
        html+='<button class="btn btn-secondary btn-sm" style="font-size:10px;padding:2px 8px;margin-right:6px;" onclick=\'previewClientInvoice('+esc(JSON.stringify(c.name))+','+esc(JSON.stringify(period))+')\'>Preview</button>';
      }
      html+='<span style="font-size:11px;font-weight:600;color:'+stColor+';min-width:80px;text-align:right;">'+stLabel+'</span>';
      html+='</div>';
      // Issues detail line below row when present
      if(hasIssues){
        html+='<div style="font-size:10px;color:#c07000;padding:2px 22px 4px;line-height:1.4;">'+
          c.issues.map(function(s){return '• '+esc(s);}).join('<br>')+'</div>';
      }
    });
    html+='</div>';
    html+='</div>';
  });
  document.getElementById('monthlyInvResults').innerHTML=html;
}
// ──────────────────────────────────────────────────────────────────
//  AUTO-GENERATE NEXT MONTH INVOICE
//  Copies a previous invoice into a new period, shifting day patterns
//  for sub-daily columns (Laundry, Shopping, etc.) so the new invoice
//  doesn't look like an exact carbon-copy. Hospital column always
//  starts empty since it's by-exception, not recurring.
// ──────────────────────────────────────────────────────────────────
// Monthly Emails period input — accept shorthand (0526, 052026, 5/26, etc.)
function onMonthlyPeriodInput(el){
  // Allow auto-format as user types digits
  var raw=el.value.replace(/\D/g,'');
  if(raw.length<=2){el.value=raw;return;}
  var formatted=raw.slice(0,2)+'/'+raw.slice(2,6);
  el.value=formatted;
}
function onMonthlyPeriodBlur(el){
  var v=el.value.trim();
  // Already in correct MM/YYYY form
  if(/^\d{2}\/\d{4}$/.test(v))return;
  var raw=v.replace(/\D/g,'');
  var mm,yyyy;
  if(raw.length===4){
    // MMYY → MM/YYYY (assume 20YY for YY<50, else 19YY)
    mm=raw.slice(0,2);
    var yy=raw.slice(2);
    yyyy=parseInt(yy)<50?'20'+yy:'19'+yy;
  } else if(raw.length===6){
    // MMYYYY
    mm=raw.slice(0,2);yyyy=raw.slice(2,6);
  } else if(raw.length===5){
    // MYYYY (single-digit month) — not strictly valid but try
    mm='0'+raw.slice(0,1);yyyy=raw.slice(1,5);
  } else if(raw.length===3){
    // MYY (single-digit month, 2-digit year)
    mm='0'+raw.slice(0,1);
    var yy3=raw.slice(1);
    yyyy=parseInt(yy3)<50?'20'+yy3:'19'+yy3;
  }
  if(mm&&yyyy){el.value=mm+'/'+yyyy;}
}

function defaultGenInvPeriod(){
  // Default to the previous month — user typically generates last month's invoices
  var d=new Date();d.setDate(1);d.setMonth(d.getMonth()-1);
  return String(d.getMonth()+1).padStart(2,'0')+'/'+d.getFullYear();
}
// Small modal that lists exactly which active clients are missing an invoice for a
// specific billing period. Clicking a name jumps to the client's detail; the "+ Create
// Invoice" button jumps straight to the new-invoice screen for that client.
function showMissingInvoicesModal(period){
  var existing=document.getElementById('missingInvModal');if(existing)existing.remove();
  var profiles=getProfiles();
  var missing=Object.keys(profiles).filter(function(name){
    var st=profiles[name].clientStatus||'active';
    if(st!=='active')return false;
    if(!clientDueForInvoice(profiles[name],period))return false; // #10: respect the service start date
    return !((profiles[name].invoices)||[]).some(function(i){return i.billingPeriod===period;});
  }).sort(function(a,b){return a.localeCompare(b);});
  var ov=document.createElement('div');
  ov.id='missingInvModal';ov.className='modal-overlay open';
  var rows=missing.map(function(name){
    var prof=profiles[name];
    return '<div style="display:flex;align-items:center;justify-content:space-between;gap:12px;padding:8px 0;border-bottom:1px solid var(--border);">'+
      '<a href="'+buildClientUrl(name)+'" class="link-plain" style="font-size:13px;font-weight:600;" onclick="closeMissingInvModal()">'+esc(name)+(prof.medicaidId?' <span style="font-weight:400;color:var(--text-muted);font-size:11px;">'+esc(prof.medicaidId)+'</span>':'')+'</a>'+
      '<button class="btn btn-primary btn-sm" style="font-size:11px;padding:4px 10px;white-space:nowrap;" onclick="activeProfileName=\''+escJsAttr(name)+'\';closeMissingInvModal();navInvoice()">+ Invoice</button>'+
    '</div>';
  }).join('');
  ov.innerHTML='<div class="modal-box" style="max-width:520px;">'+
    '<h3>Missing invoices for '+esc(period)+'</h3>'+
    (missing.length?
      '<div style="font-size:11px;color:var(--text-muted);margin:4px 0 12px;">'+missing.length+' active client'+(missing.length>1?'s':'')+'</div>'+
      '<div style="max-height:340px;overflow-y:auto;">'+rows+'</div>':
      '<div style="font-size:13px;color:var(--text-muted);padding:16px 0;">✓ All active clients have an invoice for this period.</div>'
    )+
    '<div class="modal-row" style="display:flex;gap:8px;justify-content:space-between;margin-top:16px;">'+
      (missing.length?'<button class="btn btn-secondary" onclick="closeMissingInvModal();openGenerateInvoicesModal();">Generate for all →</button>':'<span></span>')+
      '<button class="btn btn-secondary" onclick="closeMissingInvModal()">Close</button>'+
    '</div>'+
  '</div>';
  document.body.appendChild(ov);
}
function closeMissingInvModal(){var m=document.getElementById('missingInvModal');if(m)m.remove();}

function openGenerateInvoicesModal(){
  if(!isInvoiceAdmin()){showAlert('Only the account owner can generate invoices.');return;}
  var existing=document.getElementById('genInvModal');if(existing)existing.remove();
  var ov=document.createElement('div');
  ov.id='genInvModal';ov.className='modal-overlay open';
  ov.innerHTML='<div class="modal-box" style="max-width:480px;">'+
    '<h3>Generate Invoices</h3>'+
    '<label style="display:block;font-size:11px;color:#4d6c88;font-weight:600;text-transform:uppercase;letter-spacing:.4px;margin-top:12px;margin-bottom:4px;" for="genInvPeriod">Billing Period</label>'+
    '<input id="genInvPeriod" type="text" placeholder="MM/YYYY" maxlength="7" '+
      'style="padding:8px 10px;border:1px solid #d0d8e4;border-radius:5px;font-size:14px;width:140px;outline:none;" '+
      'oninput="onMonthlyPeriodInput(this)" onblur="onMonthlyPeriodBlur(this);refreshGenInvCount();" '+
      'onkeydown="if(event.key===\'Enter\'){onMonthlyPeriodBlur(this);refreshGenInvCount();doGenerateInvoices();}">'+
    '<div style="font-size:11px;color:#5c7590;margin-top:6px;">e.g. <b>04/2026</b>, <b>0426</b>, or <b>042026</b></div>'+
    '<div id="genInvCount" style="margin-top:14px;font-size:13px;color:#4a5d7a;min-height:18px;"></div>'+
    '<div class="modal-row" style="display:flex;gap:8px;justify-content:flex-end;margin-top:16px;">'+
      '<button class="btn btn-secondary" onclick="closeGenInvModal()">Cancel</button>'+
      '<button id="genInvDoBtn" class="btn btn-primary" onclick="doGenerateInvoices()">Generate</button>'+
    '</div>'+
  '</div>';
  document.body.appendChild(ov);
  var inp=document.getElementById('genInvPeriod');
  inp.value=defaultGenInvPeriod();
  refreshGenInvCount();
  setTimeout(function(){inp.focus();inp.select();},50);
}
function closeGenInvModal(){var m=document.getElementById('genInvModal');if(m)m.remove();}
function refreshGenInvCount(){
  var inp=document.getElementById('genInvPeriod');var lbl=document.getElementById('genInvCount');var btn=document.getElementById('genInvDoBtn');
  if(!inp||!lbl||!btn)return;
  var period=(inp.value||'').trim();
  if(!/^\d{2}\/\d{4}$/.test(period)){lbl.textContent='';btn.disabled=true;btn.textContent='Generate';return;}
  var eligible=findClientsEligibleForAutoGen(period);
  if(!eligible.length){lbl.innerHTML='<span style="color:#a05a00;">No clients need invoices for '+period+' (already generated, or no prior invoice to copy).</span>';btn.disabled=true;btn.textContent='Generate';}
  else{lbl.innerHTML='<b>'+eligible.length+'</b> client'+(eligible.length>1?'s':'')+' will get a new draft invoice for <b>'+period+'</b>.';btn.disabled=false;btn.textContent='Generate '+eligible.length;}
}
function doGenerateInvoices(){
  var inp=document.getElementById('genInvPeriod');if(!inp)return;
  onMonthlyPeriodBlur(inp);
  var period=(inp.value||'').trim();
  if(!/^\d{2}\/\d{4}$/.test(period)){showAlert('Enter a valid billing period (MM/YYYY).');return;}
  var eligible=findClientsEligibleForAutoGen(period);
  if(!eligible.length){refreshGenInvCount();return;}
  closeGenInvModal();
  _doAutoGenerateInvoices(eligible,period);
}
// ── Persistent undo for auto-generated invoices ──────────────────
// Stored in localStorage so it survives reloads. Each batch tracks the
// invoice IDs created + a snapshot of their data (so we can detect manual
// edits). Batches expire after 24 hours.
var AUTOGEN_UNDO_TTL_MS=24*60*60*1000;
function _getAutoGenUndoStack(){
  try{
    var arr=JSON.parse(localStorage.getItem('lhca_autogen_undo')||'[]');
    if(!Array.isArray(arr))return [];
    var now=Date.now();
    arr=arr.filter(function(b){return b&&b.when&&(now-b.when)<AUTOGEN_UNDO_TTL_MS;});
    return arr;
  }catch(e){return [];}
}
function _setAutoGenUndoStack(arr){
  try{localStorage.setItem('lhca_autogen_undo',JSON.stringify(arr||[]));}catch(e){}
}
function _pushAutoGenUndoBatch(batch){
  var arr=_getAutoGenUndoStack();arr.push(batch);_setAutoGenUndoStack(arr);renderUndoBanner();
}
// Per-batch eligibility — must still be pristine draft, otherwise refuse
function _autoGenBatchStatus(batch){
  var profiles=getProfiles();
  var ok=0,sent=0,edited=0,missing=0;
  (batch.invoices||[]).forEach(function(rec){
    var prof=profiles[rec.clientName];if(!prof||!prof.invoices){missing++;return;}
    // Match by billing period (survives the DB round-trip, which drops the client-side
    // 'id' in favour of 'dbId'); fall back to the original id before the first reload.
    var inv=prof.invoices.find(function(i){return (batch.period&&i.billingPeriod===batch.period)||i.id===rec.invoiceId;});
    if(!inv){missing++;return;}
    if(inv.status&&inv.status!=='draft'){sent++;return;}
    // Compare data hash with snapshot — if user edited the invoice, refuse undo
    if(rec.dataHash&&rec.dataHash!==_quickInvoiceHash(inv.data)){edited++;return;}
    ok++;
  });
  return {ok:ok,sent:sent,edited:edited,missing:missing};
}
function _quickInvoiceHash(data){
  if(!data)return '';
  // Lightweight hash of fields most likely to change with edits
  var s=[data.svcHH||'',data.svcMM||'',data.cplxHH||'',data.cplxMM||'',
         JSON.stringify((data.tasks&&data.tasks.svc)||[]),
         JSON.stringify((data.tasks&&data.tasks.cplx)||[])].join('|');
  // Fast 32-bit hash
  var h=0;for(var i=0;i<s.length;i++){h=((h<<5)-h+s.charCodeAt(i))|0;}
  return String(h);
}
function undoAutoGenBatch(batchId){
  var arr=_getAutoGenUndoStack();
  var idx=arr.findIndex(function(b){return b.id===batchId;});
  if(idx<0){showAlert('That undo entry has expired or already been used.');renderUndoBanner();return;}
  var batch=arr[idx];
  var status=_autoGenBatchStatus(batch);
  if(status.ok===0){
    showAlert('Cannot undo — none of the invoices are still in their original draft state. They may have been sent, edited, or deleted.');
    return;
  }
  var warnExtras='';
  if(status.sent)warnExtras+='\n• '+status.sent+' already sent (will skip)';
  if(status.edited)warnExtras+='\n• '+status.edited+' edited since generation (will skip)';
  if(status.missing)warnExtras+='\n• '+status.missing+' already removed (will skip)';
  showConfirm(
    'Remove '+status.ok+' auto-generated invoice'+(status.ok!==1?'s':'')+' for '+batch.period+'?'+
    (warnExtras?'\n\nNote — the following will not be undone:'+warnExtras:''),
    function(){
      // Collect the invoices still eligible to remove (pristine draft, unedited, present).
      var profiles=getProfiles();
      var targets=[]; // {clientName, dbId, billingPeriod, invoiceId}
      batch.invoices.forEach(function(rec){
        var prof=profiles[rec.clientName];if(!prof||!prof.invoices)return;
        var inv=prof.invoices.find(function(iv){return (batch.period&&iv.billingPeriod===batch.period)||iv.id===rec.invoiceId;});
        if(!inv)return;
        if(inv.status&&inv.status!=='draft')return;
        if(rec.dataHash&&rec.dataHash!==_quickInvoiceHash(inv.data))return;
        targets.push({clientName:rec.clientName,dbId:inv.dbId,billingPeriod:inv.billingPeriod,invoiceId:inv.id});
      });
      if(!targets.length){renderUndoBanner();return;}
      // DELETE each server-side (via the hardened deleteInvoiceAPI: server-first, remove-on-
      // confirmed-2xx, retry toast on failure). The old code spliced locally + saveProfileSP,
      // but syncNewInvoices only inserts/updates — it never deletes, so removed invoices
      // resurrected on the next reload. finalize() runs once every targeted delete confirms
      // (immediately for local-only rows, or later after a retry); a failed delete leaves its
      // row + a Retry toast and keeps the batch on the stack so it can be re-undone.
      var removed=0,settled=0;
      var finalize=function(){
        if(settled<targets.length)return;
        var arr2=_getAutoGenUndoStack().filter(function(b){return b.id!==batchId;});
        _setAutoGenUndoStack(arr2);
        logActivity('invoice','Undid auto-generation: removed '+removed+' invoice'+(removed!==1?'s':'')+' for '+batch.period);
        showAlert('✓ Removed '+removed+' auto-generated invoice'+(removed!==1?'s':'')+' for '+batch.period+'.',{title:'Undo Complete'});
        if(typeof previewMonthlyInvoices==='function')previewMonthlyInvoices();
        updateStats();renderUndoBanner();
      };
      targets.forEach(function(t){
        deleteInvoiceAPI(t.dbId,t.clientName,t.billingPeriod,function(){
          var p2=getProfiles(),prof=p2[t.clientName];
          if(prof&&prof.invoices){
            var j=t.dbId
              ? prof.invoices.findIndex(function(iv){return iv.dbId===t.dbId;})
              : prof.invoices.findIndex(function(iv){return !iv.dbId && iv.id===t.invoiceId;});
            if(j>=0){prof.invoices.splice(j,1);saveProfilesLS(p2);}
          }
          removed++;settled++;renderUndoBanner();finalize();
        });
      });
    },
    {title:'Undo Auto-Generation',okText:'Remove '+status.ok,danger:true}
  );
}
function dismissAutoGenUndoBatch(batchId){
  var arr=_getAutoGenUndoStack().filter(function(b){return b.id!==batchId;});
  _setAutoGenUndoStack(arr);renderUndoBanner();
}
function renderUndoBanner(){
  var host=document.getElementById('autogenUndoBanner');if(!host)return;
  var arr=_getAutoGenUndoStack();
  if(!arr.length){host.innerHTML='';host.style.display='none';return;}
  host.style.display='';
  host.innerHTML=arr.slice().reverse().map(function(b){
    var status=_autoGenBatchStatus(b);
    var minsAgo=Math.max(0,Math.floor((Date.now()-b.when)/60000));
    var ageLabel=minsAgo<60?(minsAgo+'m ago'):(Math.floor(minsAgo/60)+'h '+(minsAgo%60)+'m ago');
    var canUndo=status.ok>0;
    var tooltip=canUndo
      ?(status.ok+' still pristine'+(status.sent||status.edited?' — '+(status.sent?status.sent+' sent, ':'')+(status.edited?status.edited+' edited':'')+' will be skipped':''))
      :'Cannot undo — '+(status.sent?status.sent+' already sent':'')+(status.edited?(status.sent?', ':'')+status.edited+' edited':'');
    return '<div style="background:#fff8e6;border:1px solid #f0d18a;border-radius:6px;padding:10px 14px;margin-bottom:10px;display:flex;align-items:center;gap:10px;">'+
      '<span style="font-size:18px;">↩</span>'+
      '<div style="flex:1;font-size:13px;color:#5a4a1a;">'+
        '<b>'+(b.invoices||[]).length+' invoice'+((b.invoices||[]).length!==1?'s':'')+'</b> auto-generated for <b>'+esc(b.period)+'</b> · '+ageLabel+
        (canUndo?'':'<div style="font-size:11px;color:#8a6f1a;margin-top:2px;">'+esc(tooltip)+'</div>')+
      '</div>'+
      (canUndo?'<button class="btn btn-secondary btn-sm" title="'+esc(tooltip)+'" onclick="undoAutoGenBatch(\''+escJsAttr(b.id)+'\')">Undo</button>':'')+
      '<button class="btn btn-secondary btn-sm" onclick="dismissAutoGenUndoBatch(\''+escJsAttr(b.id)+'\')" title="Dismiss this undo entry">&times;</button>'+
    '</div>';
  }).join('');
}

// ── First invoice from a DHS-1210 authorization ────────────────────────────────
// The invoice's 15 service columns (order matters — it's the grid column order).
function _dhsSvcColNames(){
  return ['Bathing','Dressing','Eating','Grooming','Mobility','Toileting','Transferring','Housework',
    'Laundry','Travel Time for Laundry','Medication','Meal Preparation','Shopping','Travel Time for Shopping',
    'Hospital/Nursing Facility Stay'];
}
// Map a DHS-1210 task name to a service-column index (-1 if none). Exact match first, then the
// known DHS aliases, then a loose contains — so a task is never put in the WRONG column silently.
function _dhsMapTaskToCol(taskName){
  var cols=_dhsSvcColNames(); var t=String(taskName||'').toLowerCase().trim(); if(!t)return -1;
  for(var i=0;i<cols.length;i++){ if(cols[i].toLowerCase()===t)return i; }
  if(/travel.*laundry/.test(t))return cols.indexOf('Travel Time for Laundry');
  if(/travel.*shop/.test(t))return cols.indexOf('Travel Time for Shopping');
  if(/^shopping|shopping for (food|meds)/.test(t))return cols.indexOf('Shopping');
  for(var j=0;j<cols.length;j++){ var c=cols[j].toLowerCase(); if(c===t)continue; if(c.indexOf(t)>=0||t.indexOf(c)>=0)return j; }
  return -1;
}
// Turn a frequency phrase into the day indices to check across `days`. The form gives frequency,
// not calendar days — this is a reviewable starting pattern (the provider adjusts), not a claim.
function _dhsFreqToDays(freq, days){
  var f=String(freq||'').toLowerCase(), out=[], i;
  if(/7 days? per week|daily|every day/.test(f)){ for(i=0;i<days;i++)out.push(i); return out; }
  var wk=f.match(/(\d+)\s*days?\s*per\s*week/);
  if(wk){ var n=Math.min(7,parseInt(wk[1])||1); for(i=0;i<days;i++){ if((i%7)<n)out.push(i); } return out; }
  // Per-month frequency — numeric ("2 days/times per month") OR spelled out ("Twice per month",
  // "three times per month"), since MDHHS forms use the word forms. Spread the days evenly; the
  // count is what matters (provider adjusts exact days).
  var perMonth=null, mo=f.match(/(\d+)\s*(?:days?|times?)\s*(?:per|a)\s*month/);
  if(mo){ perMonth=parseInt(mo[1])||1; }
  else if(/\bonce\b[^.]*month|1\s*day\s*per\s*month|\bmonthly\b/.test(f)){ perMonth=1; }
  else if(/\btwice\b[^.]*month/.test(f)){ perMonth=2; }
  else if(/(?:three\s*times|\bthrice\b)[^.]*month/.test(f)){ perMonth=3; }
  else if(/four\s*times[^.]*month/.test(f)){ perMonth=4; }
  if(perMonth!=null){
    var m=Math.max(1,perMonth), step=Math.max(1,Math.floor(days/m));
    for(i=0;i<m&&i*step<days;i++)out.push(i*step);
    return out;
  }
  return [0]; // unknown frequency → mark the 1st day only, provider adjusts
}
// "HH:MM" → minutes (e.g. "02:00" → 120). 0 if unparseable.
function _dhsHmToMin(s){ var m=String(s||'').match(/(\d+):(\d+)/); return m?(parseInt(m[1],10)*60+parseInt(m[2],10)):0; }
// A stable per-month seed from "MM/YYYY" — a distinct integer per month so each month's generated
// pattern differs from the last, but regenerating the SAME month reproduces it (no randomness).
function _dhsPeriodSeed(period){ var p=String(period||'').split('/'); var mm=parseInt(p[0],10)||1, yy=parseInt(p[1],10)||2000; return yy*12+(mm-1); }
// Place `count` day-indices spread EVENLY across a `days`-long month, rotated by `seed` so the exact
// days vary month to month. Always returns `count` distinct in-range indices (or all days if count≥days).
function _dhsSpreadDays(count, days, seed){
  count=Math.max(0,Math.min(days, count|0)); if(count<=0)return [];
  if(count>=days){ var all=[]; for(var i=0;i<days;i++)all.push(i); return all; }
  var offset=(((seed|0)%days)+days)%days, used={}, out=[];
  for(var k=0;k<count;k++){
    var pos=(Math.round(k*days/count)+offset)%days;
    while(used[pos])pos=(pos+1)%days;   // resolve a rare post-rotation collision
    used[pos]=1; out.push(pos);
  }
  return out.sort(function(a,b){return a-b;});
}
// The day-indices to check for one authorized task in a `days`-long month:
//  • "7 days per week" (daily) → EVERY day the month has (owner's rule).
//  • otherwise → the authorized monthly count = Time/Month ÷ Time/Day (falls back to the frequency
//    pattern's length if those times are missing), spread evenly and varied by the period seed.
function _dhsTaskDays(task, days, seed){
  var freq=String((task&&task.freq)||'').toLowerCase();
  if(/7 days? per week|daily|every day/.test(freq)){ var all=[]; for(var i=0;i<days;i++)all.push(i); return all; }
  var pd=_dhsHmToMin(task&&task.perDay), pm=_dhsHmToMin(task&&task.perMonth);
  // FLOOR so days×Time/Day never EXCEEDS the authorized Time/Month (never document more than
  // authorized); at least 1 day for any task that has a frequency at all.
  var count=(pd>0&&pm>0)?Math.max(1,Math.floor(pm/pd)):_dhsFreqToDays(task&&task.freq, days).length;
  return _dhsSpreadDays(count, days, seed);
}
// Minutes are printed as the ".MM" half of "HH.MM" on the certified form, so a bare "5" reads as
// 50 minutes: "20 hours 5 minutes" printed "20.5" — 45 minutes of over-billed time. Always 2 digits.
function _padMin(v){ v=(v==null?'':String(v)).trim(); return /^\d$/.test(v)?('0'+v):v; }
// Build the data for an invoice from a parsed authorization (res) + client (prof) + period MM/YYYY.
// The day grid is derived from each task's authorized frequency/count (see _dhsTaskDays), varied per
// month. Returns { data, unmapped:[taskNames] } — unmapped tasks are surfaced, never dropped silently.
function _dhsBuildFirstInvoice(res, prof, period){
  var pp=String(period||'').split('/'); if(pp.length!==2||pp[1].length!==4)return null;
  var days=daysIn(pp[0],pp[1]); var cols=_dhsSvcColNames();
  var grid=[]; for(var d=0;d<days;d++){ var row=[]; for(var c=0;c<cols.length;c++)row.push(false); grid.push(row); }
  var unmapped=[]; var seed=_dhsPeriodSeed(period);
  (res.tasks||[]).forEach(function(t){
    var col=_dhsMapTaskToCol(t.task);
    if(col<0){ unmapped.push(t.task); return; }
    _dhsTaskDays(t, days, seed).forEach(function(di){ if(di>=0&&di<days)grid[di][col]=true; });
  });
  var hours=res.hours!=null?String(res.hours):'', mins=res.minutes!=null?_padMin(res.minutes):'';
  var rate=stateRate();  // invoices always bill the flat state rate ($27), not the form's printed rate
  var T=today();
  return {
    data:{
      clientName:(prof&&prof.clientName)||'', medicaidId:(prof&&prof.medicaidId)||'', worker:(prof&&prof.worker)||'',
      billingPeriod:period, hourlyRate:rate,
      svcHH:hours, svcMM:mins, cplxHH:'', cplxMM:'', p1HH:'', p1MM:'', grandHH:hours, grandMM:mins,
      dateSubmitted:T, sigDate1:T, sigDate2:T, hasComplex:false, tasks:{svc:grid, cplx:[]}
    },
    unmapped:unmapped
  };
}

// Prorate a monthly authorization for a partial first month: authorized minutes × (days served ÷
// days in month), rounded to the nearest minute. Returns {hours, minutes, daysServed, daysInMonth}.
function _proratedFirstMonth(a, effectiveDate){
  var eff=String(effectiveDate||'').split('/');
  if(eff.length!==3)return null;
  var mm=parseInt(eff[0],10), dd=parseInt(eff[1],10), yyyy=parseInt(eff[2],10);
  if(!mm||!dd||!yyyy)return null;
  var inMonth=daysIn(String(mm).padStart(2,'0'),String(yyyy));
  if(dd<=1)return null;                       // started on the 1st — nothing to prorate
  var served=inMonth-dd+1;
  var totalMin=(parseInt(a.hours,10)||0)*60+(parseInt(a.minutes,10)||0);
  var proMin=Math.round(totalMin*served/inMonth);
  return { hours:Math.floor(proMin/60), minutes:proMin%60, daysServed:served, daysInMonth:inMonth, day:dd };
}
// Create a first invoice (Draft) for the active client from their DHS-1210 authorization.
function createFirstInvoiceFromAuth(){
  if(!activeProfileName)return;
  var p=getProfiles(); var prof=p[activeProfileName];
  if(!prof||!hasAuthorization(prof)){ showAlert('No DHS-1210 authorization on file for this client.'); return; }
  // Approved hours are the billed amount — without them the invoice would certify a blank Total Time.
  if(!hasBillableAuthorization(prof)){ showAlert('This authorization has no approved HOURS on file, so an invoice built from it would have a blank Total Time.\n\nAdd the approved hours on the Authorization tab (or re-import the DHS-1210), then generate.'); return; }
  var a=prof.authorization;
  var eff=(a.effectiveDate||'').split('/');
  var period=(eff.length===3)?(eff[0]+'/'+eff[2]):'';
  if(!period){ showAlert('The authorization has no effective date, so the billing period can’t be set. Add one on the Authorization tab first.'); return; }
  if((prof.invoices||[]).some(function(i){return i.billingPeriod===period;})){ showAlert('An invoice for '+period+' already exists for this client.'); return; }
  // A mid-month start can be billed either way depending on the case, so ASK — this used to bill the
  // full month silently (_firstOfEffectiveMonth pushes the start to the 1st), which over-bills MDHHS
  // whenever service did not actually cover the whole month.
  var _pro=_proratedFirstMonth(a, a.effectiveDate);
  if(_pro){
    var _full=a.hours+':'+_padMin(a.minutes||0);
    var _part=_pro.hours+':'+_padMin(_pro.minutes);
    showConfirm(
      activeProfileName+'’s service starts on day '+_pro.day+' of '+period+', so this first month is partial.\n\n'+
      'How should it be billed?\n\n'+
      '• Full month — the whole authorized '+_full+' (correct when MDHHS pays the monthly amount regardless of start day)\n'+
      '• Prorate — '+_part+', for the '+_pro.daysServed+' of '+_pro.daysInMonth+' days actually served\n\n'+
      'Either way the invoice is created as a Draft you can edit before sending.',
      function(){ _createFirstInvoice(prof,a,period,null); },
      { title:'Partial first month', okText:'Bill full month', danger:false,
        extraText:'Prorate to '+_part, onExtra:function(){ _createFirstInvoice(prof,a,period,_pro); } }
    );
    return;
  }
  _createFirstInvoice(prof,a,period,null);
}
// `pro` is null for a full month, or the _proratedFirstMonth result to bill only the days served.
function _createFirstInvoice(prof,a,period,pro){
  var built=_dhsBuildFirstInvoice({hours:(pro?pro.hours:a.hours), minutes:(pro?pro.minutes:a.minutes),
                                   rate:a.rate, tasks:a.tasks||[]}, prof, period);
  if(!built){ showAlert('Could not build the invoice from this authorization.'); return; }
  var inv={ id:'dhs_'+Date.now()+'_'+Math.random().toString(36).slice(2,7), billingPeriod:period,
    savedAt:new Date().toLocaleString(), status:'draft', invoiceNote:'', data:built.data };
  if(!prof.invoices)prof.invoices=[];
  prof.invoices.unshift(inv);
  saveProfilesLS(p); saveProfileSP(activeProfileName, prof);
  if(typeof logActivity==='function')logActivity('invoice','First invoice created from DHS-1210 for '+activeProfileName+' ('+period+')');
  var msg='✓ Draft invoice created for '+period+' from the authorization.'+
    (pro?('\n\nPRORATED to '+pro.hours+':'+_padMin(pro.minutes)+' for the '+pro.daysServed+' of '+pro.daysInMonth+' days served.'):'')+
    '\n\nIt is a Draft — review the day grid (the checked days are a starting pattern from each task’s frequency) before sending.';
  if(built.unmapped.length)msg+='\n\n⚠ These tasks didn’t match a service column and were left OFF — add them manually: '+built.unmapped.join(', ');
  showAlert(msg,{title:'First Invoice Created'});
  if(typeof renderAuthPane==='function')renderAuthPane(false);
  if(typeof switchTab==='function')switchTab('history');
}

// Returns clients eligible for auto-gen: active in the period, missing this period's invoice, and
// (required) with a DHS authorization to build from. No authorization → no auto-generated invoice.
function findClientsEligibleForAutoGen(period){
  var profiles=getProfiles();var out=[];
  Object.keys(profiles).forEach(function(name){
    var p=profiles[name];
    if(!isInvoiceableStatus(p))return;   // In Progress / Lost / Terminated are never invoiced
    if(isCarrierClient(p))return;     // managed-care clients are billed by the carrier — never auto-generate
    if(!clientWasActiveInPeriod(p,period))return;
    if(!hasBillableAuthorization(p))return;   // need APPROVED HOURS — never generate a blank-total invoice
    var hasCurrent=(p.invoices||[]).some(function(i){return i.billingPeriod===period;});
    if(hasCurrent)return;
    // Most recent prior invoice (if any) — kept only as a fallback signal; the grid is built fresh
    // from the authorization, not copied from it.
    var prior=(p.invoices||[]).slice().sort(function(a,b){return new Date(b.savedAt||0)-new Date(a.savedAt||0);})[0];
    out.push({name:name,prevInv:prior});
  });
  return out;
}

// Bulk auto-generate: confirm, then create new draft invoices for all eligible clients
function autoGenerateMonthlyInvoices(period){
  var eligible=findClientsEligibleForAutoGen(period);
  if(!eligible.length){
    showAlert('No clients need auto-generation. Either everyone active already has an invoice for '+period+', or the active clients don’t have a DHS authorization with approved hours to build from (an authorization with no hours parsed is skipped — fix it on the Authorization tab).',{title:'Nothing to Generate'});
    return;
  }
  var names=eligible.map(function(e){return '• '+e.name;}).join('\n');
  showConfirm(
    'Auto-generate '+eligible.length+' invoice'+(eligible.length>1?'s':'')+' for '+period+'?\n\n'+
    names+'\n\n'+
    'For each client WITH a DHS authorization, this builds the grid from the authorized tasks:\n'+
    '• Daily tasks (7 days/week) checked every day of the month\n'+
    '• Weekly/monthly tasks get their authorized count, spread out and varied so each month differs\n'+
    '• Only authorized tasks are checked (others stay empty)\n'+
    '(Clients without an authorization fall back to copying last month.)\n'+
    '• Sets Date Submitted to today\n\n'+
    'All generated invoices are marked Draft — review before sending.',
    function(){_doAutoGenerateInvoices(eligible,period);},
    {title:'Auto-Generate Invoices',okText:'Generate '+eligible.length}
  );
}
function _doAutoGenerateInvoices(eligible,period){
  if(!isInvoiceAdmin()){showAlert('Only the account owner can generate invoices.');return;}
  var profiles=getProfiles();
  var generated=0,skipped=0,unmappedBy={};
  var undoBatch={id:'b_'+Date.now()+'_'+Math.random().toString(36).slice(2,8),period:period,when:Date.now(),invoices:[]};
  eligible.forEach(function(e){
    // Build the month's grid from the client's DHS authorization (correct counts per the authorized
    // frequency, varied each month). Eligibility already requires an authorization.
    var prof=profiles[e.name], a=prof&&prof.authorization, newInv=null;
    var built=(a && hasBillableAuthorization(prof)) ? _dhsBuildFirstInvoice({hours:a.hours,minutes:a.minutes,rate:a.rate,tasks:a.tasks||[]}, prof, period) : null;
    if(built) newInv={ billingPeriod:period, savedAt:new Date().toLocaleString(), status:'draft', invoiceNote:'', data:built.data };
    if(!newInv){skipped++;return;}
    // Surface tasks that matched no service column — parity with createFirstInvoiceFromAuth. Silently
    // dropping them yields a certified form with billed hours and an empty grid.
    if(built.unmapped&&built.unmapped.length)unmappedBy[e.name]=built.unmapped.slice();
    if(!newInv.id)newInv.id='auto_'+Date.now()+'_'+Math.random().toString(36).slice(2,9);
    if(!profiles[e.name].invoices)profiles[e.name].invoices=[];
    profiles[e.name].invoices.unshift(newInv);
    undoBatch.invoices.push({clientName:e.name,invoiceId:newInv.id,dataHash:_quickInvoiceHash(newInv.data)});
    generated++;
  });
  saveProfilesLS(profiles);
  batchSaveWithSummary('Auto-generate', eligible.filter(function(e){return profiles[e.name];}).map(function(e){var th=function(){return saveProfileSP(e.name,profiles[e.name],true);};th._label=e.name;return th;}));
  if(undoBatch.invoices.length)_pushAutoGenUndoBatch(undoBatch);
  logActivity('invoice','Auto-generated '+generated+' invoice'+(generated!==1?'s':'')+' for '+period);
  var _unNames=Object.keys(unmappedBy);
  var _unMsg=_unNames.length?('\n\n⚠ Some tasks didn’t match a service column and were left OFF the grid — add them manually:\n'+
    _unNames.slice(0,10).map(function(n){return '• '+n+': '+unmappedBy[n].join(', ');}).join('\n')+
    (_unNames.length>10?('\n…and '+(_unNames.length-10)+' more client(s).'):'')):'';
  showAlert(
    '✓ Auto-generated '+generated+' invoice'+(generated!==1?'s':'')+' for '+period+(skipped>0?' ('+skipped+' skipped — missing data)':'')+'.\n\n'+
    'All new invoices are marked Draft. An Undo banner will stay on the Clients page for 24 hours — use it if you change your mind.'+_unMsg,
    {title:'Auto-Generation Complete'}
  );
  if(typeof previewMonthlyInvoices==='function')previewMonthlyInvoices();
  updateStats();
}

// Send to ALL caseworkers (one email per caseworker, sequentially) — for monthly batch day
async function sendAllCaseworkerEmails(period){
  if(!spToken){
    showConfirm('You need to sign in with your Microsoft account to send emails. Click Sign In to be redirected.',function(){signIn();},{title:'Sign In Required',okText:'Sign In',danger:false});
    return;
  }
  // Build the same groups previewMonthlyInvoices uses
  var profiles=getProfiles();var cws=getCaseworkers();
  var groups={};
  Object.keys(profiles).forEach(function(name){
    var prof=profiles[name];
    if(!isInvoiceableStatus(prof))return;   // In Progress / Lost / Terminated are never invoiced
    if(isCarrierClient(prof))return;   // managed-care clients are billed by the carrier — never invoiced here
    if(!clientWasActiveInPeriod(prof,period))return;
    var rawWorker=(prof.worker||'').trim();
    var cwRec=cws.find(function(c){return c.id&&prof.caseworkerId&&String(c.id)===String(prof.caseworkerId);})||
              cws.find(function(c){return (c.name||'').trim().toLowerCase()===rawWorker.toLowerCase()&&rawWorker;})||
              {};
    if(!cwRec.email)return; // skip groups with no email
    var groupKey=cwRec.id?'cw:'+cwRec.id:(rawWorker?'nm:'+rawWorker.toLowerCase():'(No Worker Assigned)');
    var displayName=(cwRec.name||rawWorker||'(No Worker Assigned)').trim();
    if(!groups[groupKey])groups[groupKey]={email:cwRec.email,cwName:displayName,clients:[]};
    var inv=(prof.invoices||[]).find(function(i){return i.billingPeriod===period;})||null;
    var issues=validateInvoiceForSend(name,prof,inv,cwRec);
    groups[groupKey].clients.push({name:name,medicaidId:prof.medicaidId||'',invStatus:inv?(inv.status||'draft'):'none',hasIssues:issues.length>0,issues:issues});
  });
  // Determine which groups have anything sendable
  var sendable=[];
  Object.keys(groups).forEach(function(gk){
    var g=groups[gk];
    var ready=g.clients.filter(function(c){return c.invStatus==='draft' && !c.hasIssues;});
    if(ready.length)sendable.push({wname:g.cwName,group:g,readyCount:ready.length});
  });
  if(!sendable.length){
    showConfirm('No caseworkers have invoices ready to send for '+period+'. Use the per-caseworker Send Email button to see what\'s missing.',function(){},{title:'Nothing to Send',okText:'OK',danger:false});
    return;
  }
  var totalInvoices=sendable.reduce(function(s,x){return s+x.readyCount;},0);
  showConfirm(
    'Ready to send '+totalInvoices+' invoice'+(totalInvoices>1?'s':'')+' to '+sendable.length+' caseworker'+(sendable.length>1?'s':'')+' for billing period '+period+'.\n\n'+
    sendable.map(function(x){return '• '+x.wname+' — '+x.readyCount+' invoice'+(x.readyCount>1?'s':'');}).join('\n')+
    '\n\nThis will send the emails one after another. Continue?',
    async function(){
      var sentCount=0;
      var BULK_THROTTLE_MS=3000; // 3-sec delay between caseworker sends — avoids burst-send pattern that trips spam filters
      for(var i=0;i<sendable.length;i++){
        var item=sendable[i];
        // CRITICAL: only send the READY clients. Bucket the same way sendMonthlyEmail does.
        var allClients=item.group.clients;
        var alreadySent=allClients.filter(function(c){return c.invStatus==='submitted'||c.invStatus==='paid';});
        var hasIssues=allClients.filter(function(c){return c.invStatus==='draft'&&c.hasIssues;});
        var missingInvoice=allClients.filter(function(c){return !c.invStatus||c.invStatus==='none';});
        var readyToSend=allClients.filter(function(c){return c.invStatus==='draft'&&!c.hasIssues;});
        if(!readyToSend.length)continue;
        try{
          // Await DIRECTLY so each caseworker's PDF capture finishes before the next worker starts.
          // (Earlier bug: sendMonthlyEmail did fire-and-forget, causing parallel page-invoice usage and PDFs going to wrong worker.)
          var _res=await _doMonthlyEmailSend(item.group.email,item.wname,period,readyToSend,alreadySent.length,hasIssues,missingInvoice);
          // Only count a caseworker as SENT when the send actually reported success. The inner function
          // returns undefined on its failure paths, and this used to increment regardless — so on
          // billing day the closing toast counted failures as successes.
          if(_res&&_res.ok)sentCount++; else failedWorkers.push(item.wname);
          // Throttle between sends (skip after the last one)
          if(i<sendable.length-1){
            await new Promise(function(r){setTimeout(r,BULK_THROTTLE_MS);});
          }
        }catch(e){console.error('Send to '+item.wname+' failed:',e); failedWorkers.push(item.wname);}
      }
      showToast('✓ Sent batches to '+sentCount+' caseworker'+(sentCount===1?'':'s')+'.',5000);
      if(failedWorkers.length){
        showAlert('These caseworkers were NOT sent — their clients are still Draft and will reappear next run:\n\n• '+
          failedWorkers.join('\n• ')+'\n\nRe-run Send All, or send them individually.',{title:'Some batches failed',danger:true});
      }
      // Refresh preview
      previewMonthlyInvoices();
    },
    {title:'Send All Caseworker Emails',okText:'Send All',danger:false}
  );
}

async function sendMonthlyEmail(email,workerName,clients,period){
  if(!isInvoiceAdmin()){showAlert('Only the account owner can send monthly emails.');return;}
  if(!spToken){
    showConfirm('You need to sign in with your Microsoft account to send emails. Click Sign In to be redirected.',
      function(){signIn();},
      {title:'Sign In Required',okText:'Sign In',danger:false}
    );
    return;
  }
  // Bucket clients by current status:
  // - alreadySent: status submitted/paid (skip — they got the previous email)
  // - readyToSend: status draft AND no issues
  // - hasIssues: status draft AND validation issues
  // - missingInvoice: no invoice for this period
  var alreadySent=[],readyToSend=[],hasIssues=[],missingInvoice=[];
  clients.forEach(function(c){
    if(!c.invStatus||c.invStatus==='none'){missingInvoice.push(c);return;}
    if(c.invStatus==='submitted'||c.invStatus==='paid'){alreadySent.push(c);return;}
    if(c.hasIssues){hasIssues.push(c);return;}
    readyToSend.push(c);
  });

  var allWithInvoice=readyToSend.concat(hasIssues);

  function sendReadyOnly(){
    return _doMonthlyEmailSend(email,workerName,period,readyToSend,alreadySent.length,hasIssues,missingInvoice);
  }

  // Case 1: nothing has an invoice for this period at all
  if(!allWithInvoice.length){
    if(alreadySent.length===clients.length){
      showConfirm('All '+alreadySent.length+' invoice'+(alreadySent.length>1?'s':'')+' for '+period+' have already been submitted to '+workerName+'.\n\nNo new invoices to send.',function(){},{title:'Already Sent',okText:'OK',danger:false});
      return;
    }
    showConfirm('No invoices exist for '+period+' yet. Fill out the invoices on each client page first.',function(){},{title:'Nothing to Send',okText:'OK',danger:false});
    return;
  }

  // Case 2: ALL invoices have issues — block. User must fix before sending.
  if(readyToSend.length===0&&hasIssues.length>0){
    var allProblemNames=hasIssues.map(function(c){return c.name+(c.issues&&c.issues.length?' ('+c.issues[0]+')':'');}).join('\n• ');
    showConfirm(
      'Cannot send — all '+hasIssues.length+' invoice'+(hasIssues.length>1?'s have':' has')+' validation issues that must be fixed first:\n\n• '+allProblemNames+'\n\nUse the Preview button on each client to see the specific issue, or open the client and fix the missing fields.',
      function(){},
      {title:'Cannot Send',okText:'OK',danger:false}
    );
    return;
  }

  // Case 3: SOME ready + SOME with issues — push through the ready ones; the others stay queued for follow-up
  if(hasIssues.length>0){
    var problemNames=hasIssues.map(function(c){return c.name+(c.issues&&c.issues.length?' ('+c.issues[0]+')':'');}).join('\n• ');
    var readyNames=readyToSend.map(function(c){return c.name;}).join('\n• ');
    var extraInfo='';
    if(alreadySent.length)extraInfo+='\n\nAlready submitted ('+alreadySent.length+' — will not re-send):\n• '+alreadySent.map(function(c){return c.name;}).join('\n• ');
    if(missingInvoice.length)extraInfo+='\n\nNo invoice yet ('+missingInvoice.length+'):\n• '+missingInvoice.map(function(c){return c.name;}).join('\n• ');
    showConfirm(
      readyToSend.length+' of '+allWithInvoice.length+' invoices are ready to send to '+workerName+'.\n\n'+
      'Will send ('+readyToSend.length+'):\n• '+readyNames+'\n\n'+
      'Skipping for issues ('+hasIssues.length+'):\n• '+problemNames+
      extraInfo+
      '\n\nSend the '+readyToSend.length+' ready one'+(readyToSend.length>1?'s':'')+' now?',
      sendReadyOnly,
      {title:'Send Ready Invoices?',okText:'Send '+readyToSend.length+' Now',danger:false}
    );
    return;
  }

  // Case 4: all ready, no issues — send straight through
  sendReadyOnly();
}

var _monthlyEmailSendInProgress=false;
async function _doMonthlyEmailSend(email,workerName,period,readyToSend,alreadySentCount,hasIssues,missingInvoice){
  if(_monthlyEmailSendInProgress){
    console.warn('Monthly email send already in progress — refusing to start a parallel send');
    showAlert('Another email send is already running. Please wait for it to finish before sending again.');
    return {ok:false,err:'concurrent_send_blocked'};
  }
  _monthlyEmailSendInProgress=true;
  try{
    return await _doMonthlyEmailSendInner(email,workerName,period,readyToSend,alreadySentCount,hasIssues,missingInvoice);
  } finally {
    _monthlyEmailSendInProgress=false;
  }
}
async function _doMonthlyEmailSendInner(email,workerName,period,readyToSend,alreadySentCount,hasIssues,missingInvoice){
  var profiles=getProfiles();
  var withInv=readyToSend; // only the ready ones

  var po=document.getElementById('monthlyProgressOverlay');
  var pb=document.getElementById('monthlyProgressBar');
  var pl=document.getElementById('monthlyProgressLabel');
  var ps=document.getElementById('monthlyProgressSub');
  if(po)po.classList.add('open');
  if(ps)ps.textContent='Generating PDFs for '+withInv.length+' client'+(withInv.length>1?'s':'')+'…';

  // Save current invoice page state to restore later
  var savedPage=document.querySelector('.page.active');
  var savedActive=savedPage?savedPage.id:'page-home';
  var invPage=document.getElementById('page-invoice');
  // Make invoice page off-screen but renderable
  invPage.classList.add('active');
  invPage.style.position='fixed';invPage.style.left='-9999px';invPage.style.top='0';invPage.style.zIndex='-1';

  var attachments=[];
  for(var i=0;i<withInv.length;i++){
    var c=withInv[i];
    if(pb)pb.style.width=Math.round(((i)/withInv.length)*100)+'%';
    if(pl)pl.textContent='Processing: '+c.name+' ('+(i+1)+' of '+withInv.length+')';
    var prof=profiles[c.name]||{};
    var inv=(prof.invoices||[]).find(function(inv2){return inv2.billingPeriod===period;})||{};
    await loadInvoiceForCapture(c.name,inv,period);
    try{
      var base64=await captureInvoicePDF();
      var fname=c.name.replace(/[^a-z0-9]/gi,'_')+'_'+period.replace('/','_')+'.pdf';
      attachments.push({name:fname,base64:base64,clientName:c.name});
    }catch(e){console.error('PDF capture failed',e);}
  }

  invPage.classList.remove('active');
  invPage.style.position='';invPage.style.left='';invPage.style.top='';invPage.style.zIndex='';
  document.querySelectorAll('.page').forEach(function(el){el.classList.remove('active');});
  var restorePage=document.getElementById(savedActive);
  if(restorePage)restorePage.classList.add('active');

  if(pb)pb.style.width='100%';
  if(pl)pl.textContent='Sending email to '+workerName+'…';

  if(!attachments.length){
    if(po)po.classList.remove('open');
    showAlert('No PDFs could be generated. Please check the console for errors.');return;
  }

  var workerDisplay=workerName&&workerName!=='(No Worker Assigned)'?workerName:'Caseworker';
  // First name only for friendly subject/greeting
  var workerFirst=(workerDisplay.split(/\s+/)[0]||'').replace(/[^a-zA-Z\-']/g,'')||workerDisplay;
  // Period "05/2026" → "May 2026" for readability
  var _months=['January','February','March','April','May','June','July','August','September','October','November','December'];
  var _pParts=(period||'').split('/');
  var periodLabel=(_pParts.length===2 && parseInt(_pParts[0],10)>=1 && parseInt(_pParts[0],10)<=12)
    ? _months[parseInt(_pParts[0],10)-1]+' '+_pParts[1]
    : period;
  var isFollowUp=alreadySentCount>0;
  // Subject is exactly "INVOICE" (agency direction — caseworkers file/match on that word). The
  // month/year and the client list live in the email body, not the subject.
  var subj='INVOICE';
  var _list='<ul>'+attachments.map(function(a){return '<li>'+esc(a.clientName)+'</li>';}).join('')+'</ul>';
  var multi=attachments.length>1;
  var _seed=(email||'')+'|'+(period||'');
  var _mword=(periodLabel||'').split(' ')[0];
  var _apprec=_apprecLine(_seed,multi);  // plural matches count: "these" (multi) / "this" (one)
  var body;
  if(isFollowUp){
    body='<p>Hi '+esc(workerFirst)+',</p>'+
      '<p>A quick follow-up to my earlier note — attached '+(multi?'are a few additional invoices':'is one additional invoice')+' for '+periodLabel+' that I wanted to be sure reached you, listed below. Please let me know if you have any questions.'+_apprec+'</p>'+
      _list+_emailSig();
  } else {
    // Rotated, month-start opener — seeded per caseworker so a batch doesn't read identically.
    var _openers=[
      'Hope '+_mword+' is off to a good start. Attached '+(multi?'are the':'is the')+' invoice'+(multi?'s':'')+' for our shared client'+(multi?'s':'')+', listed below.',
      'Now that '+_mword+' is underway, here '+(multi?'are the':'is the')+' invoice'+(multi?'s':'')+' for our shared client'+(multi?'s':'')+' below.',
      'Attached '+(multi?'are the':'is the')+' '+periodLabel+' invoice'+(multi?'s':'')+' for our shared client'+(multi?'s':'')+', listed below.'
    ];
    body='<p>Hi '+esc(workerFirst)+',</p>'+
      '<p>'+_openers[_emailSeed(_seed,_openers.length)]+' Please review at your convenience and let me know if anything needs to be adjusted.'+_apprec+'</p>'+
      _list+_emailSig();
  }

  // Resolve supervisor CC — caseworker → supervisor_id → supervisor.email
  var cwRec=getCaseworkers().find(function(c){return (c.email||'').toLowerCase()===(email||'').toLowerCase();})||{};
  var sups=getSupervisors();
  var sup=cwRec.supervisor_id?(sups[cwRec.supervisor_id]||null):null;
  var ccEmails=(sup && sup.email)?[sup.email]:[];

  var result=await sendMailWithPDF(email,subj,body,attachments,function(done,total,label){
    if(pb)pb.style.width=Math.round((done/Math.max(total,1))*100)+'%';
    if(pl)pl.textContent=label||('Sending '+done+' of '+total+'…');
  },ccEmails);
  if(po)po.classList.remove('open');

  // HIPAA audit: log every PHI email send (success OR failure)
  logEmailSend({
    type:'mass',
    recipient:email,
    cc:ccEmails,
    caseworkerName:workerName||'(none)',
    billingPeriod:period,
    clientNames:attachments.map(function(a){return a.clientName;}),
    attachmentCount:attachments.length,
    success:!!result.ok,
    errorMsg:result.ok?null:(result.err||result.status||'unknown')
  });

  if(result.ok){
    // Mark all sent invoices as submitted
    attachments.forEach(function(a){markInvoiceSubmitted(a.clientName,period);});
    closeMonthlyInvModal();
    updateStats();
    var toastMsg='✓ '+attachments.length+(isFollowUp?' additional':'')+' invoice'+(attachments.length>1?'s':'')+' emailed to '+workerDisplay+' — status set to Submitted';
    if(hasIssues&&hasIssues.length){
      toastMsg+=' ('+hasIssues.length+' still need'+(hasIssues.length>1?'':'s')+' fixing)';
    }
    showToast(toastMsg,6000);
    // Refresh the modal results if re-opened
  }else{
    var msg2='Email failed to send.';
    if(result.status===401)msg2='Authentication error (401) — please sign out and sign back in.';
    else if(result.status===403)msg2='Permission denied (403) — sign out, sign back in, and accept the Mail.Send permission when prompted.';
    else if(result.err)msg2='Error ('+(result.status||'?')+'):\n'+result.err.slice(0,300);
    showAlert(msg2);
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// ✨ ASSISTANT — one global AI chat that can look up clients, run roster reports, and
// assemble emails/forms for the owner to review + send. The model (backend /ai-chat) decides
// which TOOL to call; the tools run here in the browser because they read localStorage and use the
// user's Microsoft token. Nothing is ever sent without the owner reviewing it in the compose modal.
// ═══════════════════════════════════════════════════════════════════════════════
var _asstMessages=[], _asstBusy=false, _asstOpen=false;

// The non-PHI facts the model needs to sign emails + greet by time of day. Sent each turn.
function _asstContext(){
  var ag=getAgencyInfo()||{};
  var hr=(new Date()).getHours(), tod=hr<12?'morning':(hr<17?'afternoon':'evening');
  return { agency:{ agency_provider_name:ag.agency_provider_name, agency_name:ag.agency_name,
    agency_phone:ag.agency_phone, agency_fax:ag.agency_fax }, today:today(), time_of_day:tod };
}

// Parse a stored date ('YYYY-MM-DD' or 'M/D/YYYY') → {y,m} for start-date filters/aggregation.
function _asstParseDate(v){
  if(!v)return null; v=String(v).trim();
  // d = day of month (1 when the value carries no day, e.g. "2026-05"). before/after compare it, so
  // an intra-month range ("started after 05/15/2026") is answered correctly instead of by month only.
  var m=v.match(/^(\d{4})-(\d{1,2})(?:-(\d{1,2}))?/); if(m)return {y:+m[1],m:+m[2],d:+(m[3]||1)};
  m=v.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/); if(m)return {y:+m[3],m:+m[1],d:+m[2]};
  return null;
}
// Sortable YYYYMMDD number for a parsed date — one comparison for y/m/d.
function _asstDateNum(d){ return d?(d.y*10000+d.m*100+(d.d||1)):0; }
// Turn "May" / "5" / "2026-05" into a month number 1-12 (0 = couldn't tell).
function _asstMonthNum(s){
  s=String(s||'').trim().toLowerCase();
  var m=s.match(/(\d{4})-(\d{1,2})/); if(m)return +m[2];
  if(/^\d{1,2}$/.test(s))return +s;
  var names=['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];
  var i=names.indexOf(s.slice(0,3)); return i>=0?i+1:0;
}
// ── Tool: find_client ── resolve a client by (partial) name → full record + caregiver + caseworker.
function _asstFindClient(args){
  var q=((args&&args.name)||'').trim().toLowerCase();
  var toks=q.split(/\s+/).filter(Boolean);
  var profs=getProfiles(), cgs=getCaregivers(), cws=getCaseworkers();
  var out=[];
  Object.keys(profs).forEach(function(key){
    var p=profs[key]; var hay=key.toLowerCase();
    if(toks.length && !toks.every(function(t){return hay.indexOf(t)>=0;})) return;
    var cg=(p.caregiverId&&cgs[p.caregiverId])?cgs[p.caregiverId]:null;
    var cw=null;
    if(p.caseworkerId)cw=cws.find(function(c){return c.id===p.caseworkerId;});
    if(!cw&&p.worker)cw=cws.find(function(c){return (c.name||'').toLowerCase()===(p.worker||'').toLowerCase();});
    out.push({ client_name:key, status:(p.clientStatus||'active'),
      county:p.county||'', address:p.street||p.address||'', city:p.city||'', state:p.state||'', zip:p.zip||'',
      phone:p.phone||'', email:p.clientEmail||p.cemail||'', dob:p.dob||'', medicaid_id:p.medicaidId||'',
      caregiver: cg?{name:cg.name||'',email:cg.email||'',phone:cg.phone||''}:null,
      caseworker: cw?{name:cw.name||'',email:cw.email||'',phone:cw.phone||''}:(p.worker?{name:p.worker,email:'',phone:''}:null) });
  });
  var capped=out.slice(0,8);
  return { matches:capped, count:out.length, truncated: out.length>capped.length };
}

// Flatten one client profile into a flat field map the query engine filters/groups over. This is the
// single place fields are exposed — add a line here and every query gets the new field for free.
function _asstClientFields(p, cgs, cws){
  var cg=(p.caregiverId&&cgs[p.caregiverId])?cgs[p.caregiverId]:null;
  var cwRec=p.caseworkerId?cws.find(function(c){return c.id===p.caseworkerId;}):null;
  var cwn=cwRec?cwRec.name:(p.worker||'');
  return {
    status:(p.clientStatus||'active'), county:(p.county||''), city:(p.city||''), state:(p.state||''),
    zip:(p.zip||''), address:(p.street||p.address||''), caregiver:(cg?cg.name:''), caseworker:(cwn||''),
    caregiver_email:(cg?(cg.email||''):''), caseworker_email:(cwRec?(cwRec.email||''):''),
    medicaid_id:(p.medicaidId||''), medicare:(p.medicare||''), dob:(p.dob||''),
    phone:(p.phone||p.homePhone||''), email:(p.clientEmail||p.cemail||''), start_date:(p.startDate||''),
    hourly_rate:(p.hourlyRate||''), live_in:(p.liveIn?'yes':'no'),
    program:(p.program==='carrier'?'carrier':'champs'), carrier:(p.carrier||'')
  };
}
function _asstFieldAlias(f){
  f=String(f||'').toLowerCase().replace(/\s+/g,'_');
  if(f==='medicaid'||f==='medicaid_number'||f==='medicaid_#')return 'medicaid_id';
  if(f==='worker'||f==='case_worker')return 'caseworker';
  if(f==='name')return 'client';
  return f;
}
// Apply one {field, op, value} filter to a flat field map. Deterministic — this is where accuracy comes from.
function _asstMatchFilter(fields, flt){
  if(!flt||!flt.field)return true;
  var field=_asstFieldAlias(flt.field);
  var op=String(flt.op||'eq').toLowerCase();
  var raw=(field in fields)?fields[field]:'';
  var s=String(raw==null?'':raw).toLowerCase().trim();
  var v=String(flt.value==null?'':flt.value).toLowerCase().trim();
  // The UI shows "In Progress" for a client STORED as 'inactive' (see the status dropdown), so a
  // filter typed as "in progress" matched nothing. Normalize both sides to the stored value.
  if(field==='status'){ var _st=function(x){ return /^in[\s-]?progress$/.test(x)?'inactive':x; }; s=_st(s); v=_st(v); }
  switch(op){
    case 'present': case 'has': case 'exists': return s!=='' && s!=='no';
    case 'missing': case 'empty': case 'none': return s==='' || s==='no';
    case 'eq': case 'is': case 'equals': return s===v;
    case 'ne': case 'not_eq': case 'not': return s!==v;
    case 'contains': case 'like': return v===''||s.indexOf(v)>=0;
    case 'in_month': case 'month': { var d=_asstParseDate(raw); return !!d && d.m===_asstMonthNum(flt.value); }
    case 'in_year': case 'year': { var d2=_asstParseDate(raw); return !!d2 && String(d2.y)===String(flt.value).replace(/\D/g,''); }
    case 'before': { var d3=_asstParseDate(raw), t3=_asstParseDate(flt.value); return !!d3&&!!t3 && _asstDateNum(d3)<_asstDateNum(t3); }
    case 'after': { var d4=_asstParseDate(raw), t4=_asstParseDate(flt.value); return !!d4&&!!t4 && _asstDateNum(d4)>_asstDateNum(t4); }
    case 'gt': return parseFloat(s)>parseFloat(v);
    case 'lt': return parseFloat(s)<parseFloat(v);
    default: return v===''?true:s.indexOf(v)>=0;
  }
}
function _asstGroupValue(fields, g){
  g=_asstFieldAlias(g);
  if(g==='start_month'||g==='month'){ var d=_asstParseDate(fields.start_date); return d?(d.y+'-'+String(d.m).padStart(2,'0')):'(no start date)'; }
  var v=(g in fields)?fields[g]:'';
  return (v===''||v==null)?'(none)':v;
}
// ── Tool: query_roster ── ONE general, deterministic query engine over the whole roster. The model
// composes {action, filters, group_by}; CODE computes the answer (count/list/group) — so the AI never
// counts in its head. Any field, any filter, any grouping — no per-question code needed.
function _asstQueryRoster(args){
  args=args||{};
  var action=String(args.action||'list').toLowerCase();
  var filters=Array.isArray(args.filters)?args.filters:[];
  var groupBy=args.group_by||'';
  var profs=getProfiles(), cgs=getCaregivers(), cws=getCaseworkers();
  var matched=[];
  Object.keys(profs).forEach(function(key){
    var f=_asstClientFields(profs[key],cgs,cws);
    if(filters.every(function(flt){return _asstMatchFilter(f,flt);})) matched.push({name:key,f:f});
  });
  if(action==='count'){
    return { action:'count', count:matched.length, matched_names:matched.slice(0,150).map(function(m){return m.name;}),
      truncated:matched.length>150 };
  }
  if(action==='group'){
    var groups={};
    matched.forEach(function(m){ var k=String(_asstGroupValue(m.f,groupBy)); groups[k]=(groups[k]||0)+1; });
    var arr=Object.keys(groups).map(function(k){return {value:k,count:groups[k]};}).sort(function(a,b){return b.count-a.count;});
    return { action:'group', group_by:_asstFieldAlias(groupBy), total:matched.length, groups:arr };
  }
  // The 80-row cap keeps the MODEL's context small. export_data passes no_cap:true so a downloaded
  // file contains every matched row — exporting 80 of 200 clients (and reporting 80 as the total)
  // is silent data loss in something the user keeps as a record.
  var rowsAll=(args.no_cap===true)?matched:matched.slice(0,80);
  var cap=rowsAll.map(function(m){ return { name:m.name, status:m.f.status, county:m.f.county,
    caregiver:m.f.caregiver, caseworker:m.f.caseworker, medicaid_id:m.f.medicaid_id, start_date:m.f.start_date }; });
  return { action:'list', count:matched.length, clients:cap, truncated:matched.length>cap.length };
}
// ── Tool: onboarding_status ── deterministic checklist of what's left to onboard/transfer a client.
function _asstOnboardingStatus(args){
  args=args||{};
  var cname=(args.client_name||'').trim();
  var profs=getProfiles(); var key=cname;
  if(!profs[key]){
    var toks=cname.toLowerCase().split(/\s+/).filter(Boolean);
    key=Object.keys(profs).find(function(k){var h=k.toLowerCase();return toks.length&&toks.every(function(t){return h.indexOf(t)>=0;});})||cname;
  }
  var p=profs[key];
  if(!p)return {error:'No client named "'+cname+'" was found.'};
  var cgs=getCaregivers(), cws=getCaseworkers();
  var cg=(p.caregiverId&&cgs[p.caregiverId])?cgs[p.caregiverId]:null;
  var cwRec=p.caseworkerId?cws.find(function(c){return c.id===p.caseworkerId;}):null;
  var cwName=cwRec?cwRec.name:(p.worker||'');
  var cwEmail=cwRec?(cwRec.email||''):'';
  var carrier=(p.program==='carrier');
  // The checklist swaps by program: a managed-care (carrier) client needs the carrier + member # +
  // a care coordinator (a caseworker record), and does NOT need a caseworker email or an MSA-4676.
  var checks = carrier ? [
    {item:'Carrier', ok:!!(p.carrier&&String(p.carrier).trim())},
    {item:'Member #', ok:!!(p.memberId&&String(p.memberId).trim())},
    {item:'Care coordinator assigned', ok:!!cwName},
    {item:'Caregiver assigned', ok:!!cg},
    {item:'Date of birth', ok:!!p.dob},
    {item:'Address', ok:!!((p.street||p.address)&&p.city&&p.zip)},
    {item:'Phone', ok:!!(p.phone||p.homePhone)},
    {item:'Start date', ok:!!p.startDate}
  ] : [
    {item:'Medicaid ID', ok:!!(p.medicaidId&&String(p.medicaidId).trim())},
    {item:'Caregiver assigned', ok:!!cg},
    {item:'Caseworker assigned', ok:!!cwName},
    {item:'Caseworker email', ok:!!cwEmail},
    {item:'Date of birth', ok:!!p.dob},
    {item:'Address', ok:!!((p.street||p.address)&&p.city&&p.zip)},
    {item:'Phone', ok:!!(p.phone||p.homePhone)},
    {item:'Start date', ok:!!p.startDate}
  ];
  // "4676 sent?" (CHAMPS only) — best-effort: scan the activity log for a 4676 email for this client.
  var sent4676=false;
  if(!carrier){ try{ var log=(typeof getAuditLog==='function')?getAuditLog():[];
    sent4676=log.some(function(e){return e && e.client===key && /4676/i.test(e.action||'');}); }catch(e){} }
  var missing=checks.filter(function(c){return !c.ok;}).map(function(c){return c.item;});
  var present=checks.filter(function(c){return c.ok;}).map(function(c){return c.item;});
  return { client:key, program:(carrier?'carrier (managed care)':'champs'), carrier:(p.carrier||''),
    missing_fields:missing, present_fields:present,
    msa_4676_sent:(carrier?'N/A — managed care':sent4676),
    all_fields_complete: missing.length===0 };
}
// ── Tool: open_email ── resolve recipient, generate/attach a form if asked, open the compose modal.
function _asstOpenEmail(args){
  args=args||{};
  var cname=(args.client_name||'').trim();
  var profs=getProfiles(); var key=cname;
  if(!profs[key]){ var f=Object.keys(profs).find(function(k){return k.toLowerCase()===cname.toLowerCase();}); if(f)key=f; }
  var p=profs[key];
  if(!p)return Promise.resolve({error:'No client named "'+cname+'" was found.'});
  // The MSA-4676 doesn't apply to managed-care (carrier) clients — their authorizations go via the carrier.
  if(p.program==='carrier' && /4676/i.test(args.attach_form||''))
    return Promise.resolve({error:key+' is a managed-care (carrier) client — the MSA-4676 doesn’t apply to them.'});
  var cgs=getCaregivers(), cws=getCaseworkers();
  var recipient=(args.recipient||'').trim(), toEmail='', recipLabel='';
  var _is4676=/4676/i.test(args.attach_form||'');
  // MSA-4676 routing follows the actual two-step workflow:
  //   1. the form is generated UNSIGNED and goes to the CAREGIVER, who signs it;
  //   2. the signed copy is filed in Documents and sent from there to the CASEWORKER
  //      (the ✉ caseworker action on the document card).
  // This tool only ever attaches a freshly GENERATED (unsigned) form, so with no recipient named it
  // is step 1 — the caregiver. 'caseworker' still works when the user asks for it explicitly.
  if(/@/.test(recipient)){ toEmail=recipient; recipLabel='recipient'; }
  else if(/case|worker|coordinator/i.test(recipient)){
    var cw=null;
    if(p.caseworkerId)cw=cws.find(function(c){return c.id===p.caseworkerId;});
    if(!cw&&p.worker)cw=cws.find(function(c){return (c.name||'').toLowerCase()===(p.worker||'').toLowerCase();});
    toEmail=cw?(cw.email||''):''; recipLabel='caseworker'+(cw&&cw.name?(' '+cw.name):'');
    if(!toEmail)return Promise.resolve({error:'No email on file for '+key+'’s caseworker'+(cw&&cw.name?(' ('+cw.name+')'):'')+'. Add it on the client, then try again.'});
  } else {
    var cg=(p.caregiverId&&cgs[p.caregiverId])?cgs[p.caregiverId]:null;
    toEmail=cg?(cg.email||''):''; recipLabel='caregiver'+(cg&&cg.name?(' '+cg.name):'');
    if(!toEmail)return Promise.resolve({error:'No email on file for '+key+'’s caregiver'+(cg&&cg.name?(' ('+cg.name+')'):'')+'. Add it on the client, then try again.'});
  }
  var subject=(args.subject||'').trim(), body=(args.body||'').trim();
  var attach=(args.attach_form||'none');
  var openIt=function(d){ _openDocEmailModal(d,toEmail,subject,body,{auditClient:key});
    var out={opened:true, to:toEmail, recipient:recipLabel, attached:(d?(d.displayName||d.name):'nothing')};
    // Tell the model what happens next so it can say it, rather than inventing a next step.
    if(_is4676&&/^caregiver/.test(recipLabel))
      out.workflow_note='This is the UNSIGNED MSA-4676 going to the caregiver to sign. Once they return it signed, file it on the client’s Documents tab and use the ✉ caseworker action there to send it on to the caseworker.';
    return out; };
  if(attach&&attach!=='none'){
    var ft=/4676/.test(attach)?'msa4676':null;
    if(!ft)return Promise.resolve({error:'I can only attach the MSA-4676 right now.'});
    return buildStateFormBytes(ft,key).then(function(bytes){
      var url=URL.createObjectURL(new Blob([bytes],{type:'application/pdf'}));
      var fname='MSA-4676_'+key.replace(/[^a-z0-9]/gi,'_')+'.pdf';
      return openIt({displayName:fname,name:fname,url:url});
    }).catch(function(e){ return {error:'Could not generate the MSA-4676: '+((e&&e.message)||e)}; });
  }
  return Promise.resolve(openIt(null));
}

// ── Tool: find_caregiver ── resolve a CAREGIVER by name → contact + the clients they serve (with
// live-in status). Live-in is stored on the CLIENT, so this is how "is <caregiver> live-in" is answered.
function _asstFindCaregiver(args){
  var q=((args&&args.name)||'').trim().toLowerCase();
  var toks=q.split(/\s+/).filter(Boolean);
  var cgs=getCaregivers(), profs=getProfiles();
  var out=[];
  Object.keys(cgs).forEach(function(id){
    var cg=cgs[id]; var nm=(cg.name||((cg.firstName||'')+' '+(cg.lastName||''))).trim();
    if(toks.length && !toks.every(function(t){return nm.toLowerCase().indexOf(t)>=0;})) return;
    var clients=Object.keys(profs).filter(function(k){return profs[k].caregiverId===id;}).map(function(k){
      var p=profs[k];
      return { name:k, live_in:(p.liveIn?'yes':'no'), status:(p.clientStatus||'active'),
        program:(p.program==='carrier'?'carrier':'champs') };
    });
    out.push({ caregiver_name:nm, email:cg.email||'', phone:cg.phone||'', champs_id:cg.champsId||cg.champs_id||'',
      client_count:clients.length, any_client_live_in:(clients.some(function(c){return c.live_in==='yes';})?'yes':'no'),
      clients:clients });
  });
  var capped=out.slice(0,8);
  return { matches:capped, count:out.length, truncated: out.length>capped.length };
}
// The billing period that can actually be "missing": the PREVIOUS month. A month's invoice is
// generated on the 1st of the NEXT month (see renderAttentionPanel), so the CURRENT month is never
// overdue — defaulting to it reported nearly every active client as unbilled.
function _asstPrevPeriod(){
  var t=today().split('/'); var d=new Date(+t[2], (+t[0])-2, 1);
  return String(d.getMonth()+1).padStart(2,'0')+'/'+d.getFullYear();
}
// Normalize a billing period to MM/YYYY from "MM/YYYY", "YYYY-MM", a month name, or "this month".
function _asstBillingPeriodNorm(v){
  v=String(v||'').trim();
  var tParts=today().split('/');
  if(/last month|previous month|prior month/i.test(v)) return _asstPrevPeriod();
  if(!v||/this month|current/i.test(v)) return tParts[0]+'/'+tParts[2];
  var m=v.match(/^(\d{1,2})\/(\d{4})$/); if(m) return m[1].padStart(2,'0')+'/'+m[2];
  m=v.match(/^(\d{4})-(\d{1,2})$/); if(m) return m[2].padStart(2,'0')+'/'+m[1];
  var mn=_asstMonthNum(v); if(mn){ var yr=(v.match(/(\d{4})/)||[])[1]||tParts[2]; return String(mn).padStart(2,'0')+'/'+yr; }
  return v;
}
// ── Tool: query_billing ── invoice questions, computed in code. Carriers excluded (billed elsewhere).
function _asstQueryBilling(args){
  args=args||{};
  var what=String(args.what||'invoices').toLowerCase();
  var status=String(args.status||'').toLowerCase();
  var period=args.period?_asstBillingPeriodNorm(args.period):'';
  var profs=getProfiles();
  if(what==='client'){
    var cname=(args.client_name||'').trim(); var key=cname;
    if(!profs[key]){ var f=Object.keys(profs).find(function(k){return k.toLowerCase()===cname.toLowerCase();}); if(f)key=f; }
    var p=profs[key]; if(!p)return {error:'No client named "'+cname+'" found.'};
    if(isCarrierClient(p))return {client:key, note:'Managed-care (carrier) client — billed through the carrier; no CRM invoices.', invoices:[]};
    var invs=(p.invoices||[]).map(function(i){return {period:i.billingPeriod||'', status:(i.status||'draft')};});
    return {client:key, count:invs.length, invoices:invs};
  }
  if(what==='unbilled'){
    // Default to the PREVIOUS month, matching the attention panel: the current month isn't billed
    // until the 1st of the next month, so defaulting to it flagged nearly every active client.
    var per=period||_asstPrevPeriod(); var missing=[];
    Object.keys(profs).forEach(function(k){ var p=profs[k];
      if(!clientDueForInvoice(p,per))return;   // excludes carriers, inactive, and pre-start-date
      if(!((p.invoices||[]).some(function(i){return i.billingPeriod===per;}))) missing.push(k);
    });
    var mrows=(args.no_cap===true)?missing:missing.slice(0,80);
    return {what:'unbilled', period:per, count:missing.length, clients:mrows, truncated:missing.length>mrows.length};
  }
  // 'invoices': list/count invoices matching period and/or status (carrier clients excluded)
  var rows=[], byStatus={draft:0,submitted:0,paid:0};
  Object.keys(profs).forEach(function(k){ var p=profs[k]; if(isCarrierClient(p))return;
    (p.invoices||[]).forEach(function(i){ var st=(i.status||'draft').toLowerCase();
      if(period && i.billingPeriod!==period)return;
      if(status && st!==status)return;
      byStatus[st]=(byStatus[st]||0)+1;
      rows.push({client:k, period:i.billingPeriod||'', status:st});
    });
  });
  var irows=(args.no_cap===true)?rows:rows.slice(0,80);
  return {what:'invoices', period:period||'any', status:status||'any', count:rows.length,
    by_status:byStatus, invoices:irows, truncated:rows.length>irows.length};
}
// ── Tool: find_caseworker ── caseworker/coordinator by name → org, contact, and clients served.
function _asstFindCaseworker(args){
  var q=((args&&args.name)||'').trim().toLowerCase();
  var toks=q.split(/\s+/).filter(Boolean);
  var cws=getCaseworkers(), profs=getProfiles();
  var out=[];
  cws.forEach(function(cw){
    var nm=(cw.name||'').trim();
    if(toks.length && !toks.every(function(t){return nm.toLowerCase().indexOf(t)>=0;})) return;
    var clients=Object.keys(profs).filter(function(k){return profs[k].caseworkerId===cw.id||(cw.name&&profs[k].worker===cw.name);})
      .map(function(k){return {name:k, status:(profs[k].clientStatus||'active'), program:(profs[k].program==='carrier'?'carrier':'champs')};});
    out.push({ caseworker_name:nm, org:cw.org||'', agency:cw.agency||'', email:cw.email||'', phone:cw.phone||'',
      client_count:clients.length, clients:clients });
  });
  var capped=out.slice(0,8);
  return { matches:capped, count:out.length, truncated:out.length>capped.length };
}
// ── Tool: create_task ── opens the task form pre-filled for the user to review + save (not auto-saved).
function _asstCreateTask(args){
  args=args||{};
  var text=(args.task||args.text||'').trim(); if(!text)return {error:'What should the task say?'};
  var clientName=(args.client_name||'').trim();
  if(clientName){ var profs=getProfiles(); if(!profs[clientName]){ var f=Object.keys(profs).find(function(k){return k.toLowerCase()===clientName.toLowerCase();}); if(f)clientName=f; } }
  if(typeof showTaskEditModal!=='function')return {error:'Task UI is unavailable.'};
  showTaskEditModal({
    title:'Add Task', subtitle:(clientName?('For client: '+clientName):'General task'),
    name:text, due:(args.due||''), note:(args.note||''), okText:'Add Task',
    onSave:function(vals){
      var todos=getTodos();
      todos.unshift({id:todoId(),text:vals.name,client:clientName||'',due:vals.due||'',note:vals.note||'',done:false,created:new Date().toLocaleString()});
      saveTodos(todos); if(typeof saveTaskAPI==='function')saveTaskAPI(todos[0]); if(typeof updateTaskBadge==='function')updateTaskBadge();
      if(typeof renderOverviewPane==='function'){try{renderOverviewPane();}catch(e){}}
      if(typeof logActivity==='function')logActivity('client','Task added'+(clientName?(' for '+clientName):'')+': '+vals.name);
    }
  });
  return { opened:true, note:'Opened a task, pre-filled, for the user to review and save.' };
}
// ── Tool: update_client ── change ONE whitelisted field on a client, AFTER a confirm dialog.
var _ASST_UPDATABLE={ phone:'phone', home_phone:'homePhone', email:'clientEmail', client_email:'clientEmail',
  street:'street', address:'street', city:'city', state:'state', zip:'zip', county:'county',
  medicaid_id:'medicaidId', medicare:'medicare', dob:'dob', hourly_rate:'hourlyRate', start_date:'startDate',
  status:'clientStatus', client_status:'clientStatus', carrier:'carrier', member_id:'memberId', program:'program' };
function _asstUpdateClient(args){
  args=args||{};
  var cname=(args.client_name||'').trim(); var profs=getProfiles(); var key=cname;
  if(!profs[key]){ var f=Object.keys(profs).find(function(k){return k.toLowerCase()===cname.toLowerCase();}); if(f)key=f; }
  var p=profs[key]; if(!p)return {error:'No client named "'+cname+'" found.'};
  var fieldRaw=String(args.field||'').toLowerCase().replace(/\s+/g,'_');
  var prop=_ASST_UPDATABLE[fieldRaw];
  if(!prop)return {error:'I can only update: phone, email, address, city, state, zip, county, medicaid ID, medicare, DOB, hourly rate, start date, status, program, carrier, or member #. (SSN and license must be edited on the client page.)'};
  var value=(args.value==null?'':String(args.value)).trim();
  if(prop==='clientStatus'){
    var sv=value.toLowerCase(); value=(sv==='in progress')?'inactive':sv;
    if(['active','inactive','lost','terminated'].indexOf(value)<0)
      return {error:'Status must be one of: Active, In Progress, Lost, Terminated.'};
    // Same gate the client pane enforces (saveClientInfo): a CHAMPS client can't go Active without a
    // DHS-1210 on file. Going through the assistant used to skip it, and an Active client with no
    // authorization is then picked up by invoice generation.
    if(value==='active' && p.program!=='carrier' && !hasAuthorization(p))
      return {error:key+' has no DHS-1210 authorization on file, so they can\'t be set Active. Import the authorization on their Authorization tab first.'};
    if(value==='active' && !(p.startDate&&String(p.startDate).trim()))
      return {error:key+' has no service start date, which is required before they can be Active.'};
  }
  if(prop==='program'){ value=(/carrier|managed/i.test(value))?'carrier':'champs'; }
  var label=(args.field||prop);
  if(typeof showConfirm!=='function')return {error:'Confirm UI unavailable.'};
  showConfirm('Set '+key+'’s '+label+' to “'+(value||'(blank)')+'”?', function(){
    p[prop]=value; saveProfilesLS(profs); if(typeof saveProfileSP==='function')saveProfileSP(key,p);
    if(typeof addAuditEntry==='function')addAuditEntry(key,'Updated '+label+' via Assistant');
    if(typeof showToast==='function')showToast('✓ Updated '+key+'’s '+label,3000);
    if(activeProfileName===key && typeof renderInfoPane==='function'){try{renderInfoPane();}catch(e){}}
  },{title:'Update Client',okText:'Update'});
  return { opened:true, note:'Asked the user to confirm the change (not applied until they click Update).' };
}
// ── Export helpers ── build + download a file from a list of rows.
function _asstTriggerDownload(blob, fname){
  var url=URL.createObjectURL(blob); var a=document.createElement('a'); a.href=url; a.download=fname;
  document.body.appendChild(a); a.click(); setTimeout(function(){if(a.parentNode)a.parentNode.removeChild(a); URL.revokeObjectURL(url);}, 2000);
}
function _asstDownloadCsv(cols, rows, fname){
  var q=function(v){ v=(v==null?'':String(v)); return /[",\n]/.test(v)?'"'+v.replace(/"/g,'""')+'"':v; };
  var lines=[cols.map(q).join(',')].concat(rows.map(function(r){return r.map(q).join(',');}));
  _asstTriggerDownload(new Blob([lines.join('\r\n')],{type:'text/csv'}), fname);
}
function _asstDownloadXlsx(cols, rows, fname, sheet){
  if(typeof XLSX==='undefined')throw new Error('Excel library not loaded yet — try again in a moment.');
  var ws=XLSX.utils.aoa_to_sheet([cols].concat(rows)); var wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, (sheet||'Export').replace(/[^\w ]/g,'').slice(0,28)||'Export');
  XLSX.writeFile(wb, fname);   // XLSX triggers the download itself
}
async function _asstDownloadPdf(cols, rows, title, fname){
  if(!window.PDFLib)throw new Error('PDF library still loading — try again in a moment.');
  var doc=await PDFLib.PDFDocument.create();
  var font=await doc.embedFont(PDFLib.StandardFonts.Helvetica), bold=await doc.embedFont(PDFLib.StandardFonts.HelveticaBold);
  // Landscape + an 8pt face: the portrait/9pt layout clipped most cells to ~15 characters, which
  // truncated ordinary client names and dates in a file the user keeps as a record.
  var pageW=792, pageH=612, margin=36, size=8, lineH=13, colW=(pageW-2*margin)/cols.length, page=null, y=0;
  var cap=Math.max(8,Math.floor(colW/4.4));
  var header=function(){ page=doc.addPage([pageW,pageH]); y=pageH-margin;
    page.drawText(String(title||'Export'),{x:margin,y:y,size:13,font:bold}); y-=22;
    cols.forEach(function(c,i){ page.drawText(String(c),{x:margin+i*colW,y:y,size:size,font:bold}); }); y-=lineH; };
  header();
  rows.forEach(function(r){ if(y<margin+lineH)header();
    r.forEach(function(v,i){ var s=(v==null?'':String(v)); if(s.length>cap)s=s.slice(0,cap-1)+'…';
      page.drawText(s,{x:margin+i*colW,y:y,size:size,font:font}); }); y-=lineH; });
  _asstTriggerDownload(new Blob([await doc.save()],{type:'application/pdf'}), fname);
}
// ── Tool: export_data ── turn a roster/billing list into a downloadable Excel/CSV/PDF.
function _asstExportData(args){
  args=args||{};
  var source=String(args.source||'clients').toLowerCase();
  var format=String(args.format||'excel').toLowerCase();
  var cols=[], rows=[], title='';
  if(source==='invoices'||source==='billing'){
    var r=_asstQueryBilling({what:'invoices', status:args.status, period:args.period, no_cap:true});
    cols=['Client','Period','Status']; rows=(r.invoices||[]).map(function(i){return [i.client,i.period,i.status];});
    title='Invoices'+(args.period?(' '+args.period):'')+(args.status?(' - '+args.status):'');
  } else if(source==='unbilled'){
    var u=_asstQueryBilling({what:'unbilled', period:args.period, no_cap:true});
    cols=['Client']; rows=(u.clients||[]).map(function(n){return [n];}); title='Unbilled '+(u.period||'');
  } else {
    var q=_asstQueryRoster({action:'list', filters:(Array.isArray(args.filters)?args.filters:[]), no_cap:true});
    cols=['Name','Status','County','Caregiver','Caseworker','Medicaid ID','Start Date'];
    rows=(q.clients||[]).map(function(c){return [c.name,c.status,c.county,c.caregiver,c.caseworker,c.medicaid_id,c.start_date];});
    title='Clients';
  }
  if(!rows.length)return {error:'Nothing to export — the query returned no rows.'};
  var fname=(title.replace(/[^a-z0-9]+/gi,'_').replace(/^_|_$/g,'')||'export')+'_'+today().replace(/\//g,'-');
  try{
    if(format==='csv'){ _asstDownloadCsv(cols,rows,fname+'.csv'); }
    else if(format==='pdf'){ return _asstDownloadPdf(cols,rows,title,fname+'.pdf').then(function(){return {exported:true,format:'pdf',count:rows.length,file:fname+'.pdf'};},function(e){return {error:'PDF export failed: '+((e&&e.message)||e)};}); }
    else { _asstDownloadXlsx(cols,rows,fname+'.xlsx',title); format='excel'; }
  }catch(e){ return {error:'Export failed: '+((e&&e.message)||e)}; }
  return {exported:true, format:format, count:rows.length, file:fname};
}
// ── Tool: roster_digest ── a "what needs my attention" scan: deterministic data-gap + billing checks.
function _asstRosterDigest(){
  var profs=getProfiles(), cgs=getCaregivers(), cws=getCaseworkers();
  var period=_asstPrevPeriod();   // the month that can actually be overdue (see _asstPrevPeriod)
  var active=0, champs=0, carrier=0;
  var missingMedicaid=[], noCaregiver=[], noCaseworker=[], noCwEmail=[], unbilled=[], carrierGaps=[], missingDob=[];
  Object.keys(profs).forEach(function(k){ var p=profs[k];
    var st=(p.clientStatus||'active'); var isActive=(st==='active'); if(isActive)active++;
    var isCarrier=(p.program==='carrier'); if(isCarrier)carrier++; else champs++;
    var cg=(p.caregiverId&&cgs[p.caregiverId])?cgs[p.caregiverId]:null;
    if(!cg && isActive) noCaregiver.push(k);
    var cwRec=p.caseworkerId?cws.find(function(c){return c.id===p.caseworkerId;}):null;
    var cwn=cwRec?cwRec.name:(p.worker||'');
    if(!cwn && isActive) noCaseworker.push(k);
    if(isActive && !p.dob) missingDob.push(k);
    if(!isCarrier){
      if(isActive && !(p.medicaidId&&String(p.medicaidId).trim())) missingMedicaid.push(k);
      if(isActive && cwRec && !(cwRec.email||'')) noCwEmail.push(k);
      if(clientDueForInvoice(p,period) && !((p.invoices||[]).some(function(i){return i.billingPeriod===period;}))) unbilled.push(k);
    } else if(isActive && (!(p.carrier&&String(p.carrier).trim()) || !(p.memberId&&String(p.memberId).trim()))){
      carrierGaps.push(k);
    }
  });
  var g=function(a){return {count:a.length, clients:a.slice(0,25), truncated:a.length>25};};
  return { period:period, active_clients:active, champs_clients:champs, carrier_clients:carrier,
    missing_medicaid_id:g(missingMedicaid), no_caregiver:g(noCaregiver), no_caseworker_or_coordinator:g(noCaseworker),
    caseworker_missing_email:g(noCwEmail), unbilled_this_period:g(unbilled),
    carrier_missing_carrier_or_member:g(carrierGaps), missing_dob:g(missingDob) };
}
function _asstRunTool(name,args){
  try{
    if(name==='find_client')return Promise.resolve(_asstFindClient(args));
    if(name==='query_roster')return Promise.resolve(_asstQueryRoster(args));
    if(name==='roster_digest')return Promise.resolve(_asstRosterDigest());
    if(name==='find_caregiver')return Promise.resolve(_asstFindCaregiver(args));
    if(name==='find_caseworker')return Promise.resolve(_asstFindCaseworker(args));
    if(name==='query_billing')return Promise.resolve(_asstQueryBilling(args));
    if(name==='onboarding_status')return Promise.resolve(_asstOnboardingStatus(args));
    if(name==='create_task')return Promise.resolve(_asstCreateTask(args));
    if(name==='update_client')return Promise.resolve(_asstUpdateClient(args));
    if(name==='export_data')return Promise.resolve(_asstExportData(args));
    if(name==='open_email')return Promise.resolve(_asstOpenEmail(args));
    return Promise.resolve({error:'Unknown tool: '+name});
  }catch(e){ return Promise.resolve({error:(e&&e.message)||String(e)}); }
}

// ── The chat loop: send → model → (run tools → loop) → show reply ──
async function _asstSend(text){
  text=(text||'').trim(); if(!text||_asstBusy)return;
  if(typeof spToken==='undefined'||!spToken){ _asstRenderMsg('assistant','Please sign in first (Settings) so I can look up clients and send email.'); return; }
  // The backend rejects a conversation over 40 turns (ai-chat.js), and this array only ever grew —
  // so after ~10 tool-using questions the assistant answered every message with "Conversation too
  // long" and the only way out was a page reload. Keep the recent window instead, always starting on
  // a user turn so the tool_call/tool_result pairing stays valid.
  if(_asstMessages.length>28){
    var _cut=_asstMessages.length-24;
    while(_cut<_asstMessages.length && _asstMessages[_cut].role!=='user')_cut++;
    if(_cut<_asstMessages.length)_asstMessages=_asstMessages.slice(_cut);
  }
  _asstMessages.push({role:'user',content:text});
  _asstRenderMsg('user',text);
  _asstBusy=true; _asstSetTyping(true);
  try{
    var guard=0, answered=false;
    while(guard++<6){
      var resp=await fetch(API_BASE+'/ai-chat',{method:'POST',headers:apiHeaders(),
        body:JSON.stringify({messages:_asstMessages,context:_asstContext()})});
      if(!resp.ok){ var je=await resp.json().catch(function(){return{};}); throw new Error(je.error||('HTTP '+resp.status)); }
      var msg=(await resp.json()).message;
      _asstMessages.push(msg);
      if(msg.tool_calls&&msg.tool_calls.length){
        if(msg.content)_asstRenderMsg('assistant',msg.content);
        for(var i=0;i<msg.tool_calls.length;i++){
          var tc=msg.tool_calls[i], a={}; try{a=JSON.parse(tc.function.arguments||'{}');}catch(e){a={};}
          _asstToolNote(tc.function.name,a);
          var result=await _asstRunTool(tc.function.name,a);
          _asstMessages.push({role:'tool',tool_call_id:tc.id,name:tc.function.name,content:JSON.stringify(result)});
        }
        continue;
      }
      _asstRenderMsg('assistant',msg.content||'(no response)');
      answered=true;
      break;
    }
    // The loop can hit its turn limit while the model is still calling tools — say so rather than
    // leaving the user staring at a typing indicator that resolved into nothing.
    if(!answered)_asstRenderMsg('assistant','I ran several steps but didn’t reach a final answer. Try asking for one thing at a time, or narrowing the question.');
  }catch(e){ _asstRenderMsg('assistant','⚠️ '+((e&&e.message)||e)); }
  finally{ _asstBusy=false; _asstSetTyping(false); }
}

// ── UI ──
function _asstScroll(){ var m=document.getElementById('asst-msgs'); if(m)m.scrollTop=m.scrollHeight; }
function _asstRenderMsg(role,text){
  var m=document.getElementById('asst-msgs'); if(!m)return;
  var b=document.createElement('div'); b.className='asst-msg asst-'+role;
  b.innerHTML=esc(text||'').replace(/\n/g,'<br>');
  m.appendChild(b); _asstScroll();
}
function _asstToolNote(name,args){
  var m=document.getElementById('asst-msgs'); if(!m)return;
  var label=name==='find_client'?('Looking up '+((args&&args.name)||'client')+'…')
    :name==='query_roster'?'Checking the roster…'
    :name==='roster_digest'?'Scanning for anything that needs attention…'
    :name==='find_caregiver'?('Looking up caregiver '+((args&&args.name)||'')+'…')
    :name==='find_caseworker'?('Looking up caseworker '+((args&&args.name)||'')+'…')
    :name==='query_billing'?'Checking invoices…'
    :name==='create_task'?'Setting up a task…'
    :name==='update_client'?('Preparing an update'+((args&&args.client_name)?(' for '+args.client_name):'')+'…')
    :name==='export_data'?'Building your file…'
    :name==='onboarding_status'?('Checking onboarding'+((args&&args.client_name)?(' for '+args.client_name):'')+'…')
    :name==='open_email'?('Preparing an email'+((args&&args.client_name)?(' for '+args.client_name):'')+'…')
    :('Working…');
  var b=document.createElement('div'); b.className='asst-tool'; b.textContent='⚙ '+label;
  m.appendChild(b); _asstScroll();
}
function _asstSetTyping(on){
  var m=document.getElementById('asst-msgs'); if(!m)return;
  var ex=document.getElementById('asst-typing');
  if(on){ if(!ex){ var t=document.createElement('div'); t.id='asst-typing'; t.className='asst-msg asst-assistant asst-typing'; t.innerHTML='<span></span><span></span><span></span>'; m.appendChild(t); _asstScroll(); } }
  else if(ex){ ex.parentNode.removeChild(ex); }
  var send=document.getElementById('asst-send'); if(send)send.disabled=on;
}
function _asstToggle(){
  var p=document.getElementById('asst-panel'); if(!p)return;
  _asstOpen=!_asstOpen; p.classList.toggle('open',_asstOpen);
  if(_asstOpen){ var i=document.getElementById('asst-input'); if(i)setTimeout(function(){i.focus();},60); }
}
function _asstInit(){
  try{
    if(typeof document==='undefined'||!document.body||document.getElementById('asst-fab'))return;
    var css=''
      +'#asst-fab{position:fixed;right:20px;bottom:20px;z-index:9998;width:56px;height:56px;border:none;border-radius:50%;'
      +'background:linear-gradient(135deg,#2b6fb3,#1a4e86);color:#fff;font-size:24px;cursor:pointer;box-shadow:0 6px 20px rgba(20,60,110,.35);transition:transform .15s;}'
      +'#asst-fab:hover{transform:scale(1.07);}'
      +'#asst-panel{position:fixed;right:20px;bottom:88px;z-index:9999;width:390px;max-width:calc(100vw - 32px);height:560px;max-height:calc(100vh - 120px);'
      +'background:#fff;border:1px solid #d7e2ef;border-radius:14px;box-shadow:0 18px 50px rgba(20,50,90,.28);display:flex;flex-direction:column;overflow:hidden;'
      +'opacity:0;pointer-events:none;transform:translateY(12px);transition:opacity .18s,transform .18s;}'
      +'#asst-panel.open{opacity:1;pointer-events:auto;transform:translateY(0);}'
      +'.asst-head{background:linear-gradient(135deg,#2b6fb3,#1a4e86);color:#fff;padding:12px 14px;display:flex;align-items:center;gap:8px;}'
      +'.asst-head b{font-size:15px;} .asst-head .asst-sub{font-size:11px;opacity:.85;font-weight:400;}'
      +'.asst-x{margin-left:auto;background:none;border:none;color:#fff;font-size:20px;cursor:pointer;opacity:.85;line-height:1;}'
      +'#asst-msgs{flex:1;overflow-y:auto;padding:14px;background:#f6f9fd;display:flex;flex-direction:column;gap:8px;}'
      +'.asst-msg{max-width:85%;padding:9px 12px;border-radius:12px;font-size:13.5px;line-height:1.45;white-space:normal;word-wrap:break-word;}'
      +'.asst-user{align-self:flex-end;background:#2b6fb3;color:#fff;border-bottom-right-radius:4px;}'
      +'.asst-assistant{align-self:flex-start;background:#fff;border:1px solid #dbe6f2;color:#213547;border-bottom-left-radius:4px;}'
      +'.asst-tool{align-self:flex-start;font-size:11.5px;color:#6a86a3;font-style:italic;padding:1px 4px;}'
      +'.asst-typing{display:flex;gap:4px;align-items:center;} .asst-typing span{width:6px;height:6px;border-radius:50%;background:#9db4cd;animation:asstb 1s infinite;}'
      +'.asst-typing span:nth-child(2){animation-delay:.2s;} .asst-typing span:nth-child(3){animation-delay:.4s;}'
      +'@keyframes asstb{0%,60%,100%{opacity:.3;}30%{opacity:1;}}'
      +'.asst-chips{display:flex;flex-wrap:wrap;gap:6px;padding:0 14px 8px;background:#f6f9fd;}'
      +'.asst-chip{font-size:11.5px;border:1px solid #cfe0f2;background:#fff;color:#2b5f96;border-radius:14px;padding:4px 9px;cursor:pointer;}'
      +'.asst-chip:hover{background:#eaf3fc;}'
      +'.asst-inbar{display:flex;gap:8px;padding:10px;border-top:1px solid #e3ecf6;background:#fff;}'
      +'#asst-input{flex:1;resize:none;border:1px solid #cdd9e8;border-radius:10px;padding:9px 11px;font-size:13.5px;font-family:inherit;max-height:96px;}'
      +'#asst-send{border:none;background:#2b6fb3;color:#fff;border-radius:10px;padding:0 15px;font-size:14px;cursor:pointer;}'
      +'#asst-send:disabled{opacity:.5;cursor:default;}'
      +'@media(max-width:480px){#asst-panel{right:8px;left:8px;width:auto;bottom:80px;height:calc(100vh - 100px);}}';
    var st=document.createElement('style'); st.textContent=css; document.head.appendChild(st);

    var fab=document.createElement('button'); fab.id='asst-fab'; fab.type='button'; fab.title='Assistant'; fab.textContent='✨';
    var ag=getAgencyInfo()||{}; var first=((ag.agency_provider_name||'').split(' ')[0])||'there';
    var panel=document.createElement('div'); panel.id='asst-panel';
    panel.innerHTML=''
      +'<div class="asst-head"><b>✨ Assistant</b><span class="asst-sub">looks up clients · drafts forms · reports</span>'
      +'<button class="asst-x" id="asst-close" title="Close">×</button></div>'
      +'<div id="asst-msgs"></div>'
      +'<div class="asst-chips">'
      +'<button class="asst-chip" data-q="Send the MSA-4676 for ">Send a 4676…</button>'
      +'<button class="asst-chip" data-q="How many active clients do I have?">Active clients?</button>'
      +'<button class="asst-chip" data-q="Which clients have no caregiver assigned?">No caregiver?</button>'
      +'</div>'
      +'<div class="asst-inbar"><textarea id="asst-input" rows="1" placeholder="Ask me anything… e.g. “send the 4676 for Jane Doe to her caseworker”"></textarea>'
      +'<button id="asst-send" type="button">Send</button></div>';
    document.body.appendChild(fab); document.body.appendChild(panel);

    _asstRenderMsg('assistant','Hi '+first+'! I can pull up a client, generate & send their forms (like the MSA-4676), or answer quick roster questions. What do you need?');

    fab.addEventListener('click',_asstToggle);
    document.getElementById('asst-close').addEventListener('click',_asstToggle);
    var input=document.getElementById('asst-input'), send=document.getElementById('asst-send');
    var submit=function(){ var v=input.value; input.value=''; input.style.height='auto'; _asstSend(v); };
    send.addEventListener('click',submit);
    input.addEventListener('keydown',function(e){ if(e.key==='Enter'&&!e.shiftKey){ e.preventDefault(); submit(); } });
    input.addEventListener('input',function(){ input.style.height='auto'; input.style.height=Math.min(input.scrollHeight,96)+'px'; });
    Array.prototype.forEach.call(panel.querySelectorAll('.asst-chip'),function(c){
      c.addEventListener('click',function(){ input.value=c.getAttribute('data-q'); input.focus(); input.dispatchEvent(new Event('input')); });
    });
  }catch(e){ if(typeof console!=='undefined')console.warn('[Assistant] init skipped:',e&&e.message); }
}
// app.js runs at the end of <body>, so the DOM is ready; init immediately (guarded for the jsdom tests).
_asstInit();
