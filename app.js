// ============================================================
//  STATE
// ============================================================
var SVC=15,CPLX=9,active=31,clipboard={};
var activeProfileName=null,lastLoadedStates=null;
var pendingSigTarget=null,sigDrawing=false,sigCanvas=null,sigCtx=null;
var draftTimer=null;
var unsavedChanges=false;

// ============================================================
//  AZURE FUNCTIONS API CONFIG
// ============================================================
var API_BASE    = 'https://liberty-crm-api-cyb3dkhnd2e7a3cy.centralus-01.azurewebsites.net/api';
var API_APP_ID  = '0c1627c1-c186-4e46-b919-e4a12f2f3952'; // Easy Auth app registration
var _apiToken   = null; // cached Bearer token, refreshed automatically

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
var SP_DOC_LIB   = 'ClientDocuments'; // kept for SharePoint doc library (if used)
var SP_SITE      = 'https://libertybellhealth.sharepoint.com/sites/Invoice'; // kept for doc upload
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
    setTimeout(refreshApiToken, Math.max(ttl, 60000));
  } catch(e) {
    console.warn('API token silent refresh failed, falling back to redirect:', e);
    // Redirect works on mobile (popup is blocked on iOS Safari)
    try {
      await msalInstance.acquireTokenRedirect({ scopes: [API_SCOPE] });
    } catch(e2) { console.warn('API token redirect failed:', e2); }
  }
}
function getIdMap() {
  try { return JSON.parse(localStorage.getItem('lhca_id_map') || '{}'); } catch(e) { return {}; }
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
function navHome(){showPage('home');bc([{l:'Clients'}]);document.getElementById('topbarActions').innerHTML='';unsavedChanges=false;renderClientTable();updateStats();renderSidebarClients();}

// ============================================================
//  HASH ROUTER — enables right-click "Open in new tab" for records
//  URL format: #/client/<name>, #/caregiver/<id>, #/caseworker/<id>,
//              #/forms, #/tasks, #/caregivers, #/caseworkers, #/settings
// ============================================================
function buildClientUrl(name){return '#/client/'+encodeURIComponent(name);}
function buildCaregiverUrl(id){return '#/caregiver/'+encodeURIComponent(id);}
function buildCaseworkerUrl(id){return '#/caseworker/'+encodeURIComponent(id);}

function routeFromHash(){
  var hash=(window.location.hash||'').replace(/^#\/?/, '');
  if(!hash){if(typeof navHome==='function')navHome();return;}
  var parts=hash.split('/').map(decodeURIComponent);
  var route=parts[0];
  try{
    if(route==='client'&&parts[1]){
      activeProfileName=parts[1];
      if(typeof navDetail==='function')navDetail(parts[1],parts[2]||null);
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

// ============================================================
//  GLOBAL SEARCH
// ============================================================
function runGlobalSearch(q){
  var res=document.getElementById('globalResults');if(!q||q.length<2){res.style.display='none';return;}
  q=q.toLowerCase();
  var profiles=getProfiles(),cgs=getCaregivers(),todos=getTodos();
  var sections=[];
  // Clients
  var clientHits=Object.keys(profiles).filter(function(n){return n.toLowerCase().includes(q)||(profiles[n].medicaidId||'').toLowerCase().includes(q);}).slice(0,5);
  if(clientHits.length){
    sections.push('<div class="gr-section">Clients</div>');
    clientHits.forEach(function(n){sections.push('<div class="gr-item" onclick="closeGlobalSearch();navDetail(\''+esc(n)+'\')"><span class="gr-label">'+esc(n)+'</span><span class="gr-meta">'+(profiles[n].medicaidId||'')+'</span></div>');});
  }
  // Invoices
  var invHits=[];
  Object.keys(profiles).forEach(function(name){(profiles[name].invoices||[]).forEach(function(inv){if((inv.billingPeriod||'').includes(q.toUpperCase())||name.toLowerCase().includes(q))invHits.push({name:name,inv:inv});});});
  invHits=invHits.slice(0,4);
  if(invHits.length){
    sections.push('<div class="gr-section">Invoices</div>');
    invHits.forEach(function(r){sections.push('<div class="gr-item" onclick="closeGlobalSearch();navDetail(\''+esc(r.name)+'\',\'history\')"><span class="gr-label">'+esc(r.inv.billingPeriod)+' — '+esc(r.name)+'</span><span class="gr-meta">'+(r.inv.status||'draft')+'</span></div>');});
  }
  // Tasks
  var taskHits=todos.filter(function(t){return !t.done&&(t.text||'').toLowerCase().includes(q);}).slice(0,3);
  if(taskHits.length){
    sections.push('<div class="gr-section">Tasks</div>');
    taskHits.forEach(function(t){sections.push('<div class="gr-item" onclick="closeGlobalSearch();navTasks()"><span class="gr-label">'+esc(t.text)+'</span><span class="gr-meta">'+(t.client?t.client:'')+(t.due?' · '+t.due:'')+'</span></div>');});
  }
  // Caregivers
  var cgHits=Object.keys(cgs).filter(function(id){return (cgs[id].name||'').toLowerCase().includes(q);}).slice(0,3);
  if(cgHits.length){
    sections.push('<div class="gr-section">Caregivers</div>');
    cgHits.forEach(function(id){sections.push('<div class="gr-item" onclick="closeGlobalSearch();navCaregivers()"><span class="gr-label">'+esc(cgs[id].name)+'</span><span class="gr-meta">'+esc(cgs[id].status||'')+'</span></div>');});
  }
  if(!sections.length)sections.push('<div class="gr-empty">No results for "'+esc(q)+'"</div>');
  res.innerHTML=sections.join('');res.style.display='block';
}
function closeGlobalSearch(){
  var gs=document.getElementById('globalSearch');if(gs)gs.value='';
  var res=document.getElementById('globalResults');if(res)res.style.display='none';
}
document.addEventListener('click',function(e){if(!e.target.closest('.global-search-wrap'))closeGlobalSearch();});

// ============================================================
//  PRINT & EMAIL INVOICE
// ============================================================
function printThenEmail(){
  var prof=getProfiles()[activeProfileName]||{};
  var cwRec3=getCaseworkers().find(function(c){return c.id===prof.caseworkerId||c.name===prof.worker;})||{};
  var agentEmail=cwRec3.email||document.getElementById('activeAgentEmail').value||'';
  var clientName=document.getElementById('clientName').value||activeProfileName||'Client';
  var bp=document.getElementById('billingPeriod').value||'';
  // Step 1: print dialog (user saves as PDF)
  var overlay=document.createElement('div');
  overlay.style.cssText='position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.55);z-index:9999;display:flex;align-items:center;justify-content:center;';
  overlay.innerHTML='<div style="background:#fff;border-radius:8px;padding:24px 28px;max-width:400px;text-align:center;box-shadow:0 8px 32px rgba(0,0,0,0.2);">'+
    '<div style="font-size:15px;font-weight:700;color:#1a2b45;margin-bottom:8px;">Step 1 of 2 — Save as PDF</div>'+
    '<div style="font-size:13px;color:#4a6a8a;line-height:1.6;margin-bottom:16px;">The print dialog will open. Choose <strong>Save as PDF</strong> as your printer, then click OK below to open the email composer.</div>'+
    '<button onclick="document.body.removeChild(this.closest(\'div\').parentNode.parentNode);_openEmailAfterPrint(\''+esc(agentEmail)+'\',\''+esc(clientName)+'\',\''+esc(bp)+'\')" style="background:#185FA5;color:#fff;border:none;border-radius:5px;padding:9px 20px;font-size:13px;font-family:Arial,sans-serif;font-weight:600;cursor:pointer;margin-right:8px;">OK — Open Email Composer</button>'+
    '<button onclick="document.body.removeChild(this.closest(\'div\').parentNode.parentNode)" style="background:#f0f3f7;color:#1a2b45;border:none;border-radius:5px;padding:9px 14px;font-size:13px;font-family:Arial,sans-serif;cursor:pointer;">Cancel</button>'+
    '</div>';
  document.body.appendChild(overlay);
  window.print();
}
function _openEmailAfterPrint(toAddr,clientName,bp){
  var subj='Home Help Invoice'+(clientName?' — '+clientName:'')+(bp?' — '+bp:'');
  var body='Dear Caseworker,\n\nPlease find the attached Home Help Agency Invoice'+(clientName?' for '+clientName:'')+(bp?' for the billing period '+bp:'')+'.\n\nPlease review and process at your earliest convenience. Do not hesitate to contact us with any questions.\n\nThank you,\nThomas Jaboro\nLiberty Home Care Assistance\n(248) 291-4106';
  openEmailComposer(toAddr,subj,body);
}
function navNewClient(){
  showPage('new-client');bc([{l:'Clients',fn:navHome},{l:'New Client'}]);document.getElementById('topbarActions').innerHTML='';
  ['nc-first','nc-middle','nc-last','nc-nickname','nc-medicaid','nc-rate','nc-dl','nc-ssn','nc-phone','nc-cemail','nc-street','nc-city','nc-state','nc-zip','nc-county','nc-start-date','nc-worker-search','nc-worker-val'].forEach(function(id){var e=document.getElementById(id);if(e)e.value='';});
  var drop=document.getElementById('nc-worker-drop');if(drop)drop.style.display='none';
  populateCaregiverSelect('nc-caregiver','');
}
function navDetail(name,tab){
  aiTrack('ClientRecordOpened',{client:name,tab:tab||'info'});
  activeProfileName=name;var prof=getProfiles()[name];if(!prof)return;
  showPage('detail');
  var ini=name.split(' ').map(function(w){return w[0]||'';}).slice(0,2).join('').toUpperCase();
  document.getElementById('detailAvatar').textContent=ini;
  document.getElementById('detailName').textContent=name+(prof.nickname?' ('+prof.nickname+')':'');
  var st=prof.clientStatus||'active';
  document.getElementById('detailMeta').innerHTML=(prof.medicaidId?'Medicaid: '+prof.medicaidId:'No Medicaid ID')+(prof.phone?' &nbsp;·&nbsp; '+prof.phone:'')+
    ' &nbsp;<span class="cs-badge cs-'+st+'">'+st.charAt(0).toUpperCase()+st.slice(1)+'</span>';
  bc([{l:'Clients',fn:navHome},{l:name}]);
  document.getElementById('topbarActions').innerHTML='';
  switchTab(tab||'overview');renderSidebarClients();
}
function navInvoice(loadSpecific){
  if(!activeProfileName){showAlert('Select a client first.');return;}
  var prof=getProfiles()[activeProfileName];
  if(loadSpecific){
    // Opening a saved invoice file
    showPage('invoice');
    document.getElementById('invClientTag').textContent=activeProfileName;
    document.getElementById('saveInvoiceBtn').style.display='inline-block';
    bc([{l:'Clients',fn:navHome},{l:activeProfileName,fn:function(){navDetail(activeProfileName);}},{l:'Invoice'}]);
    document.getElementById('topbarActions').innerHTML='';
    document.getElementById('dupWarning').style.display='none';
    applyFullInvoice(loadSpecific);
    syncBillingPeriodFields();
    if(draftTimer)clearInterval(draftTimer);
    return;
  }
  // Show choice modal — New or Copy from last?
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
  document.getElementById('newInvChoiceModal').classList.remove('open');
  var prof=getProfiles()[activeProfileName];
  showPage('invoice');
  document.getElementById('invClientTag').textContent=activeProfileName;
  document.getElementById('saveInvoiceBtn').style.display='inline-block';
  bc([{l:'Clients',fn:navHome},{l:activeProfileName,fn:function(){navDetail(activeProfileName);}},{l:'Invoice'}]);
  document.getElementById('topbarActions').innerHTML='';
  document.getElementById('dupWarning').style.display='none';
  // Kill any stale draft
  try{localStorage.removeItem('lhca_draft_'+activeProfileName);}catch(e){}
  if(draftTimer)clearInterval(draftTimer);
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
    document.getElementById('billingPeriod').value='';
    document.getElementById('billingPeriod2').value='';
    var T=today();
    document.getElementById('dateSubmitted').value=T;
    document.getElementById('sigDate1').value=T;
    document.getElementById('sigDate2').value=T;
    ['svcHH','svcMM','cplxHH','cplxMM','p1HH','p1MM','grandHH','grandMM'].forEach(function(id){var e=document.getElementById(id);if(e)e.value='';});
    rebuild(31);
    resetSigArea(1);resetSigArea(2);
  }
  startDraftAutosave();
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
}
function navSettings(){showPage('settings');bc([{l:'Settings'}]);document.getElementById('topbarActions').innerHTML='';renderSigSettings();updateSettingsAuth();renderEmailAuditTable();if(typeof loadSigningTemplates==='function')loadSigningTemplates();}
function navTasks(){showPage('tasks');bc([{l:'Tasks'}]);document.getElementById('topbarActions').innerHTML='';populateTodoClientSelect();renderTodos();}
function navReports(){showPage('reports');bc([{l:'Reports'}]);document.getElementById('topbarActions').innerHTML='';renderReports();}

// ============================================================
//  SIDEBAR
// ============================================================
function toggleSidebar(){document.getElementById('sidebar').classList.toggle('collapsed');}
function renderSidebarClients(){
  var list=document.getElementById('sbClientList'),profiles=getProfiles(),keys=Object.keys(profiles);
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
    row.innerHTML='<div class="sb-avatar">'+ini+'</div><div class="sb-client-info"><div class="sb-client-name">'+esc(sbDisplay)+'</div><div class="sb-client-meta">'+(profiles[name].medicaidId||'No ID')+'</div></div>';
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
  keys.forEach(function(k){var invs=(p[k].invoices)||[];ti+=invs.length;invs.forEach(function(inv){if(!inv.status||inv.status==='draft'||inv.status==='submitted')outstanding++;});});
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
function renderAttentionPanel(){
  var panel=document.getElementById('attentionPanel');if(!panel)return;
  var p=getProfiles(),items=[];
  var d=new Date();
  // Flag PREVIOUS month — current month is generated on the 1st of next month, so it's not "missing" yet
  var prev=new Date(d.getFullYear(),d.getMonth()-1,1);
  var prevPeriod=String(prev.getMonth()+1).padStart(2,'0')+'/'+prev.getFullYear();
  var active=Object.keys(p).filter(function(k){return !p[k].clientStatus||p[k].clientStatus==='active';});
  var missingPrev=active.filter(function(k){return clientWasActiveInPeriod(p[k],prevPeriod) && !((p[k].invoices)||[]).some(function(i){return i.billingPeriod===prevPeriod;});});
  if(missingPrev.length){
    items.push({cls:'attn-warn',count:missingPrev.length,label:missingPrev.length+' active client'+(missingPrev.length>1?'s':'')+' missing invoice for '+prevPeriod+' — click 🔄 Generate Invoices to create',fn:'openGenerateInvoicesModal()'});
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
function renderClientTable(forceStatus){
  var profiles=getProfiles();
  var q=((document.getElementById('clientSearch')?document.getElementById('clientSearch').value:'')||'').toLowerCase();
  var filterStatus=forceStatus||(document.getElementById('filterStatus')?document.getElementById('filterStatus').value:'active');
  var filterCg=(document.getElementById('filterCaregiver')&&document.getElementById('filterCaregiver').value)||'';
  var sortBy=(document.getElementById('filterSort')&&document.getElementById('filterSort').value)||'name';
  var outstandingOnly=document.getElementById('filterOutstanding')&&document.getElementById('filterOutstanding').checked;
  var cgs=getCaregivers();
  var keys=Object.keys(profiles).filter(function(k){
    var st=profiles[k].clientStatus||'active';
    var matchStatus=filterStatus==='all'||st===filterStatus;
    var matchQ=!q||k.toLowerCase().includes(q)||(profiles[k].medicaidId||'').toLowerCase().includes(q);
    var matchCg=!filterCg||profiles[k].caregiverId===filterCg;
    var matchOut=!outstandingOnly||(profiles[k].invoices||[]).some(function(i){return !i.status||i.status==='draft'||i.status==='submitted';});
    return matchStatus&&matchQ&&matchCg&&matchOut;
  });
  if(sortBy==='name')keys.sort(function(a,b){return a.localeCompare(b);});
  else if(sortBy==='recent')keys.sort(function(a,b){
    var la=(profiles[a].invoices||[])[0],lb=(profiles[b].invoices||[])[0];
    if(!la&&!lb)return 0;if(!la)return 1;if(!lb)return -1;
    return new Date(lb.savedAt)-new Date(la.savedAt);
  });
  var tbody=document.getElementById('clientTableBody'),empty=document.getElementById('clientTableEmpty');
  if(!tbody)return;tbody.innerHTML='';
  if(!keys.length){if(empty)empty.style.display='block';if(tbody)tbody.innerHTML='';return;}
  if(empty)empty.style.display='none';
  keys.forEach(function(name){
    var prof=profiles[name];
    var st=prof.clientStatus||'active';
    var cgName=prof.caregiverId&&cgs[prof.caregiverId]?cgs[prof.caregiverId].name:'—';
    var invs=prof.invoices||[];
    var lastInv=invs.length?invs[0].billingPeriod:'—';
    var open=invs.filter(function(i){return !i.status||i.status==='draft'||i.status==='submitted';}).length;
    var rate=prof.hourlyRate?'$'+prof.hourlyRate+'/hr':'—';
    var checked=bulkSelected[name]?'checked':'';
    var tr=document.createElement('tr');
    var hrefCl=buildClientUrl(name);
    tr.innerHTML=
      '<td style="width:32px;" onclick="event.stopPropagation()"><input type="checkbox" '+checked+' onchange="toggleBulkClient(\''+esc(name)+'\',this)" style="width:13px;height:13px;cursor:pointer;"></td>'+
      '<td><a href="'+hrefCl+'" style="text-decoration:none;color:inherit;display:block;"><div class="ct-name">'+esc(name)+(prof.nickname?'<span style="font-weight:normal;color:#8ca0b4;"> ('+esc(prof.nickname)+')</span>':'')+'</div><div class="ct-id">'+(prof.medicaidId||'No Medicaid ID')+'</div></a></td>'+
      '<td><span class="cs-badge cs-'+st+'">'+st.charAt(0).toUpperCase()+st.slice(1)+'</span></td>'+
      '<td style="color:#4a6a8a;font-size:12px;">'+esc(cgName)+'</td>'+
      '<td style="font-size:12px;color:#4a6a8a;">'+esc(lastInv)+'</td>'+
      '<td>'+(open?'<span style="color:#c47800;font-weight:700;">'+open+'</span>':'<span style="color:#ccc;">0</span>')+'</td>'+
      '<td style="font-size:12px;color:#4a6a8a;">'+esc(rate)+'</td>'+
      '<td onclick="event.stopPropagation()"><button class="ct-action-btn" onclick="activeProfileName=\''+esc(name)+'\';navInvoice()">+ Invoice</button></td>';
    tr.addEventListener('click',function(e){
      if(e.target.closest('a')||e.target.tagName==='BUTTON'||e.target.tagName==='INPUT')return;
      navDetail(name);
    });
    tbody.appendChild(tr);
  });
}
function renderClientGrid(){renderClientTable();}
function toggleBulkClient(name,cb){
  if(cb.checked)bulkSelected[name]=true;else delete bulkSelected[name];
  var count=Object.keys(bulkSelected).length;
  var bar=document.getElementById('ctBulkBar');
  if(bar){bar.classList.toggle('visible',count>0);var lbl=document.getElementById('ctBulkCount');if(lbl)lbl.textContent=count+' selected';}
}
function clearBulkSelect(){bulkSelected={};var bar=document.getElementById('ctBulkBar');if(bar)bar.classList.remove('visible');renderClientTable();}
function bulkSetStatus(status){
  var names=Object.keys(bulkSelected);if(!names.length)return;
  var p=getProfiles(),changed=0;
  names.forEach(function(name){
    if(!p[name])return;
    (p[name].invoices||[]).forEach(function(inv){if((inv.status||'draft')!=='paid'){inv.status=status;changed++;}});
    saveProfilesLS(p);saveProfileSP(name,p[name]);
    addAuditEntry(name,'Bulk status update: open invoices set to '+status);
  });
  logActivity('status','Bulk update: '+changed+' invoices set to '+status);
  clearBulkSelect();updateStats();
  showAlert(changed+' invoice'+(changed!==1?'s':'')+' updated to '+status+'.');
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
  ['overview','info','history','notes','docs','audit'].forEach(function(t){
    var dtab=document.getElementById('dtab-'+t);
    var dpane=document.getElementById('dpane-'+t);
    if(dtab)dtab.classList.toggle('active',t===tab);
    if(dpane)dpane.classList.toggle('active',t===tab);
  });
  if(tab==='overview')renderOverviewPane();
  if(tab==='info')renderInfoPane();
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
  // Caseworker lookup
  var cwRec=getCaseworkers().find(function(c){return c.id===prof.caseworkerId||c.name===prof.worker;})||null;
  var cwName=cwRec?cwRec.name:(prof.worker||'');
  // Address
  var addrStr=(prof.street||'')+( prof.city?', '+prof.city:'')+(prof.state?' '+prof.state:'')+(prof.zip?' '+prof.zip:'');
  // Client tasks
  var clientTasks=getTodos().filter(function(t){return t.client===activeProfileName&&!t.done;});
  var tasksHtml='';
  if(clientTasks.length){
    tasksHtml=clientTasks.slice(0,4).map(function(t){
      var overdue=t.due&&new Date(t.due)<new Date();
      return '<div class="ov-inv-row" style="cursor:pointer;" onclick="navTasks()" title="Go to Tasks">'+
        '<span style="font-size:11px;'+(overdue?'color:#b03030;font-weight:600;':'color:#1a2b45;')+'">'+esc(t.text)+'</span>'+
        (t.due?'<span style="font-size:10px;color:'+(overdue?'#b03030':'#8ca0b4')+';">'+t.due+'</span>':'')+
      '</div>';
    }).join('');
  } else {
    tasksHtml='<div style="color:#8ca0b4;font-size:12px;padding:4px 0;">No open tasks.</div>';
  }
  // Recent invoices — read-only badges
  var recentInvHtmlRO='';
  if(!invoices.length){recentInvHtmlRO='<div style="color:#8ca0b4;font-size:12px;padding:6px 0;">No invoices yet.</div>';}
  else{invoices.slice(0,4).forEach(function(inv){
    var st=inv.status||'draft';
    var daysSince=Math.floor((Date.now()-new Date(inv.savedAt))/86400000);
    var overdueFlag=(st==='submitted'&&daysSince>30)?'<span style="background:#ff8c00;color:#fff;font-size:10px;padding:1px 6px;border-radius:8px;margin-left:4px;">'+daysSince+'d</span>':'';
    recentInvHtmlRO+='<div class="ov-inv-row" style="cursor:pointer;" onclick="switchTab(\'history\')" title="Open Invoices tab">'+
      '<span class="inv-badge" style="min-width:60px;text-align:center;">'+esc(inv.billingPeriod)+'</span>'+
      '<span class="inv-badge st-'+st+'" style="min-width:70px;text-align:center;">'+st.charAt(0).toUpperCase()+st.slice(1)+overdueFlag+'</span>'+
      '<span style="flex:1;color:#8ca0b4;font-size:11px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">'+esc(inv.invoiceNote||'')+'</span>'+
      '<span style="font-size:11px;color:#6b8dae;white-space:nowrap;">'+esc(inv.savedAt.split(',')[0])+'</span></div>';
  });}
  pane.innerHTML='<div class="overview-grid">'+
    '<div class="ov-card"><h4>Client Info</h4>'+
      '<div class="ov-row"><span class="ov-label">Medicaid ID</span><span class="ov-value">'+(prof.medicaidId||'—')+'</span></div>'+
      '<div class="ov-row"><span class="ov-label">Hourly Rate</span><span class="ov-value">'+(prof.hourlyRate?'$'+prof.hourlyRate+'/hr':'—')+'</span></div>'+
      '<div class="ov-row"><span class="ov-label">Phone</span><span class="ov-value">'+(prof.phone||'—')+'</span></div>'+
      (addrStr.trim()?'<div class="ov-row"><span class="ov-label">Address</span><span class="ov-value">'+esc(addrStr.trim())+'</span></div>':'')+
      '<div class="ov-row"><span class="ov-label">Caregiver</span>'+(cgName&&prof.caregiverId?'<span class="ov-value" style="color:#185FA5;cursor:pointer;text-decoration:underline;" onclick="navCaregivers();setTimeout(function(){openCgDetail(\''+esc(prof.caregiverId)+'\');},50)">'+esc(cgName)+'</span>':'<span class="ov-value">'+esc(cgName||'Unassigned')+'</span>')+'</div>'+
      (prof.liveIn?'<div class="ov-row"><span class="ov-label">Live-In</span><span class="ov-value" style="display:inline-block;background:#fff3cd;color:#856404;font-size:11px;font-weight:600;padding:2px 10px;border-radius:10px;border:1px solid #ffeaa7;">YES</span></div>':'')+
      '<div class="ov-row"><span class="ov-label">Caseworker</span>'+(cwRec?'<span class="ov-value" style="color:#185FA5;cursor:pointer;text-decoration:underline;" onclick="navCaseworkers();setTimeout(function(){openCwDetail(\''+esc(cwRec.id)+'\');},50)">'+esc(cwName)+'</span>':'<span class="ov-value">'+esc(cwName||'—')+'</span>')+'</div>'+
      (prof.startDate?'<div class="ov-row"><span class="ov-label">Service Start</span><span class="ov-value">'+esc(prof.startDate)+'</span></div>':'')+
    '</div>'+
    '<div class="ov-card"><h4>Invoice Summary</h4>'+
      '<div class="ov-row"><span class="ov-label">Total Saved</span><span class="ov-value">'+invoices.length+'</span></div>'+
      '<div class="ov-row"><span class="ov-label">Paid</span><span class="ov-value" style="color:#1e7e34;">'+paid+'</span></div>'+
      '<div class="ov-row"><span class="ov-label">Submitted</span><span class="ov-value" style="color:#1565a0;">'+submitted+'</span></div>'+
      '<div class="ov-row"><span class="ov-label">Draft / Unsent</span><span class="ov-value" style="color:#666;">'+draft+'</span></div>'+
      '<div style="margin-top:10px;border-top:1px solid #f0f3f7;padding-top:10px;">'+recentInvHtmlRO+'</div>'+
    '</div>'+
    '<div class="ov-card"><h4>Tasks <span style="font-size:10px;color:#8ca0b4;font-weight:normal;text-transform:none;letter-spacing:0;">('+clientTasks.length+' open)</span>'+
      '<div style="margin-left:auto;display:flex;gap:6px;">'+
        '<button class="btn btn-secondary btn-sm" style="font-size:10px;" onclick="addTaskForClient(\''+esc(activeProfileName)+'\')">+ Add Task</button>'+
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

function cycleStatus(period,el){
  var states=['draft','submitted','paid'];
  var p=getProfiles();if(!p[activeProfileName])return;
  var inv=p[activeProfileName].invoices.find(function(i){return i.billingPeriod===period;});
  if(!inv)return;
  var cur=inv.status||'draft',idx=states.indexOf(cur),next=states[(idx+1)%states.length];
  inv.status=next;
  el.textContent=next.charAt(0).toUpperCase()+next.slice(1);
  el.className='inv-status inv-status-'+next;
  saveProfilesLS(p);saveProfileSP(activeProfileName,p[activeProfileName]);
  logActivity('status','Invoice '+period+' for '+activeProfileName+' marked '+next);
  updateStats();
}
function cycleStatusOverview(sel){
  var period=sel.dataset.period,next=sel.value;
  sel.className='status-select st-'+next;
  var p=getProfiles();if(!p[activeProfileName])return;
  var inv=p[activeProfileName].invoices.find(function(i){return i.billingPeriod===period;});
  if(!inv)return;
  inv.status=next;
  saveProfilesLS(p);saveProfileSP(activeProfileName,p[activeProfileName]);
  logActivity('status','Invoice '+period+' for '+activeProfileName+' marked '+next);
  updateStats();
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
        '<button class="btn btn-secondary btn-sm" onclick="event.stopPropagation();navDetail(\''+esc(r.client)+'\')">Profile</button>'+
        '<button class="btn btn-secondary btn-sm" onclick="event.stopPropagation();openInvFromModal(\''+esc(r.client)+'\','+r.idx+')">Open Invoice →</button>'+
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

  // Helper to create simple text field
  function mkField(id,label,val,full){
    var d=document.createElement('div');d.className='info-field'+(full?' full':'');
    d.innerHTML='<label>'+label+'</label><input id="'+id+'" value="'+esc(val)+'" oninput="unsavedChanges=true;">';
    g.appendChild(d);
  }
  // Helper for divider
  function mkDivider(label){
    var div=document.createElement('div');div.className='form-section-divider full';div.innerHTML='<span>'+label+'</span>';g.appendChild(div);
  }

  // Name row: First / Middle / Last
  var dName=document.createElement('div');dName.className='info-field-row full';dName.style.gridTemplateColumns='1fr 1fr 1fr';
  dName.innerHTML='<div class="info-field"><label>First Name *</label><input id="ei-first" value="'+esc(storedFirst)+'" oninput="unsavedChanges=true;"></div>'+
    '<div class="info-field"><label>Middle Name</label><input id="ei-middle" value="'+esc(prof.middleName||'')+'" oninput="unsavedChanges=true;"></div>'+
    '<div class="info-field"><label>Last Name *</label><input id="ei-last" value="'+esc(storedLast)+'" oninput="unsavedChanges=true;"></div>';
  g.appendChild(dName);

  // Nickname + Status row
  var dNickSt=document.createElement('div');dNickSt.className='info-field-row full';
  dNickSt.innerHTML='<div class="info-field"><label>Nickname / Goes By</label><input id="ei-nickname" value="'+esc(prof.nickname||'')+'" oninput="unsavedChanges=true;"></div>'+
    '<div class="info-field"><label>Client Status</label><select id="ei-status">'+
      ['active','inactive','lost','terminated'].map(function(s){return '<option value="'+s+'"'+((prof.clientStatus||'active')===s?' selected':'')+'>'+s.charAt(0).toUpperCase()+s.slice(1)+'</option>';}).join('')+
    '</select></div>';
  g.appendChild(dNickSt);

  mkField('ei-medicaid','Medicaid ID',prof.medicaidId||'',false);
  // Date of Birth — needed on DHS-390 / MDHHS-6200 / MSA-4676 state forms
  var dDob=document.createElement('div');dDob.className='info-field';
  dDob.innerHTML='<label>Date of Birth</label><input id="ei-dob" type="date" value="'+esc(prof.dob||'')+'" oninput="unsavedChanges=true;">';
  g.appendChild(dDob);
  var dGender=document.createElement('div');dGender.className='info-field';
  var gv=prof.gender||'';
  dGender.innerHTML='<label>Gender</label><select id="ei-gender" onchange="unsavedChanges=true;">'+
    '<option value=""'+(gv===''?' selected':'')+'>—</option>'+
    '<option value="Male"'+(gv==='Male'?' selected':'')+'>Male</option>'+
    '<option value="Female"'+(gv==='Female'?' selected':'')+'>Female</option>'+
  '</select>';
  g.appendChild(dGender);
  mkField('ei-rate','Hourly Rate',prof.hourlyRate||'',false);
  mkField('ei-dl',"Driver's License #",prof.driversLicense||'',false);
  // SSN masked by default with Show/Hide toggle
  var dSsn=document.createElement('div');dSsn.className='info-field';
  dSsn.innerHTML='<label>Social Security #</label>'+
    '<div style="display:flex;gap:4px;align-items:center;">'+
      '<input id="ei-ssn" type="password" autocomplete="off" value="'+esc(prof.ssn||'')+'" style="flex:1;" oninput="unsavedChanges=true;">'+
      '<button type="button" class="btn btn-secondary btn-sm" onclick="toggleMask(\'ei-ssn\',this)" style="padding:4px 8px;font-size:11px;white-space:nowrap;">Show</button>'+
    '</div>';
  g.appendChild(dSsn);
  mkField('ei-phone','Client Phone',prof.phone||'',false);
  mkField('ei-cemail','Client Email',prof.clientEmail||'',false);
  mkField('ei-street','Street',prof.street||'',true);

  // City + State row
  var dCityState=document.createElement('div');dCityState.className='info-field-row full';
  dCityState.innerHTML='<div class="info-field"><label>City</label><input id="ei-city" value="'+esc(prof.city||'')+'" oninput="unsavedChanges=true;"></div>'+
    '<div class="info-field"><label>State</label><input id="ei-state" value="'+esc(prof.state||'')+'" oninput="unsavedChanges=true;"></div>';
  g.appendChild(dCityState);

  // ZIP + County row
  var dZipCounty=document.createElement('div');dZipCounty.className='info-field-row full';
  dZipCounty.innerHTML='<div class="info-field"><label>ZIP</label><input id="ei-zip" value="'+esc(prof.zip||'')+'" oninput="unsavedChanges=true;lookupZip(\'ei-zip\',\'ei-city\',\'ei-state\',\'ei-county\')"></div>'+
    '<div class="info-field"><label>County</label><input id="ei-county" value="'+esc(prof.county||'')+'" oninput="unsavedChanges=true;"></div>';
  g.appendChild(dZipCounty);

  mkDivider('Assignments');

  // Service Start Date — moved ABOVE caregiver per user
  var dStart=document.createElement('div');dStart.className='info-field full';
  dStart.innerHTML='<label>Service Start Date <span style="font-weight:400;font-size:11px;color:#8ca0b4;">(prevents missing-invoice warnings for months before this date)</span></label><input type="date" id="ei-start-date" value="'+esc(prof.startDate||'')+'" oninput="unsavedChanges=true;">';
  g.appendChild(dStart);

  // Caregiver + Live-In side by side
  var dCgRow=document.createElement('div');dCgRow.className='info-field-row full';
  var cgName='';var cgsMap=getCaregivers();if(prof.caregiverId&&cgsMap[prof.caregiverId])cgName=cgsMap[prof.caregiverId].name;
  dCgRow.innerHTML='<div class="info-field"><label>Assigned Caregiver</label>'+
      '<div style="display:flex;align-items:center;gap:6px;position:relative;">'+
        '<div style="flex:1;position:relative;">'+
          '<input id="ei-caregiver-search" placeholder="Click to browse, or type to search…" maxlength="80" autocomplete="off" value="'+esc(cgName)+'" oninput="cgSearch(this,\'ei-caregiver-val\',\'ei-caregiver-drop\');unsavedChanges=true;" onfocus="cgSearch(this,\'ei-caregiver-val\',\'ei-caregiver-drop\')" onblur="setTimeout(function(){var d=document.getElementById(\'ei-caregiver-drop\');if(d)d.style.display=\'none\';},200)" style="width:100%;padding:7px 10px;border:1px solid #d0d8e4;border-radius:5px;font-size:13px;font-family:Arial,sans-serif;outline:none;">'+
          '<input type="hidden" id="ei-caregiver-val" value="'+esc(prof.caregiverId||'')+'">'+
          '<div id="ei-caregiver-drop" style="display:none;position:absolute;top:100%;left:0;right:0;background:#fff;border:1px solid #d0d8e4;border-radius:0 0 6px 6px;z-index:200;max-height:180px;overflow-y:auto;box-shadow:0 4px 12px rgba(0,0,0,0.1);"></div>'+
        '</div>'+
        (prof.caregiverId?'<button class="btn btn-secondary btn-sm" style="white-space:nowrap;" onclick="navCaregivers();setTimeout(function(){openCgDetail(\''+esc(prof.caregiverId)+'\');},50)">Open</button>':'')+
      '</div>'+
    '</div>'+
    '<div class="info-field" style="display:flex;align-items:flex-end;padding-bottom:8px;">'+
      '<label style="display:flex;align-items:center;gap:6px;cursor:pointer;font-weight:500;margin:0;"><input type="checkbox" id="ei-live-in" '+(prof.liveIn?'checked':'')+' onchange="unsavedChanges=true;" style="width:14px;height:14px;cursor:pointer;"> Live-In Caregiver</label>'+
    '</div>';
  g.appendChild(dCgRow);

  // Caseworker searchable autocomplete + Open button
  var dCw=document.createElement('div');dCw.className='info-field full';
  var cwName=prof.worker||'';
  dCw.innerHTML='<label>Caseworker</label>'+
    '<div style="display:flex;align-items:center;gap:6px;position:relative;">'+
      '<div style="flex:1;position:relative;">'+
        '<input id="ei-worker-search" placeholder="Click to browse, or type to search…" maxlength="80" autocomplete="off" value="'+esc(cwName)+'" oninput="cwSearch(this,\'ei-worker-val\',\'ei-worker-drop\');unsavedChanges=true;" onfocus="cwSearch(this,\'ei-worker-val\',\'ei-worker-drop\')" onblur="setTimeout(function(){var d=document.getElementById(\'ei-worker-drop\');if(d)d.style.display=\'none\';},200)" style="width:100%;padding:7px 10px;border:1px solid #d0d8e4;border-radius:5px;font-size:13px;font-family:Arial,sans-serif;outline:none;">'+
        '<input type="hidden" id="ei-worker-val" value="'+esc(prof.caseworkerId||'')+'">'+
        '<div id="ei-worker-drop" style="display:none;position:absolute;top:100%;left:0;right:0;background:#fff;border:1px solid #d0d8e4;border-radius:0 0 6px 6px;z-index:200;max-height:180px;overflow-y:auto;box-shadow:0 4px 12px rgba(0,0,0,0.1);"></div>'+
      '</div>'+
      (prof.caseworkerId?'<button class="btn btn-secondary btn-sm" style="white-space:nowrap;" onclick="navCaseworkers();setTimeout(function(){openCwDetail(\''+esc(prof.caseworkerId)+'\');},50)">Open</button>':'')+
    '</div>';
  g.appendChild(dCw);
}
function saveClientInfo(){
  if(!activeProfileName)return;
  var p=getProfiles();
  var rec=p[activeProfileName];
  var first=(document.getElementById('ei-first').value||'').trim();
  var middle=(document.getElementById('ei-middle').value||'').trim();
  var last=(document.getElementById('ei-last').value||'').trim();
  var nickname=(document.getElementById('ei-nickname').value||'').trim();
  var newName=((first+' '+last).trim())||activeProfileName;
  rec.firstName=first;rec.middleName=middle;rec.lastName=last;rec.nickname=nickname;
  rec.clientName=newName;rec.medicaidId=document.getElementById('ei-medicaid').value;
  var dobEl=document.getElementById('ei-dob');if(dobEl)rec.dob=dobEl.value||'';
  var genderEl=document.getElementById('ei-gender');if(genderEl)rec.gender=genderEl.value||'';
  rec.hourlyRate=document.getElementById('ei-rate').value;
  var dlEl=document.getElementById('ei-dl');if(dlEl)rec.driversLicense=dlEl.value;
  var ssnEl=document.getElementById('ei-ssn');if(ssnEl)rec.ssn=ssnEl.value;
  rec.phone=document.getElementById('ei-phone').value;rec.clientEmail=document.getElementById('ei-cemail').value;
  rec.street=document.getElementById('ei-street').value;
  rec.city=document.getElementById('ei-city').value;
  rec.state=document.getElementById('ei-state').value;
  rec.zip=document.getElementById('ei-zip').value;
  rec.county=document.getElementById('ei-county').value;
  // Searchable autocomplete fields
  var cgValEl=document.getElementById('ei-caregiver-val');rec.caregiverId=cgValEl?cgValEl.value:'';
  var liveInEl=document.getElementById('ei-live-in');if(liveInEl)rec.liveIn=liveInEl.checked;
  var startDateEl=document.getElementById('ei-start-date');if(startDateEl)rec.startDate=startDateEl.value||'';
  var cwValEl=document.getElementById('ei-worker-val');var cwSearchEl=document.getElementById('ei-worker-search');
  rec.caseworkerId=cwValEl?cwValEl.value:'';rec.worker=cwSearchEl?cwSearchEl.value:'';
  var statusEl=document.getElementById('ei-status');if(statusEl)rec.clientStatus=statusEl.value;
  if(newName!==activeProfileName){p[newName]=rec;delete p[activeProfileName];activeProfileName=newName;}
  saveProfilesLS(p);saveProfileSP(activeProfileName,p[activeProfileName]);
  unsavedChanges=false;
  logActivity('edit','Profile updated for '+activeProfileName);
  addAuditEntry(activeProfileName,'Profile information updated');
  var displayName=activeProfileName+(rec.nickname?' ('+rec.nickname+')':'');
  document.getElementById('detailName').textContent=displayName;
  var st=rec.clientStatus||'active';
  document.getElementById('detailMeta').innerHTML=(rec.medicaidId?'Medicaid: '+rec.medicaidId:'No Medicaid ID')+(rec.phone?' &nbsp;·&nbsp; '+rec.phone:'')+' &nbsp;<span class="cs-badge cs-'+st+'">'+st.charAt(0).toUpperCase()+st.slice(1)+'</span>';
  renderSidebarClients();
  var btn=document.getElementById('saveInfoBtn');btn.textContent='Saved';setTimeout(function(){btn.textContent='Save Changes';},1800);
}
function deleteClient(){
  if(!activeProfileName)return;
  var name=activeProfileName;
  showConfirm('Delete "'+name+'" and all their data? This cannot be undone.',function(){
    aiTrack('ClientDeleted',{client:name});
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
  c.innerHTML='<div style="margin-bottom:12px;display:flex;align-items:center;justify-content:space-between;"><span style="font-size:13px;color:#6b8dae;">'+invoices.length+' invoice'+(invoices.length!==1?'s':'')+' saved</span></div>';
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
    saveProfilesLS(p2);saveProfileSP(activeProfileName,p2[activeProfileName]);updateInvoiceStatusAPI(inv2.dbId, next);
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
    p2[activeProfileName].invoices[idx].invoiceNote=note||'';
    saveProfilesLS(p2);renderInvHistory();renderOverviewPane();
  },{title:'Invoice Note',okText:'Save Note'});
}
function cycleStatusByIdx(idx,el){
  var states=['draft','submitted','paid'];
  var p=getProfiles();if(!p[activeProfileName]||!p[activeProfileName].invoices[idx])return;
  var cur=p[activeProfileName].invoices[idx].status||'draft',next=states[(states.indexOf(cur)+1)%states.length];
  var period=p[activeProfileName].invoices[idx].billingPeriod;
  function apply(){
    var p2=getProfiles();if(!p2[activeProfileName]||!p2[activeProfileName].invoices[idx])return;
    p2[activeProfileName].invoices[idx].status=next;
    el.textContent=next.charAt(0).toUpperCase()+next.slice(1);el.className='inv-status inv-status-'+next;
    saveProfilesLS(p2);saveProfileSP(activeProfileName,p2[activeProfileName]);
    logActivity('status','Invoice '+period+' for '+activeProfileName+' marked '+next);
    updateStats();
  }
  if(next==='paid'){
    showConfirm('Mark invoice '+period+' as PAID?\n\nPaid invoices are locked from edits. Only mark Paid after payment is actually received.',apply,{title:'Mark as Paid?',okText:'Mark as Paid'});
  } else if(next==='submitted'&&cur==='draft'){
    showConfirm('Mark invoice '+period+' as SUBMITTED?\n\nThis indicates the invoice has been sent. Only do this if you have actually emailed/delivered it.',apply,{title:'Mark as Submitted?',okText:'Mark as Submitted'});
  } else {
    apply();
  }
}
function saveInvNote(input){
  var p=getProfiles(),idx=parseInt(input.dataset.idx);
  if(p[activeProfileName]&&p[activeProfileName].invoices&&p[activeProfileName].invoices[idx]){p[activeProfileName].invoices[idx].invoiceNote=input.value;saveProfilesLS(p);}
}
function deleteInv(idx){
  var p=getProfiles(),inv=p[activeProfileName].invoices[idx];
  if(!inv)return;
  showConfirm('Permanently delete the '+inv.billingPeriod+' invoice for '+activeProfileName+'? This cannot be undone.',function(){doDeleteInv(idx);},{title:'Delete Invoice',okText:'Delete'});
}
function doDeleteInv(idx){
  var p=getProfiles(),inv=p[activeProfileName].invoices[idx];
  if(!inv)return;
  deleteInvoiceAPI(inv.dbId,activeProfileName,inv.billingPeriod);
  p[activeProfileName].invoices.splice(idx,1);saveProfilesLS(p);saveProfileSP(activeProfileName,p[activeProfileName]);renderInvHistory();
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
  var t=null;
  ta._nl=function(){clearTimeout(t);t=setTimeout(function(){var p2=getProfiles();if(p2[activeProfileName]){p2[activeProfileName].clientNotes=ta.value;saveProfilesLS(p2);var f=document.getElementById('notesSavedFlash');f.style.display='inline';setTimeout(function(){f.style.display='none';},2000);}},600);};
  ta.addEventListener('input',ta._nl);
  var c=document.getElementById('invNotesContent'),invoices=(prof&&prof.invoices)?prof.invoices:[];
  if(!invoices.length){c.innerHTML='<div style="color:#8ca0b4;font-size:13px;">No saved invoices yet.</div>';return;}
  c.innerHTML='';
  invoices.forEach(function(inv,idx){
    var row=document.createElement('div');row.style.cssText='display:flex;align-items:center;gap:10px;margin-bottom:8px;';
    row.innerHTML='<span style="min-width:72px;font-size:12px;font-weight:600;color:#185FA5;">'+esc(inv.billingPeriod)+'</span>'+
      '<input style="flex:1;padding:6px 10px;border:1px solid #d0d8e4;border-radius:5px;font-size:12px;font-family:Arial,sans-serif;outline:none;color:#1a2b45;" value="'+esc(inv.invoiceNote||'')+'" placeholder="Note…" data-idx="'+idx+'">';
    c.appendChild(row);
  });
  c.querySelectorAll('input[data-idx]').forEach(function(inp){
    inp.addEventListener('change',function(){var p2=getProfiles(),i=parseInt(inp.dataset.idx);if(p2[activeProfileName]&&p2[activeProfileName].invoices[i]){p2[activeProfileName].invoices[i].invoiceNote=inp.value;saveProfilesLS(p2);}});
  });
}

// ============================================================
//  DOCUMENTS (Azure Blob Storage)
// ============================================================
function getHcClientId(){
  var prof=getProfiles()[activeProfileName];
  return prof&&prof._dbId?prof._dbId:null;
}
function renderDocsPane(){
  var c=document.getElementById('docsContent');c.innerHTML='';
  if(!activeProfileName)return;
  var clientId=getHcClientId();
  c.innerHTML=
    '<div class="doc-upload-card">'+
      '<div class="doc-upload-head">'+
        '<h4>Client Documents</h4>'+
        '<p>Upload SSN cards, driver\'s licenses, insurance cards, authorizations, etc.</p>'+
      '</div>'+
      '<div class="doc-upload-row">'+
        '<div class="doc-upload-fields">'+
          '<label>Category</label>'+
          '<select id="hcDocCategory">'+
            '<option value="Other">Other</option>'+
            '<option value="SSN_Card">SSN Card</option>'+
            '<option value="Drivers_License">Driver\'s License</option>'+
            '<option value="Insurance_Card">Insurance Card</option>'+
            '<option value="Medicare_Card">Medicare Card</option>'+
            '<option value="Medicaid_Card">Medicaid Card</option>'+
            '<option value="Authorization">Authorization</option>'+
          '</select>'+
          '<label style="margin-top:8px;">File</label>'+
          '<input type="file" id="hcDocFileInput" accept=".pdf,.jpg,.jpeg,.png,.doc,.docx" multiple>'+
        '</div>'+
        '<div class="doc-upload-actions">'+
          '<button class="btn btn-primary" onclick="uploadHcDoc()">Upload</button>'+
          '<input type="file" id="docScanInput" accept="image/*" capture="environment" style="display:none;" onchange="handleDocScan(this)">'+
          '<button class="btn btn-secondary" onclick="document.getElementById(\'docScanInput\').click()">Scan / Photo</button>'+
        '</div>'+
      '</div>'+
      '<span id="hcDocStatus" class="doc-upload-status"></span>'+
    '</div>'+
    '<div id="hcDocList"><div style="color:#8ca0b4;font-size:13px;">Loading...</div></div>';
  if(clientId){loadHcDocs(clientId);}
  else{document.getElementById('hcDocList').innerHTML='<div style="color:#8ca0b4;font-size:12px;">Save this client to the database first before uploading documents.</div>';}
}
function loadHcDocs(clientId){
  fetch(API_BASE+'/documents?clientType=homecare&clientId='+clientId,{headers:apiHeaders()})
  .then(function(r){return r.json();})
  .then(function(docs){renderHcDocList(clientId,docs||[]);})
  .catch(function(){renderHcDocList(clientId,[]);});
}
function renderHcDocList(clientId,docs){
  var list=document.getElementById('hcDocList');if(!list)return;
  if(!docs.length){list.innerHTML='<div style="color:#8ca0b4;font-size:13px;padding:4px 0;">No documents yet.</div>';return;}
  list.innerHTML='';
  var categoryLabels={SSN_Card:'SSN Card',Drivers_License:"Driver's License",Insurance_Card:'Insurance Card',Medicare_Card:'Medicare Card',Medicaid_Card:'Medicaid Card',Authorization:'Authorization',Other:'Other'};
  docs.forEach(function(d){
    var kb=d.size?Math.round(d.size/1024)+'KB':'';
    var ext=(d.name||'').split('.').pop().toLowerCase();
    var isImg=['jpg','jpeg','png','gif'].indexOf(ext)>=0;
    var icon=(ext||"").toUpperCase().slice(0,4);
    // parse category prefix from filename: "SSN_Card__filename.pdf"
    var parts=d.name.split('__');
    var cat=parts.length>1?parts[0]:'Other';
    var displayName=parts.length>1?parts.slice(1).join('__'):d.name;
    var catLabel=categoryLabels[cat]||cat;
    var card=document.createElement('div');
    card.style.cssText='display:flex;align-items:center;gap:10px;padding:8px 10px;border:1px solid #e1e8f0;border-radius:6px;margin-bottom:6px;background:#fafbfc;';
    card.innerHTML=
      '<span style="display:inline-block;min-width:34px;padding:3px 6px;background:#e8eef5;color:#1a3a5c;border-radius:4px;font-size:10px;font-weight:600;text-align:center;letter-spacing:.3px;">'+(icon||'FILE')+'</span>'+
      '<div style="flex:1;min-width:0;">'+
        '<div style="font-size:12px;font-weight:600;color:#1a3a5c;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;"><a href="'+d.url+'" target="_blank" style="color:#1a3a5c;text-decoration:none;">'+esc(displayName)+'</a></div>'+
        '<div style="font-size:11px;color:#8ca0b4;">'+
          '<span style="background:#dbeafe;color:#1e40af;padding:1px 7px;border-radius:10px;font-size:10px;font-weight:600;margin-right:6px;">'+esc(catLabel)+'</span>'+
          kb+
        '</div>'+
      '</div>'+
      '<button class="btn btn-danger btn-sm" style="padding:3px 10px;font-size:11px;" onclick="deleteHcDoc('+clientId+',\''+encodeURIComponent(d.name)+'\')">✕</button>';
    list.appendChild(card);
  });
}
function uploadHcDoc(){
  var clientId=getHcClientId();
  if(!clientId){showAlert('Save this client first before uploading documents.');return;}
  var input=document.getElementById('hcDocFileInput');
  if(!input||!input.files||!input.files.length){showAlert('Please select a file first.');return;}
  var cat=(document.getElementById('hcDocCategory')&&document.getElementById('hcDocCategory').value)||'Other';
  var status=document.getElementById('hcDocStatus');
  status.textContent='Uploading...';
  var fd=new FormData();
  fd.append('clientType','homecare');fd.append('clientId',clientId);
  Array.from(input.files).forEach(function(f){
    var prefixedFile=new File([f],cat+'__'+f.name,{type:f.type});
    fd.append('file',prefixedFile);
  });
  var fileNames=Array.from(input.files).map(function(f){return f.name;}).join(', ');
  fetch(API_BASE+'/documents',{method:'POST',headers:authUploadHeaders(),body:fd})
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
  var cat=(document.getElementById('hcDocCategory')&&document.getElementById('hcDocCategory').value)||'Other';
  var status=document.getElementById('hcDocStatus');
  if(status)status.textContent='Uploading scanned image…';
  var fd=new FormData();
  fd.append('clientType','homecare');fd.append('clientId',clientId);
  var f=input.files[0];
  var prefixedFile=new File([f],cat+'__'+f.name,{type:f.type});
  fd.append('file',prefixedFile);
  fetch(API_BASE+'/documents',{method:'POST',headers:authUploadHeaders(),body:fd})
  .then(function(){
    aiTrack('DocumentUploaded',{clientType:'homecare',clientId:clientId,category:cat,files:f.name,source:'scan'});
    if(status)status.textContent='';input.value='';loadHcDocs(clientId);
  })
  .catch(function(e){if(status)status.textContent='Upload failed: '+e;});
}
function deleteHcDoc(clientId,encodedName){
  showConfirm('Delete this document?',function(){
    fetch(API_BASE+'/documents?clientType=homecare&clientId='+clientId+'&name='+encodedName,{method:'DELETE',headers:apiHeaders()})
    .then(function(){loadHcDocs(clientId);}).catch(function(e){showAlert('Delete failed: '+e);});
  },{title:'Delete Document',okText:'Delete'});
}

// ── CAREGIVER DOCUMENTS ──────────────────────────────────────
function loadCgDocs(cgId){
  var sec=document.getElementById('cgDocsSection');if(!sec)return;
  sec.innerHTML='<div style="font-weight:600;font-size:12px;margin-bottom:8px;color:#1a3a5c;">Documents <span style="font-weight:normal;color:#8ca0b4;font-size:11px;">(SSN card, license, certs)</span></div>'+
    '<div id="cgDocList" style="margin-bottom:8px;"><div style="color:#8ca0b4;font-size:12px;">Loading...</div></div>'+
    '<div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;">'+
    '<input type="file" id="cgDocFileInput" accept=".pdf,.jpg,.jpeg,.png,.doc,.docx" multiple style="font-size:11px;flex:1;min-width:0;">'+
    '<button class="btn btn-primary btn-sm" onclick="uploadCgDoc(\''+cgId+'\')">Upload</button>'+
    '</div>'+
    '<span id="cgDocStatus" style="font-size:11px;color:#666;"></span>';
  fetch(API_BASE+'/documents?clientType=caregiver&clientId='+cgId,{headers:apiHeaders()})
  .then(function(r){return r.json();})
  .then(function(docs){
    var list=document.getElementById('cgDocList');if(!list)return;
    if(!docs||!docs.length){list.innerHTML='<div style="color:#8ca0b4;font-size:12px;">No documents yet.</div>';return;}
    list.innerHTML='';
    docs.forEach(function(d){
      var kb=d.size?Math.round(d.size/1024)+'KB':'';
      var div=document.createElement('div');
      div.style.cssText='display:flex;align-items:center;gap:6px;padding:5px 0;border-bottom:1px solid #f0f0f0;font-size:12px;';
      div.innerHTML='<a href="'+d.url+'" target="_blank" style="flex:1;color:#1a3a5c;text-decoration:none;word-break:break-all;">'+esc(d.name)+'</a>'+
        '<span style="color:#8ca0b4;font-size:11px;">'+kb+'</span>'+
        '<button class="btn btn-danger btn-sm" style="padding:2px 8px;font-size:10px;" onclick="deleteCgDoc(\''+cgId+'\',\''+encodeURIComponent(d.name)+'\')">✕</button>';
      list.appendChild(div);
    });
  }).catch(function(){var l=document.getElementById('cgDocList');if(l)l.innerHTML='<div style="color:#e74c3c;font-size:12px;">Could not load documents.</div>';});
}
function uploadCgDoc(cgId){
  var input=document.getElementById('cgDocFileInput');
  if(!input||!input.files||!input.files.length){showAlert('Please select a file first.');return;}
  var status=document.getElementById('cgDocStatus');status.textContent='Uploading...';
  var fd=new FormData();fd.append('clientType','caregiver');fd.append('clientId',cgId);
  var cgFileNames=Array.from(input.files).map(function(f){return f.name;}).join(', ');
  Array.from(input.files).forEach(function(f){fd.append('file',f);});
  fetch(API_BASE+'/documents',{method:'POST',headers:authUploadHeaders(),body:fd})
  .then(function(){
    aiTrack('DocumentUploaded',{clientType:'caregiver',clientId:cgId,files:cgFileNames});
    status.textContent='';input.value='';loadCgDocs(cgId);
  })
  .catch(function(e){status.textContent='Upload failed: '+e;});
}
function deleteCgDoc(cgId,encodedName){
  showConfirm('Delete this document?',function(){
    fetch(API_BASE+'/documents?clientType=caregiver&clientId='+cgId+'&name='+encodedName,{method:'DELETE',headers:apiHeaders()})
    .then(function(){loadCgDocs(cgId);}).catch(function(e){showAlert('Delete failed: '+e);});
  },{title:'Delete Document',okText:'Delete'});
}

// ============================================================
//  NEW CLIENT
// ============================================================
function createClient(){
  var first=(document.getElementById('nc-first').value||'').trim();
  var middle=(document.getElementById('nc-middle').value||'').trim();
  var last=(document.getElementById('nc-last').value||'').trim();
  var nickname=(document.getElementById('nc-nickname').value||'').trim();
  if(!first||!last){showAlert('First Name and Last Name are required.');return;}
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
  p[name].medicaidId=document.getElementById('nc-medicaid').value.trim();
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
  p[name].caregiverId=document.getElementById('nc-caregiver').value;
  p[name].liveIn=document.getElementById('nc-live-in').checked;
  p[name].startDate=document.getElementById('nc-start-date').value||'';
  p[name].clientStatus=document.getElementById('nc-status').value||'active';
  p[name].hasComplex=false;if(!p[name].tasks)p[name].tasks=captureStates();
  saveProfilesLS(p);saveProfileSP(name,p[name]);logActivity('client','New client added: '+name);navDetail(name);
}

// ============================================================
//  CAREGIVERS
// ============================================================
function getCaregivers(){try{return JSON.parse(localStorage.getItem('lhca_caregivers')||'{}');}catch(e){return{};}}
function saveCaregiversLS(cg){localStorage.setItem('lhca_caregivers',JSON.stringify(cg));}
function cgId(){return 'cg_'+Date.now()+'_'+Math.random().toString(36).slice(2,7);}

function renderCaregiverGrid(){
  var cgs=getCaregivers();
  var q=(document.getElementById('cgSearch')?document.getElementById('cgSearch').value:'').toLowerCase();
  var filterStatus=(document.getElementById('cgFilterStatus')&&document.getElementById('cgFilterStatus').value)||'active';
  var profiles=getProfiles();
  var tbody=document.getElementById('cgTableBody');if(!tbody)return;tbody.innerHTML='';
  var ids=Object.keys(cgs).filter(function(id){
    var cg=cgs[id];
    var st=cg.status||'active';
    var matchStatus=filterStatus==='all'||st===filterStatus;
    var matchQ=!q||(cg.name||'').toLowerCase().includes(q)||(cg.phone||'').includes(q);
    return matchStatus&&matchQ;
  });
  ids.sort(function(a,b){return (cgs[a].name||'').localeCompare(cgs[b].name||'');});
  var empty=document.getElementById('cgTableEmpty');
  if(!ids.length){if(empty)empty.style.display='block';return;}
  if(empty)empty.style.display='none';
  ids.forEach(function(id){
    var cg=cgs[id];
    var st=cg.status||'active';
    var clientCount=Object.keys(profiles).filter(function(k){return profiles[k].caregiverId===id;}).length;
    var displayName=(cg.firstName&&cg.lastName)?(cg.firstName+(cg.middleName?' '+cg.middleName:'')+' '+cg.lastName).trim():(cg.name||id);
    var tr=document.createElement('tr');
    var hrefCg=buildCaregiverUrl(id);
    tr.innerHTML=
      '<td><a href="'+hrefCg+'" style="text-decoration:none;color:inherit;display:block;"><div class="ct-name">'+esc(displayName)+(cg.nickname?'<span style="font-weight:normal;color:#8ca0b4;"> ('+esc(cg.nickname)+')</span>':'')+'</div>'+
        '<div class="ct-id">'+(cg.email||'No email')+'</div></a></td>'+
      '<td><span class="cs-badge cs-'+st+'">'+st.charAt(0).toUpperCase()+st.slice(1)+'</span></td>'+
      '<td style="color:#4a6a8a;font-size:12px;">'+esc(cg.emptype||'—')+'</td>'+
      '<td style="color:#4a6a8a;font-size:12px;">'+esc(cg.phone||'—')+'</td>'+
      '<td style="font-size:12px;color:#4a6a8a;">'+esc(cg.hireDate||'—')+'</td>'+
      '<td style="font-size:12px;">'+clientCount+'</td>'+
      '<td onclick="event.stopPropagation()"><button class="ct-action-btn" onclick="event.stopPropagation();editCaregiver(\''+id+'\')">Edit</button></td>';
    tr.addEventListener('click',function(e){
      // Don't double-fire if user clicked the <a> (which the browser navigates via hashchange)
      if(e.target.closest('a')||e.target.tagName==='BUTTON')return;
      openCgDetail(id);
    });
    tbody.appendChild(tr);
  });
}
function bulkDeleteCaregivers(){
  var checked=Array.from(document.querySelectorAll('.cg-select:checked'));
  if(!checked.length){showAlert('No caregivers selected.');return;}
  showConfirm('Delete '+checked.length+' caregiver'+(checked.length>1?'s':'')+'? This cannot be undone.',function(){
    var cgs=getCaregivers();
    checked.forEach(function(cb){
      var id=cb.dataset.id;
      delete cgs[id];
      deleteCaregiverAPI(id);
    });
    saveCaregiversLS(cgs);renderCaregiverGrid();updateStats();
  },{title:'Bulk Delete Caregivers',okText:'Delete All'});
}
function toggleAllCaregivers(cb){
  document.querySelectorAll('.cg-select').forEach(function(c){c.checked=cb.checked;});
}
function showNewCaregiverForm(){
  document.getElementById('cgFormWrap').style.display='block';
  document.getElementById('cgFormTitle').textContent='New Caregiver';
  document.getElementById('cg-editing-id').value='';
  ['cg-first','cg-middle','cg-last','cg-nickname','cg-phone','cg-email','cg-dl','cg-ssn','cg-street','cg-city','cg-state','cg-zip','cg-county','cg-dob','cg-gender','cg-hire','cg-pay','cg-hours','cg-certs','cg-ec-name','cg-ec-phone','cg-champs','cg-notes'].forEach(function(id){var e=document.getElementById(id);if(e)e.value='';});
  document.getElementById('cg-status').value='active';document.getElementById('cg-emptype').value='full-time';
  document.getElementById('cgDeleteBtn').style.display='none';
  var cgDocSec=document.getElementById('cgDocsSection');if(cgDocSec)cgDocSec.style.display='none';
  document.getElementById('cgFormWrap').scrollIntoView({behavior:'smooth'});
}
function editCaregiver(id){
  var cg=getCaregivers()[id];if(!cg)return;
  showNewCaregiverForm();
  document.getElementById('cgFormTitle').textContent='Edit Caregiver';
  document.getElementById('cg-editing-id').value=id;
  // Populate split name fields (fall back to splitting cg.name if firstName not stored)
  var cgFirst=cg.firstName||'';var cgLast=cg.lastName||'';
  if(!cgFirst&&!cgLast){var cgParts=(cg.name||'').split(' ');cgFirst=cgParts[0]||'';cgLast=cgParts.slice(1).join(' ')||'';}
  document.getElementById('cg-first').value=cgFirst;
  document.getElementById('cg-middle').value=cg.middleName||'';
  document.getElementById('cg-last').value=cgLast;
  document.getElementById('cg-nickname').value=cg.nickname||'';
  document.getElementById('cg-phone').value=cg.phone||'';
  document.getElementById('cg-email').value=cg.email||'';
  var cgDlEl=document.getElementById('cg-dl');if(cgDlEl)cgDlEl.value=cg.driversLicense||'';
  var cgSsnEl=document.getElementById('cg-ssn');if(cgSsnEl)cgSsnEl.value=cg.ssn||'';
  document.getElementById('cg-street').value=cg.street||cg.address||'';
  document.getElementById('cg-city').value=cg.city||'';
  document.getElementById('cg-state').value=cg.state||'';
  document.getElementById('cg-zip').value=cg.zip||'';
  document.getElementById('cg-county').value=cg.county||'';
  var cgDobEl=document.getElementById('cg-dob');if(cgDobEl)cgDobEl.value=cg.dob||cg.dateOfBirth||'';
  var cgChampsEl=document.getElementById('cg-champs');if(cgChampsEl)cgChampsEl.value=cg.champsId||cg.champs_id||'';
  var cgGenderEl=document.getElementById('cg-gender');if(cgGenderEl)cgGenderEl.value=cg.gender||'';
  document.getElementById('cg-hire').value=cg.hireDate||'';document.getElementById('cg-pay').value=cg.payRate||'';
  document.getElementById('cg-hours').value=cg.maxHours||'';document.getElementById('cg-certs').value=cg.certs||'';
  document.getElementById('cg-ec-name').value=cg.ecName||'';document.getElementById('cg-ec-phone').value=cg.ecPhone||'';
  document.getElementById('cg-notes').value=cg.notes||'';
  document.getElementById('cg-status').value=cg.status||'active';document.getElementById('cg-emptype').value=cg.emptype||'full-time';
  document.getElementById('cgDeleteBtn').style.display='inline-block';
  var cgDocSec=document.getElementById('cgDocsSection');
  if(cgDocSec){cgDocSec.style.display='block';loadCgDocs(id);}
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
    payRate:document.getElementById('cg-pay').value,maxHours:document.getElementById('cg-hours').value,
    certs:document.getElementById('cg-certs').value,ecName:document.getElementById('cg-ec-name').value,
    ecPhone:document.getElementById('cg-ec-phone').value,
    champsId:(document.getElementById('cg-champs')||{}).value||'',
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
function populateCaregiverSelect(selId,selectedId){
  var sel=document.getElementById(selId);if(!sel)return;
  sel.innerHTML='<option value="">— None assigned —</option>';
  var cgs=getCaregivers();
  Object.keys(cgs).forEach(function(id){
    var o=document.createElement('option');o.value=id;o.textContent=cgs[id].name+(cgs[id].status!=='active'?' ('+cgs[id].status+')':'');
    if(id===selectedId)o.selected=true;
    sel.appendChild(o);
  });
}

function populateCaseworkerSelect(selId, selectedName){
  var sel=document.getElementById(selId);if(!sel)return;
  sel.innerHTML='<option value="">— None assigned —</option>';
  getCaseworkers().forEach(function(cw){
    var o=document.createElement('option');
    o.value=cw.name;o.textContent=cw.name+(cw.agency?' ('+cw.agency+')':'');
    if(cw.name===selectedName)o.selected=true;
    sel.appendChild(o);
  });
  // allow typing a custom name if not in list
  if(selectedName&&!getCaseworkers().find(function(c){return c.name===selectedName;})){
    var o=document.createElement('option');o.value=selectedName;o.textContent=selectedName+' (custom)';o.selected=true;sel.appendChild(o);
  }
}

// --- Caregiver detail view ---
var activeCgId=null;
function openCgDetail(id){
  var cg=getCaregivers()[id];if(!cg)return;
  activeCgId=id;
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
    if(!arr.length){host.innerHTML='<div style="padding:14px;color:#8ca0b4;text-align:center;">No templates uploaded yet. Pick a PDF above to add your first one.</div>';return;}
    host.innerHTML=arr.map(function(t){
      var when=new Date(t.created_at).toLocaleDateString();
      return '<div style="display:flex;align-items:center;gap:10px;padding:10px 14px;border-bottom:1px solid #f0f3f7;">'+
        '<div style="flex:1;min-width:0;">'+
          '<div style="font-weight:600;color:#1a2b45;">'+esc(t.name)+(t.version?' <span style="color:#8ca0b4;font-weight:400;font-size:11px;">('+esc(t.version)+')</span>':'')+'</div>'+
          '<div style="font-size:11px;color:#8ca0b4;">Uploaded '+when+(t.is_active?'':' · Inactive')+'</div>'+
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
  host.innerHTML='<div style="padding:14px;color:#8ca0b4;text-align:center;font-size:12px;">Loading…</div>';
  try{
    var resp=await fetch(API_BASE+'/signing/list?caregiverId='+encodeURIComponent(activeCgId),{headers:apiHeaders()});
    if(!resp.ok)throw new Error('HTTP '+resp.status);
    var arr=await resp.json();
    if(!arr.length){host.innerHTML='<div style="padding:14px;color:#8ca0b4;text-align:center;font-size:12px;">No signing requests yet. Use the Send for Signature button above to create one.</div>';return;}
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
          '<div style="font-weight:600;font-size:12px;color:#1a2b45;">'+esc(r.template_name||'Document')+(r.template_version?' <span style="color:#8ca0b4;font-weight:400;">('+esc(r.template_version)+')</span>':'')+'</div>'+
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
async function resendSigningRequest(id){
  if(!spToken){showAlert('Sign in with Microsoft first to send the email.');return;}
  try{
    var cg=getCaregivers()[activeCgId]||{};
    var resp=await fetch(API_BASE+'/signing/'+id+'/resend',{method:'POST',headers:apiHeaders()});
    var data=await resp.json();
    if(!resp.ok)throw new Error(data.error||'HTTP '+resp.status);
    var subject='Reminder: please sign your document';
    var body='<p>Hi '+esc(data.recipientName||cg.name||'')+',</p>'+
      '<p>Here\'s a fresh link to sign your document. The previous link is no longer valid.</p>'+
      '<p><a href="'+data.signUrl+'" style="background:#185FA5;color:#fff;padding:10px 16px;border-radius:5px;text-decoration:none;display:inline-block;">Open &amp; Sign</a></p>'+
      '<p style="font-size:13px;color:#444;">When you open the link, you\'ll be asked to verify your date of birth.</p>'+
      '<p style="font-size:12px;color:#666;">Or paste:<br><span style="word-break:break-all;">'+data.signUrl+'</span></p>'+
      '<p style="font-size:12px;color:#666;">Expires '+new Date(data.expiresAt).toLocaleDateString()+'.</p>';
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

  // Fetch active templates
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
    '<label style="display:block;font-size:11px;color:#6b8dae;font-weight:600;text-transform:uppercase;letter-spacing:.4px;margin-bottom:4px;">Document</label>'+
    '<select id="sendSigTemplate" style="width:100%;padding:8px 10px;border:1px solid #d0d8e4;border-radius:5px;font-size:13px;outline:none;background:#fff;margin-bottom:12px;">'+tplOptions+'</select>'+
    '<label style="display:block;font-size:11px;color:#6b8dae;font-weight:600;text-transform:uppercase;letter-spacing:.4px;margin-bottom:4px;">Recipient Date of Birth <span style="font-weight:400;text-transform:none;letter-spacing:0;color:#a05a00;">(used to verify identity at signing — 3 wrong attempts locks the link)</span></label>'+
    '<input id="sendSigDob" type="date" value="'+esc(prefillDob)+'" style="width:200px;padding:8px 10px;border:1px solid #d0d8e4;border-radius:5px;font-size:13px;outline:none;margin-bottom:14px;">'+
    '<div style="font-size:11px;color:#8ca0b4;margin-bottom:14px;">Link expires in 14 days. A copy of every signed document is auto-CC\'d to <b>tommy@mybellcare.com</b>.</div>'+
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
    var subject='Please sign: '+(document.getElementById('sendSigTemplate').selectedOptions[0].textContent||'Document');
    var body='<p>Hi '+esc(cg.name||'')+',</p>'+
      '<p>Liberty Home Care needs your signature on a document. Please click the secure link below to review and sign.</p>'+
      '<p><a href="'+data.signUrl+'" style="background:#185FA5;color:#fff;padding:10px 16px;border-radius:5px;text-decoration:none;display:inline-block;">Open &amp; Sign</a></p>'+
      '<p style="font-size:13px;color:#444;">When you open the link, you\'ll be asked to verify your date of birth before viewing the document. This protects you in case the email is forwarded or intercepted.</p>'+
      '<p style="font-size:12px;color:#666;">Or paste this URL into your browser:<br><span style="word-break:break-all;">'+data.signUrl+'</span></p>'+
      '<p style="font-size:12px;color:#666;">This link expires '+new Date(data.expiresAt).toLocaleDateString()+' and can only be used once.</p>'+
      '<p style="font-size:12px;color:#999;">— Liberty Home Care Assistance</p>';
    if(spToken){
      var emailResp=await sendMailWithPDF(cg.email,subject,body,[]);
      if(!emailResp.ok){
        // Backend already created the request — show the URL so admin can copy/paste
        errEl.innerHTML='Created the request, but email failed to send: '+esc(emailResp.err||emailResp.status||'unknown')+'<br><br>You can manually share this link:<br><a href="'+data.signUrl+'" target="_blank">'+esc(data.signUrl)+'</a>';
        errEl.style.display='block';
        btn.textContent='Send Link';btn.disabled=false;
        return;
      }
    } else {
      errEl.innerHTML='Sign in with Microsoft first to email the link automatically.<br><br>For now, you can manually share this link:<br><a href="'+data.signUrl+'" target="_blank">'+esc(data.signUrl)+'</a>';
      errEl.style.display='block';
      btn.textContent='Send Link';btn.disabled=false;
      return;
    }
    closeSendSigModal();
    showAlert('✓ Signing link emailed to '+cg.email+'.\n\nThey have '+days+' days to sign. You can track status from the caregiver detail page.',{title:'Sent'});
    logActivity('signing','Sent signing request to '+cg.name+' for template #'+tplId);
  } catch(e){
    errEl.textContent='Error: '+(e.message||e);
    errEl.style.display='block';
    btn.textContent='Send Link';btn.disabled=false;
  }
}
function editCaregiverFromDetail(){
  if(!activeCgId)return;
  document.getElementById('cgDetailView').style.display='none';
  document.getElementById('cgGridView').style.display='none';
  editCaregiver(activeCgId);
  setTimeout(function(){
    var form=document.getElementById('cgFormWrap');
    if(form){form.style.display='block';form.scrollIntoView({behavior:'smooth'});}
  },50);
}
function deleteCaregiverFromDetail(){
  var cg=getCaregivers()[activeCgId];
  if(!cg)return;
  showConfirm('Delete caregiver "'+cg.name+'"? This cannot be undone.',function(){_doDeleteCaregiverFromDetail();},{title:'Delete Caregiver',okText:'Delete'});
}
function _doDeleteCaregiverFromDetail(){
  var cgs=getCaregivers();delete cgs[activeCgId];saveCaregiversLS(cgs);deleteCaregiverAPI(activeCgId);
  aiTrack('CaregiverDeleted',{caregiverId:activeCgId,caregiverName:cg.name||activeCgId});
  showCgGrid();
}
function switchCgTab(tab){
  ['overview','info','clients','notes','docs','signatures','audit'].forEach(function(t){
    var tb=document.getElementById('cgtab-'+t);
    var pn=document.getElementById('cgpane-'+t);
    if(tb)tb.classList.toggle('active',t===tab);
    if(pn)pn.classList.toggle('active',t===tab);
  });
  if(tab==='overview')renderCgOverviewPane();
  if(tab==='info')renderCgInfoPane();
  if(tab==='clients')renderCgClientsPane();
  if(tab==='notes')renderCgNotesPane();
  if(tab==='docs')renderCgDocsPane();
  if(tab==='signatures'&&typeof loadCgSigningRequests==='function')loadCgSigningRequests();
  if(tab==='audit')renderCgAuditPane();
}
function renderCgOverviewPane(){
  if(!activeCgId)return;
  var cg=getCaregivers()[activeCgId];
  var pane=document.getElementById('cgpane-overview');
  if(!cg||!pane)return;
  var profiles=getProfiles();
  var assigned=Object.keys(profiles).filter(function(k){return profiles[k].caregiverId===activeCgId;});
  var addrStr=(cg.street||cg.address||'').trim();
  var certsStr=cg.certifications||cg.certs||'';
  pane.innerHTML='<div class="overview-grid">'+
    '<div class="ov-card"><h4>Contact Info</h4>'+
      (cg.phone?'<div class="ov-row"><span class="ov-label">Phone</span><span class="ov-value">'+esc(cg.phone)+'</span></div>':'')+
      (cg.email?'<div class="ov-row"><span class="ov-label">Email</span><span class="ov-value">'+esc(cg.email)+'</span></div>':'')+
      (addrStr?'<div class="ov-row"><span class="ov-label">Address</span><span class="ov-value">'+esc(addrStr)+'</span></div>':'')+
      (cg.ecName?'<div class="ov-row"><span class="ov-label">Emergency Contact</span><span class="ov-value">'+esc(cg.ecName)+(cg.ecPhone?' · '+cg.ecPhone:'')+'</span></div>':'')+
    '</div>'+
    '<div class="ov-card"><h4>Employment</h4>'+
      (cg.emptype?'<div class="ov-row"><span class="ov-label">Type</span><span class="ov-value">'+esc(cg.emptype)+'</span></div>':'')+
      (cg.hireDate?'<div class="ov-row"><span class="ov-label">Hire Date</span><span class="ov-value">'+esc(cg.hireDate)+'</span></div>':'')+
      (cg.payRate?'<div class="ov-row"><span class="ov-label">Pay Rate</span><span class="ov-value">$'+esc(cg.payRate)+'/hr</span></div>':'')+
      (cg.maxHours?'<div class="ov-row"><span class="ov-label">Max Hours/Week</span><span class="ov-value">'+esc(cg.maxHours)+'</span></div>':'')+
    '</div>'+
    '<div class="ov-card"><h4>Certifications</h4>'+
      (certsStr?'<div style="font-size:13px;color:#1a2b45;line-height:1.5;white-space:pre-wrap;">'+esc(certsStr)+'</div>':'<div style="color:#8ca0b4;font-size:12px;">No certifications on file.</div>')+
    '</div>'+
    '<div class="ov-card"><h4>Assigned Clients</h4>'+
      '<div class="ov-row" style="cursor:pointer;" onclick="switchCgTab(\'clients\')">'+
        '<span class="ov-label">Clients</span>'+
        '<span class="ov-value" style="color:#185FA5;font-size:18px;font-weight:700;">'+assigned.length+'</span>'+
      '</div>'+
      (assigned.length?'<div style="font-size:12px;color:#4a6a8a;margin-top:6px;">'+assigned.slice(0,3).map(function(n){return esc(n);}).join(', ')+(assigned.length>3?' + '+(assigned.length-3)+' more':'')+'</div>':'<div style="font-size:12px;color:#8ca0b4;">No clients assigned.</div>')+
    '</div>'+
  '</div>';
}
function renderCgNotesPane(){
  var cg=getCaregivers()[activeCgId];
  var c=document.getElementById('cgNotesContent');
  if(!c||!cg)return;
  c.innerHTML='<textarea id="cgNotesArea" style="width:100%;min-height:200px;padding:12px;border:1px solid #d0d8e4;border-radius:6px;font-size:13px;font-family:Arial,sans-serif;outline:none;resize:vertical;max-width:620px;">'+esc(cg.notes||'')+'</textarea>'+
    '<div style="margin-top:10px;"><button class="btn btn-primary" onclick="saveCgNotes()">Save Notes</button></div>';
}
function saveCgNotes(){
  var cgs=getCaregivers();
  if(!cgs[activeCgId])return;
  var area=document.getElementById('cgNotesArea');
  if(!area)return;
  cgs[activeCgId].notes=area.value;
  saveCaregiversLS(cgs);
  saveCaregiverAPI(activeCgId,cgs[activeCgId]);
  var btn=document.querySelector('#cgNotesContent .btn-primary');
  if(btn){btn.textContent='Saved';setTimeout(function(){btn.textContent='Save Notes';},1800);}
}
function renderCgAuditPane(){
  var pane=document.getElementById('cgpane-audit');
  var cg=getCaregivers()[activeCgId];
  if(!pane||!cg)return;
  pane.innerHTML='<div style="color:#8ca0b4;font-size:13px;padding:8px 0;">Loading…</div>';
  if(spToken){
    fetch(API_BASE+'/audit?client='+encodeURIComponent(cg.name||activeCgId)+'&limit=100',{headers:apiHeaders()})
      .then(function(r){return r.ok?r.json():Promise.reject();})
      .then(function(rows){
        pane.innerHTML='';
        if(!rows.length){pane.innerHTML='<div style="color:#8ca0b4;font-size:13px;padding:16px 0;">No audit entries yet.</div>';return;}
        var wrap=document.createElement('div');wrap.style.cssText='max-width:600px;';
        rows.forEach(function(e){
          var row=document.createElement('div');row.className='audit-row';
          var ts=e.created_at?new Date(e.created_at).toLocaleString():e.ts||'';
          row.innerHTML='<span class="audit-icon">—</span><div><div class="audit-text">'+esc(e.action)+'</div><div class="audit-who">'+esc(e.who)+' · '+esc(ts)+'</div></div>';
          wrap.appendChild(row);
        });
        pane.appendChild(wrap);
      })
      .catch(function(){pane.innerHTML='<div style="color:#8ca0b4;font-size:13px;padding:16px 0;">No audit entries yet.</div>';});
  } else {
    pane.innerHTML='<div style="color:#8ca0b4;font-size:13px;padding:16px 0;">Sign in to view audit log.</div>';
  }
}
function renderCgInfoPane(){
  if(!activeCgId)return;
  var cg=getCaregivers()[activeCgId];
  var c=document.getElementById('cgInfoContent');c.innerHTML='';
  if(!cg)return;

  var g=document.createElement('div');g.className='info-grid';
  c.appendChild(g);

  function mkF(id,label,val,full){
    var d=document.createElement('div');d.className='info-field'+(full?' full':'');
    d.innerHTML='<label>'+label+'</label><input id="'+id+'" value="'+esc(val||'')+'">';
    g.appendChild(d);
  }
  function mkDiv(label){
    var d=document.createElement('div');d.className='form-section-divider full';d.innerHTML='<span>'+label+'</span>';g.appendChild(d);
  }
  function mkRow(children){
    var d=document.createElement('div');d.className='info-field-row full';d.innerHTML=children;g.appendChild(d);
  }

  // Name row (3 cols)
  var dName=document.createElement('div');dName.className='info-field-row full';dName.style.gridTemplateColumns='1fr 1fr 1fr';
  var storedFirst=cg.firstName||cg.first_name||(cg.name||'').split(' ')[0]||'';
  var storedLast=cg.lastName||cg.last_name||(cg.name||'').split(' ').slice(1).join(' ')||'';
  dName.innerHTML='<div class="info-field"><label>First Name *</label><input id="cgi-first" value="'+esc(storedFirst)+'"></div>'+
    '<div class="info-field"><label>Middle Name</label><input id="cgi-middle" value="'+esc(cg.middleName||cg.middle_name||'')+'"></div>'+
    '<div class="info-field"><label>Last Name *</label><input id="cgi-last" value="'+esc(storedLast)+'"></div>';
  g.appendChild(dName);

  // DOB + Gender row (moved up — needed for state forms / identity)
  mkRow('<div class="info-field"><label>Date of Birth</label><input id="cgi-dob" type="date" value="'+esc(cg.dob||cg.dateOfBirth||'')+'"></div>'+
    '<div class="info-field"><label>Gender</label><select id="cgi-gender"><option value=""'+(!cg.gender?' selected':'')+'>—</option><option value="Male"'+(cg.gender==='Male'?' selected':'')+'>Male</option><option value="Female"'+(cg.gender==='Female'?' selected':'')+'>Female</option></select></div>');

  // Nickname + Status row
  mkRow('<div class="info-field"><label>Nickname</label><input id="cgi-nickname" value="'+esc(cg.nickname||'')+'"></div>'+
    '<div class="info-field"><label>Status</label><select id="cgi-status"><option value="active"'+((!cg.status||cg.status==="active")?" selected":"")+'>Active</option><option value="inactive"'+(cg.status==="inactive"?" selected":"")+'>Inactive</option><option value="terminated"'+(cg.status==="terminated"?" selected":"")+'>Terminated</option></select></div>');

  mkF('cgi-phone','Phone',cg.phone,false);
  mkF('cgi-email','Email',cg.email,false);
  mkF('cgi-dl',"Driver's License #",cg.driversLicense,false);
  // SSN masked with Show/Hide toggle
  var cgSsnDiv=document.createElement('div');cgSsnDiv.className='info-field';
  cgSsnDiv.innerHTML='<label>Social Security #</label>'+
    '<div style="display:flex;gap:4px;align-items:center;">'+
      '<input id="cgi-ssn" type="password" autocomplete="off" value="'+esc(cg.ssn||'')+'" style="flex:1;">'+
      '<button type="button" class="btn btn-secondary btn-sm" onclick="toggleMask(\'cgi-ssn\',this)" style="padding:4px 8px;font-size:11px;white-space:nowrap;">Show</button>'+
    '</div>';
  g.appendChild(cgSsnDiv);
  mkF('cgi-street','Street',cg.street||cg.address,true);
  mkRow('<div class="info-field"><label>City</label><input id="cgi-city" value="'+esc(cg.city||'')+'"></div>'+
    '<div class="info-field"><label>State</label><input id="cgi-state" value="'+esc(cg.state||'')+'"></div>');
  mkRow('<div class="info-field"><label>ZIP</label><input id="cgi-zip" value="'+esc(cg.zip||'')+'" oninput="lookupZip(\'cgi-zip\',\'cgi-city\',\'cgi-state\',\'cgi-county\')"></div>'+
    '<div class="info-field"><label>County</label><input id="cgi-county" value="'+esc(cg.county||'')+'"></div>');

  mkDiv('Employment');
  mkRow('<div class="info-field"><label>Hire Date</label><input id="cgi-hire" type="date" value="'+esc(cg.hireDate||'')+'"></div>'+
    '<div class="info-field"><label>Employment Type</label><select id="cgi-emptype"><option value="full-time"'+(cg.emptype==='full-time'?' selected':'')+'>Full-Time</option><option value="part-time"'+(cg.emptype==='part-time'?' selected':'')+'>Part-Time</option><option value="per-diem"'+(cg.emptype==='per-diem'?' selected':'')+'>Per Diem</option></select></div>');
  mkRow('<div class="info-field"><label>Pay Rate ($/hr)</label><input id="cgi-pay" value="'+esc(cg.payRate||'')+'"></div>'+
    '<div class="info-field"><label>Max Hours/Week</label><input id="cgi-hours" value="'+esc(cg.maxHours||'')+'"></div>');

  var certDiv=document.createElement('div');certDiv.className='info-field full';
  certDiv.innerHTML='<label>Certifications &amp; Training</label><textarea id="cgi-certs" rows="2" style="width:100%;padding:7px 10px;border:1px solid #d0d8e4;border-radius:5px;font-size:13px;font-family:Arial,sans-serif;outline:none;resize:vertical;">'+esc(cg.certs||cg.certifications||'')+'</textarea>';
  g.appendChild(certDiv);

  mkDiv('Emergency Contact');
  mkRow('<div class="info-field"><label>Contact Name</label><input id="cgi-ecname" value="'+esc(cg.ecName||'')+'"></div>'+
    '<div class="info-field"><label>Contact Phone</label><input id="cgi-ecphone" value="'+esc(cg.ecPhone||'')+'"></div>');

  // Save + Delete buttons
  var actions=document.createElement('div');actions.style.cssText='margin-top:16px;display:flex;gap:8px;';
  actions.innerHTML='<button class="btn btn-primary" id="cgSaveInfoBtn" onclick="saveCgInfoPane()">Save Changes</button>'+
    '<button class="btn btn-danger btn-sm" onclick="deleteCaregiverFromDetail()" style="padding:6px 14px;">Delete Caregiver</button>';
  c.appendChild(actions);
}
function saveCgInfoPane(){
  if(!activeCgId)return;
  var cgs=getCaregivers();var cg=cgs[activeCgId];if(!cg)return;
  var first=(document.getElementById('cgi-first').value||'').trim();
  var last=(document.getElementById('cgi-last').value||'').trim();
  if(!first||!last){showAlert('First and last name are required.');return;}
  cg.firstName=first;cg.middleName=document.getElementById('cgi-middle').value;cg.lastName=last;
  cg.name=(first+(document.getElementById('cgi-middle').value?' '+document.getElementById('cgi-middle').value:'')+' '+last).trim();
  cg.nickname=document.getElementById('cgi-nickname').value;
  cg.status=document.getElementById('cgi-status').value;
  cg.phone=document.getElementById('cgi-phone').value;cg.email=document.getElementById('cgi-email').value;
  var cgiDl=document.getElementById('cgi-dl');if(cgiDl)cg.driversLicense=cgiDl.value;
  var cgiSsn=document.getElementById('cgi-ssn');if(cgiSsn)cg.ssn=cgiSsn.value;
  cg.street=document.getElementById('cgi-street').value;cg.city=document.getElementById('cgi-city').value;
  cg.state=document.getElementById('cgi-state').value;cg.zip=document.getElementById('cgi-zip').value;cg.county=document.getElementById('cgi-county').value;
  var cgiDob=document.getElementById('cgi-dob');if(cgiDob)cg.dob=cgiDob.value;
  var cgiGender=document.getElementById('cgi-gender');if(cgiGender)cg.gender=cgiGender.value;
  cg.hireDate=document.getElementById('cgi-hire').value;cg.emptype=document.getElementById('cgi-emptype').value;
  cg.payRate=document.getElementById('cgi-pay').value;cg.maxHours=document.getElementById('cgi-hours').value;
  cg.certifications=document.getElementById('cgi-certs').value;
  cg.ecName=document.getElementById('cgi-ecname').value;cg.ecPhone=document.getElementById('cgi-ecphone').value;
  saveCaregiversLS(cgs);saveCaregiverAPI(activeCgId,cg);
  // Update header
  document.getElementById('cgDetailName').textContent=cg.name;
  var st=cg.status||'active';
  document.getElementById('cgDetailMeta').innerHTML=esc(cg.emptype||'')+(cg.payRate?' · $'+cg.payRate+'/hr':'')+' &nbsp;<span class="cs-badge cs-'+st+'">'+st.charAt(0).toUpperCase()+st.slice(1)+'</span>';
  var btn=document.getElementById('cgSaveInfoBtn');if(btn){btn.textContent='Saved ✓';setTimeout(function(){btn.textContent='Save Changes';},1800);}
  addAuditEntry(cg.name,'Caregiver profile updated');
  renderCaregiverGrid();
}
function renderCgClientsPane(){
  if(!activeCgId)return;
  var c=document.getElementById('cgClientsContent');
  var profiles=getProfiles();
  var assigned=Object.keys(profiles).filter(function(k){return profiles[k].caregiverId===activeCgId;});
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
        '<div style="font-size:11px;color:#6b8dae;">'+(prof.medicaidId||'No Medicaid ID')+(prof.phone?' · '+prof.phone:'')+'</div>'+
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
  c.innerHTML=
    '<div class="doc-upload-card">'+
      '<div class="doc-upload-head">'+
        '<h4>Documents for '+esc(cgName)+'</h4>'+
        '<p>Upload SSN card, driver\'s license, certifications, etc.</p>'+
      '</div>'+
      '<div class="doc-upload-row">'+
        '<div class="doc-upload-fields">'+
          '<label>Category</label>'+
          '<select id="cgDocCategory">'+
            '<option value="Other">Other</option>'+
            '<option value="SSN_Card">SSN Card</option>'+
            '<option value="Drivers_License">Driver\'s License</option>'+
            '<option value="Insurance_Card">Insurance Card</option>'+
            '<option value="Certification">Certification</option>'+
            '<option value="Background_Check">Background Check</option>'+
            '<option value="I9_W4">I-9 / W-4</option>'+
          '</select>'+
          '<label style="margin-top:8px;">File</label>'+
          '<input type="file" id="cgDocFileInput2" accept=".pdf,.jpg,.jpeg,.png,.doc,.docx" multiple>'+
        '</div>'+
        '<div class="doc-upload-actions">'+
          '<button class="btn btn-primary" onclick="uploadCgDocAzure()">Upload</button>'+
          '<input type="file" id="cgDocScanInput" accept="image/*" capture="environment" style="display:none;" onchange="handleCgDocScan(this)">'+
          '<button class="btn btn-secondary" onclick="document.getElementById(\'cgDocScanInput\').click()">Scan / Photo</button>'+
        '</div>'+
      '</div>'+
      '<span id="cgDocUploadStatus" class="doc-upload-status"></span>'+
    '</div>'+
    '<div id="cgDocListAzure"><div style="color:#8ca0b4;font-size:13px;">Loading...</div></div>';
  loadCgDocsAzure(activeCgId);
}
function loadCgDocsAzure(cgId){
  fetch(API_BASE+'/documents?clientType=caregiver&clientId='+cgId,{headers:apiHeaders()})
  .then(function(r){return r.json();})
  .then(function(docs){renderCgDocListAzure(cgId,docs||[]);})
  .catch(function(){renderCgDocListAzure(cgId,[]);});
}
function renderCgDocListAzure(cgId,docs){
  var list=document.getElementById('cgDocListAzure');if(!list)return;
  if(!docs.length){list.innerHTML='<div style="color:#8ca0b4;font-size:13px;padding:4px 0;">No documents yet.</div>';return;}
  list.innerHTML='';
  var categoryLabels={SSN_Card:'SSN Card',Drivers_License:"Driver's License",Insurance_Card:'Insurance Card',Certification:'Certification',Background_Check:'Background Check',I9_W4:'I-9 / W-4',Other:'Other'};
  docs.forEach(function(d){
    var kb=d.size?Math.round(d.size/1024)+'KB':'';
    var ext=(d.name||'').split('.').pop().toLowerCase();
    var isImg=['jpg','jpeg','png','gif'].indexOf(ext)>=0;
    var icon=(ext||"").toUpperCase().slice(0,4);
    // parse category prefix from filename: "SSN_Card__filename.pdf"
    var parts=d.name.split('__');
    var cat=parts.length>1?parts[0]:'Other';
    var displayName=parts.length>1?parts.slice(1).join('__'):d.name;
    var catLabel=categoryLabels[cat]||cat;
    var card=document.createElement('div');
    card.style.cssText='display:flex;align-items:center;gap:10px;padding:8px 10px;border:1px solid #e1e8f0;border-radius:6px;margin-bottom:6px;background:#fafbfc;';
    card.innerHTML=
      '<span style="display:inline-block;min-width:34px;padding:3px 6px;background:#e8eef5;color:#1a3a5c;border-radius:4px;font-size:10px;font-weight:600;text-align:center;letter-spacing:.3px;">'+(icon||'FILE')+'</span>'+
      '<div style="flex:1;min-width:0;">'+
        '<div style="font-size:12px;font-weight:600;color:#1a3a5c;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;"><a href="'+d.url+'" target="_blank" style="color:#1a3a5c;text-decoration:none;">'+esc(displayName)+'</a></div>'+
        '<div style="font-size:11px;color:#8ca0b4;">'+
          '<span style="background:#dbeafe;color:#1e40af;padding:1px 7px;border-radius:10px;font-size:10px;font-weight:600;margin-right:6px;">'+esc(catLabel)+'</span>'+
          kb+
        '</div>'+
      '</div>'+
      '<button class="btn btn-danger btn-sm" style="padding:3px 10px;font-size:11px;" onclick="deleteCgDocAzure(\''+cgId+'\',\''+encodeURIComponent(d.name)+'\')">✕</button>';
    list.appendChild(card);
  });
}
function uploadCgDocAzure(){
  var input=document.getElementById('cgDocFileInput2');
  if(!input||!input.files||!input.files.length){showAlert('Please select a file first.');return;}
  var cat=document.getElementById('cgDocCategory').value||'Other';
  var status=document.getElementById('cgDocUploadStatus');status.textContent='Uploading...';
  var fd=new FormData();fd.append('clientType','caregiver');fd.append('clientId',activeCgId);
  Array.from(input.files).forEach(function(f){
    // prefix filename with category
    var prefixedFile=new File([f],cat+'__'+f.name,{type:f.type});
    fd.append('file',prefixedFile);
  });
  fetch(API_BASE+'/documents',{method:'POST',headers:authUploadHeaders(),body:fd})
  .then(function(){status.textContent='';input.value='';loadCgDocsAzure(activeCgId);})
  .catch(function(e){status.textContent='Upload failed: '+e;});
}
function deleteCgDocAzure(cgId,encodedName){
  showConfirm('Delete this document?',function(){
    fetch(API_BASE+'/documents?clientType=caregiver&clientId='+cgId+'&name='+encodedName,{method:'DELETE',headers:apiHeaders()})
    .then(function(){loadCgDocsAzure(cgId);}).catch(function(e){showAlert('Delete failed: '+e);});
  },{title:'Delete Document',okText:'Delete'});
}
function handleCgDocScan(input){
  if(!activeCgId){showAlert('Open a caregiver first.');return;}
  if(!input||!input.files||!input.files.length)return;
  var cat=(document.getElementById('cgDocCategory')&&document.getElementById('cgDocCategory').value)||'Other';
  var status=document.getElementById('cgDocUploadStatus');
  if(status)status.textContent='Uploading scanned image…';
  var fd=new FormData();fd.append('clientType','caregiver');fd.append('clientId',activeCgId);
  var f=input.files[0];
  var prefixedFile=new File([f],cat+'__'+f.name,{type:f.type});
  fd.append('file',prefixedFile);
  fetch(API_BASE+'/documents',{method:'POST',headers:authUploadHeaders(),body:fd})
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
function loadSignaturesAPI(){
  if(!spToken)return;
  fetch(API_BASE+'/signatures',{headers:apiHeaders()})
    .then(function(r){return r.ok?r.json():Promise.reject(r.status);})
    .then(function(rows){
      // rows: [{id,label,data_url,created_at}]
      var sigs=rows.map(function(r){return {id:r.id,label:r.label,data:r.data_url};});
      saveSigsLS(sigs);
      renderSigSettings();
    })
    .catch(function(e){console.error('Signatures load error:',e);});
}
function renderSigSettings(){
  var sigs=getSigs(),list=document.getElementById('sigStoredList');list.innerHTML='';
  if(!sigs.length){list.innerHTML='<div style="color:#8ca0b4;font-size:13px;">No signatures saved yet.</div>';return;}
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
  // Replace handlers
  var newOk=okBtn.cloneNode(true),newCancel=cancelBtn.cloneNode(true);
  okBtn.parentNode.replaceChild(newOk,okBtn);cancelBtn.parentNode.replaceChild(newCancel,cancelBtn);
  newOk.addEventListener('click',function(){modal.classList.remove('open');if(typeof onSave==='function')onSave(inp.value);});
  newCancel.addEventListener('click',function(){modal.classList.remove('open');});
  modal.classList.add('open');
  setTimeout(function(){inp.focus();inp.select();},50);
}

// Toggle masking on a sensitive input (SSN). Auto re-masks after 8 seconds.
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
    // Remove from DB
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
  // Reset upload state
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
      // Second confirmation
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
function emailSuggest(input,suggestId){
  var val=input.value,at=val.indexOf('@'),suggest=document.getElementById(suggestId);
  if(!suggest)return;
  if(at===-1||val.slice(at+1).length===0){
    // Show domain suggestions after @
    if(at!==-1){
      var domains=['michigan.gov','mdhhs.mi.gov','gmail.com'];
      var prefix=val.slice(0,at+1);
      suggest.innerHTML='';
      domains.forEach(function(d){
        var div=document.createElement('div');div.textContent=prefix+d;
        div.addEventListener('mousedown',function(e){e.preventDefault();input.value=prefix+d;suggest.style.display='none';});
        suggest.appendChild(div);
      });
      suggest.style.display='block';
    } else {suggest.style.display='none';}
  } else {
    // partial domain typed
    var typed=val.slice(at+1);
    var domains=['michigan.gov','mdhhs.mi.gov','gmail.com'].filter(function(d){return d.startsWith(typed)&&d!==typed;});
    var prefix=val.slice(0,at+1);
    suggest.innerHTML='';
    domains.forEach(function(d){
      var div=document.createElement('div');div.textContent=prefix+d;
      div.addEventListener('mousedown',function(e){e.preventDefault();input.value=prefix+d;suggest.style.display='none';});
      suggest.appendChild(div);
    });
    suggest.style.display=domains.length?'block':'none';
  }
}
function hideEmailSuggest(id){setTimeout(function(){var el=document.getElementById(id);if(el)el.style.display='none';},150);}

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
      var newName = input.value.trim();
      var newId = cwId();
      var newCw = { id: newId, name: newName, first_name: '', last_name: '', agency: '', phone: '', email: '', fax: '', street: '', city: '', state: '', zip: '', county: '', notes: '' };
      var arr = getCaseworkers(); arr.push(newCw); saveCaseworkersLS(arr);
      saveCaseworkerAPI(newCw);
      if(hidden)hidden.value = newId;
    });
    drop.appendChild(addItem);
  }
  if(!matches.length && !val){
    var noItem=document.createElement('div');noItem.style.cssText='padding:8px 12px;font-size:13px;color:#8ca0b4;';noItem.textContent='No caseworkers yet';drop.appendChild(noItem);
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
  if (!ids.length) {
    var noItem = document.createElement('div');
    noItem.style.cssText = 'padding:8px 12px;font-size:13px;color:#8ca0b4;';
    noItem.textContent = val ? 'No caregivers found' : 'No caregivers yet';
    drop.appendChild(noItem);
  }
  drop.style.display = 'block';
}
// Focus handler — opens the dropdown showing all options
function cgSearchFocus(input, hiddenId, dropId){ cgSearch(input, hiddenId, dropId); }

// ============================================================
//  ACTIVITY LOG
// ============================================================
function getActivity(){try{return JSON.parse(localStorage.getItem('lhca_activity')||'[]');}catch(e){return[];}}
function logActivity(type,text){
  var log=getActivity();
  log.unshift({type:type,text:text,ts:new Date().toLocaleString()});
  if(log.length>40)log=log.slice(0,40);
  try{localStorage.setItem('lhca_activity',JSON.stringify(log));}catch(e){}
  // Persist to DB (fire-and-forget)
  if(spToken){
    fetch(API_BASE+'/audit',{method:'POST',headers:apiHeaders(),body:JSON.stringify({event_type:type,client_name:'',action:text,who:currentUserEmail()})})
      .catch(function(e){console.error('Audit log error:',e);});
  }
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
function getAuditLog(){try{return JSON.parse(localStorage.getItem('lhca_audit')||'[]');}catch(e){return[];}}
function addAuditEntry(clientName,action){
  var who=currentUserEmail();
  var log=getAuditLog();
  log.unshift({client:clientName,action:action,who:who,ts:new Date().toLocaleString()});
  if(log.length>200)log=log.slice(0,200);
  try{localStorage.setItem('lhca_audit',JSON.stringify(log));}catch(e){}
  // Persist to DB (fire-and-forget)
  if(spToken){
    fetch(API_BASE+'/audit',{method:'POST',headers:apiHeaders(),body:JSON.stringify({event_type:'audit',client_name:clientName,action:action,who:who})})
      .catch(function(e){console.error('Audit save error:',e);});
  }
}
function renderAuditPane(){
  if(!activeProfileName)return;
  var pane=document.getElementById('dpane-audit');pane.innerHTML='';
  if(spToken){
    // Load from DB for the active client
    pane.innerHTML='<div style="color:#8ca0b4;font-size:13px;padding:8px 0;">Loading audit history…</div>';
    fetch(API_BASE+'/audit?client='+encodeURIComponent(activeProfileName)+'&limit=100',{headers:apiHeaders()})
      .then(function(r){return r.ok?r.json():Promise.reject(r.status);})
      .then(function(rows){
        pane.innerHTML='';
        if(!rows.length){pane.innerHTML='<div style="color:#8ca0b4;font-size:13px;padding:16px 0;">No audit entries for this client yet.</div>';return;}
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
  if(!log.length){pane.innerHTML='<div style="color:#8ca0b4;font-size:13px;padding:16px 0;">No audit entries for this client yet.</div>';return;}
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
  // Populate client picker
  if(clientSel){
    var profs=getProfiles();
    var names=Object.keys(profs).filter(function(n){return !profs[n].clientStatus||profs[n].clientStatus==='active';}).sort(function(a,b){return a.localeCompare(b);});
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
    saveTodos(getTodos().filter(function(t){return !t.done;}));renderTodos();updateTaskBadge();
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
function sendEmailMailto(toAddr){
  var prof=getProfiles()[activeProfileName]||{};
  var cwRec4=getCaseworkers().find(function(c){return c.id===prof.caseworkerId||c.name===prof.worker;})||{};
  var addr=toAddr||cwRec4.email||'';
  var clientName=activeProfileName||'';
  var subj='Re: '+clientName;
  window.location.href='mailto:'+encodeURIComponent(addr)+'?subject='+encodeURIComponent(subj);
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
    if(!open.length){openC.innerHTML='<div style="color:#8ca0b4;font-size:13px;padding:8px 0;">No open tasks.</div>';}
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
            (isFollowUp?'<div style="font-size:10px;color:#8ca0b4;margin-bottom:2px;">↳ Follow-up</div>':'')+
            '<div class="todo-text" style="font-weight:600;">'+esc(t.text)+'</div>'+
            (t.note?'<div style="font-size:11px;color:#4a6a8a;margin-top:4px;padding:4px 6px;background:#e8f0fb;border-radius:3px;white-space:pre-wrap;">'+esc(t.note)+'</div>':'')+
            '<div class="todo-meta">'+
              (t.due?'<span class="'+(overdue?'todo-due-overdue':'')+'">Due: '+t.due+'</span>':'')+
              (!filterClient&&t.client?'<span class="todo-link" onclick="navDetail(\''+esc(t.client)+'\')">'+esc(t.client)+'</span>':'')+
            '</div>'+
          '</div>'+
          '<button class="btn btn-secondary btn-sm" onclick="editTodoTask(\''+t.id+'\')" style="padding:3px 7px;font-size:11px;" title="Edit task name, due date, or note">Edit</button>'+
          '<button class="btn btn-secondary btn-sm" onclick="addFollowUpTask(\''+(isFollowUp?(t.parentId||t.id):t.id)+'\')" style="padding:3px 7px;font-size:11px;" title="Add a follow-up that comes after this one">+ Follow-up</button>'+
          '<button class="btn btn-danger btn-sm" onclick="deleteTodo(\''+t.id+'\')" style="padding:3px 7px;font-size:11px;">Remove</button>';
        return item;
      }

      parents.forEach(function(parent){
        var fups=followUpsByParent[parent.id]||[];
        // Parent row
        var parentRow=renderTaskRow(parent);
        // If has follow-ups, append a small toggle bar inside the parent row
        if(fups.length){
          var toggleBar=document.createElement('div');
          toggleBar.style.cssText='margin-top:8px;padding:8px 12px;background:#eef3fb;border:1px solid #d0dbe9;border-radius:5px;cursor:pointer;font-size:12px;color:#185FA5;font-weight:600;display:flex;justify-content:space-between;align-items:center;user-select:none;width:100%;transition:background 0.15s;';
          var toggleLabel=document.createElement('span');
          toggleLabel.style.cssText='display:flex;align-items:center;gap:8px;';
          toggleLabel.innerHTML='<span class="fup-arrow" style="display:inline-block;font-size:14px;transition:transform 0.18s;">▶</span><span>'+fups.length+' follow-up'+(fups.length>1?'s':'')+'</span>';
          var clearBtn=document.createElement('span');clearBtn.style.cssText='font-size:11px;color:#8ca0b4;font-weight:400;';clearBtn.textContent='click to expand';
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
    if(!done.length){doneC.innerHTML='<div style="color:#8ca0b4;font-size:13px;padding:8px 0;">Nothing completed yet.</div>';return;}
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
function editTodoNote(id){
  var todos=getTodos(),t=todos.find(function(x){return String(x.id)===String(id);});if(!t)return;
  showPrompt('Note for this task:',t.note||'',function(note){
    var todos2=getTodos(),t2=todos2.find(function(x){return String(x.id)===String(id);});if(!t2)return;
    t2.note=note||'';saveTodos(todos2);saveTaskAPI(t2);renderTodos();
  },{title:'Task Note',okText:'Save Note'});
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
  // Sync to API
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
function openEmailComposer(toAddr,subject,body){
  document.getElementById('emailTo').value=toAddr||'';
  document.getElementById('emailSubject').value=subject||'';
  document.getElementById('emailBody').value=body||'';
  document.getElementById('emailComposerModal').classList.add('open');
}
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
  // --- Invoice Status Summary ---
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
        '<td style="width:200px;"><div class="report-bar"><div class="report-bar-fill" style="width:'+pct+'%;background:'+color+';"></div></div><span style="font-size:10px;color:#8ca0b4;">'+pct+'%</span></td></tr>';
    }).join('')+
    '<tr class="total-row"><td>Total</td><td>'+allInvoices.length+'</td><td></td></tr></tbody></table>';
  c.appendChild(s1);
  // --- Per-Client Invoice Breakdown ---
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
      return '<tr><td style="cursor:pointer;color:#185FA5;" onclick="navDetail(\''+esc(r.name)+'\')">'+esc(r.name)+'</td><td>'+r.total+'</td>'+
        '<td style="color:#1e7e34;font-weight:600;">'+r.paid+'</td><td style="color:'+(r.open?'#b07800':'#aaa')+';font-weight:600;">'+r.open+'</td></tr>';
    }).join('')+'</tbody></table>':'<div style="color:#8ca0b4;font-size:13px;">No client data yet.</div>');
  c.appendChild(s2);
  // --- Monthly Invoice Volume ---
  var byMonth={};
  allInvoices.forEach(function(r){var bp=r.inv.billingPeriod||'Unknown';byMonth[bp]=(byMonth[bp]||0)+1;});
  var months=Object.keys(byMonth).sort(function(a,b){
    var pa=a.split('/'),pb=b.split('/');
    if(pa.length<2||pb.length<2)return 0;
    return (parseInt(pb[1])*12+parseInt(pb[0]))-(parseInt(pa[1])*12+parseInt(pa[0]));
  }).slice(0,12);
  var s3=document.createElement('div');s3.className='report-section';
  var maxV=months.reduce(function(m,k){return Math.max(m,byMonth[k]);},1);
  s3.innerHTML='<h3>Monthly Invoice Volume <span style="font-size:11px;font-weight:normal;color:#8ca0b4;">(last 12 months)</span></h3>'+
    (months.length?'<table class="report-table"><thead><tr><th>Period</th><th>Count</th><th>Visual</th></tr></thead><tbody>'+
    months.map(function(m){
      var n=byMonth[m],pct=Math.round(n/maxV*100);
      return '<tr><td style="font-weight:500;">'+esc(m)+'</td><td>'+n+'</td>'+
        '<td style="width:200px;"><div class="report-bar"><div class="report-bar-fill" style="width:'+pct+'%;"></div></div></td></tr>';
    }).join('')+'</tbody></table>':'<div style="color:#8ca0b4;font-size:13px;">No invoices yet.</div>');
  c.appendChild(s3);
  // --- Missing Invoices for Month ---
  var s4=document.createElement('div');s4.className='report-section';
  // Get all billing periods sorted recent first
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
    return !((profiles[name].invoices)||[]).some(function(i){return i.billingPeriod===defaultPeriod;});
  });
  s4.innerHTML='<h3>Clients Missing Invoice for Month</h3>'+
    '<div style="display:flex;align-items:center;gap:10px;margin-bottom:12px;flex-wrap:wrap;">'+
      '<label style="font-size:12px;color:#6b8dae;">Check period:</label>'+
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
    return !((profiles[name].invoices)||[]).some(function(i){return i.billingPeriod===period;});
  });
  var list=document.getElementById('missingReportList');if(!list)return;
  if(!missing.length){list.innerHTML='<div style="color:#1e7e34;font-size:13px;font-weight:600;">✓ All clients have an invoice for '+esc(period)+'</div>';return;}
  list.innerHTML='<div style="font-size:12px;color:#b07800;font-weight:600;margin-bottom:8px;">'+missing.length+' client'+(missing.length>1?'s':'')+' missing invoice for '+esc(period)+':</div>'+
    missing.map(function(name){
      return '<div style="display:flex;align-items:center;gap:10px;padding:6px 0;border-bottom:1px solid #f0f3f7;">'+
        '<span style="cursor:pointer;color:#185FA5;font-size:13px;" onclick="navDetail(\''+esc(name)+'\')">'+esc(name)+'</span>'+
        '<button class="btn btn-primary btn-sm" style="font-size:11px;" onclick="activeProfileName=\''+esc(name)+'\';navInvoice()">+ Create Invoice</button>'+
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
function getProfiles(){try{return JSON.parse(localStorage.getItem('lhca_profiles')||'{}');}catch(e){return{};}}
function saveProfilesLS(p){localStorage.setItem('lhca_profiles',JSON.stringify(p));}
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
//  /Liberty Home Care Backups/. Keeps 12 most recent (auto-deletes older).
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
  if(btn&&!silent){btn.disabled=true;btn.textContent='Uploading…';}
  try{
    // Build the backup payload (same logic as JSON export)
    var p=getProfiles();
    if(!Object.keys(p).length){throw new Error('No clients yet.');}
    var copy=JSON.parse(JSON.stringify(p));
    if(!includeFullSSN){
      Object.keys(copy).forEach(function(name){
        var prof=copy[name];if(prof.ssn){
          var digits=String(prof.ssn).replace(/\D/g,'');
          prof.ssn=digits.length>=4?'***-**-'+digits.slice(-4):'***-**-****';
        }
      });
    }
    var payload={
      _exportedAt:new Date().toISOString(),
      _exportedBy:(window.signedInEmail||'unknown'),
      _includesFullSSN:!!includeFullSSN,
      _appVersion:'liberty-homecare-v13',
      caregivers:getCaregivers(),
      caseworkers:getCaseworkers(),
      signatures:getSigs(),
      clients:copy
    };
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

function doExportProfiles(includeFullSSN){
  var p=getProfiles();
  // Deep clone so we don't mutate live data
  var copy=JSON.parse(JSON.stringify(p));
  if(!includeFullSSN){
    Object.keys(copy).forEach(function(name){
      var prof=copy[name];if(prof.ssn){
        var digits=String(prof.ssn).replace(/\D/g,'');
        prof.ssn=digits.length>=4?'***-**-'+digits.slice(-4):'***-**-****';
      }
    });
  }
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
      'Status':p.status||'active',
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
        'Hours (HH.MM)':((inv.data&&inv.data.svcHH)||'')+'.'+((inv.data&&inv.data.svcMM)||''),
        'Hourly Rate':(inv.data&&inv.data.hourlyRate)||'27.00',
        'Caseworker':p.worker||'',
        'Note':inv.invoiceNote||''
      });
    });
  });
  // Sheet 3: Caregivers (SSN masked to last 4)
  var caregiversRows=Object.keys(caregivers).map(function(id){
    var c=caregivers[id];
    return {'ID':id,'Name':c.name||'','First':c.firstName||'','Last':c.lastName||'','Status':c.status||'','Role':c.emptype||'','Phone':c.phone||'','Email':c.email||'',"Driver's License":c.driversLicense||'','SSN (masked)':maskSSN(c.ssn),'Hire Date':c.hireDate||'','Notes':(c.notes||'').slice(0,500)};
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
          // Convert to bytes
          var bin=atob(base64);
          var bytes=new Uint8Array(bin.length);
          for(var k=0;k<bin.length;k++)bytes[k]=bin.charCodeAt(k);
          var pdfName=(inv.billingPeriod||'invoice').replace(/\//g,'_')+'.pdf';
          clientFolder.file(pdfName,bytes);
        }catch(e){console.error('PDF gen failed for '+name+' '+inv.billingPeriod,e);}
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
function importProfiles(ev){
  var file=ev.target.files[0];if(!file)return;
  var r=new FileReader();r.onload=function(e){
    try{
      var imp=JSON.parse(e.target.result),ex=getProfiles(),keys=Object.keys(imp);
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
        saveProfilesLS(Object.assign({},ex,imp));renderSidebarClients();renderClientGrid();updateStats();
        showConfirm('Imported '+keys.length+' client'+(keys.length>1?'s':'')+'.',function(){},{title:'Import Complete',okText:'OK',danger:false});
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
function esc(s){return String(s||'').replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;');}

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
  // Hourly rate on invoice is always the government billing rate ($27), not caregiver pay rate
  document.getElementById('hourlyRate').value='27.00';
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
  var f=['clientName','medicaidId','worker','billingPeriod','svcHH','svcMM','cplxHH','cplxMM','p1HH','p1MM','grandHH','grandMM','dateSubmitted','sigDate1','sigDate2'];
  f.forEach(function(id){var el=document.getElementById(id);if(el&&data[id]!==undefined)el.value=data[id];});
  // Hourly rate is always the government billing rate
  document.getElementById('hourlyRate').value='27.00';
  // Bill To: always from caseworker billing code, not whatever was saved on the invoice
  var profCur=(activeProfileName&&getProfiles()[activeProfileName])||{};
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
function startDraftAutosave(){
  if(draftTimer)clearInterval(draftTimer);
  draftTimer=setInterval(function(){
    if(!activeProfileName)return;
    var d=captureFullInvoice();
    if(!d.billingPeriod)return; // don't save empty drafts
    try{localStorage.setItem('lhca_draft_'+activeProfileName,JSON.stringify(d));}catch(e){}
    var badge=document.getElementById('draftBadge');
    if(badge){badge.style.display='inline-block';badge.textContent='● Draft autosaved';setTimeout(function(){badge.style.display='none';},2000);}
  },30000);
}

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
    // Second cell: plain day number
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
  // Strip non-digits
  var raw=el.value.replace(/\D/g,'');
  // Auto-format as user types
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
function onMonthPickerChange(el){
  // Legacy — kept for compatibility if ever re-added
  var val=el.value;if(!val)return;
  var parts=val.split('-'),y=parts[0],m=parts[1];
  var formatted=m+'/'+y;
  document.getElementById('billingPeriod').value=formatted;
  document.getElementById('billingPeriod2').value=formatted;
  rebuild(daysIn(m,y));
  checkDuplicatePeriod(formatted);
}
function syncMonthPickerFromHidden(){syncBillingPeriodFields();}
function checkDuplicatePeriod(bp){
  var warn=document.getElementById('dupWarning');if(!warn||!activeProfileName)return;
  var p=getProfiles();if(!p[activeProfileName])return;
  var ex=((p[activeProfileName].invoices)||[]).find(function(i){return i.billingPeriod===bp;});
  if(ex){
    var st=ex.status||'draft';
    var msg=st==='paid'?
      '[Locked] A <strong>Paid</strong> invoice already exists for '+bp+'. It is locked and cannot be overwritten.':
      'Warning: An invoice for '+bp+' already exists (status: '+st+'). Saving will overwrite it.';
    warn.innerHTML=msg;warn.style.display='block';
  } else {warn.style.display='none';}
}
function onBillingInput(el){onBillingTextInput(el);}
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
    states[key].forEach(function(rowState,i){if(!rows[i])return;var cells=rows[i].querySelectorAll('td.mc');rowState.forEach(function(on,j){if(cells[j])cells[j].classList.toggle('on',on);});});
  });
}

// FIX #6: copyMonth preserves clientName and medicaidId
// ── Toast ─────────────────────────────────────────────────────
function showToast(msg,ms){
  var t=document.getElementById('lhcaToast');if(!t){t=document.createElement('div');t.id='lhcaToast';t.className='lhca-toast';document.body.appendChild(t);}
  t.textContent=msg;t.classList.add('show');
  clearTimeout(t._tid);t._tid=setTimeout(function(){t.classList.remove('show');},ms||3500);
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
    pdf.text((data.svcHH||'')+'.'+(data.svcMM||''),x0+5+lblWidth+10,y+11);
  } else {
    pdf.setFontSize(9);
    var lblY=y+14;
    function dualText(label,val,xPos){
      pdf.setFont('helvetica','bold');pdf.text(label,xPos,lblY);
      var w=pdf.getTextWidth(label);
      pdf.setFont('helvetica','normal');pdf.text(val,xPos+w+4,lblY);
    }
    dualText('Total Time for Services Above:',(data.cplxHH||'')+'.'+(data.cplxMM||''),x0+5);
    dualText('Total Time from Previous Page:',(data.p1HH||'')+'.'+(data.p1MM||''),x0+200);
    dualText('Total Time for Billing Period:',(data.grandHH||'')+'.'+(data.grandMM||''),x0+395);
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
async function sendMailWithPDF(toEmail,subject,bodyHtml,attachments,onProgress){
  if(!spToken)return {ok:false,err:'Not signed in'};
  // Calculate total attachment size in bytes (base64 inflates ~33%, so actual=base64Length*0.75)
  var totalBytes=0;
  attachments.forEach(function(a){totalBytes+=Math.floor(a.base64.length*0.75);});
  var THRESHOLD_BYTES=3.5*1024*1024; // stay safely under Graph's 4 MB sendMail cap
  if(totalBytes<=THRESHOLD_BYTES){
    return _sendMailDirect(toEmail,subject,bodyHtml,attachments);
  }
  console.log('[Graph] Total attachments '+(totalBytes/1024/1024).toFixed(2)+' MB exceeds 3.5 MB — using upload-session flow');
  return _sendMailUploadSession(toEmail,subject,bodyHtml,attachments,onProgress);
}

// Fast path — direct sendMail with attachments inline (≤4 MB total)
async function _sendMailDirect(toEmail,subject,bodyHtml,attachments){
  var msg={subject:subject,body:{contentType:'HTML',content:bodyHtml},toRecipients:[{emailAddress:{address:toEmail}}],attachments:attachments.map(function(a){return{'@odata.type':'#microsoft.graph.fileAttachment',name:a.name,contentType:'application/pdf',contentBytes:a.base64};})};
  try{
    var resp=await fetch('https://graph.microsoft.com/v1.0/me/sendMail',{method:'POST',headers:{'Authorization':'Bearer '+spToken,'Content-Type':'application/json'},body:JSON.stringify({message:msg,saveToSentItems:true})});
    if(resp.ok||resp.status===202||resp.status===204)return {ok:true};
    var errText=await resp.text();
    console.error('Graph sendMail failed',resp.status,errText);
    return {ok:false,status:resp.status,err:errText};
  }catch(e){console.error('Graph email error:',e);return {ok:false,err:e.message};}
}

// Upload-session path — for emails > 4 MB. Creates draft, uploads each attachment via session, then sends.
async function _sendMailUploadSession(toEmail,subject,bodyHtml,attachments,onProgress){
  function progress(done,total,label){if(typeof onProgress==='function')onProgress(done,total,label);}
  try{
    // 1) Create draft message (no attachments yet)
    progress(0,attachments.length,'Creating draft…');
    var draftBody={subject:subject,body:{contentType:'HTML',content:bodyHtml},toRecipients:[{emailAddress:{address:toEmail}}]};
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
  document.getElementById('hourlyRate').value='27.00';
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
  var arr=getEmailAuditLog();
  if(!arr.length){wrap.innerHTML='<div style="padding:14px;color:#8ca0b4;text-align:center;">No log entries yet.</div>';return;}
  // Newest first
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
  var arr=getEmailAuditLog();
  if(!arr.length){showAlert('No audit log entries to download.');return;}
  function csvEscape(v){var s=String(v==null?'':v);return /[",\n]/.test(s)?'"'+s.replace(/"/g,'""')+'"':s;}
  var header=['Timestamp','SentBy','Type','Recipient','CaseworkerName','BillingPeriod','AttachmentCount','ClientNames','Success','ErrorMsg'].join(',');
  var lines=arr.map(function(r){return [
    r.timestamp,r.sentBy,r.type,r.recipient,r.caseworkerName,r.billingPeriod,r.attachmentCount,(r.clientNames||[]).join('; '),r.success?'YES':'NO',r.errorMsg||''
  ].map(csvEscape).join(',');});
  var csv=header+'\n'+lines.join('\n');
  var blob=new Blob([csv],{type:'text/csv;charset=utf-8;'});
  var url=URL.createObjectURL(blob);
  var a=document.createElement('a');a.href=url;
  a.download='email_audit_log_'+new Date().toISOString().slice(0,10)+'.csv';
  document.body.appendChild(a);a.click();document.body.removeChild(a);
  setTimeout(function(){URL.revokeObjectURL(url);},2000);
}

function logEmailSend(entry){
  // entry: { recipient, caseworkerName, billingPeriod, clientNames[], attachmentCount, sentBy, success, errorMsg? }
  var record=Object.assign({
    timestamp:new Date().toISOString(),
    sentBy:(window.signedInEmail||'unknown')
  },entry);
  var arr=getEmailAuditLog();arr.push(record);saveEmailAuditLog(arr);
  // Mirror to App Insights for centralized tracking
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
  saveProfileSP(clientName,p[clientName]);
  if(inv.dbId){
    fetch(API_BASE+'/invoices/'+inv.dbId+'/status',{method:'PATCH',headers:apiHeaders(),body:JSON.stringify({status:'submitted'})}).catch(function(e){console.error(e);});
  }
}

// ── Send single invoice email ─────────────────────────────────
async function sendEmail(){
  var cn=document.getElementById('clientName').value.trim();
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
    var subj='Home Help Agency Invoice'+(cn?' – '+cn:'')+(bp?' – '+bp:'');
    var body='<p>Dear '+(w||'Caseworker')+',</p>'+
      '<p>Please find attached the Home Help Agency Invoice'+(cn?' for <b>'+cn+'</b>':'')+(bp?' for billing period <b>'+bp+'</b>':'')+'. Please review and process at your earliest convenience.</p>'+
      '<p>Please do not hesitate to contact us with any questions.</p>'+
      '<p>Thank you,<br><b>Thomas Jaboro</b><br>Liberty Home Care Assistance<br>(248) 291-4106</p>';
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
  document.getElementById('hourlyRate').value='27.00';document.getElementById('billTo').value=bt;
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
  // Multiple sigs — show picker
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
  // Trim padding
  var trimmed=trimCanvasToContent(sigCanvas,4);if(trimmed)data=trimmed;
  var label=document.getElementById('sigLabel').value.trim()||'Signature';
  var id=sigId();
  var sigs=getSigs();sigs.push({id:id,label:label,data:data});saveSigsLS(sigs);
  // Persist to DB
  if(spToken){
    fetch(API_BASE+'/signatures',{method:'POST',headers:apiHeaders(),body:JSON.stringify({id:id,label:label,data_url:data})})
      .catch(function(e){console.error('Sig save error:',e);});
  }
  closeSigModal();
  if(pendingSigTarget!==null){stampSignatureData(pendingSigTarget,data);}
  if(document.getElementById('page-settings').classList.contains('active'))renderSigSettings();
}
function clearSigPad(){if(sigCtx)sigCtx.clearRect(0,0,sigCanvas.width,sigCanvas.height);}
function closeSigModal(){document.getElementById('sigModal').classList.remove('open');}
function changeSignature(){placeSignature(1);}// Alias for topbar button

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
    // AND the row above/below it are NOT both equally dense (so we don't erase real glyph rows)
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
  if(spToken){
    fetch(API_BASE+'/signatures',{method:'POST',headers:apiHeaders(),body:JSON.stringify({id:id,label:label,data_url:data})})
      .catch(function(e){console.error('Sig save error:',e);});
  }
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
  // Shrink to fit if too wide
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
  // Persist to DB
  if(spToken){
    fetch(API_BASE+'/signatures',{method:'POST',headers:apiHeaders(),body:JSON.stringify({id:id,label:label,data_url:data})})
      .catch(function(e){console.error('Sig save error:',e);});
  }
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
  // Restore complex section
  var cs=document.getElementById('complexSection');
  if(cs.getAttribute('data-print-hidden')==='1'){
    cs.removeAttribute('data-print-hidden');
    if(document.getElementById('showComplex').checked)cs.style.display='block';
  }
  // Restore column headers to screen version
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
    cache: { cacheLocation: 'localStorage' },
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
          setTimeout(refreshApiToken, Math.max(ttl, 60000));
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
  // Remove all PHI/PII from localStorage — required for HIPAA compliance
  var phiKeys = ['lhca_profiles','lhca_caregivers','lhca_signatures','lhca_sig',
    'lhca_caseworkers','lhca_todos','lhca_activity','lhca_audit','lhca_id_map'];
  phiKeys.forEach(function(k){try{localStorage.removeItem(k);}catch(e){}});
  Object.keys(localStorage).filter(function(k){return k.startsWith('lhca_draft_');})
    .forEach(function(k){try{localStorage.removeItem(k);}catch(e){}});
}
function signOut() {
  aiTrack('UserSignOut');
  clearPHIFromStorage();
  spToken = null;
  _apiToken = null;
  msalInstance.logoutPopup({ redirectUri: REDIRECT_URI });
}
function updateAuthUI(on) {
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
  document.body.classList.add('db-syncing');
}
function syncEnd(){
  window._dbSyncPending=Math.max(0,window._dbSyncPending-1);
  if(window._dbSyncPending===0){
    document.body.classList.remove('db-syncing');
  }
}

// ── LOAD all clients + invoices from Azure SQL (bulk fetch) ───────────────
// One round-trip instead of N+1. Backend returns clients with invoices nested.
function loadProfilesAPI() {
  syncStart();
  fetch(API_BASE + '/homecare-clients-with-invoices', { headers: apiHeaders() })
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
        var seenByPeriod = {};
        var dedupedInvs = [];
        invs.forEach(function (inv) {
          var key = (inv.billing_period || '').trim();
          if (!key) { dedupedInvs.push(inv); return; }
          if (seenByPeriod[key]) {
            console.warn('[invoices] Duplicate for ' + name + ' ' + key + ' — deleting old DB id ' + inv.id);
            fetch(API_BASE + '/invoices/' + inv.id, { method: 'DELETE', headers: apiHeaders() })
              .catch(function (e) { console.error('Failed to delete duplicate invoice:', e); });
            return;
          }
          seenByPeriod[key] = true;
          dedupedInvs.push(inv);
        });
        var mappedInvs = dedupedInvs.map(function (inv) {
          var data = null; try { data = JSON.parse(inv.invoice_data || 'null'); } catch (e) {}
          return {
            dbId: inv.id, billingPeriod: inv.billing_period || '',
            status: inv.status || 'draft', invoiceNote: inv.invoice_note || '',
            savedAt: inv.saved_at ? new Date(inv.saved_at).toLocaleString() : '', data: data,
          };
        });
        profiles[name] = {
          clientName: name, firstName: c.first_name || '', lastName: c.last_name || '',
          middleName: c.middle_name || '', nickname: c.nickname || '',
          medicaidId: c.medicaid_id || '', hourlyRate: c.hourly_rate || '',
          worker: c.worker || '', caseworkerId: c.caseworker_id || '',
          street: c.street || '', city: c.city || '', state: c.state || '',
          zip: c.zip || '', county: c.county || '',
          phone: c.phone || '', clientEmail: c.client_email || '', caregiverId: c.caregiver_id || '',
          dob: c.dob || '', gender: c.gender || '',
          driversLicense: c.drivers_license || '', ssn: c.ssn || '',
          startDate: c.start_date || '', liveIn: !!c.live_in,
          clientStatus: c.client_status || 'active', hasComplex: !!c.has_complex,
          clientNotes: c.client_notes || '', auditLog: [], _dbId: c.id, invoices: mappedInvs,
        };
      });
      saveProfilesLS(profiles);
      localStorage.setItem('lhca_id_map', JSON.stringify(idMap));
      var now = new Date().toLocaleString(); localStorage.setItem('lhca_last_synced', now);
      renderSidebarClients(); renderClientTable(); updateStats();
      var lsl = document.getElementById('lastSyncedLabel');
      if (lsl) lsl.textContent = 'Last synced: ' + now;
      syncEnd();
    })
    .catch(function (e) {
      console.error('Bulk load error:', e);
      renderSidebarClients(); renderClientTable(); updateStats();
      syncEnd();
    });

  // Also load caregivers, caseworkers, tasks, and signatures from API
  loadCaregiversAPI();
  loadCaseworkersAPI();
  loadTasksAPI();
  loadSignaturesAPI();
}

// ── SAVE client profile to Azure SQL ────────────────────────
function saveProfileSP(name, data) {
  var idMap = getIdMap();
  var dbId = data._dbId || idMap[name];
  var body = {
    id: dbId || undefined, client_name: name,
    first_name: data.firstName || '', last_name: data.lastName || '',
    middle_name: data.middleName || '', nickname: data.nickname || '',
    medicaid_id: data.medicaidId || '', hourly_rate: data.hourlyRate || '',
    worker: data.worker || '', caseworker_id: data.caseworkerId || '',
    street: data.street || '', city: data.city || '', state: data.state || '',
    zip: data.zip || '', county: data.county || '',
    phone: data.phone || '', client_email: data.clientEmail || '', caregiver_id: data.caregiverId || '',
    client_status: data.clientStatus || 'active', has_complex: data.hasComplex ? 1 : 0,
    // Newly persisted fields
    dob: data.dob || '', gender: data.gender || '',
    drivers_license: data.driversLicense || '', ssn: data.ssn || '',
    start_date: data.startDate || '', live_in: data.liveIn ? 1 : 0,
    client_notes: data.clientNotes || '', audit_json: JSON.stringify(data.auditLog || []),
  };
  fetch(API_BASE + '/homecare-clients', { method: 'POST', headers: apiHeaders(), body: JSON.stringify(body) })
    .then(function (r) { return r.json(); })
    .then(function (result) {
      if (!dbId && result.id) {
        idMap[name] = result.id; localStorage.setItem('lhca_id_map', JSON.stringify(idMap));
        var p = getProfiles(); if (p[name]) { p[name]._dbId = result.id; saveProfilesLS(p); }
      }
      aiTrack('ClientInfoUpdated',{clientName:name,clientStatus:body.client_status});
      var now = new Date().toLocaleString(); localStorage.setItem('lhca_last_synced', now);
      var lsl = document.getElementById('lastSyncedLabel'); if (lsl) lsl.textContent = 'Last synced: ' + now;
      // Sync any invoices not yet in DB
      syncNewInvoices(name, data);
    })
    .catch(function (e) { console.error('Save profile error:', e); });
}
function syncNewInvoices(name, data) {
  var idMap = getIdMap(); var clientDbId = idMap[name];
  if (!clientDbId || !data.invoices) return;
  data.invoices.forEach(function (inv, idx) {
    if (inv.dbId) return; // already saved
    fetch(API_BASE + '/invoices', {
      method: 'POST', headers: apiHeaders(),
      body: JSON.stringify({
        homecare_client_id: clientDbId, billing_period: inv.billingPeriod || '',
        status: inv.status || 'draft', invoice_note: inv.invoiceNote || '',
        invoice_data: inv.data ? JSON.stringify(inv.data) : null,
      }),
    }).then(function (r) { return r.json(); }).then(function (result) {
      if (result.id) {
        var p2 = getProfiles();
        if (p2[name] && p2[name].invoices && p2[name].invoices[idx]) {
          p2[name].invoices[idx].dbId = result.id; saveProfilesLS(p2);
        }
      }
    }).catch(function (e) { console.error('Sync invoice error:', e); });
  });
}

// ── DELETE client from Azure SQL ─────────────────────────────
function deleteProfileSP(name) {
  var idMap = getIdMap(); var dbId = idMap[name]; if (!dbId) return;
  fetch(API_BASE + '/homecare-clients/' + dbId, { method: 'DELETE', headers: apiHeaders() })
    .then(function () { delete idMap[name]; localStorage.setItem('lhca_id_map', JSON.stringify(idMap)); })
    .catch(function (e) { console.error('Delete profile error:', e); });
}

// ── INVOICE status update via API ───────────────────────────
function updateInvoiceStatusAPI(dbId, status) {
  if (!dbId) return;
  fetch(API_BASE + '/invoices/' + dbId + '/status', {
    method: 'PATCH', headers: apiHeaders(), body: JSON.stringify({ status: status }),
  }).catch(function (e) { console.error('Status update error:', e); });
}
function deleteInvoiceAPI(dbId, clientName, billingPeriod) {
  if (!dbId) return;
  aiTrack('InvoiceDeleted',{invoiceDbId:dbId,clientName:clientName||'',billingPeriod:billingPeriod||''});
  fetch(API_BASE + '/invoices/' + dbId, { method: 'DELETE', headers: apiHeaders() })
    .catch(function (e) { console.error('Delete invoice error:', e); });
}

// ── CAREGIVERS API ───────────────────────────────────────────
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
          street: cg.street || '', city: cg.city || '',
          state: cg.state || '', zip: cg.zip || '', county: cg.county || '',
          dob: cg.dob || '', driversLicense: cg.drivers_license || '', ssn: cg.ssn || '',
          notes: cg.notes || ''
        };
      });
      if (Object.keys(obj).length) saveCaregiversLS(obj);
      syncEnd();
    }).catch(function (e) { console.error('Load caregivers error:', e); syncEnd(); });
}
function saveCaregiverAPI(id, cg) {
  fetch(API_BASE + '/caregivers', {
    method: 'POST', headers: apiHeaders(),
    body: JSON.stringify({
      id: id, name: cg.name || '',
      first_name: cg.firstName || '', last_name: cg.lastName || '',
      middle_name: cg.middleName || '', nickname: cg.nickname || '',
      role: cg.emptype || '', emptype: cg.emptype || '',
      phone: cg.phone || '', email: cg.email || '',
      start_date: cg.hireDate || '', hire_date: cg.hireDate || '',
      status: cg.status || 'active', notes: cg.notes || '',
      // Address
      street: cg.street || cg.address || '', city: cg.city || '',
      state: cg.state || '', zip: cg.zip || '', county: cg.county || '',
      // Employment
      pay_rate: cg.payRate || '', max_hours: cg.maxHours || '',
      certifications: cg.certs || cg.certifications || '',
      ec_name: cg.ecName || '', ec_phone: cg.ecPhone || '',
      champs_id: cg.champsId || '', gender: cg.gender || '',
      // Identity (sensitive — these fields are masked in exports too)
      dob: cg.dob || '', drivers_license: cg.driversLicense || '', ssn: cg.ssn || ''
    }),
  }).catch(function (e) { console.error('Save caregiver error:', e); });
}
function deleteCaregiverAPI(id) {
  fetch(API_BASE + '/caregivers/' + id, { method: 'DELETE', headers: apiHeaders() })
    .catch(function (e) { console.error('Delete caregiver error:', e); });
}

// ── TASKS API ────────────────────────────────────────────────
function loadTasksAPI() {
  fetch(API_BASE + '/tasks?source=homecare', { headers: apiHeaders() })
    .then(function (r) { return r.json(); })
    .then(function (tasks) {
      var todos = tasks.map(function (t) {
        // CRITICAL: id MUST be a string to match the format used by todoId() (e.g. 'td_xxx_yyy').
        // If id is a number, all task button click handlers (deleteTodo / toggleTodo / etc.)
        // fail their strict-equality lookup and silently do nothing.
        return {
          id: String(t.id), dbId: t.id, text: t.task_text, done: !!t.done,
          due: t.due_date ? t.due_date.split('T')[0] : '', client: t.client_name || '',
          priority: t.priority || 'normal', note: t.note || '',
          parentId: t.parent_id ? String(t.parent_id) : null,
        };
      });
      if (todos.length) saveTodos(todos);
      // Re-render if user is currently on the tasks page so DB-loaded tasks become clickable
      if(document.getElementById('page-tasks')&&document.getElementById('page-tasks').classList.contains('active')){
        renderTodos();updateTaskBadge();
      }
    }).catch(function (e) { console.error('Load tasks error:', e); });
}
function saveTaskAPI(todo) {
  fetch(API_BASE + '/tasks', {
    method: 'POST', headers: apiHeaders(),
    body: JSON.stringify({
      id: todo.dbId || undefined, text: todo.text, done: todo.done ? 1 : 0,
      due: todo.due || null, client: todo.client || '', priority: todo.priority || 'normal', source: 'homecare',
      parent_id: todo.parentId || null, note: todo.note || null,
    }),
  }).then(function (r) { return r.json(); })
    .then(function (result) { if (!todo.dbId && result.id) { todo.dbId = result.id; } })
    .catch(function (e) { console.error('Save task error:', e); });
}
function deleteTaskAPI(dbId) {
  if (!dbId) return;
  fetch(API_BASE + '/tasks/' + dbId, { method: 'DELETE', headers: apiHeaders() })
    .catch(function (e) { console.error('Delete task error:', e); });
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
    .then(function(r){ return r.json(); })
    .then(function(rows){
      var arr = (rows||[]).map(function(c){
        return { id:c.id, name:c.name||'', first_name:c.first_name||'', last_name:c.last_name||'',
                 agency:c.agency||'', phone:c.phone||'', email:c.email||'', fax:c.fax||'',
                 street:c.street||'', city:c.city||'', state:c.state||'', zip:c.zip||'', county:c.county||'',
                 notes:c.notes||'' };
      });
      saveCaseworkersLS(arr);
      if (typeof renderCaseworkerList === 'function' && document.getElementById('cwList')) renderCaseworkerList();
      syncEnd();
    })
    .catch(function(e){ console.error('Load caseworkers error:', e); syncEnd(); });
}
function saveCaseworkerAPI(cw){
  return fetch(API_BASE + '/caseworkers', {
    method: 'POST', headers: apiHeaders(), body: JSON.stringify(cw),
  }).catch(function(e){ console.error('Save caseworker error:', e); });
}
function deleteCaseworkerAPI(id){
  return fetch(API_BASE + '/caseworkers/' + encodeURIComponent(id), {
    method: 'DELETE', headers: apiHeaders(),
  }).catch(function(e){ console.error('Delete caseworker error:', e); });
}
function cwId(){return 'cw_'+Date.now()+'_'+Math.random().toString(36).slice(2,7);}
function navCaseworkers(){
  showPage('caseworkers');bc([{l:'Caseworkers'}]);document.getElementById('topbarActions').innerHTML='';
  showCwGrid();
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
      // Split name into first/last
      var nameParts=(cw.name||'').trim().split(' ');
      var firstName=nameParts[0]||'';
      var lastName=nameParts.slice(1).join(' ')||'';
      document.getElementById('cw-first-name').value=cw.first_name||firstName;
      document.getElementById('cw-middle-name').value=cw.middle_name||'';
      document.getElementById('cw-last-name').value=cw.last_name||lastName;
      document.getElementById('cw-nickname').value=cw.nickname||'';
      document.getElementById('cw-agency').value=cw.agency||'';
      document.getElementById('cw-phone').value=cw.phone||'';
      document.getElementById('cw-fax').value=cw.fax||'';
      document.getElementById('cw-email').value=cw.email||'';
      document.getElementById('cw-street').value=cw.street||'';
      document.getElementById('cw-city').value=cw.city||'';
      document.getElementById('cw-state').value=cw.state||'';
      document.getElementById('cw-zip').value=cw.zip||'';
      document.getElementById('cw-county').value=cw.county||'';
      document.getElementById('cw-notes').value=cw.notes||'';
    }
    document.getElementById('cwDeleteBtn').style.display='inline-block';
  } else {
    ['cw-first-name','cw-middle-name','cw-last-name','cw-nickname','cw-agency','cw-phone','cw-fax','cw-email','cw-street','cw-city','cw-state','cw-zip','cw-county','cw-notes'].forEach(function(fid){var e=document.getElementById(fid);if(e)e.value='';});
    document.getElementById('cwDeleteBtn').style.display='none';
  }
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
    id:editingId||cwId(),name:name,first_name:firstName,middle_name:middleName,last_name:lastName,nickname:nickname,
    agency:document.getElementById('cw-agency').value,
    phone:document.getElementById('cw-phone').value,
    fax:document.getElementById('cw-fax').value,
    email:document.getElementById('cw-email').value,
    street:document.getElementById('cw-street').value,
    city:document.getElementById('cw-city').value,
    state:document.getElementById('cw-state').value,
    zip:document.getElementById('cw-zip').value,
    county:document.getElementById('cw-county').value,
    notes:document.getElementById('cw-notes').value
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
function renderCaseworkerList(){
  var cws=getCaseworkers();
  var q=(document.getElementById('cwSearch')?document.getElementById('cwSearch').value:'').toLowerCase();
  var profiles=getProfiles();
  var tbody=document.getElementById('cwTableBody');if(!tbody)return;tbody.innerHTML='';
  var filtered=cws.filter(function(cw){
    return !q||(cw.name||'').toLowerCase().includes(q)||(cw.agency||'').toLowerCase().includes(q);
  });
  filtered.sort(function(a,b){return (a.name||'').localeCompare(b.name||'');});
  var empty=document.getElementById('cwTableEmpty');
  if(!filtered.length){if(empty)empty.style.display='block';return;}
  if(empty)empty.style.display='none';
  filtered.forEach(function(cw){
    var clientCount=Object.keys(profiles).filter(function(k){return profiles[k].caseworkerId===cw.id||profiles[k].worker===cw.name;}).length;
    var hrefCw=buildCaseworkerUrl(cw.id);
    var tr=document.createElement('tr');
    tr.innerHTML=
      '<td><a href="'+hrefCw+'" style="text-decoration:none;color:inherit;display:block;"><div class="ct-name">'+esc(cw.name||'')+'</div><div class="ct-id">'+esc(cw.agency||'No agency')+'</div></a></td>'+
      '<td style="color:#4a6a8a;font-size:12px;">'+esc(cw.phone||'—')+'</td>'+
      '<td style="color:#4a6a8a;font-size:12px;">'+esc(cw.email||'—')+'</td>'+
      '<td style="color:#4a6a8a;font-size:12px;">'+esc(cw.county||'—')+'</td>'+
      '<td style="font-size:12px;">'+clientCount+'</td>'+
      '<td onclick="event.stopPropagation()"><button class="ct-action-btn" onclick="event.stopPropagation();showCaseworkerForm(\''+cw.id+'\')">Edit</button></td>';
    tr.addEventListener('click',function(e){
      if(e.target.closest('a')||e.target.closest('button')||e.target.closest('input'))return;
      openCwDetail(cw.id);
    });
    tbody.appendChild(tr);
  });
}

// ============================================================
//  CASEWORKER DETAIL VIEW
// ============================================================
var activeCwId=null;
function openCwDetail(id){
  var cw=getCaseworkers().find(function(c){return c.id===id;});
  if(!cw)return;
  activeCwId=id;
  document.getElementById('cwGridView').style.display='none';
  document.getElementById('cwFormWrap').style.display='none';
  document.getElementById('cwDetailView').style.display='block';
  var ini=(cw.name||'?').split(' ').map(function(w){return w[0]||'';}).slice(0,2).join('').toUpperCase();
  document.getElementById('cwDetailAvatar').textContent=ini;
  document.getElementById('cwDetailName').textContent=cw.name||'';
  document.getElementById('cwDetailMeta').innerHTML=esc(cw.agency||'')+(cw.phone?' · '+esc(cw.phone):'');
  switchCwTab('overview');
}
function showCwGrid(){
  activeCwId=null;
  var dv=document.getElementById('cwDetailView');if(dv)dv.style.display='none';
  var gv=document.getElementById('cwGridView');if(gv)gv.style.display='';
  hideCaseworkerForm();
  renderCaseworkerList();
}
function switchCwTab(tab){
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
  // Build clickable client list
  var clientListHtml='';
  if(assignedNames.length){
    clientListHtml='<div style="display:flex;flex-direction:column;gap:4px;margin-top:6px;">'+
      assignedNames.map(function(name){
        var p=profiles[name]||{};
        var st=p.clientStatus||'active';
        var stColor=st==='active'?'#1e7e34':st==='inactive'?'#888':'#a83232';
        return '<div onclick="navDetail(\''+esc(name).replace(/'/g,"\\'")+'\')" style="display:flex;align-items:center;gap:8px;padding:6px 8px;background:#f7faff;border:1px solid #e1e5ea;border-radius:5px;cursor:pointer;font-size:12px;" onmouseover="this.style.borderColor=\'#b0c8e8\'" onmouseout="this.style.borderColor=\'#e1e5ea\'">'+
          '<span style="flex:1;color:#185FA5;font-weight:500;">'+esc(name)+'</span>'+
          (p.medicaidId?'<span style="color:#8ca0b4;">'+esc(p.medicaidId)+'</span>':'')+
          '<span style="color:'+stColor+';font-size:10px;font-weight:600;text-transform:uppercase;">'+st+'</span>'+
        '</div>';
      }).join('')+
    '</div>';
  } else {
    clientListHtml='<div style="font-size:12px;color:#8ca0b4;margin-top:6px;">No clients assigned.</div>';
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

  var g=document.createElement('div');g.className='info-grid';c.appendChild(g);

  function mkF(id,label,val,full){
    var d=document.createElement('div');d.className='info-field'+(full?' full':'');
    d.innerHTML='<label>'+label+'</label><input id="'+id+'" value="'+esc(val||'')+'">';g.appendChild(d);
  }
  function mkDiv(label){var d=document.createElement('div');d.className='form-section-divider full';d.innerHTML='<span>'+label+'</span>';g.appendChild(d);}
  function mkRow(html){var d=document.createElement('div');d.className='info-field-row full';d.innerHTML=html;g.appendChild(d);}

  // Name row
  var firstName=cw.first_name||(cw.name||'').split(' ')[0]||'';
  var lastName=cw.last_name||(cw.name||'').split(' ').slice(1).join(' ')||'';
  var dName=document.createElement('div');dName.className='info-field-row full';dName.style.gridTemplateColumns='1fr 1fr 1fr';
  dName.innerHTML='<div class="info-field"><label>First Name *</label><input id="cwi-first" value="'+esc(firstName)+'"></div>'+
    '<div class="info-field"><label>Middle Name</label><input id="cwi-middle" value="'+esc(cw.middle_name||'')+'"></div>'+
    '<div class="info-field"><label>Last Name *</label><input id="cwi-last" value="'+esc(lastName)+'"></div>';
  g.appendChild(dName);

  mkF('cwi-agency','Agency',cw.agency,true);
  mkRow('<div class="info-field"><label>Phone</label><input id="cwi-phone" value="'+esc(cw.phone||'')+'"></div>'+
    '<div class="info-field"><label>Fax</label><input id="cwi-fax" value="'+esc(cw.fax||'')+'"></div>');
  mkF('cwi-email','Email',cw.email,true);

  mkDiv('Address');
  mkF('cwi-street','Street',cw.street,true);
  mkRow('<div class="info-field"><label>City</label><input id="cwi-city" value="'+esc(cw.city||'')+'"></div>'+
    '<div class="info-field"><label>State</label><input id="cwi-state" value="'+esc(cw.state||'')+'"></div>');
  mkRow('<div class="info-field"><label>ZIP</label><input id="cwi-zip" value="'+esc(cw.zip||'')+'" oninput="lookupZip(\'cwi-zip\',\'cwi-city\',\'cwi-state\',\'cwi-county\')"></div>'+
    '<div class="info-field"><label>County</label><input id="cwi-county" value="'+esc(cw.county||'')+'"></div>');

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
  cw.first_name=first;cw.middle_name=document.getElementById('cwi-middle').value;cw.last_name=last;
  cw.name=(first+(document.getElementById('cwi-middle').value?' '+document.getElementById('cwi-middle').value:'')+' '+last).trim();
  cw.agency=document.getElementById('cwi-agency').value;
  cw.phone=document.getElementById('cwi-phone').value;cw.fax=document.getElementById('cwi-fax').value;
  cw.email=document.getElementById('cwi-email').value;
  cw.street=document.getElementById('cwi-street').value;cw.city=document.getElementById('cwi-city').value;
  cw.state=document.getElementById('cwi-state').value;cw.zip=document.getElementById('cwi-zip').value;cw.county=document.getElementById('cwi-county').value;
  saveCaseworkersLS(arr);saveCaseworkerAPI(cw);
  document.getElementById('cwDetailName').textContent=cw.name;
  document.getElementById('cwDetailMeta').innerHTML=esc(cw.agency||'')+(cw.phone?' · '+cw.phone:'');
  var btn=document.getElementById('cwSaveInfoBtn');if(btn){btn.textContent='Saved ✓';setTimeout(function(){btn.textContent='Save Changes';},1800);}
  addAuditEntry(cw.name,'Caseworker profile updated');
  renderCaseworkerList();
}
function showCwEditForm(){
  if(!activeCwId)return;
  document.getElementById('cwDetailView').style.display='none';
  document.getElementById('cwGridView').style.display='none';
  showCaseworkerForm(activeCwId);
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
        '<div style="font-size:11px;color:#6b8dae;">'+(prof.medicaidId||'No Medicaid ID')+(prof.phone?' · '+prof.phone:'')+'</div>'+
      '</div>'+
      '<span class="cs-badge cs-'+st+'">'+st.charAt(0).toUpperCase()+st.slice(1)+'</span>'+
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
  c.innerHTML='<textarea id="cwNotesArea" style="width:100%;min-height:200px;padding:12px;border:1px solid #d0d8e4;border-radius:6px;font-size:13px;font-family:Arial,sans-serif;outline:none;resize:vertical;max-width:620px;">'+esc(cw.notes||'')+'</textarea>'+
    '<div style="margin-top:10px;"><button class="btn btn-primary" onclick="saveCwNotes()">Save Notes</button></div>';
}
function saveCwNotes(){
  var arr=getCaseworkers();
  var cw=arr.find(function(c){return c.id===activeCwId;});
  if(!cw)return;
  var area=document.getElementById('cwNotesArea');
  if(!area)return;
  cw.notes=area.value;
  saveCaseworkersLS(arr);
  saveCaseworkerAPI(cw);
  var btn=document.querySelector('#cwNotesContent .btn-primary');
  if(btn){btn.textContent='Saved';setTimeout(function(){btn.textContent='Save Notes';},1800);}
}
function renderCwDocsPane(){
  if(!activeCwId)return;
  var cw=getCaseworkers().find(function(c){return c.id===activeCwId;});
  var c=document.getElementById('cwDocsContent');
  if(!c||!cw)return;
  c.innerHTML=
    '<div class="doc-upload-card">'+
      '<div class="doc-upload-head">'+
        '<h4>Documents for '+esc(cw.name||'Caseworker')+'</h4>'+
        '<p>Upload letters, authorization docs, etc.</p>'+
      '</div>'+
      '<div class="doc-upload-row">'+
        '<div class="doc-upload-fields">'+
          '<label>File</label>'+
          '<input type="file" id="cwDocFileInput" accept=".pdf,.jpg,.jpeg,.png,.doc,.docx" multiple>'+
        '</div>'+
        '<div class="doc-upload-actions">'+
          '<button class="btn btn-primary" onclick="uploadCwDoc()">Upload</button>'+
          '<input type="file" id="cwDocScanInput" accept="image/*" capture="environment" style="display:none;" onchange="handleCwDocScan(this)">'+
          '<button class="btn btn-secondary" onclick="document.getElementById(\'cwDocScanInput\').click()">Scan / Photo</button>'+
        '</div>'+
      '</div>'+
      '<span id="cwDocUploadStatus" class="doc-upload-status"></span>'+
    '</div>'+
    '<div id="cwDocListAzure"><div style="color:#8ca0b4;font-size:13px;">Loading...</div></div>';
  fetch(API_BASE+'/documents?clientType=caseworker&clientId='+activeCwId,{headers:apiHeaders()})
    .then(function(r){return r.json();})
    .then(function(docs){renderCwDocListAzure(activeCwId,docs||[]);})
    .catch(function(){renderCwDocListAzure(activeCwId,[]);});
}
function renderCwDocListAzure(cwId,docs){
  var list=document.getElementById('cwDocListAzure');if(!list)return;
  if(!docs.length){list.innerHTML='<div style="color:#8ca0b4;font-size:13px;padding:4px 0;">No documents yet.</div>';return;}
  list.innerHTML='';
  docs.forEach(function(d){
    var kb=d.size?Math.round(d.size/1024)+'KB':'';
    var ext=(d.name||'').split('.').pop().toLowerCase();
    var isImg=['jpg','jpeg','png','gif'].indexOf(ext)>=0;
    var icon=(ext||"").toUpperCase().slice(0,4);
    var card=document.createElement('div');
    card.style.cssText='display:flex;align-items:center;gap:10px;padding:8px 10px;border:1px solid #e1e8f0;border-radius:6px;margin-bottom:6px;background:#fafbfc;';
    card.innerHTML=
      '<span style="display:inline-block;min-width:34px;padding:3px 6px;background:#e8eef5;color:#1a3a5c;border-radius:4px;font-size:10px;font-weight:600;text-align:center;letter-spacing:.3px;">'+(icon||'FILE')+'</span>'+
      '<div style="flex:1;min-width:0;">'+
        '<div style="font-size:12px;font-weight:600;color:#1a3a5c;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;"><a href="'+d.url+'" target="_blank" style="color:#1a3a5c;text-decoration:none;">'+esc(d.name)+'</a></div>'+
        '<div style="font-size:11px;color:#8ca0b4;">'+kb+'</div>'+
      '</div>'+
      '<button class="btn btn-danger btn-sm" style="padding:3px 10px;font-size:11px;" onclick="deleteCwDoc(\''+cwId+'\',\''+encodeURIComponent(d.name)+'\')">✕</button>';
    list.appendChild(card);
  });
}
function uploadCwDoc(){
  var input=document.getElementById('cwDocFileInput');
  if(!input||!input.files||!input.files.length){showAlert('Please select a file first.');return;}
  var status=document.getElementById('cwDocUploadStatus');status.textContent='Uploading...';
  var fd=new FormData();fd.append('clientType','caseworker');fd.append('clientId',activeCwId);
  Array.from(input.files).forEach(function(f){fd.append('file',f);});
  fetch(API_BASE+'/documents',{method:'POST',headers:authUploadHeaders(),body:fd})
    .then(function(){status.textContent='';input.value='';renderCwDocsPane();})
    .catch(function(e){status.textContent='Upload failed: '+e;});
}
function deleteCwDoc(cwId,encodedName){
  showConfirm('Delete this document?',function(){
    fetch(API_BASE+'/documents?clientType=caseworker&clientId='+cwId+'&name='+encodedName,{method:'DELETE',headers:apiHeaders()})
      .then(function(){renderCwDocsPane();}).catch(function(e){showAlert('Delete failed: '+e);});
  },{title:'Delete Document',okText:'Delete'});
}
function handleCwDocScan(input){
  if(!activeCwId){showAlert('Open a caseworker first.');return;}
  if(!input||!input.files||!input.files.length)return;
  var status=document.getElementById('cwDocUploadStatus');
  if(status)status.textContent='Uploading scanned image…';
  var fd=new FormData();fd.append('clientType','caseworker');fd.append('clientId',activeCwId);
  fd.append('file',input.files[0]);
  fetch(API_BASE+'/documents',{method:'POST',headers:authUploadHeaders(),body:fd})
    .then(function(){if(status)status.textContent='';input.value='';renderCwDocsPane();})
    .catch(function(e){if(status)status.textContent='Upload failed: '+e;});
}
function renderCwAuditPane(){
  if(!activeCwId)return;
  var pane=document.getElementById('cwpane-audit');
  var cw=getCaseworkers().find(function(c){return c.id===activeCwId;});
  if(!pane||!cw)return;
  pane.innerHTML='<div style="color:#8ca0b4;font-size:13px;padding:8px 0;">Loading…</div>';
  if(spToken){
    fetch(API_BASE+'/audit?client='+encodeURIComponent(cw.name||activeCwId)+'&limit=100',{headers:apiHeaders()})
      .then(function(r){return r.ok?r.json():Promise.reject();})
      .then(function(rows){
        pane.innerHTML='';
        if(!rows.length){pane.innerHTML='<div style="color:#8ca0b4;font-size:13px;padding:16px 0;">No audit entries yet.</div>';return;}
        var wrap=document.createElement('div');wrap.style.cssText='max-width:600px;';
        rows.forEach(function(e){
          var row=document.createElement('div');row.className='audit-row';
          var ts=e.created_at?new Date(e.created_at).toLocaleString():e.ts||'';
          row.innerHTML='<span class="audit-icon">—</span><div><div class="audit-text">'+esc(e.action)+'</div><div class="audit-who">'+esc(e.who)+' · '+esc(ts)+'</div></div>';
          wrap.appendChild(row);
        });
        pane.appendChild(wrap);
      })
      .catch(function(){pane.innerHTML='<div style="color:#8ca0b4;font-size:13px;padding:16px 0;">No audit entries yet.</div>';});
  } else {
    pane.innerHTML='<div style="color:#8ca0b4;font-size:13px;padding:16px 0;">Sign in to view audit log.</div>';
  }
}

// ============================================================
//  KEYBOARD SHORTCUTS & NAVIGATION GUARD
// ============================================================
document.addEventListener('keydown',function(e){
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
  if(unsavedChanges){e.preventDefault();e.returnValue='';}
});

// ============================================================
//  AUTO SESSION TIMEOUT (HIPAA — 15 min inactivity)
// ============================================================
var SESSION_TIMEOUT_MS  = 15 * 60 * 1000; // sign out after 15 min idle
var SESSION_WARN_MS     = 13 * 60 * 1000; // warn at 13 min
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
    updateAuthUI(false);
    showToast('Signed out automatically due to inactivity.', 5000);
  }, SESSION_TIMEOUT_MS);
}
// Restart timer on any user interaction
['mousemove','mousedown','keydown','touchstart','scroll','click'].forEach(function(ev){
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
  var titles={dhs390:'DHS-390 — Adult Services Application',dhs4771:'DHS-4771 — FICA Tax Authorization',mdhhs6200:'MDHHS-6200 — Medical Needs Certification',msa4676:'MSA-4676 — Home Help Services Agreement',bphasa2421:'BPHASA-2421 — Live-In Caregiver Attestation'};
  showPage('form-fill');
  bc([{l:'Forms',fn:navForms},{l:titles[type]||type}]);
  ['sb-home','sb-caregivers','sb-settings','sb-tasks','sb-reports','sb-caseworkers','sb-forms'].forEach(function(id){var el=document.getElementById(id);if(el)el.classList.remove('active');});
  var fb=document.getElementById('sb-forms');if(fb)fb.classList.add('active');
  document.getElementById('formFillTitle').textContent=titles[type]||'';
  // Single-pane PDF view — auto-fill from CRM on load, user can switch clients here
  var profsForPicker=getProfiles();
  var clientNames=Object.keys(profsForPicker).filter(function(n){var p=profsForPicker[n];return !p.clientStatus||p.clientStatus==='active';}).sort();
  var clientOpts='<option value=""'+(!activeFormClientName?' selected':'')+'>— Pick a client —</option>'+
    clientNames.map(function(n){return '<option value="'+esc(n)+'"'+(n===activeFormClientName?' selected':'')+'>'+esc(n)+'</option>';}).join('');
  document.getElementById('formFillContent').innerHTML=
    '<div style="max-width:1100px;margin:0 auto;">'+
      '<div style="font-size:12px;color:#5a7296;margin-bottom:8px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px;">'+
        '<div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;">'+
          '<label style="font-size:12px;font-weight:600;color:#1a2b45;white-space:nowrap;">Pre-fill for:</label>'+
          '<select id="sfClientPicker" onchange="changeStateFormClient(this.value)" style="padding:5px 8px;border:1px solid #d0d8e4;border-radius:5px;font-size:12px;background:#fff;min-width:220px;">'+clientOpts+'</select>'+
          (activeFormClientName?'<span style="color:#1a7740;font-size:11px;">✓ pre-filled</span>':'<span style="color:#a05a00;font-size:11px;">type directly, or pick a client to auto-fill</span>')+
        '</div>'+
        '<span id="sfPreviewStatusText" style="font-size:11px;color:#6b8dae;">Loading…</span>'+
      '</div>'+
      '<iframe id="sfPreviewFrame" style="width:100%;height:calc(100vh - 130px);border:1px solid #d0d8e4;border-radius:8px;background:#f0f3f7;" title="Form"></iframe>'+
    '</div>';
  scheduleSfPreview(0);
  // Topbar buttons
  var topbar=document.querySelector('.form-topbar');
  if(topbar){
    var titleSpan=document.getElementById('formFillTitle');
    topbar.innerHTML='';
    var back=document.createElement('button');back.className='btn btn-secondary btn-sm';back.textContent='← Back to Forms';back.onclick=navForms;
    topbar.appendChild(back);
    if(titleSpan)topbar.appendChild(titleSpan);
    var spacer=document.createElement('div');spacer.style.flex='1';topbar.appendChild(spacer);
    var dl=document.createElement('button');dl.className='btn btn-primary';dl.textContent='⬇ Download PDF';dl.title='Download the auto-filled PDF for printing and signature. Upload the signed scan via the Documents tab.';dl.onclick=downloadStateFormPdf;
    topbar.appendChild(dl);
  }
}

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
function hookStateFormLivePreview(){
  var formEl=document.querySelector('.sf-form');
  if(!formEl)return;
  formEl.addEventListener('input',scheduleSfPreview);
  formEl.addEventListener('change',scheduleSfPreview);
  // Initial render
  scheduleSfPreview(0);
}
function scheduleSfPreview(delay){
  clearTimeout(_sfPreviewTimer);
  var ms=(delay==null||typeof delay==='object')?350:delay;
  var s=document.querySelector('.sf-preview-status');var t=document.getElementById('sfPreviewStatusText');
  if(s){s.classList.remove('idle');}if(t)t.textContent='Updating preview…';
  _sfPreviewTimer=setTimeout(renderSfPreview,ms);
}
// Each state form's left-pane input IDs → dict keys. Lets us read live-edited
// values from the form on the left and flow them into the PDF preview/save.
var STATE_FORM_INPUT_MAPS={
  dhs390:{
    dhs390_cname:'client_name', dhs390_county:'client_county', dhs390_date:'today_date',
    dhs390_wname:'worker_name', dhs390_wphone:'worker_phone',
    dhs390_fname:'client_name', dhs390_dob:'client_dob', dhs390_mid:'medicaid_id',
    dhs390_addr:'client_address', dhs390_city:'client_city', dhs390_st:'client_state', dhs390_zip:'client_zip',
    dhs390_phone:'client_phone', dhs390_tty:'client_tty', dhs390_email:'client_email',
    dhs390_sigdate:'signature_date'
  },
  dhs4771:{
    dhs4771_off1:'agency_office_addr1', dhs4771_off2:'agency_office_addr2', dhs4771_offcity:'agency_office_csz',
    dhs4771_cn:'client_name', dhs4771_case:'case_number', dhs4771_cid:'medicaid_id', dhs4771_county:'client_county',
    dhs4771_asw:'worker_name', dhs4771_aswph:'worker_phone',
    dhs4771_pname:'client_name', dhs4771_date:'today_date',
    dhs4771_addr:'client_address', dhs4771_city:'client_city', dhs4771_st:'client_state', dhs4771_zip:'client_zip'
  },
  mdhhs6200:{
    m62_pname:'client_name', m62_dob:'client_dob',
    m62_cname:'client_name', m62_log:'log_number', m62_rid:'medicaid_id',
    m62_wname:'worker_name', m62_wemail:'worker_email', m62_wphone:'worker_phone',
    m62_county:'client_county', m62_fax:'worker_fax',
    m62_lastseen:'last_seen', m62_diag:'diagnosis', m62_equdet:'equipment_details', m62_resolved:'resolved_date',
    m62_sigpname:'provider_name', m62_sigdate:'signature_date'
  },
  msa4676:{
    msa_cname:'client_name', msa_mid:'medicaid_id',
    msa_addr:'client_address', msa_city:'client_city', msa_st:'client_state', msa_zip:'client_zip',
    msa_phone:'client_phone', msa_county:'client_county',
    msa_cgname:'caregiver_full_name', msa_cgphone:'caregiver_phone',
    msa_start:'start_date', msa_date:'today_date'
  },
  bphasa2421:{
    bp_cg_first:'caregiver_first_name', bp_cg_last:'caregiver_last_name',
    bp_cg_addr:'caregiver_address', bp_cg_city:'caregiver_city', bp_cg_st:'caregiver_state', bp_cg_zip:'caregiver_zip',
    bp_cg_email:'caregiver_email', bp_cg_phone:'caregiver_phone', bp_cg_champs:'caregiver_champs_id',
    bp_b_first:'client_first_name', bp_b_last:'client_last_name', bp_b_mid:'medicaid_id',
    bp_b_addr:'client_address', bp_b_city:'client_city', bp_b_st:'client_state', bp_b_zip:'client_zip',
    bp_sig_date:'signature_date'
  }
};
// Convert YYYY-MM-DD (HTML date-input format) → MM/DD/YYYY (state-form format)
function _normalizeDate(v){
  if(!v)return '';
  var m=String(v).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if(m)return m[2]+'/'+m[3]+'/'+m[1];
  return v;
}

// Hard-coded agency info — appears as Section 2 (provider info) on MSA-4676.
// If anything here changes (new owner / address / CHAMPS ID), update this block.
var AGENCY_INFO={
  agency_provider_name:'Thomas Jaboro',
  agency_provider_id:'6221933',
  agency_address:'2741 Balsam Way Dr',
  agency_city:'Sterling Heights',
  agency_state:'MI',
  agency_zip:'48314',
  agency_phone:'(248) 291-4106',
  agency_relationship:'N/A - Agency'
};
// Set of dict keys that hold dates and should be normalized to MM/DD/YYYY
var _DATE_DICT_KEYS={client_dob:1,caregiver_dob:1,signature_date:1,today_date:1,last_seen:1,resolved_date:1,start_date:1};

// Build the data dictionary the renderer fills AcroForm fields from.
// CRM data is the default; any non-empty value typed into the left-pane form
// overrides it (so live edits flow into the PDF preview).
function _buildFormDataDict(){
  var prof=activeFormClientName?(getProfiles()[activeFormClientName]||{}):{};
  var cw=getCaseworkers().find(function(c){return (c.name||'').toLowerCase()===(prof.worker||'').toLowerCase();})||{};
  var cgs=getCaregivers();
  var assignedCg=(prof.caregiverId&&cgs[prof.caregiverId])||{};
  var td=today();
  var fullName=activeFormClientName||'';
  var fParts=fullName.split(' ');
  var firstN=fParts[0]||'',lastN=fParts.slice(1).join(' ')||'';
  var cgFull=assignedCg.name||'';
  var cgParts=cgFull.split(' ');
  var cgFirst=assignedCg.firstName||cgParts[0]||'';
  var cgLast=assignedCg.lastName||cgParts.slice(1).join(' ')||'';
  var dict={
    // Client
    client_name:fullName, client_first_name:firstN, client_last_name:lastN,
    client_dob:_normalizeDate(prof.dob||''), medicaid_id:prof.medicaidId||'', case_number:'', recipient_id:prof.medicaidId||'',
    client_address:prof.street||prof.address||'', client_city:prof.city||'', client_state:prof.state||'MI', client_zip:prof.zip||'',
    client_phone:prof.phone||'', client_email:prof.clientEmail||prof.cemail||'', client_county:prof.county||'',
    // Caseworker
    worker_name:cw.name||prof.worker||'', worker_phone:cw.phone||'', worker_email:cw.email||'', worker_fax:'',
    // Caregiver
    caregiver_first_name:cgFirst, caregiver_last_name:cgLast, caregiver_full_name:cgFull,
    caregiver_dob:_normalizeDate(assignedCg.dob||''), caregiver_address:assignedCg.street||assignedCg.address||'',
    caregiver_city:assignedCg.city||'', caregiver_state:assignedCg.state||'MI', caregiver_zip:assignedCg.zip||'',
    caregiver_phone:assignedCg.phone||'', caregiver_email:assignedCg.email||'',
    caregiver_champs_id:assignedCg.champsId||assignedCg.champs_id||'',
    // Agency (hardcoded — used as Section 2 on MSA-4676)
    agency_provider_name:AGENCY_INFO.agency_provider_name,
    agency_provider_id:AGENCY_INFO.agency_provider_id,
    agency_address:AGENCY_INFO.agency_address,
    agency_city:AGENCY_INFO.agency_city,
    agency_state:AGENCY_INFO.agency_state,
    agency_zip:AGENCY_INFO.agency_zip,
    agency_phone:AGENCY_INFO.agency_phone,
    agency_relationship:AGENCY_INFO.agency_relationship,
    // Common
    today_date:td, signature_date:td, log_number:''
  };
  // Final pass: normalize any date-key in the dict that's still YYYY-MM-DD (defensive)
  Object.keys(_DATE_DICT_KEYS).forEach(function(k){if(dict[k])dict[k]=_normalizeDate(dict[k]);});
  return dict;
}
// Per-form override maps — exact field-name → data-key. Used when the
// fuzzy matcher can't read the field name (e.g. MSA-4676's garbled names
// caused by the PDF's custom font encoding).
var STATE_FORM_FIELD_MAPS={
  // ── DHS-4771: FICA Tax Authorization ──────────────────────
  // Adobe-detected names match the printed labels; explicit map
  // because fuzzy rules misclassify "Client ID" as nothing and
  // "ASW Telephone Number" as client phone.
  dhs4771:{
    'Client Name':'client_name',
    'Case Number':'case_number',
    'Client ID':'medicaid_id',
    'County':'client_county',
    'Adult Services Worker ASW':'worker_name',
    'ASW Telephone Number':'worker_phone',
    'Printed Name':'client_name',
    'Date':'today_date',
    'Address':'client_address',
    'City':'client_city',
    'State':'client_state',
    'Zip Code':'client_zip'
  },
  // ── BPHASA-2421 ───────────────────────────────────────────
  // Adobe used "Row1" suffix on Section 1 (Caregiver) fields and
  // "_2" on Section 2 (Beneficiary) fields. Disambiguating explicitly.
  bphasa2421:{
    // Section 1: CAREGIVER (top of form, y≈633-538)
    'First NameRow1':'caregiver_first_name',
    'Last NameRow1':'caregiver_last_name',
    'Street AddressRow1':'caregiver_address',
    'CityRow1':'caregiver_city',
    'StateRow1':'caregiver_state',
    'Zip CodeRow1':'caregiver_zip',
    'Email AddressRow1':'caregiver_email',
    'Phone NumberRow1':'caregiver_phone',
    'CHAMPS Provider ID NumberRow1':'caregiver_champs_id',
    // Section 2: BENEFICIARY / CLIENT (lower, y≈468-420)
    'First NameRow1_2':'client_first_name',
    'Last NameRow1_2':'client_last_name',
    'Medicaid ID NumberRow1':'medicaid_id',
    'Street AddressRow1_2':'client_address',
    'CityRow1_2':'client_city',
    'StateRow1_2':'client_state',
    'Zip CodeRow1_2':'client_zip'
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
        } else if(typeof field.select==='function'&&typeof field.check!=='function'){
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

async function downloadStateFormPdf(){
  try{
    showToast('Generating final (flattened) PDF…',2000);
    var bytes=await _renderStateFormBytes({flatten:true});
    var def=STATE_FORM_OVERLAYS[activeFormType]||{};
    var clientTag=(activeFormClientName||'client').replace(/[^a-z0-9]/gi,'_');
    var fname=(def.title||activeFormType)+'_'+clientTag+'_'+today().replace(/\//g,'-')+'.pdf';
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
 *    fields: [{ inputId:'dhs4771_x', page:0, x:.., y:.., size:10, maxWidth?:.. }, ...],
 *    signature?: { inputId:'dhs4771_sig_present', page, x, y, w, h }  // optional, draws stored sig PNG
 *  }
 * ────────────────────────────────────────────────────────────────── */
var STATE_FORM_OVERLAYS={
  dhs4771:{
    file:'/forms/DHS-4771.pdf',
    title:'DHS-4771_FICA_Authorization',
    // Coordinates derived from the actual PDF text positions
    fields:[
      // County office block — overlay directly on top of "COUNTY STREET ADDRESS 1" etc placeholders
      {inputId:'dhs4771_off1',    page:0, x:33,  y:625, size:10, maxWidth:300, coverRect:{x:30,y:621,w:280,h:13}},
      {inputId:'dhs4771_off2',    page:0, x:33,  y:610, size:10, maxWidth:300, coverRect:{x:30,y:606,w:280,h:13}},
      {inputId:'dhs4771_offcity', page:0, x:33,  y:595, size:10, maxWidth:300, coverRect:{x:30,y:591,w:280,h:13}},
      // Client info row — labels @ y=503, value goes ~18pt below
      {inputId:'dhs4771_cn',      page:0, x:30,  y:485, size:10, maxWidth:155},
      {inputId:'dhs4771_case',    page:0, x:193, y:485, size:10, maxWidth:115},
      {inputId:'dhs4771_cid',     page:0, x:314, y:485, size:10, maxWidth:115},
      {inputId:'dhs4771_county',  page:0, x:433, y:485, size:10, maxWidth:140},
      // Worker row — labels @ y=470
      {inputId:'dhs4771_asw',     page:0, x:30,  y:452, size:10, maxWidth:275},
      {inputId:'dhs4771_aswph',   page:0, x:314, y:452, size:10, maxWidth:260},
      // Signature row — labels @ y=215, value goes 18pt below at y=197
      {inputId:'dhs4771_pname',   page:0, x:249, y:197, size:10, maxWidth:200},
      {inputId:'dhs4771_date',    page:0, x:469, y:197, size:10, maxWidth:100},
      // Address row — labels @ y=174
      {inputId:'dhs4771_addr',    page:0, x:30,  y:156, size:10, maxWidth:275},
      {inputId:'dhs4771_city',    page:0, x:314, y:156, size:10, maxWidth:145},
      {inputId:'dhs4771_st',      page:0, x:469, y:156, size:10, maxWidth:40},
      {inputId:'dhs4771_zip',     page:0, x:517, y:156, size:10, maxWidth:65}
    ],
    signature:null
  },
  // ── DHS-390: Adult Services Application ──────────────────
  dhs390:{
    file:'/forms/DHS-390.pdf',
    title:'DHS-390_Adult_Services_Application',
    fields:[
      // Section 1 — labels @ y=305
      {inputId:'dhs390_cname',    page:0, x:30,  y:287, size:10, maxWidth:175},
      {inputId:'dhs390_log',      page:0, x:219, y:287, size:10, maxWidth:140},
      {inputId:'dhs390_rid',      page:0, x:379, y:287, size:10, maxWidth:200},
      // Row at y=272
      {inputId:'dhs390_county',   page:0, x:30,  y:254, size:10, maxWidth:270},
      {inputId:'dhs390_date',     page:0, x:314, y:254, size:10, maxWidth:265},
      // Row at y=240
      {inputId:'dhs390_wname',    page:0, x:30,  y:222, size:10, maxWidth:270},
      {inputId:'dhs390_wphone',   page:0, x:314, y:222, size:10, maxWidth:265},
      // Section 2 Client Info — label @ y=184
      {inputId:'dhs390_fname',    page:0, x:30,  y:166, size:10, maxWidth:550},
      // Row at y=151
      {inputId:'dhs390_dob',      page:0, x:30,  y:133, size:10, maxWidth:270},
      {inputId:'dhs390_mid',      page:0, x:314, y:133, size:10, maxWidth:265},
      // Row at y=119 (address)
      {inputId:'dhs390_addr',     page:0, x:30,  y:101, size:10, maxWidth:270},
      {inputId:'dhs390_city',     page:0, x:314, y:101, size:10, maxWidth:145},
      {inputId:'dhs390_st',       page:0, x:469, y:101, size:10, maxWidth:40},
      {inputId:'dhs390_zip',      page:0, x:517, y:101, size:10, maxWidth:65},
      // Row at y=87
      {inputId:'dhs390_phone',    page:0, x:30,  y:69,  size:10, maxWidth:120},
      {inputId:'dhs390_tty',      page:0, x:156, y:69,  size:10, maxWidth:220},
      {inputId:'dhs390_email',    page:0, x:385, y:69,  size:10, maxWidth:195},
      // Page 2 signature date — TODO: capture page-1 positions when needed
      {inputId:'dhs390_sigdate',  page:1, x:430, y:160, size:10, maxWidth:120}
    ],
    signature:null
  },
  // ── MDHHS-6200: Adult Services Medical Needs Certification ─
  mdhhs6200:{
    file:'/forms/MDHHS-6200.pdf',
    title:'MDHHS-6200_Medical_Needs_Cert',
    fields:[
      // Section 1 — Patient name/DOB labels @ y=333
      {inputId:'m62_pname',       page:0, x:30,  y:315, size:10, maxWidth:365},
      {inputId:'m62_dob',         page:0, x:409, y:315, size:10, maxWidth:160},
      // Section 1 signature row — Printed Name + Sig Date labels @ y=262
      {inputId:'m62_sigpname',    page:0, x:267, y:244, size:10, maxWidth:200},
      {inputId:'m62_sigdate',     page:0, x:482, y:244, size:10, maxWidth:90},
      // Section 2 — labels @ y=198
      {inputId:'m62_cname',       page:0, x:30,  y:180, size:10, maxWidth:270},
      {inputId:'m62_log',         page:0, x:314, y:180, size:10, maxWidth:125},
      {inputId:'m62_rid',         page:0, x:453, y:180, size:10, maxWidth:120},
      // Row at y=165
      {inputId:'m62_wname',       page:0, x:30,  y:147, size:10, maxWidth:175},
      {inputId:'m62_wemail',      page:0, x:219, y:147, size:10, maxWidth:220},
      {inputId:'m62_wphone',      page:0, x:453, y:147, size:10, maxWidth:120},
      // Row at y=133
      {inputId:'m62_county',      page:0, x:30,  y:115, size:10, maxWidth:270},
      {inputId:'m62_fax',         page:0, x:314, y:115, size:10, maxWidth:265},
      // Section 3 — Date Patient Last Seen "A" @ y=78 (this is a small bottom line)
      {inputId:'m62_lastseen',    page:0, x:200, y:78,  size:10, maxWidth:380},
      // Page 1 — Diagnosis B @ y=758 (right after the "B" label, value goes below at y=735)
      {inputId:'m62_diag',        page:1, x:47,  y:740, size:10, maxWidth:520},
      // Page 1 — Adaptive equipment details (F section, y=542 label)
      {inputId:'m62_equdet',      page:1, x:130, y:542, size:10, maxWidth:430},
      // Page 1 — Resolved (D, y=690)
      {inputId:'m62_resolved',    page:1, x:47,  y:670, size:10, maxWidth:520}
    ],
    signature:null
  },
  // ── MSA-4676: Home Help Services Agreement ────────────────
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
  },
  // ── BPHASA-2421: Live-In Caregiver Attestation ────────────
  bphasa2421:{
    file:'/forms/BPHASA-2421.pdf',
    title:'BPHASA-2421_Live-In_Caregiver',
    fields:[
      // Page 1 (the form itself; instructions are page 0)
      // Purpose checkboxes — labels @ y=694; checkbox is left of each label
      {inputId:'bp_purpose_initial', page:1, x:243, y:692, size:11, checkedText:'X', maxWidth:10},
      {inputId:'bp_purpose_addr',    page:1, x:349, y:692, size:11, checkedText:'X', maxWidth:10},
      {inputId:'bp_purpose_renew',   page:1, x:464, y:692, size:11, checkedText:'X', maxWidth:10},
      // Caregiver name row — labels @ y=668
      {inputId:'bp_cg_first',        page:1, x:30,  y:650, size:10, maxWidth:255},
      {inputId:'bp_cg_last',         page:1, x:295, y:650, size:10, maxWidth:285},
      // Caregiver address row — labels @ y=620
      {inputId:'bp_cg_addr',         page:1, x:30,  y:602, size:10, maxWidth:275},
      {inputId:'bp_cg_city',         page:1, x:313, y:602, size:10, maxWidth:105},
      {inputId:'bp_cg_st',           page:1, x:425, y:602, size:10, maxWidth:65},
      {inputId:'bp_cg_zip',          page:1, x:497, y:602, size:10, maxWidth:80},
      // Caregiver contact row — labels @ y=574
      {inputId:'bp_cg_email',        page:1, x:30,  y:556, size:10, maxWidth:180},
      {inputId:'bp_cg_phone',        page:1, x:217, y:556, size:10, maxWidth:180},
      {inputId:'bp_cg_champs',       page:1, x:406, y:556, size:10, maxWidth:170},
      // Beneficiary name row — labels @ y=503
      {inputId:'bp_b_first',         page:1, x:30,  y:485, size:10, maxWidth:170},
      {inputId:'bp_b_last',          page:1, x:206, y:485, size:10, maxWidth:190},
      {inputId:'bp_b_mid',           page:1, x:402, y:485, size:10, maxWidth:175},
      // Beneficiary address row — labels @ y=454
      {inputId:'bp_b_addr',          page:1, x:30,  y:436, size:10, maxWidth:185},
      {inputId:'bp_b_city',          page:1, x:217, y:436, size:10, maxWidth:185},
      {inputId:'bp_b_st',            page:1, x:406, y:436, size:10, maxWidth:90},
      {inputId:'bp_b_zip',           page:1, x:497, y:436, size:10, maxWidth:80},
      // Program checkboxes — labels @ y=403; checkbox sits left of each
      {inputId:'bp_prog_bh',         page:1, x:115, y:401, size:11, checkedText:'X', maxWidth:10},
      {inputId:'bp_prog_hh',         page:1, x:240, y:401, size:11, checkedText:'X', maxWidth:10},
      {inputId:'bp_prog_mc',         page:1, x:325, y:401, size:11, checkedText:'X', maxWidth:10},
      {inputId:'bp_prog_mhl',        page:1, x:410, y:401, size:11, checkedText:'X', maxWidth:10},
      // Caregiver signature/date row — labels @ y=255
      {inputId:'bp_sig_date',        page:1, x:219, y:237, size:10, maxWidth:170}
    ],
    signature:null
  }
};

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
    // Also offer download
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

/* Helpers */
function sfi(id,val,ph,w){
  return '<input id="'+id+'" value="'+(val||'')+'" placeholder="'+(ph||'')+'" style="'+(w?'width:'+w+';':'width:100%;')+'border:none;border-bottom:1px solid #aaa;background:transparent;font-family:Times New Roman,Times,serif;font-size:9.5pt;outline:none;padding:1px 2px;">';
}
function sfcb(id){
  return '<input type="checkbox" id="'+id+'" style="width:13px;height:13px;margin-right:4px;cursor:pointer;vertical-align:middle;">';
}

// Clean labeled field for the new split-view forms
function sff(label,id,val,opts){
  opts=opts||{};
  var t=opts.type||'text';
  var ph=opts.placeholder||'';
  var input;
  if(t==='date')input='<input type="date" id="'+id+'" value="'+esc(val||'')+'">';
  else if(t==='checkbox')input='<input type="checkbox" id="'+id+'"'+(val?' checked':'')+'>';
  else if(t==='select'){
    var opts2=opts.options||[];
    input='<select id="'+id+'">'+opts2.map(function(o){var s=(typeof o==='string')?{v:o,l:o}:o;return '<option value="'+esc(s.v)+'"'+(s.v===(val||'')?' selected':'')+'>'+esc(s.l)+'</option>';}).join('')+'</select>';
  }
  else if(t==='textarea')input='<textarea id="'+id+'" rows="'+(opts.rows||3)+'" placeholder="'+esc(ph)+'">'+esc(val||'')+'</textarea>';
  else input='<input type="text" id="'+id+'" value="'+esc(val||'')+'" placeholder="'+esc(ph)+'"'+(opts.maxlength?' maxlength="'+opts.maxlength+'"':'')+'>';
  return '<div class="sf-field">'+
    (t==='checkbox'
      ? '<label style="flex-direction:row;align-items:center;gap:6px;text-transform:none;letter-spacing:0;font-size:13px;font-weight:400;color:#1a2b45;">'+input+esc(label)+'</label>'
      : '<label>'+esc(label)+'</label>'+input
    )+
    '</div>';
}
// Section heading + row helpers
function sfh(t){return '<div class="sf-section-h">'+esc(t)+'</div>';}
function sfr(){var args=Array.prototype.slice.call(arguments);var n=args.length;return '<div class="sf-row sf-row-'+n+'">'+args.join('')+'</div>';}

/* DHS-390 */
function buildDHS390(prof,cw){
  var name=activeFormClientName||'';
  var td=today();
  return ''+
    sfh('Section 1 — Departmental Use')+
    sfr(sff('Case Name','dhs390_cname',name),sff('Log Number','dhs390_log',''))+
    sfr(sff('Recipient ID','dhs390_rid',prof.medicaidId||''),sff('County','dhs390_county',prof.county||''))+
    sfr(sff('Worker Name','dhs390_wname',cw.name||prof.worker||''),sff('Worker Phone','dhs390_wphone',cw.phone||''))+
    sfr(sff('Date','dhs390_date',td,{type:'date'}))+
    sfh('Section 2 — Client Information')+
    sfr(sff('Full Name of Applicant','dhs390_fname',name))+
    sfr(sff('Date of Birth','dhs390_dob','',{type:'date'}),sff('Medicaid / Recipient ID','dhs390_mid',prof.medicaidId||''))+
    sfr(sff('Street Address','dhs390_addr',prof.street||prof.address||''))+
    sfr(sff('City','dhs390_city',prof.city||''),sff('State','dhs390_st',prof.state||'MI'),sff('Zip','dhs390_zip',prof.zip||''))+
    sfr(sff('Phone','dhs390_phone',prof.phone||''),sff('TTY','dhs390_tty',''))+
    sfr(sff('Email','dhs390_email',prof.clientEmail||prof.cemail||''))+
    sfh('Section 3 — Programs Requested')+
    '<div class="sf-checks">'+
      sff('Home Help','dhs390_hh',false,{type:'checkbox'})+
      sff('Adult Community Placement','dhs390_acp',false,{type:'checkbox'})+
      sff('Other Services','dhs390_os',false,{type:'checkbox'})+
    '</div>'+
    sfh('Section 4 — Living Arrangement')+
    '<div class="sf-checks">'+
      sff('Alone','dhs390_alone',false,{type:'checkbox'})+
      sff('With spouse','dhs390_spouse',false,{type:'checkbox'})+
      sff('With children under 18','dhs390_ch',false,{type:'checkbox'})+
      sff('With others','dhs390_others',false,{type:'checkbox'})+
      sff('Adult foster care','dhs390_foster',false,{type:'checkbox'})+
    '</div>'+
    sfr(sff('Spouse Name (if applicable)','dhs390_spname',''),sff('# Children','dhs390_chn',''))+
    sfr(sff('Guardian Name (if any)','dhs390_guaname',''))+
    sfh('Signature')+
    sfr(sff('Signature Date','dhs390_sigdate',td,{type:'date'}));
}

/* DHS-4771 */
function buildDHS4771(prof,cw){
  var name=activeFormClientName||'';
  var td=today();
  var addr=[prof.street||prof.address||''].join('');
  return ''+
    sfh('MDHHS County Office')+
    sfr(sff('Address line 1','dhs4771_off1','',{placeholder:'e.g. Wayne County DHHS'}))+
    sfr(sff('Address line 2','dhs4771_off2',''))+
    sfr(sff('City, State ZIP','dhs4771_offcity',''))+
    sfh('Client Information')+
    sfr(sff('Client Name','dhs4771_cn',name),sff('County','dhs4771_county',prof.county||''))+
    sfr(sff('Case Number','dhs4771_case',''),sff('Client ID (Medicaid)','dhs4771_cid',prof.medicaidId||''))+
    sfh('Adult Services Worker')+
    sfr(sff('ASW Name','dhs4771_asw',cw.name||prof.worker||''),sff('ASW Phone','dhs4771_aswph',cw.phone||''))+
    sfh('Authorization & Signature')+
    sfr(sff('Printed Name','dhs4771_pname',name),sff('Signature Date','dhs4771_date',td))+
    sfr(sff('Address','dhs4771_addr',addr))+
    sfr(sff('City','dhs4771_city',prof.city||''),sff('State','dhs4771_st',prof.state||'MI'),sff('Zip','dhs4771_zip',prof.zip||''));
}

/* MDHHS-6200 */
function buildMDHHS6200(prof,cw){
  var name=activeFormClientName||'';
  var td=today();
  return ''+
    sfh('Section 1 — Patient Information')+
    sfr(sff("Patient's Name",'m62_pname',name),sff("Date of Birth",'m62_dob','',{type:'date'}))+
    sfh('Section 2 — Departmental Use')+
    sfr(sff('Case Name','m62_cname',name),sff('Log Number','m62_log',''))+
    sfr(sff('Recipient ID','m62_rid',prof.medicaidId||''),sff('County','m62_county',prof.county||''))+
    sfr(sff('Worker Name','m62_wname',cw.name||prof.worker||''),sff('Worker Email','m62_wemail',cw.email||''))+
    sfr(sff('Worker Phone','m62_wphone',cw.phone||''),sff('Return Fax','m62_fax',''))+
    sfh('Section 3 — Medical Assessment')+
    sfr(sff('Date Last Seen','m62_lastseen','',{type:'date'}))+
    sfr(sff('Diagnosis / Conditions','m62_diag','',{type:'textarea',rows:3}))+
    '<div class="sf-checks">'+
      sff('Chronic / ongoing','m62_chry',false,{type:'checkbox'})+
      sff('Nonambulatory','m62_noamby',false,{type:'checkbox'})+
      sff('Adaptive equipment','m62_equy',false,{type:'checkbox'})+
    '</div>'+
    sfr(sff('Adaptive Equipment Details','m62_equdet',''))+
    sfr(sff('Resolved Date (if not chronic)','m62_resolved','',{type:'date'}))+
    sfh('Provider Signature')+
    sfr(sff('Provider Printed Name','m62_sigpname',''),sff('Signature Date','m62_sigdate',td,{type:'date'}));
}

/* MSA-4676 */
function buildMSA4676(prof,cw){
  var name=activeFormClientName||'';
  var td=today();
  return ''+
    sfh('Beneficiary Information')+
    sfr(sff('Client Name','msa_cname',name),sff('Medicaid ID','msa_mid',prof.medicaidId||''))+
    sfr(sff('Address','msa_addr',prof.street||prof.address||''))+
    sfr(sff('City','msa_city',prof.city||''),sff('State','msa_st',prof.state||'MI'),sff('Zip','msa_zip',prof.zip||''))+
    sfr(sff('Phone','msa_phone',prof.phone||''),sff('County','msa_county',prof.county||''))+
    sfh('Caregiver / Provider Information')+
    sfr(sff('Caregiver Name','msa_cgname',''),sff('Caregiver Phone','msa_cgphone',''))+
    sfr(sff('Service Start Date','msa_start','',{type:'date'}))+
    sfh('Signature')+
    sfr(sff('Signature Date','msa_date',td,{type:'date'}));
}

/* BPHASA-2421 — Live-In Caregiver Attestation */
function buildBPHASA2421(prof,cw){
  var name=activeFormClientName||'';
  var td=today();
  var cgs=getCaregivers();
  var assignedCg=(prof.caregiverId&&cgs[prof.caregiverId])||{};
  var cgFirst=assignedCg.firstName||'';
  var cgLast=assignedCg.lastName||'';
  if(!cgFirst&&assignedCg.name){var parts=assignedCg.name.split(' ');cgFirst=parts[0]||'';cgLast=parts.slice(1).join(' ')||'';}
  var bFirst='',bLast='';
  if(name){var bp=name.split(' ');bFirst=bp[0]||'';bLast=bp.slice(1).join(' ')||'';}
  return ''+
    sfh('Purpose of Attestation')+
    '<div class="sf-checks">'+
      sff('Initial Request','bp_purpose_initial',false,{type:'checkbox'})+
      sff('Address Change','bp_purpose_addr',false,{type:'checkbox'})+
      sff('Renewal','bp_purpose_renew',false,{type:'checkbox'})+
    '</div>'+
    sfh('Section 1 — Caregiver Information')+
    sfr(sff('First Name','bp_cg_first',cgFirst),sff('Last Name','bp_cg_last',cgLast))+
    sfr(sff('Street Address','bp_cg_addr',assignedCg.street||assignedCg.address||''))+
    sfr(sff('City','bp_cg_city',assignedCg.city||''),sff('State','bp_cg_st',assignedCg.state||'MI'),sff('Zip','bp_cg_zip',assignedCg.zip||''))+
    sfr(sff('Email','bp_cg_email',assignedCg.email||''),sff('Phone','bp_cg_phone',assignedCg.phone||''))+
    sfr(sff('CHAMPS Provider ID','bp_cg_champs',''))+
    sfh('Section 2 — Beneficiary Information')+
    sfr(sff('First Name','bp_b_first',bFirst),sff('Last Name','bp_b_last',bLast),sff('Medicaid ID','bp_b_mid',prof.medicaidId||''))+
    sfr(sff('Street Address','bp_b_addr',prof.street||prof.address||''))+
    sfr(sff('City','bp_b_city',prof.city||''),sff('State','bp_b_st',prof.state||'MI'),sff('Zip','bp_b_zip',prof.zip||''))+
    '<div class="sf-checks">'+
      sff('Behavioral Health','bp_prog_bh',false,{type:'checkbox'})+
      sff('Home Help','bp_prog_hh',true,{type:'checkbox'})+
      sff('MI Choice','bp_prog_mc',false,{type:'checkbox'})+
      sff('MI Health Link','bp_prog_mhl',false,{type:'checkbox'})+
    '</div>'+
    sfh('Section 3 — Caregiver Signature')+
    sfr(sff('Signature Date','bp_sig_date',td,{type:'date'}));
}

// ============================================================
//  SESSION 3 — MONTHLY INVOICE EMAILS
// ============================================================
function openMonthlyInvModal(){
  var modal=document.getElementById('monthlyInvModal');if(!modal)return;
  modal.classList.add('open');
  // Default to previous month
  var d=new Date();d.setMonth(d.getMonth()-1);
  var m=String(d.getMonth()+1).padStart(2,'0'),y=d.getFullYear();
  var inp=document.getElementById('monthlyInvPeriod');
  if(inp&&!inp.value)inp.value=m+'/'+y;
  document.getElementById('monthlyInvResults').innerHTML='<div style="color:#8ca0b4;font-size:13px;text-align:center;padding:24px 0;">Enter a billing period and click <b>Preview</b> to see caseworker email groups.</div>';
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
  // Signature: at least one must exist locally so the PDF auto-places it
  if(!getSigs().length)issues.push('Missing signature — PDF will export with no signature on it');
  return issues;
}

function previewMonthlyInvoices(){
  var period=(document.getElementById('monthlyInvPeriod').value||'').trim();
  if(!period||period.length<7){showAlert('Enter a billing period in MM/YYYY format.');return;}
  var profiles=getProfiles();
  var cws=getCaseworkers();
  // Group ALL active clients by worker name
  var groups={};
  Object.keys(profiles).forEach(function(name){
    var prof=profiles[name];
    if(prof.status==='inactive'||prof.status==='terminated'||prof.status==='lost')return;
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
    document.getElementById('monthlyInvResults').innerHTML='<div style="color:#8ca0b4;font-size:13px;text-align:center;padding:20px;">No active clients found.</div>';
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
  var html='<div style="font-size:12px;color:#6b8dae;margin-bottom:14px;display:flex;align-items:center;gap:14px;flex-wrap:wrap;">'+
    '<span><b>'+groupKeys.length+'</b> caseworker group'+(groupKeys.length!==1?'s':'')+' for billing period <b>'+esc(period)+'</b></span>'+
    '<span style="color:#1e7e34;">✓ '+totalReady+' ready</span>'+
    (totalIssues?'<span style="color:#c07000;">⚠ '+totalIssues+' issues</span>':'')+
    (totalSent?'<span style="color:#1565a0;">✓ '+totalSent+' already sent</span>':'')+
    (totalEmpty?'<span style="color:#888;">— '+totalEmpty+' missing</span>':'')+
    '<div style="margin-left:auto;display:flex;gap:8px;">'+
      (eligibleAutoGen>0?'<button class="btn btn-secondary btn-sm" onclick="autoGenerateMonthlyInvoices(\''+esc(period)+'\')" style="white-space:nowrap;" title="Copy each missing client\'s last invoice into '+esc(period)+' with day-shifted patterns">🔄 Auto-Generate '+eligibleAutoGen+'</button>':'')+
      (totalCaseworkersSendable>1?'<button class="btn btn-primary btn-sm" id="sendAllCwBtn" onclick="sendAllCaseworkerEmails(\''+esc(period)+'\')" style="white-space:nowrap;">Send All ('+totalCaseworkersSendable+' caseworkers)</button>':'')+
    '</div>'+
  '</div>';
  groupKeys.forEach(function(gk){
    var g=groups[gk];
    var wname=g.cwName||gk;
    var email=g.email||'';
    var hasInv=g.clients.filter(function(c){return c.inv;}).length;
    var missingInv=g.clients.length-hasInv;
    html+='<div class="cw-email-card">';
    // Header row
    html+='<div class="cw-email-hdr">';
    html+='<div style="width:34px;height:34px;border-radius:50%;background:#e8f0f9;color:#185FA5;font-size:12px;font-weight:700;display:flex;align-items:center;justify-content:center;flex-shrink:0;">'+
      esc((wname.split(' ').map(function(w){return w[0]||'';}).slice(0,2).join('').toUpperCase()||'?'))+
    '</div>';
    html+='<div><div class="cw-email-name">'+esc(wname)+'</div>';
    if(email)html+='<div style="font-size:11px;color:#6b8dae;">'+esc(email)+'</div>';
    html+='</div>';
    html+='<div style="flex:1;"></div>';
    if(missingInv>0)html+='<span style="font-size:11px;color:#c07000;font-weight:600;margin-right:8px;">'+missingInv+' missing invoice'+(missingInv>1?'s':'')+'</span>';
    if(email){
      var clientsJson=JSON.stringify(g.clients.map(function(c){return {name:c.name,medicaidId:c.prof.medicaidId||'',invStatus:c.inv?(c.inv.status||'draft'):'none',hasIssues:(c.issues||[]).length>0,issues:c.issues||[]};}));
      html+='<button class="btn btn-primary btn-sm" style="white-space:nowrap;" onclick=\'sendMonthlyEmail('+JSON.stringify(email)+','+JSON.stringify(wname)+','+clientsJson+','+JSON.stringify(period)+')\'>Send Email</button>';
    } else {
      html+='<span style="font-size:11px;color:#b03030;font-style:italic;margin-right:8px;">No email on file</span>';
      html+='<button class="btn btn-secondary btn-sm" onclick="closeMonthlyInvModal();navCaseworkers()">Add Email</button>';
    }
    html+='</div>';
    // Client list
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
        (c.prof.medicaidId?'<span style="color:#8ca0b4;font-size:11px;margin-right:6px;">ID: '+esc(c.prof.medicaidId)+'</span>':'');
      // Preview button (only if invoice exists)
      if(c.inv){
        html+='<button class="btn btn-secondary btn-sm" style="font-size:10px;padding:2px 8px;margin-right:6px;" onclick=\'previewClientInvoice('+JSON.stringify(c.name)+','+JSON.stringify(period)+')\'>Preview</button>';
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
  // Insert / after MM
  var formatted=raw.slice(0,2)+'/'+raw.slice(2,6);
  el.value=formatted;
}
function onMonthlyPeriodBlur(el){
  var v=el.value.trim();
  // Already in correct MM/YYYY form
  if(/^\d{2}\/\d{4}$/.test(v))return;
  // Strip non-digits and try to interpret
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
function openGenerateInvoicesModal(){
  var existing=document.getElementById('genInvModal');if(existing)existing.remove();
  var ov=document.createElement('div');
  ov.id='genInvModal';ov.className='modal-overlay open';
  ov.innerHTML='<div class="modal-box" style="max-width:480px;">'+
    '<h3>🔄 Generate Invoices</h3>'+
    '<div style="font-size:13px;color:#4a5d7a;margin:8px 0 14px;">Copies each active client\'s most recent invoice into the chosen billing period — weekly tasks shifted by one day, hospital column cleared, marked as draft.</div>'+
    '<label style="display:block;font-size:11px;color:#6b8dae;font-weight:600;text-transform:uppercase;letter-spacing:.4px;margin-bottom:4px;">Billing Period</label>'+
    '<input id="genInvPeriod" type="text" placeholder="MM/YYYY" maxlength="7" '+
      'style="padding:8px 10px;border:1px solid #d0d8e4;border-radius:5px;font-size:14px;width:140px;outline:none;" '+
      'oninput="onMonthlyPeriodInput(this)" onblur="onMonthlyPeriodBlur(this);refreshGenInvCount();" '+
      'onkeydown="if(event.key===\'Enter\'){onMonthlyPeriodBlur(this);refreshGenInvCount();doGenerateInvoices();}">'+
    '<div style="font-size:11px;color:#8ca0b4;margin-top:6px;">e.g. <b>04/2026</b>, <b>0426</b>, or <b>042026</b></div>'+
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
    // Drop expired
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
    var inv=prof.invoices.find(function(i){return i.id===rec.invoiceId;});
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
      var profiles=getProfiles();
      var removed=0;
      batch.invoices.forEach(function(rec){
        var prof=profiles[rec.clientName];if(!prof||!prof.invoices)return;
        var i=prof.invoices.findIndex(function(inv){return inv.id===rec.invoiceId;});
        if(i<0)return;
        var inv=prof.invoices[i];
        if(inv.status&&inv.status!=='draft')return;
        if(rec.dataHash&&rec.dataHash!==_quickInvoiceHash(inv.data))return;
        prof.invoices.splice(i,1);removed++;
      });
      saveProfilesLS(profiles);
      batch.invoices.forEach(function(rec){if(profiles[rec.clientName])saveProfileSP(rec.clientName,profiles[rec.clientName]);});
      // Remove this batch from the stack
      var arr2=_getAutoGenUndoStack();
      arr2=arr2.filter(function(b){return b.id!==batchId;});
      _setAutoGenUndoStack(arr2);
      logActivity('invoice','Undid auto-generation: removed '+removed+' invoice'+(removed!==1?'s':'')+' for '+batch.period);
      showAlert('✓ Removed '+removed+' auto-generated invoice'+(removed!==1?'s':'')+' for '+batch.period+'.',{title:'Undo Complete'});
      if(typeof previewMonthlyInvoices==='function')previewMonthlyInvoices();
      updateStats();renderUndoBanner();
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
      (canUndo?'<button class="btn btn-secondary btn-sm" title="'+esc(tooltip)+'" onclick="undoAutoGenBatch(\''+esc(b.id)+'\')">Undo</button>':'')+
      '<button class="btn btn-secondary btn-sm" onclick="dismissAutoGenUndoBatch(\''+esc(b.id)+'\')" title="Dismiss this undo entry">&times;</button>'+
    '</div>';
  }).join('');
}

function generateNextMonthInvoiceData(prevInv,newPeriod){
  if(!prevInv||!prevInv.data)return null;
  var prevPeriod=(prevInv.data&&prevInv.data.billingPeriod)||prevInv.billingPeriod;
  if(!prevPeriod)return null;
  var prevParts=prevPeriod.split('/'),newParts=newPeriod.split('/');
  if(prevParts.length!==2||newParts.length!==2)return null;
  var prevDays=daysIn(prevParts[0],prevParts[1]);
  var newDays=daysIn(newParts[0],newParts[1]);

  var prevSvc=(prevInv.data.tasks&&prevInv.data.tasks.svc)||[];
  var prevCplx=(prevInv.data.tasks&&prevInv.data.tasks.cplx)||[];

  // Service columns: 15 total (Bathing through Hospital). Hospital is the LAST column.
  var SVC_COLS=15, HOSP_IDX=SVC_COLS-1;

  // Decide per-column action based on check count in previous month:
  //  - Hospital column → always clear (by-exception entry)
  //  - 0 checks → keep at 0
  //  - 1-25 checks → SHIFT by +1 day (covers 1x/wk through ~6x/wk patterns)
  //  - 26+ checks → daily — keep every day checked (gov gives X days/wk authorization)
  var colAction=[];
  for(var c=0;c<SVC_COLS;c++){
    if(c===HOSP_IDX){colAction.push('clear');continue;}
    var count=0;
    for(var d=0;d<prevDays;d++){if(prevSvc[d]&&prevSvc[d][c])count++;}
    if(count===0)colAction.push('clear');
    else if(count>=26)colAction.push('daily');
    else colAction.push('shift');
  }

  // Build new svc[] for newDays — shift by +1 day means take prev[d-1] for shifted columns
  var newSvc=[];
  for(var d=0;d<newDays;d++){
    var row=[];
    for(var c=0;c<SVC_COLS;c++){
      var act=colAction[c];
      if(act==='clear'){row.push(false);}
      else if(act==='daily'){row.push(true);}
      else if(act==='shift'){
        var srcDay=d-1;
        if(srcDay<0)row.push(!!(prevSvc[prevDays-1]&&prevSvc[prevDays-1][c])); // wrap last day to first
        else if(srcDay>=prevDays)row.push(false);
        else row.push(!!(prevSvc[srcDay]&&prevSvc[srcDay][c]));
      }
      else { // 'keep' — copy same day
        if(d<prevDays)row.push(!!(prevSvc[d]&&prevSvc[d][c]));
        else row.push(false);
      }
    }
    newSvc.push(row);
  }

  // Complex tasks: copy as-is, truncated/extended to newDays
  var newCplx=[];
  var cplxCols=(prevCplx[0]&&prevCplx[0].length)||9;
  for(var d=0;d<newDays;d++){
    if(d<prevDays&&prevCplx[d])newCplx.push(prevCplx[d].slice());
    else { var blank=[];for(var k=0;k<cplxCols;k++)blank.push(false);newCplx.push(blank); }
  }

  // Build new invoice data — preserve hours/totals/etc from prev, update dates
  var T=today();
  var newData=Object.assign({},prevInv.data,{
    billingPeriod:newPeriod,
    dateSubmitted:T,
    sigDate1:T,
    sigDate2:T,
    tasks:{svc:newSvc,cplx:newCplx}
  });
  return {
    billingPeriod:newPeriod,
    savedAt:new Date().toLocaleString(),
    status:'draft',
    invoiceNote:'',
    data:newData
  };
}

// Returns list of clients eligible for auto-gen (active, missing invoice for period, has a prior)
function findClientsEligibleForAutoGen(period){
  var profiles=getProfiles();var out=[];
  Object.keys(profiles).forEach(function(name){
    var p=profiles[name];
    if(p.status==='inactive'||p.status==='terminated'||p.status==='lost')return;
    if(p.clientStatus==='inactive'||p.clientStatus==='terminated'||p.clientStatus==='lost')return;
    if(!clientWasActiveInPeriod(p,period))return;
    var hasCurrent=(p.invoices||[]).some(function(i){return i.billingPeriod===period;});
    if(hasCurrent)return;
    // Find most recent prior invoice (highest savedAt)
    var prior=(p.invoices||[]).slice().sort(function(a,b){return new Date(b.savedAt||0)-new Date(a.savedAt||0);})[0];
    if(!prior)return;
    out.push({name:name,prevInv:prior});
  });
  return out;
}

// Bulk auto-generate: confirm, then create new draft invoices for all eligible clients
function autoGenerateMonthlyInvoices(period){
  var eligible=findClientsEligibleForAutoGen(period);
  if(!eligible.length){
    showAlert('No clients need auto-generation. Either everyone has an invoice for '+period+', or no one has a prior invoice to copy from.',{title:'Nothing to Generate'});
    return;
  }
  var names=eligible.map(function(e){return '• '+e.name;}).join('\n');
  showConfirm(
    'Auto-generate '+eligible.length+' invoice'+(eligible.length>1?'s':'')+' for '+period+'?\n\n'+
    names+'\n\n'+
    'For each client, this copies their most recent invoice and:\n'+
    '• Keeps daily tasks (Bathing, Dressing, etc.) checked every day\n'+
    '• Shifts weekly tasks (Laundry, Shopping, Travel) by 1 day so it varies from last month\n'+
    '• Leaves Hospital column empty (must be added manually if needed)\n'+
    '• Sets Date Submitted to today\n\n'+
    'All generated invoices are marked Draft — review before sending.',
    function(){_doAutoGenerateInvoices(eligible,period);},
    {title:'Auto-Generate Invoices',okText:'Generate '+eligible.length}
  );
}
function _doAutoGenerateInvoices(eligible,period){
  var profiles=getProfiles();
  var generated=0,skipped=0;
  var undoBatch={id:'b_'+Date.now()+'_'+Math.random().toString(36).slice(2,8),period:period,when:Date.now(),invoices:[]};
  eligible.forEach(function(e){
    var newInv=generateNextMonthInvoiceData(e.prevInv,period);
    if(!newInv){skipped++;return;}
    if(!newInv.id)newInv.id='auto_'+Date.now()+'_'+Math.random().toString(36).slice(2,9);
    if(!profiles[e.name].invoices)profiles[e.name].invoices=[];
    profiles[e.name].invoices.unshift(newInv);
    undoBatch.invoices.push({clientName:e.name,invoiceId:newInv.id,dataHash:_quickInvoiceHash(newInv.data)});
    generated++;
  });
  saveProfilesLS(profiles);
  eligible.forEach(function(e){if(profiles[e.name])saveProfileSP(e.name,profiles[e.name]);});
  if(undoBatch.invoices.length)_pushAutoGenUndoBatch(undoBatch);
  logActivity('invoice','Auto-generated '+generated+' invoice'+(generated!==1?'s':'')+' for '+period);
  showAlert(
    '✓ Auto-generated '+generated+' invoice'+(generated!==1?'s':'')+' for '+period+(skipped>0?' ('+skipped+' skipped — missing data)':'')+'.\n\n'+
    'All new invoices are marked Draft. An Undo banner will stay on the Clients page for 24 hours — use it if you change your mind.',
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
    if(prof.status==='inactive'||prof.status==='terminated'||prof.status==='lost')return;
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
          await _doMonthlyEmailSend(item.group.email,item.wname,period,readyToSend,alreadySent.length,hasIssues,missingInvoice);
          sentCount++;
        }catch(e){console.error('Send to '+item.wname+' failed:',e);}
      }
      showToast('✓ Sent batches to '+sentCount+' caseworker'+(sentCount===1?'':'s')+'.',5000);
      // Refresh preview
      previewMonthlyInvoices();
    },
    {title:'Send All Caseworker Emails',okText:'Send All',danger:false}
  );
}

async function sendMonthlyEmail(email,workerName,clients,period){
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
    console.warn('Monthly email send already in progress — refusing to start a parallel send to '+workerName);
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

  // Show progress
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
    }catch(e){console.error('PDF capture failed for '+c.name,e);}
  }

  // Restore page state
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
  // Follow-up vs first send wording
  var isFollowUp=alreadySentCount>0;
  var subj=(isFollowUp?'Additional Home Help Agency Invoices — ':'Home Help Agency Invoices — ')+period;
  var body;
  if(isFollowUp){
    body='<p>Dear '+workerDisplay+',</p>'+
      '<p>I apologize for the delay — please find attached '+(attachments.length>1?'some additional invoices':'an additional invoice')+' for the billing period <b>'+period+'</b> that '+(attachments.length>1?'were':'was')+' not included in my earlier email. Please confirm receipt.</p>'+
      '<p><b>Additional Clients</b></p><ul>'+
      attachments.map(function(a){return '<li>'+a.clientName+'</li>';}).join('')+
      '</ul><p>Please review and process at your earliest convenience.</p>'+
      '<p>Thank you,<br>Thomas Jaboro<br>Liberty Home Care Assistance<br>(248) 291-4106</p>';
  } else {
    body='<p>Dear '+workerDisplay+',</p>'+
      '<p>You will find invoice'+(attachments.length>1?'s':'')+' attached for our shared client'+(attachments.length>1?'s':'')+' for the billing period <b>'+period+'</b>. Please confirm receipt of these documents.</p>'+
      '<p><b>Clients</b></p><ul>'+
      attachments.map(function(a){return '<li>'+a.clientName+'</li>';}).join('')+
      '</ul><p>Please review and process at your earliest convenience.</p>'+
      '<p>Thank you,<br>Thomas Jaboro<br>Liberty Home Care Assistance<br>(248) 291-4106</p>';
  }

  var result=await sendMailWithPDF(email,subj,body,attachments,function(done,total,label){
    if(pb)pb.style.width=Math.round((done/Math.max(total,1))*100)+'%';
    if(pl)pl.textContent=label||('Sending '+done+' of '+total+'…');
  });
  if(po)po.classList.remove('open');

  // HIPAA audit: log every PHI email send (success OR failure)
  logEmailSend({
    type:'mass',
    recipient:email,
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
