"""One-off: turns the claude.ai artifact (ops-test.html) into the stand-alone page (site/index.html)."""
import re, pathlib
root = pathlib.Path(__file__).resolve().parent.parent
s = (root / 'ops-test.html').read_text()

def R(a, b, c=1):
    global s
    assert s.count(a) == c, (s.count(a), a[:100]); s = s.replace(a, b)
def RX(a, b):
    global s
    n = len(re.findall(a, s, re.S)); assert n == 1, (n, a[:80]); s = re.sub(a, lambda m: b, s, flags=re.S)

# ---------------------------------------------------------------- document head
R('''<title>Catalyst Operations Test</title>
<meta name="claude:viewport-fit-sides" content="cover">
<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1,user-scalable=no,viewport-fit=cover">''',
'''<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1,user-scalable=no,viewport-fit=cover">
<title>Catalyst Operations Test</title>
<meta name="robots" content="noindex">
<meta name="description" content="Marking sheet for the Catalyst music studio operations test.">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-title" content="Ops Test">
<meta name="apple-mobile-web-app-status-bar-style" content="default">
<meta name="theme-color" content="#ffffff" media="(prefers-color-scheme: light)">
<meta name="theme-color" content="#141824" media="(prefers-color-scheme: dark)">
<link rel="manifest" href="manifest.webmanifest">
<link rel="icon" href="icon.svg" type="image/svg+xml">
<link rel="apple-touch-icon" href="icon-180.png">
<script src="config.js"></script>''')
R('</style>\n\n<header class="bar">', '''*,::before,::after{box-sizing:border-box}
body{margin:0;font-size:15px}
:root{padding-bottom:env(safe-area-inset-bottom,0px)}
[hidden]{display:none!important}
</style>
</head>
<body>
<header class="bar">''')
R('boot();\n</script>', 'boot();\n</script>\n</body>\n</html>')

# ---------------------------------------------------------------- state, data layer and boot
grading = s[s.index('const skillsFor='):s.index('async function setReveal')]
i = s.index('/* ------------------------------------------------------------------ state */')
j = s.index('/* ------------------------------------------------------------------ render */')
s = s[:i] + r'''/* ------------------------------------------------------------------ state */
// the back end address comes from config.js (or ?api=… while testing); without one the page runs as a preview
const API=new URLSearchParams(location.search).get('api')||window.OPS_API||'';
const S={token:ls.get('ops.token',''),me:ls.get('ops.me',''),role:ls.get('ops.role',''),demo:!API,tests:{},teachers:{},pending:ls.get('ops.pending',{}),view:'roster',open:null,day:ls.get('ops.day',null)||today(),all:false,days:{},dayWant:{},confirm:null,cal:null,sync:'ok',ready:false,loading:false,published:{}};

/* ---- each teacher marks on their own: a student keeps one entry per teacher under by.<teacher name>.
   Until a date is revealed the back end sends other teachers' entries only as {_progress: skills marked}. */
function entries(t){const e={};for(const [k,v] of Object.entries(t.by||{}))if(v)e[k]=v;return e}
const mine=t=>!S.me?{}:entries(t)[S.me]||{};
const canMark=()=>!!S.me&&(S.demo||!!S.token);
const teacherName=n=>n;

'''+grading+r'''async function setReveal(day,on){
  // keep the choice on screen until the back end confirms it, so a refresh in between cannot flip it back
  S.dayWant[day]=on;S.days={...S.days,[day]:{revealed:on}};S.confirm=null;render();
  const done=()=>toast(on?'Grades revealed for '+fmtDay(day):'Grades hidden again');
  if(S.demo){delete S.dayWant[day];done();return}
  const j=await api('setReveal',{day,on});
  if(!j.ok){delete S.dayWant[day];S.days={...S.days,[day]:{revealed:!on}};render();toast(j.error||'Could not save. Try again.');return}
  done();await load();
  if(on)publishResults(day);
}

/* ------------------------------------------------------------------ back end */
async function api(action,body={}){
  const ctl=new AbortController(),tm=setTimeout(()=>ctl.abort(),30000);
  try{
    const r=await fetch(API,{method:'POST',headers:{'content-type':'text/plain;charset=utf-8'},body:JSON.stringify({action,token:S.token,...body}),signal:ctl.signal});
    const j=await r.json();
    if(!j.ok&&j.code==='auth'&&action!=='login')signOut('Please sign in again.');
    return j;
  }catch(e){return {ok:false,net:true,error:'No connection. Your marks are kept on this phone and saved when the connection is back.'}}
  finally{clearTimeout(tm)}
}
async function load(){
  if(S.demo||!S.token||S.loading)return;
  S.loading=true;const j=await api('state');S.loading=false;
  if(!j.ok){if(j.net)setSync('err');return}
  S.me=j.me.name;S.role=j.me.role;ls.set('ops.me',S.me);ls.set('ops.role',S.role);
  const tests={};
  for(const st of j.students)tests[st.id]={name:st.name,studio:st.studio,day:st.day,createdAt:st.createdAt,addedBy:st.addedBy,by:{}};
  for(const e of j.entries)if(tests[e.sid])tests[e.sid].by[e.teacher]=e.data||{_progress:e.progress||0};
  S.tests=tests;
  const days={};for(const [d,v] of Object.entries(j.days))days[d]={revealed:!!v};
  for(const [d,w] of Object.entries(S.dayWant)){if(!!days[d]?.revealed===w)delete S.dayWant[d];else days[d]={revealed:w}}
  S.days=days;S.teachers=Object.fromEntries(j.teachers.map(n=>[n,{name:n}]));
  // drop queued changes for students that no longer exist
  for(const id of Object.keys(S.pending))if(!tests[id]){delete S.pending[id]}
  ls.set('ops.pending',S.pending);
  if(S.sync==='err'&&!Object.keys(S.pending).length)setSync('ok');
  S.ready=true;refresh();
  if(revealed(S.day))publishResults(S.day);
}
// keep every phone up to date: a quiet refresh every 10 seconds while the page is on screen
setInterval(()=>{if(document.visibilityState==='visible'&&!flushing&&!Object.keys(S.pending).length)load()},10000);
document.addEventListener('visibilitychange',()=>{if(document.visibilityState==='visible')load()});

/* ------------------------------------------------------------------ write queue: taps save in the background */
function view(id){return S.pending[id]?merge(S.tests[id],S.pending[id]):S.tests[id]}
function allTests(){return Object.keys(S.tests).map(id=>({id,...view(id)})).filter(t=>t.name&&!t.deleted)}
function patch(id,obj){
  S.pending[id]=merge(S.pending[id],obj);ls.set('ops.pending',S.pending);
  refresh();flushSoon();
}
const patchMine=(id,obj)=>{if(!canMark())return;patch(id,{by:{[S.me]:obj}})};
let flushT=null,flushing=false,backoff=1000;
function flushSoon(d=400){clearTimeout(flushT);flushT=setTimeout(flush,d)}
async function flush(){
  if(flushing)return;const ids=Object.keys(S.pending);if(!ids.length){setSync('ok');return}
  if(S.demo){for(const id of ids){S.tests[id]=merge(S.tests[id],S.pending[id])}S.pending={};ls.set('ops.pending',{});setSync('ok');refresh();return}
  flushing=true;setSync('busy');
  for(const id of ids){
    const p=S.pending[id];if(!p)continue;
    const steps=[];
    if(p.day)steps.push(['updateStudent',{id,day:p.day}]);
    const m=p.by&&p.by[S.me];if(m)steps.push(['saveEntry',{sid:id,patch:m}]);
    let failed=false;
    for(const [a,b] of steps){
      const j=await api(a,b);
      if(!j.ok){
        if(j.net||j.code==='busy'){failed=true;break}
        // the back end refused it (e.g. the student was removed): drop it and say why
        toast(j.error||'That change could not be saved.');break;
      }
    }
    if(failed){flushing=false;setSync('err');flushSoon(backoff=Math.min(backoff*2,20000));return}
    S.tests[id]=merge(S.tests[id],p);
    if(S.pending[id]===p)delete S.pending[id];
    ls.set('ops.pending',S.pending);
  }
  backoff=1000;flushing=false;
  if(Object.keys(S.pending).length)flushSoon(50);else{setSync('ok');load()}
}
function setSync(s){S.sync=s;const el=$('#sync');if(!el)return;el.dataset.s=s;el.querySelector('span').textContent=s==='ok'?(S.demo?'Preview only':'Saved'):s==='busy'?'Saving…':'Not saved yet, retrying'}
addEventListener('online',()=>flushSoon(10));
// a tab closed with marks still waiting gets a warning
addEventListener('beforeunload',e=>{if(!S.demo&&Object.keys(S.pending).length){e.preventDefault();e.returnValue=''}});

/* ------------------------------------------------------------------ results into the Google Sheet */
async function publishResults(day){
  if(S.demo||!revealed(day))return;
  const rows=resultsRows(allTests().filter(t=>t.day===day));
  const key=JSON.stringify(rows);if(S.published[day]===key)return;
  const j=await api('publishResults',{day,rows});if(j.ok)S.published[day]=key;
}

/* ------------------------------------------------------------------ boot */
async function boot(){
  if(S.demo){startDemo();return}
  if(!S.token){S.view='login';renderBar();render();return}
  renderBar();render();
  await load();
  flushSoon(10);
}
function signOut(msg){
  if(S.token&&!msg)api('logout');
  S.token='';S.me='';S.role='';S.tests={};S.days={};S.teachers={};
  ls.set('ops.token','');ls.set('ops.me','');ls.set('ops.role','');
  if(msg)toast(msg);go('login');
}
function startDemo(){
  S.demo=true;S.pending={};S.me='Gunnar';S.role='admin';S.ready=true;
  S.teachers={'Gunnar':{name:'Gunnar'},'Anna':{name:'Anna'}};
  const d=today();
  const nx=new Date();nx.setDate(nx.getDate()+6);const later=nx.getFullYear()+'-'+String(nx.getMonth()+1).padStart(2,'0')+'-'+String(nx.getDate()).padStart(2,'0');
  const all=v=>Object.fromEntries(CRIT.map(c=>[c.id,v(c)]));
  const mk=(name,day,by)=>{const id=rid();S.tests[id]={name,studio:'Studio A',day,createdAt:Date.now()-Math.random()*1e6,by}};
  mk('Example · Lena',d,{'Gunnar':{marks:{p1:'y',p2:'y',p3:'y',m1:'y',m2:'h',m3:'y',m4:'y',r1:'y',t1:'y',t2:'g',t3:'y'}}});
  mk('Example · Jonas',d,{'Gunnar':{marks:all(c=>c.id==='t2'?'h':'y')},'Anna':{marks:all(c=>c.id==='m2'||c.id==='c4'?'g':'y'),safety:{a:{t:Date.now(),text:'Phantom switched with fader up'}}}});
  mk('Example · Mira',d,{'Gunnar':{marks:{p1:'y',p2:'n',m1:'y'},safety:{a:{t:Date.now()-6e5,text:'Speakers switched with level up'}}}});
  mk('Example · Noah',later,{});mk('Example · Ida',later,{});
  S.view=new URLSearchParams(location.search).has('login')?'login':'roster';
  setSync('ok');renderBar();render();
}

''' + s[j:]

# ---------------------------------------------------------------- header + routing
R("""${S.view==='test'||S.view==='signin'&&S.teachers[S.me]?`<button class="iconbtn" id="back" aria-label="Back to students">←</button>`:''}""",
  """${S.view==='test'||S.view==='account'?`<button class="iconbtn" id="back" aria-label="Back to students">←</button>`:''}""")
R("""const nm=S.me?(S.teachers[S.me]?.name||S.meName||''):'';""","""const nm=S.view==='login'?'':S.me;""")
R("""$('#whoBtn')?.addEventListener('click',()=>{if(!S.demo)go('signin')});""","""$('#whoBtn')?.addEventListener('click',()=>go('account'));""")
R("""function render(){S.view==='signin'?renderSignin():S.view==='test'&&view(S.open)?renderSheet():renderRoster()}""",
  """function render(){S.view==='login'?renderLogin():S.view==='account'?renderAccount():S.view==='test'&&view(S.open)?renderSheet():renderRoster()}""")
R("""function refresh(){if(S.view==='test')syncSheet();else if(S.view==='roster')render()}""",
  """function refresh(){if(S.view==='test')syncSheet();else if(S.view==='roster')render();else if(S.view==='account')renderAccount(true)}""")

# ---------------------------------------------------------------- sign-in page (the landing page) and account page
RX(r"function renderSignin\(\)\{.*?\n\}\n", r'''function renderLogin(){
  const last=ls.get('ops.lastName','');
  $('#app').innerHTML=`<div class="landing">
    <section class="hero">
      <span class="eyebrow">Catalyst Berlin · Music Studios · CAP 2026</span>
      <h1>Mark the studio operations test on your phone.</h1>
      <p>One sheet per student, the same 18 skills for everyone, and a grade worked out the same way every time. Your marks stay yours until the test day is over.</p>
      <div class="mini" aria-hidden="true">
        <div class="mini-q">Understands phantom power <span class="tag core">Core</span></div>
        <div class="marks"><button type="button" data-v="y" aria-pressed="true" tabindex="-1"><i></i>Pass</button><button type="button" data-v="h" aria-pressed="false" tabindex="-1"><i></i>Hint</button><button type="button" data-v="g" aria-pressed="false" tabindex="-1"><i></i>Got there</button><button type="button" data-v="n" aria-pressed="false" tabindex="-1"><i></i>Failed</button></div>
      </div>
    </section>

    <form class="card login" id="loginForm" autocomplete="on">
      <b class="login-h">Sign in</b>
      <div class="field"><label for="lName">Your name</label><input type="text" id="lName" name="username" autocomplete="username" autocapitalize="words" value="${esc(last)}" placeholder="As Gunnar set it up for you"></div>
      <div class="field"><label for="lPass">Password</label><input type="password" id="lPass" name="password" autocomplete="current-password" placeholder="The password you were given"></div>
      <button class="primary" type="submit" id="lBtn">Sign in</button>
      <p class="readout err" id="lErr" hidden></p>
      ${S.demo?'<p class="readout">Preview: any name and password opens the example students.</p>':''}
    </form>

    <section class="how">
      <h2>How a test day runs</h2>
      <ol class="steps">
        <li><b>Pick the date</b><span>Swipe the week strip at the top and tap the test day. Students already added for that day are there for every teacher.</span></li>
        <li><b>Add your students</b><span>Paste the names, one per line. Anyone on the team can add them, ahead of time if you like.</span></li>
        <li><b>Mark as they work</b><span>Tap Pass, Hint, Got there or Failed for each skill. Tap a critical error the moment it happens. Every tap saves by itself.</span></li>
        <li><b>Reveal and download</b><span>When the day is done, reveal the grades. Everyone sees the results table, and it is copied into the results spreadsheet too.</span></li>
      </ol>
    </section>

    <section class="how">
      <h2>How the grade is worked out</h2>
      <div class="formula">
        <div><b>94</b><span>score</span></div><i>−</i><div><b>10</b><span>per critical error</span></div><i>=</i><div class="f-final"><b>84</b><span>final · Merit</span></div>
      </div>
      <div class="scale"><span style="--w:60">Resit</span><span style="--w:15">Pass 60</span><span style="--w:15">Merit 75</span><span style="--w:10">Dist. 90</span></div>
      <ul class="facts">
        <li>Each skill's mark is the average of every teacher who marked it, and every skill counts the same.</li>
        <li>A critical error counts once, however many teachers logged it.</li>
        <li>A Core skill that averages under 50 means Resit.</li>
        <li>Grades stay hidden from everyone until the day is revealed.</li>
      </ul>
    </section>

    <section class="how">
      <h2>Good to know</h2>
      <ul class="facts">
        <li><b>Install it.</b> On iPhone: Share ▸ Add to Home Screen. It then opens full screen like an app.</li>
        <li><b>Bad wifi is fine.</b> Marks wait on your phone and save as soon as the connection is back. The dot next to your name turns green when everything is saved.</li>
        <li><b>Forgot your password?</b> Ask Gunnar to reset it. You can change it yourself once you are signed in.</li>
      </ul>
    </section>
  </div>`;
  const f=$('#loginForm');
  if(!last)$('#lName').focus();
  f.onsubmit=async e=>{e.preventDefault();
    const name=$('#lName').value.trim(),password=$('#lPass').value,err=$('#lErr'),btn=$('#lBtn');
    err.hidden=true;if(!name||!password){err.textContent='Type your name and password.';err.hidden=false;return}
    if(S.demo){go('roster');return}
    btn.disabled=true;btn.textContent='Signing in…';
    const j=await api('login',{name,password});
    btn.disabled=false;btn.textContent='Sign in';
    if(!j.ok){err.textContent=j.error;err.hidden=false;return}
    S.token=j.token;S.me=j.name;S.role=j.role;ls.set('ops.token',S.token);ls.set('ops.me',S.me);ls.set('ops.role',S.role);ls.set('ops.lastName',j.name);
    S.view='roster';renderBar();render();await load();toast('Signed in as '+j.name);
  };
}

function renderAccount(quiet){
  if(quiet&&document.activeElement&&document.activeElement.closest('#app form'))return;
  const admin=S.role==='admin',names=Object.keys(S.teachers).sort((a,b)=>a.localeCompare(b));
  $('#app').innerHTML=`<div class="landing">
    <section class="card acct">
      <span class="eyebrow">Signed in as</span>
      <b class="login-h">${esc(S.me)}</b>
      <span class="readout" style="margin:0">${admin?'Admin: you can add and remove teachers.':'Teacher'}</span>
      <button class="ghost" id="outBtn" type="button">Sign out</button>
    </section>
    <form class="card login" id="pwForm" autocomplete="off">
      <b class="login-h">Change your password</b>
      <input type="text" name="username" value="${esc(S.me)}" autocomplete="username" hidden>
      <div class="field"><label for="pwOld">Current password</label><input type="password" id="pwOld" autocomplete="current-password"></div>
      <div class="field"><label for="pwNew">New password (6 characters or more)</label><input type="password" id="pwNew" autocomplete="new-password"></div>
      <button class="primary" type="submit">Change password</button>
      <p class="readout err" id="pwErr" hidden></p>
    </form>
    ${admin?`<form class="card login" id="tForm" autocomplete="off">
      <b class="login-h">Teachers</b>
      <ul class="tlist">${names.map(n=>`<li><span>${esc(n)}${n===S.me?' (you)':''}</span>${n===S.me?'':`<button class="more" type="button" data-rm="${esc(n)}">Remove</button>`}</li>`).join('')}</ul>
      <div class="field"><label for="tName">Add a teacher, or reset a password</label><input type="text" id="tName" placeholder="Name they sign in with" autocapitalize="words"></div>
      <div class="field"><label for="tPass">Password for them (6 characters or more)</label><input type="text" id="tPass" placeholder="Tell them in person" autocomplete="off"></div>
      <label class="check"><input type="checkbox" id="tAdmin"> Admin (can manage teachers)</label>
      <button class="primary" type="submit">Save teacher</button>
      <p class="readout err" id="tErr" hidden></p>
    </form>`:''}
  </div>`;
  $('#outBtn').onclick=()=>signOut();
  $('#pwForm').onsubmit=async e=>{e.preventDefault();const err=$('#pwErr');err.hidden=true;
    if(S.demo){toast('Preview only');return}
    const j=await api('changePassword',{current:$('#pwOld').value,next:$('#pwNew').value});
    if(!j.ok){err.textContent=j.error;err.hidden=false;return}
    S.token=j.token;ls.set('ops.token',S.token);$('#pwOld').value=$('#pwNew').value='';toast('Password changed')};
  if(!admin)return;
  let rmArm='';
  $('#tForm').onclick=async e=>{const b=e.target.closest('[data-rm]');if(!b)return;const n=b.dataset.rm;
    if(rmArm!==n){rmArm=n;b.textContent='Tap to confirm';return}
    if(S.demo){toast('Preview only');return}
    const j=await api('removeTeacher',{name:n});toast(j.ok?n+' removed':j.error);if(j.ok){delete S.teachers[n];renderAccount()}};
  $('#tForm').onsubmit=async e=>{e.preventDefault();const err=$('#tErr');err.hidden=true;
    const name=$('#tName').value.trim(),password=$('#tPass').value;
    if(S.demo){toast('Preview only');return}
    const j=await api('addTeacher',{name,password,role:$('#tAdmin').checked?'admin':'teacher'});
    if(!j.ok){err.textContent=j.error;err.hidden=false;return}
    toast(j.reset?'Password reset for '+j.name:j.name+' can now sign in');S.teachers[j.name]={name:j.name};renderAccount()};
}
''')

# ---------------------------------------------------------------- roster
R("""   ${S.demo?`<div class="banner"><b>Preview.</b> These students and the second teacher are examples and nothing here is saved. Open the page from its claude.ai link to mark real students.</div>`:''}
   ${!S.demo&&S.ready&&!S.me?`<div class="banner">Sign in to claude.ai to mark. Your marks are saved under your own account.</div>`:''}
   ${!S.writable?`<div class="banner">You can view results but not mark. Ask the page owner for Contributor access.</div>`:''}""",
"""   ${S.demo?`<div class="banner"><b>Preview.</b> These students and the second teacher are examples and nothing here is saved.</div>`:''}
   ${!S.demo&&!S.ready?`<div class="banner">Loading the students…</div>`:''}""")
R("""       <li>Each teacher signs in with their own claude.ai account and marks on their own phone. You never see another teacher's marks while you mark.</li>""",
  """       <li>Each teacher signs in with their own name and password and marks on their own phone. You never see another teacher's marks while you mark.</li>""")
R("""       <li>Grades stay hidden while you mark. When a test day is done, tap <b>Reveal grades</b> to see the results table and download it.</li>""",
  """       <li>Grades stay hidden while you mark. When a test day is done, tap <b>Reveal grades</b> to see the results table and download it. The same table is written into the Results tab of the spreadsheet.</li>""")
R("""  $('#addForm').onsubmit=e=>{e.preventDefault();if(!canMark())return;
    const names=$('#addNames').value.split('\\n').map(s=>s.trim()).filter(Boolean);if(!names.length){toast('Type at least one name');return}
    const studio=$('#addStudio').value.trim(),day=S.day;ls.set('ops.studio',studio);
    names.forEach((name,i)=>{const id=rid();S.pending[id]=merge(S.pending[id],{name,studio,day,createdAt:Date.now()+i,addedBy:S.me,by:{}})});
    ls.set('ops.pending',S.pending);flushSoon(10);render();toast((names.length>1?`Added ${names.length} students`:`Added ${names[0]}`)+' to '+fmtDay(day));
  };""",
"""  $('#addForm').onsubmit=async e=>{e.preventDefault();if(!canMark())return;
    const names=$('#addNames').value.split('\\n').map(s=>s.trim()).filter(Boolean);if(!names.length){toast('Type at least one name');return}
    const studio=$('#addStudio').value.trim(),day=S.day;ls.set('ops.studio',studio);
    const msg=(names.length>1?`Added ${names.length} students`:`Added ${names[0]}`)+' to '+fmtDay(day);
    if(S.demo){names.forEach((name,i)=>{S.tests[rid()]={name,studio,day,createdAt:Date.now()+i,by:{}}});render();toast(msg);return}
    const btn=e.target.querySelector('button[type=submit]');btn.disabled=true;btn.textContent='Adding…';
    const j=await api('addStudents',{names,studio,day});
    btn.disabled=false;
    if(!j.ok){btn.textContent='Add to '+fmtDay(day);toast(j.error);return}
    $('#addNames').value='';await load();toast(msg);
  };""")

# ---------------------------------------------------------------- exports: rows shared by the CSV and the Results sheet
R("""function resultsCSV(list){
  const head=""","""function resultsCSV(list){return toCSV([RESULTS_HEAD,...resultsRows(list)])}
const RESULTS_HEAD=""")
R("""  const rows=byDate(list).map(t=>{const g=gradeOf(t),rv=revealed(t.day),es=entries(t);""","""function resultsRows(list){
  return byDate(list).map(t=>{const g=gradeOf(t),rv=revealed(t.day),es=entries(t);""")
R("""g.complete?'Final':'Unfinished',fb]});
  return toCSV([head,...rows]);
}""","""g.complete?'Final':'Unfinished',fb]});
}""")
RX(r"async function offerFile\(name,csv\)\{.*?\n\}\n", r"""function offerFile(name,csv){
  try{
    const url=URL.createObjectURL(new Blob(['﻿'+csv],{type:'text/csv;charset=utf-8'}));
    const a=document.createElement('a');a.href=url;a.download=name;document.body.appendChild(a);a.click();a.remove();
    setTimeout(()=>URL.revokeObjectURL(url),4000);toast('Downloaded '+name);
  }catch(e){copyText(csv)}
  if(name.includes('results')&&!name.includes('all'))publishResults(S.day);
}
""")

# ---------------------------------------------------------------- student sheet: deleting asks the back end, which refuses if another teacher marked
R("""    patch(S.open,{deleted:true});toast('Student deleted');go('roster')};""",
  """    if(S.demo){delete S.tests[S.open];toast('Student removed');go('roster');return}
    const j=await api('updateStudent',{id:S.open,deleted:true});
    if(!j.ok){toast(j.error);return}
    delete S.tests[S.open];delete S.pending[S.open];ls.set('ops.pending',S.pending);toast('Student removed');go('roster');load()};""")
R("""let delArm=0;$('#delBtn').onclick=()=>{""","""let delArm=0;$('#delBtn').onclick=async()=>{""")
R("""(others.length?' Also marking: '+others.map(([uid,x])=>teacherName(uid)+' ('+scoreOf(x,t).n+'/'+sc.total+')').join(', ')+'.':'')""",
  """(others.length?' Also marking: '+others.map(([uid,x])=>teacherName(uid)+' ('+(x._progress??scoreOf(x,t).n)+'/'+sc.total+')').join(', ')+'.':'')""")

# ---------------------------------------------------------------- landing styles
R(".toast{", """.landing{display:flex;flex-direction:column;gap:18px;padding-block:18px 40px}
.hero{display:flex;flex-direction:column;gap:12px;padding-top:6px}
.hero h1{font-size:clamp(28px,8vw,40px);line-height:1.05;font-weight:800;letter-spacing:-.025em;margin:0}
.hero p{margin:0;font-size:16px;color:var(--mute);max-width:46ch}
.mini{background:var(--panel);border:1px solid var(--line);border-radius:var(--r);padding:12px;display:flex;flex-direction:column;gap:10px;box-shadow:0 10px 30px -18px var(--led);pointer-events:none}
.mini-q{font-weight:700;font-size:14px;display:flex;align-items:center;gap:8px}
.login{display:flex;flex-direction:column;gap:12px;padding:18px;border-color:var(--led)}
.login-h{font-size:20px;font-weight:800;letter-spacing:-.01em}
.err{color:var(--no);margin:0}
.acct{display:flex;flex-direction:column;gap:6px;padding:18px}
.acct .ghost{align-self:flex-start;margin-top:8px}
.how{display:flex;flex-direction:column;gap:12px}
.how h2{font-size:18px;font-weight:800;letter-spacing:-.01em}
.steps{list-style:none;margin:0;padding:0;display:flex;flex-direction:column;gap:2px;counter-reset:st}
.steps li{counter-increment:st;display:grid;grid-template-columns:34px 1fr;column-gap:10px;padding:10px 0;border-top:1px solid var(--line)}
.steps li::before{content:counter(st,decimal-leading-zero);grid-row:span 2;font-family:var(--mono);font-size:12px;font-weight:700;color:var(--led);padding-top:3px}
.steps b{font-size:15px}.steps span{font-size:14px;color:var(--mute)}
.formula{display:flex;align-items:center;gap:8px}
.formula div{flex:1;background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:10px;display:flex;flex-direction:column;align-items:center;gap:2px;min-width:0}
.formula div b{font-family:var(--mono);font-size:24px;line-height:1}
.formula div span{font-size:11px;color:var(--mute);text-align:center}
.formula i{font-style:normal;font-family:var(--mono);font-size:18px;color:var(--mute)}
.formula .f-final{border-color:var(--ok);background:var(--ok-soft)}.formula .f-final b{color:var(--ok)}
.scale{display:flex;height:26px;border-radius:8px;overflow:hidden;font-family:var(--mono);font-size:10px;font-weight:700}
.scale span{flex:var(--w);display:flex;align-items:center;justify-content:center;white-space:nowrap;overflow:hidden}
.scale span:nth-child(1){background:var(--no-soft);color:var(--no)}.scale span:nth-child(2){background:var(--sunk);color:var(--ink)}
.scale span:nth-child(3){background:var(--ok-soft);color:var(--ok)}.scale span:nth-child(4){background:var(--led-soft);color:var(--led)}
.facts{margin:0;padding-left:18px;display:flex;flex-direction:column;gap:6px;font-size:14px;color:var(--mute)}
.facts b{color:var(--ink)}
.tlist{list-style:none;margin:0;padding:0;display:flex;flex-direction:column}
.tlist li{display:flex;justify-content:space-between;align-items:center;padding:8px 0;border-bottom:1px solid var(--line);font-size:15px}
input[type=password]{background:var(--sunk);border:1px solid var(--line);border-radius:10px;padding:10px 12px;width:100%;font-size:16px}
.toast{""")

# the date strip and the account button only make sense once signed in
R("""   ${S.view==='roster'?dateStrip():''}`;""","""   ${S.view==='roster'?dateStrip():''}`;
  if(S.view==='login')$('#sync').hidden=true;""")
(root / 'site' / 'index.html').write_text(s)
print('site/index.html written', len(s))
