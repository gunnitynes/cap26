/** @OnlyCurrentDoc  — this script can only open the spreadsheet it lives in, never your other files. */

/**
 * CATALYST OPERATIONS TEST — the private back end for the marking page.
 *
 *   The page (on GitHub Pages) talks only to this script. The script lives inside one Google Sheet on your
 *   account, and that Sheet is where everything is kept. Nobody needs a Google or claude.ai account:
 *   teachers sign in with their name and a password you give them.
 *
 *   Ops Test ▸ Add or reset a teacher…   creates a sign-in (passwords are stored only as salted hashes)
 *   Ops Test ▸ Remove a teacher…         signs them out everywhere and blocks the name
 *
 *   Every teacher sees and changes only their own marks. Other teachers' marks for a date are sent to
 *   the page only after someone reveals the grades for that date — this script enforces it, not the page.
 *
 * Sheets (Ops Test ▸ Set up sheets):
 *   Teachers   Name | Role | Password hash | Salt | Added
 *   Students   ID | Test date | Student | Studio | Added by | Added at | Deleted
 *   Marks      Student ID | Teacher | Marks (JSON) | Updated            one row per student per teacher
 *   Days       Test date | Grades revealed | By | At
 *   Results    written by the page when grades are revealed or exported: the tally, one row per student
 */

const SHEETS = {
  TEACHERS: { name: 'Teachers', header: ['Name', 'Role', 'Password hash', 'Salt', 'Added'] },
  STUDENTS: { name: 'Students', header: ['ID', 'Test date', 'Student', 'Studio', 'Added by', 'Added at', 'Deleted'] },
  MARKS:    { name: 'Marks',    header: ['Student ID', 'Teacher', 'Marks (JSON)', 'Updated'] },
  DAYS:     { name: 'Days',     header: ['Test date', 'Grades revealed', 'By', 'At'] },
  RESULTS:  { name: 'Results',  header: ['Test date', 'Student', 'Studio', 'Teachers', 'Skills marked', 'Score %', 'Critical errors',
                                         'Deduction', 'Final %', 'Grade', 'Core skills under 50%', 'Status', 'Feedback', 'Updated'] },
};
const SESSION_DAYS = 60;
const HASH_ROUNDS = 400;
const MAX_TRIES = 8;          // wrong passwords per name before a 10-minute pause
const MARK_VALUES = { y: 1, h: 1, g: 1, n: 1 };

// ─────────────────────────────────────────────────────────────── sheet menu

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Ops Test')
    .addItem('Set up sheets', 'setup')
    .addItem('Add or reset a teacher…', 'menuAddTeacher')
    .addItem('Remove a teacher…', 'menuRemoveTeacher')
    .addToUi();
}

function setup() {
  Object.keys(SHEETS).forEach(k => sheet_(SHEETS[k]));
  // dates stay text, so Sheets never turns 2026-11-04 into a date with a time zone
  sheet_(SHEETS.STUDENTS).getRange('B:B').setNumberFormat('@');
  sheet_(SHEETS.DAYS).getRange('A:A').setNumberFormat('@');
  sheet_(SHEETS.RESULTS).getRange('A:A').setNumberFormat('@');
  const first = SpreadsheetApp.getActive().getSheets()[0];
  if (first.getName() === 'Sheet1' || first.getName() === 'Tabelle1') {
    if (first.getLastRow() === 0) SpreadsheetApp.getActive().deleteSheet(first);
  }
  try { SpreadsheetApp.getUi().alert('Sheets are ready. Next: Ops Test ▸ Add or reset a teacher… (add yourself first, as admin).'); } catch (e) {}
}

function menuAddTeacher() {
  const ui = SpreadsheetApp.getUi();
  const n = ui.prompt('Add or reset a teacher', 'Teacher name, as they will type it to sign in:', ui.ButtonSet.OK_CANCEL);
  if (n.getSelectedButton() !== ui.Button.OK) return;
  const p = ui.prompt('Password for ' + n.getResponseText().trim(), 'At least 6 characters. Give it to the teacher in person.', ui.ButtonSet.OK_CANCEL);
  if (p.getSelectedButton() !== ui.Button.OK) return;
  const admin = ui.alert('Admin?', 'Admins can add and remove teachers from the page as well. Make this teacher an admin?', ui.ButtonSet.YES_NO);
  const r = setTeacher_(n.getResponseText(), p.getResponseText(), admin === ui.Button.YES ? 'admin' : 'teacher');
  ui.alert(r.ok ? (r.name + ' can now sign in.') : r.error);
}

function menuRemoveTeacher() {
  const ui = SpreadsheetApp.getUi();
  const n = ui.prompt('Remove a teacher', 'Name of the teacher to remove (their marks stay in the sheet):', ui.ButtonSet.OK_CANCEL);
  if (n.getSelectedButton() !== ui.Button.OK) return;
  const r = removeTeacher_(n.getResponseText());
  ui.alert(r.ok ? 'Removed. They are signed out everywhere.' : r.error);
}

// ─────────────────────────────────────────────────────────────── web API

function doGet() {
  return ContentService.createTextOutput('Catalyst Operations Test back end is running.');
}

/** The page sends POST text/plain {action, token, ...}; every answer is JSON {ok, ...} or {ok:false, error}. */
function doPost(e) {
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    out = route_(req);
  } catch (err) {
    out = { ok: false, error: 'Something went wrong on the server: ' + (err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function route_(req) {
  const a = String(req.action || '');
  if (a === 'ping') return { ok: true };
  if (a === 'login') return login_(req.name, req.password);
  const me = auth_(req.token);
  if (!me) return { ok: false, error: 'Please sign in again.', code: 'auth' };
  switch (a) {
    case 'state':          return state_(me);
    case 'addStudents':    return withLock_(() => addStudents_(me, req));
    case 'updateStudent':  return withLock_(() => updateStudent_(me, req));
    case 'saveEntry':      return withLock_(() => saveEntry_(me, req));
    case 'setReveal':      return withLock_(() => setReveal_(me, req));
    case 'publishResults': return withLock_(() => publishResults_(me, req));
    case 'changePassword': return withLock_(() => changePassword_(me, req));
    case 'logout':         PropertiesService.getScriptProperties().deleteProperty('s_' + req.token); return { ok: true };
    case 'listTeachers':   return admin_(me) || { ok: true, teachers: teachers_().map(t => ({ name: t.name, role: t.role })) };
    case 'addTeacher':     return admin_(me) || withLock_(() => setTeacher_(req.name, req.password, req.role === 'admin' ? 'admin' : 'teacher'));
    case 'removeTeacher':  return admin_(me) || (norm_(req.name) === norm_(me.name) ? { ok: false, error: 'You cannot remove yourself.' } : withLock_(() => removeTeacher_(req.name)));
  }
  return { ok: false, error: 'Unknown action.' };
}

// ─────────────────────────────────────────────────────────────── sign-in

function login_(name, password) {
  name = clean_(name, 40); password = String(password || '');
  if (!name || !password) return { ok: false, error: 'Type your name and password.' };
  const cache = CacheService.getScriptCache(), key = 'fail_' + norm_(name).replace(/[^a-z0-9]/g, '_');
  const fails = Number(cache.get(key) || 0);
  if (fails >= MAX_TRIES) return { ok: false, error: 'Too many wrong passwords. Wait 10 minutes and try again.' };
  const t = teachers_().find(x => norm_(x.name) === norm_(name));
  if (!t || hash_(password, t.salt) !== t.hash) {
    cache.put(key, String(fails + 1), 600);
    return { ok: false, error: 'That name and password do not match.' };
  }
  cache.remove(key);
  const token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('s_' + token, JSON.stringify({ name: t.name, exp: Date.now() + SESSION_DAYS * 864e5 }));
  return { ok: true, token, name: t.name, role: t.role };
}

function auth_(token) {
  if (!token || !/^[0-9a-f]{64}$/.test(String(token))) return null;
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty('s_' + token);
  if (!raw) return null;
  const s = JSON.parse(raw);
  if (s.exp < Date.now()) { props.deleteProperty('s_' + token); return null; }
  const t = teachers_().find(x => norm_(x.name) === norm_(s.name));
  if (!t) { props.deleteProperty('s_' + token); return null; }
  return { name: t.name, role: t.role, token };
}

function admin_(me) { return me.role === 'admin' ? null : { ok: false, error: 'Only an admin can do that.' }; }

function setTeacher_(name, password, role) {
  name = clean_(name, 40); password = String(password || '');
  if (!name) return { ok: false, error: 'Type a name.' };
  if (password.length < 6) return { ok: false, error: 'The password needs at least 6 characters.' };
  const sh = sheet_(SHEETS.TEACHERS), rows = sh.getDataRange().getValues();
  const salt = Utilities.getUuid().replace(/-/g, '');
  const row = [name, role, hash_(password, salt), salt, new Date()];
  for (let i = 1; i < rows.length; i++) {
    if (norm_(rows[i][0]) === norm_(name)) {
      sh.getRange(i + 1, 1, 1, 5).setValues([row]);
      endSessions_(name);
      return { ok: true, name, reset: true };
    }
  }
  sh.appendRow(row);
  return { ok: true, name };
}

function removeTeacher_(name) {
  const sh = sheet_(SHEETS.TEACHERS), rows = sh.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (norm_(rows[i][0]) === norm_(name)) { sh.deleteRow(i + 1); endSessions_(rows[i][0]); return { ok: true }; }
  }
  return { ok: false, error: 'No teacher with that name.' };
}

function changePassword_(me, req) {
  const t = teachers_().find(x => norm_(x.name) === norm_(me.name));
  if (!t || hash_(String(req.current || ''), t.salt) !== t.hash) return { ok: false, error: 'Your current password is not right.' };
  const r = setTeacher_(me.name, req.next, t.role);
  if (!r.ok) return r;
  return login_(me.name, req.next);
}

function endSessions_(name) {
  const props = PropertiesService.getScriptProperties(), all = props.getProperties();
  Object.keys(all).forEach(k => {
    if (k.indexOf('s_') !== 0) return;
    try { if (norm_(JSON.parse(all[k]).name) === norm_(name)) props.deleteProperty(k); } catch (e) { props.deleteProperty(k); }
  });
}

function hash_(password, salt) {
  let h = salt + '|' + password;
  for (let i = 0; i < HASH_ROUNDS; i++) {
    h = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + '|' + password, Utilities.Charset.UTF_8));
  }
  return h;
}

function teachers_() {
  return rows_(SHEETS.TEACHERS).filter(r => r[0]).map(r => ({ name: String(r[0]).trim(), role: String(r[1] || 'teacher'), hash: String(r[2]), salt: String(r[3]) }));
}

// ─────────────────────────────────────────────────────────────── data

function state_(me) {
  const students = rows_(SHEETS.STUDENTS).filter(r => r[0] && !r[6]).map(r => ({
    id: String(r[0]), day: day_(r[1]), name: String(r[2]), studio: String(r[3] || ''), addedBy: String(r[4] || ''), createdAt: +new Date(r[5]) || 0,
  }));
  const dayOf = {}; students.forEach(s => dayOf[s.id] = s.day);
  const days = {}; rows_(SHEETS.DAYS).forEach(r => { if (r[0]) days[day_(r[0])] = r[1] === true || r[1] === 'yes' || r[1] === 'TRUE'; });
  const entries = [];
  rows_(SHEETS.MARKS).forEach(r => {
    const sid = String(r[0]), teacher = String(r[1]);
    if (!(sid in dayOf)) return;
    let data = {}; try { data = JSON.parse(r[2] || '{}'); } catch (e) {}
    if (norm_(teacher) === norm_(me.name) || days[dayOf[sid]]) entries.push({ sid, teacher, data });
    else entries.push({ sid, teacher, progress: Object.keys(data.marks || {}).filter(k => data.marks[k] in MARK_VALUES).length });
  });
  return { ok: true, me: { name: me.name, role: me.role }, students, entries, days, teachers: teachers_().map(t => t.name) };
}

function addStudents_(me, req) {
  const day = checkDay_(req.day); if (!day) return { ok: false, error: 'Pick a test date first.' };
  const names = (Array.isArray(req.names) ? req.names : []).map(n => clean_(n, 80)).filter(Boolean).slice(0, 60);
  if (!names.length) return { ok: false, error: 'Type at least one name.' };
  const studio = clean_(req.studio, 40), sh = sheet_(SHEETS.STUDENTS), now = new Date();
  const rows = names.map(n => [Utilities.getUuid().slice(0, 13), day, n, studio, me.name, now, '']);
  const start = sh.getLastRow() + 1;
  sh.getRange(start, 2, rows.length, 1).setNumberFormat('@');
  sh.getRange(start, 1, rows.length, 7).setValues(rows);
  return { ok: true, ids: rows.map(r => r[0]) };
}

function updateStudent_(me, req) {
  const sh = sheet_(SHEETS.STUDENTS), rows = sh.getDataRange().getValues();
  const i = rows.findIndex((r, k) => k > 0 && String(r[0]) === String(req.id));
  if (i < 1 || rows[i][6]) return { ok: false, error: 'That student is no longer on the list.' };
  if (req.deleted) {
    const others = rows_(SHEETS.MARKS).filter(r => String(r[0]) === String(req.id) && norm_(r[1]) !== norm_(me.name))
      .filter(r => { try { return Object.keys(JSON.parse(r[2] || '{}').marks || {}).length; } catch (e) { return false; } });
    if (others.length && me.role !== 'admin') return { ok: false, error: others[0][1] + ' has marked this student, so only an admin can remove them.' };
    sh.getRange(i + 1, 7).setValue('yes (' + me.name + ', ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') + ')');
    return { ok: true };
  }
  if (req.day !== undefined) {
    const day = checkDay_(req.day); if (!day) return { ok: false, error: 'That date is not valid.' };
    sh.getRange(i + 1, 2).setNumberFormat('@').setValue(day);
  }
  if (req.studio !== undefined) sh.getRange(i + 1, 4).setValue(clean_(req.studio, 40));
  return { ok: true };
}

/** Merges a patch into this teacher's own row. null removes a key. Only known fields and mark values get in. */
function saveEntry_(me, req) {
  const sid = String(req.sid || '');
  const exists = rows_(SHEETS.STUDENTS).some(r => String(r[0]) === sid && !r[6]);
  if (!exists) return { ok: false, error: 'That student is no longer on the list.' };
  const p = req.patch || {}, clean = {};
  if (p.marks) { clean.marks = {}; Object.keys(p.marks).slice(0, 200).forEach(k => { const v = p.marks[k]; if (v === null || v in MARK_VALUES) clean.marks[clean_(k, 40)] = v; }); }
  if (p.notes) { clean.notes = {}; Object.keys(p.notes).slice(0, 200).forEach(k => { clean.notes[clean_(k, 40)] = p.notes[k] === null ? null : String(p.notes[k]).slice(0, 1000); }); }
  if (p.safety) { clean.safety = {}; Object.keys(p.safety).slice(0, 50).forEach(k => { const v = p.safety[k]; clean.safety[clean_(k, 40)] = v === null ? null : { t: Number(v.t) || Date.now(), text: clean_(v.text, 200) }; }); }
  if (p.summary !== undefined) clean.summary = String(p.summary || '').slice(0, 3000);
  const sh = sheet_(SHEETS.MARKS), rows = sh.getDataRange().getValues();
  const i = rows.findIndex((r, k) => k > 0 && String(r[0]) === sid && norm_(r[1]) === norm_(me.name));
  let data = {}; if (i > 0) { try { data = JSON.parse(rows[i][2] || '{}'); } catch (e) {} }
  data = merge_(data, clean);
  const row = [sid, me.name, JSON.stringify(data), new Date()];
  if (i > 0) sh.getRange(i + 1, 1, 1, 4).setValues([row]); else sh.appendRow(row);
  return { ok: true };
}

function setReveal_(me, req) {
  const day = checkDay_(req.day); if (!day) return { ok: false, error: 'That date is not valid.' };
  const sh = sheet_(SHEETS.DAYS), rows = sh.getDataRange().getValues();
  const row = [day, req.on ? 'yes' : 'no', me.name, new Date()];
  const i = rows.findIndex((r, k) => k > 0 && day_(r[0]) === day);
  if (i > 0) sh.getRange(i + 1, 1, 1, 4).setValues([row]);
  else { const n = sh.getLastRow() + 1; sh.getRange(n, 1).setNumberFormat('@'); sh.getRange(n, 1, 1, 4).setValues([row]); }
  return { ok: true };
}

/** The page sends the tally rows for one date; they replace that date's rows in Results. */
function publishResults_(me, req) {
  const day = checkDay_(req.day); if (!day) return { ok: false, error: 'That date is not valid.' };
  const sh = sheet_(SHEETS.RESULTS), width = SHEETS.RESULTS.header.length - 1;
  const rows = (Array.isArray(req.rows) ? req.rows : []).slice(0, 300).map(r => {
    const a = (Array.isArray(r) ? r : []).slice(0, width).map(v => typeof v === 'number' ? v : String(v == null ? '' : v).slice(0, 3000));
    while (a.length < width) a.push('');
    a[0] = day; a.push(new Date()); return a;
  });
  const all = sh.getDataRange().getValues();
  for (let i = all.length - 1; i >= 1; i--) if (day_(all[i][0]) === day) sh.deleteRow(i + 1);
  if (rows.length) {
    const start = sh.getLastRow() + 1;
    sh.getRange(start, 1, rows.length, 1).setNumberFormat('@');
    sh.getRange(start, 1, rows.length, width + 1).setValues(rows);
  }
  return { ok: true, rows: rows.length };
}

// ─────────────────────────────────────────────────────────────── helpers

function sheet_(def) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(def.name);
  if (!sh) {
    sh = ss.insertSheet(def.name);
    sh.getRange(1, 1, 1, def.header.length).setValues([def.header]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function rows_(def) { const v = sheet_(def).getDataRange().getValues(); return v.slice(1); }
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return { ok: false, error: 'The sheet is busy. Try again in a moment.', code: 'busy' };
  try { return fn(); } finally { lock.releaseLock(); }
}
function merge_(a, b) {
  const o = Object.assign({}, a || {});
  Object.keys(b || {}).forEach(k => {
    const v = b[k];
    if (v === null) delete o[k];
    else if (v && typeof v === 'object' && !Array.isArray(v) && o[k] && typeof o[k] === 'object') o[k] = merge_(o[k], v);
    else if (v && typeof v === 'object' && !Array.isArray(v)) o[k] = merge_({}, v);
    else o[k] = v;
  });
  return o;
}
function norm_(s) { return String(s || '').trim().toLowerCase().replace(/\s+/g, ' '); }
function clean_(s, max) { return String(s == null ? '' : s).replace(/[\u0000-\u001f]/g, ' ').trim().replace(/\s+/g, ' ').slice(0, max); }
function day_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v || '').trim().slice(0, 10);
}
function checkDay_(v) { const d = String(v || ''); return /^\d{4}-\d{2}-\d{2}$/.test(d) ? d : ''; }
