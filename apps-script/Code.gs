/** @OnlyCurrentDoc  — this script can only open the spreadsheet it lives in, never your other files. */

/**
 * CATALYST OPERATIONS TEST — the private back end for the marking page.
 *
 *   The page (on GitHub Pages) talks only to this script, and this script keeps everything in this Sheet.
 *   Teachers need no Google account to mark: on the page they pick their name and type the team password.
 *
 *   Adding a teacher: type their name in the Teachers tab. That's all: their name appears on the sign-in page.
 *   The team password is set once, on the Start here tab (it is hidden as soon as you type it).
 *
 *   The page is for marking: each teacher sees only their own marks there. Grading happens later, here,
 *   in the Grades tab, which combines every teacher's marks.
 *   (Anyone who can open this Sheet can of course see every tab.)
 *
 * Tabs (Ops Test ▸ Set up / tidy the sheet):
 *   Start here  what this is and how to add a teacher
 *   Teachers    Name | Role | Last sign-in
 *   Students    one row per student and test date; type or paste names here (a date picker on the date column)
 *   Marks       one row per student per teacher, every skill readable; the last column is the page's own copy
 *   Grades      the overview: every student's combined grade, stages, progress and errors, rebuilt by the script
 */

const PAGE_URL = 'https://gunnitynes.github.io/cap26/';
const SESSION_DAYS = 400;       // a phone stays signed in for the school year
const HASH_ROUNDS = 400;
const MAX_TRIES = 8;            // wrong passwords per name before a 10-minute pause
const MIN_PASSWORD = 6;
const SET = '●●●●●● set';       // what Start here shows once the team password is stored
const TEAM_LABEL = 'Team password';
const POINTS = { y: 1, h: 0.75, g: 0.5, n: 0 };
const MARK_LABEL = { y: 'Pass', h: 'Hint', g: 'Got there', n: 'Failed' };

// The skills, in test order. Keep the ids in step with CRIT in docs/index.html; the labels are only column headers here.
const SKILLS = [
  ['p1', 'Power-up order'], ['p2', 'CR level down/mute for connections'], ['p3', 'Safe levels, methodical'],
  ['m1', 'Correct mic, placed'], ['m2', 'Wall box → right channel'], ['m3', 'Healthy desk signal'], ['m4', 'Phantom power'],
  ['r1', 'Console → speakers'],
  ['t1', 'PT track in/out'], ['t2', 'PT record level'], ['t3', '1-2-3 test, other channel'], ['t4', 'Playback path'],
  ['c1', 'Usable cue mix'], ['c2', 'Pre vs post fader'], ['c3', 'Cue → headphone amp'], ['c4', 'Cue to a Live room'],
  ['s1', 'Shutdown order'], ['s2', 'Left as found'],
];
// Core skills (a combined mark under 50% means Resit) and the six stages, as on the page.
const CORE = { p1: 1, p2: 1, m3: 1, m4: 1, r1: 1, t1: 1, t3: 1, t4: 1, c1: 1, c2: 1, c3: 1, s1: 1 };
const STAGES = [['p', 'Power'], ['m', 'Mic'], ['r', 'Route'], ['t', 'Record'], ['c', 'Cue'], ['s', 'Shut down']];
const BANDS = [['Grade A', 75, '#0f7a57', '#d8f1e6'], ['Pass', 60, '#5a6276', '#eef0f4'], ['Resit', 0, '#bf2b2b', '#fbe0e0']];
const ERR_POINTS = 10;

const STUDIO_LIST = ['K3', 'K4'];   // the studios a test can be in
const DATE_FMT = 'dd.mm.yyyy';
const BLUE = '#0b43e0', INK = '#10131c', PAPER = '#f4f6fa', MUTE = '#5a6276';
const SHEETS = {
  START:    { name: 'Start here', header: [] },
  TEACHERS: { name: 'Teachers', header: ['Name', 'Role', 'Last sign-in'], widths: [220, 120, 170] },
  STUDENTS: { name: 'Students', header: ['Test date', 'Student', 'Studio', 'Added by', 'Added at', 'Removed', 'ID'],
              widths: [100, 200, 110, 120, 140, 220, 110] },
  MARKS:    { name: 'Marks', header: ['Test date', 'Student', 'Teacher', 'Skills marked', 'Teacher score %', 'Critical errors']
                .concat(SKILLS.map(s => s[1])).concat(['Feedback', 'Updated', 'Student ID', 'Page data (do not edit)']),
              widths: [100, 170, 120, 100, 110, 220].concat(SKILLS.map(() => 92)).concat([260, 140, 110, 120]) },
  GRADES:   { name: 'Grades', header: [] },
};
const T = { NAME: 0, ROLE: 1, LAST: 2 };

// ─────────────────────────────────────────────────────────────── sheet menu and layout

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Ops Test')
    .addItem('Set up / tidy the sheet', 'setup')
    .addItem('Set the team password…', 'menuTeamPassword')
    .addItem('Add a teacher…', 'menuAddTeacher')
    .addItem('Remove a teacher…', 'menuRemoveTeacher')
    .addSeparator()
    .addItem('Refresh the Grades tab', 'buildGrades_')
    .addToUi();
}

/** Builds every tab, moves older layouts over without losing anything, and styles it all. Safe to run any time. */
function setup() {
  const ss = SpreadsheetApp.getActive();
  migrateTeachers_(); migrateStudents_(); migrateMarks_(); migrateDates_();
  Object.keys(SHEETS).forEach(k => { if (k !== 'START' && k !== 'GRADES') sheet_(SHEETS[k]); });
  ['Results', 'Days'].forEach(n => { const old = ss.getSheetByName(n); if (old) ss.deleteSheet(old); });   // no longer used: grading happens in Grades
  buildStart_();
  style_();
  buildGrades_();
  // tidy up the empty first sheet Google creates
  ['Sheet1', 'Tabelle1', 'Blatt1'].forEach(n => { const s = ss.getSheetByName(n); if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s); });
  try { SpreadsheetApp.getUi().alert(teamHash_() ? 'The sheet is set up.' : 'The sheet is set up. Next: type the team password on the Start here tab.'); } catch (e) {}
}

function buildStart_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEETS.START.name);
  if (!sh) sh = ss.insertSheet(SHEETS.START.name, 0);
  storeTeamPassword_();                       // keep a password typed there before the tab is rebuilt
  sh.clear();
  const rows = [
    ['Catalyst Operations Test', ''],
    ['Marking sheet for the music studio operations test · CAP 2026', ''],
    ['', ''],
    ['The page', PAGE_URL],
    ['', ''],
    [TEAM_LABEL, ''],
    ['', 'Everyone signs in with this one password. Type it in the cell above (6 characters or more); it is hidden as soon as you press Enter. Type a new one over the dots to change it.'],
    ['Add a teacher', 'Type their name in the Teachers tab. It then appears on the sign-in page: they pick it and type the team password.'],
    ['Add students', 'Open the Students tab. Type the test date, the student\'s name and the studio, one student per row.'],
    ['', 'Pasting a whole list? Paste the names into the Student column and type the date once at the top: rows below take the date and studio from the row above.'],
    ['', 'They appear on the page within a few seconds. Teachers can also add students on the page itself.'],
    ['Remove a student', 'Type anything (for example "absent") in their Removed column. Their row stays, so nothing is lost.'],
    ['Remove a teacher', 'Delete their row in Teachers. Their marks stay in Marks.'],
    ['', ''],
    ['What the tabs hold', ''],
    ['Teachers', 'Who can sign in. Their names are the list on the sign-in page.'],
    ['Students', 'Everyone booked for a test: date, name, studio. Type or paste here, or add them on the page.'],
    ['Marks', 'One row per student per teacher, every skill readable. The last column is the page\'s own copy; leave it alone.'],
    ['Grades', 'The overview: every student\'s grade, progress, a colour per test stage and their critical errors. Kept up to date by itself.'],
    ['', ''],
    ['How the grade works', 'Pass 100 · Hint 75 · Got there 50 · Failed 0. Each skill is averaged over every teacher who marked it, every skill counts the same, each different critical error takes off 10, a Core skill under 50 means Resit. Grade A 75 and up, Pass 60–74, Resit under 60.'],
  ];
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.setColumnWidth(1, 190); sh.setColumnWidth(2, 640);
  sh.getRange(1, 1, rows.length, 2).setFontFamily('Arial').setFontSize(11).setVerticalAlignment('top').setWrap(true).setFontColor(INK);
  sh.getRange('A1').setFontSize(22).setFontWeight('bold').setFontColor(BLUE);
  sh.getRange('A2').setFontColor(MUTE);
  sh.getRange('A1:B1').merge(); sh.getRange('A2:B2').merge();
  sh.getRange('B4').setFormula('=HYPERLINK("' + PAGE_URL + '","' + PAGE_URL + '")').setFontWeight('bold');
  const tr = rows.findIndex(r => r[0] === TEAM_LABEL) + 1;
  sh.getRange(tr, 2).setValue(teamHash_() ? SET : '').setBackground('#fff8d6').setFontWeight('bold').setFontFamily('Courier New')
    .setBorder(true, true, true, true, null, null, BLUE, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(tr + 1, 2).setFontColor(MUTE);
  rows.forEach((r, i) => { if (r[0] && i > 1) sh.getRange(i + 1, 1).setFontWeight('bold'); });
  const tabsAt = rows.findIndex(r => r[0] === 'What the tabs hold') + 1;   // (row numbers move when rows are added above)
  sh.getRange(tabsAt, 1, 1, 2).setBackground(PAPER);
  sh.getRange(tabsAt + 1, 1, 5, 1).setFontColor(BLUE);
  sh.setHiddenGridlines(true);
  sh.setTabColor(BLUE);
  sh.setFrozenRows(0);
}

function style_() {
  Object.keys(SHEETS).forEach(k => {
    if (k === 'START' || k === 'GRADES') return;
    const def = SHEETS[k], sh = sheet_(def), n = def.header.length;
    sh.getRange(1, 1, 1, n).setValues([def.header]).setFontWeight('bold').setFontColor('#ffffff').setBackground(BLUE)
      .setVerticalAlignment('middle').setWrap(true);
    sh.setRowHeight(1, 40);
    sh.setFrozenRows(1);
    (def.widths || []).forEach((w, i) => sh.setColumnWidth(i + 1, w));
    const body = sh.getRange(2, 1, Math.max(1, sh.getMaxRows() - 1), n);
    body.setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');
    try { sh.getBandings().forEach(b => b.remove()); sh.getRange(1, 1, sh.getMaxRows(), n).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false); } catch (e) {}
    sh.getRange(1, 1, 1, n).setBackground(BLUE).setFontColor('#ffffff');
  });
  const t = sheet_(SHEETS.TEACHERS);
  t.getRange(2, T.ROLE + 1, t.getMaxRows() - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['teacher', 'admin'], true).setAllowInvalid(false).build());
  t.getRange(2, T.LAST + 1, t.getMaxRows() - 1, 1).setNumberFormat('dd.MM.yyyy HH:mm').setFontColor(MUTE);
  t.getRange(2, T.NAME + 1, t.getMaxRows() - 1, 1).setFontWeight('bold');
  const st = sheet_(SHEETS.STUDENTS);
  st.getRange(2, 1, st.getMaxRows() - 1, 1).setNumberFormat(DATE_FMT).setHorizontalAlignment('left').setDataValidation(
    SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false)
      .setHelpText('A test date, like 04.11.2026. Double-click for a calendar.').build());
  st.getRange(2, 5, st.getMaxRows() - 1, 1).setNumberFormat('dd.MM.yyyy HH:mm');
  st.getRange(2, 2, st.getMaxRows() - 1, 1).setFontWeight('bold');
  st.getRange(2, 3, st.getMaxRows() - 1, 1).setHorizontalAlignment('center').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(STUDIO_LIST, true).setAllowInvalid(false).setHelpText('K3 or K4').build());
  st.getRange(2, 4, st.getMaxRows() - 1, 3).setFontColor(MUTE);
  st.hideColumns(7);
  const m = sheet_(SHEETS.MARKS), mh = SHEETS.MARKS.header.length;
  m.getRange(2, 1, m.getMaxRows() - 1, 1).setNumberFormat(DATE_FMT).setHorizontalAlignment('left'); m.setFrozenColumns(3);
  m.getRange(2, mh - 2, m.getMaxRows() - 1, 1).setNumberFormat('dd.MM.yyyy HH:mm');
  m.hideColumns(mh - 1, 2);
  m.getRange(1, 7, 1, SKILLS.length).setTextRotation(0).setHorizontalAlignment('center');
  m.getRange(2, 7, m.getMaxRows() - 1, SKILLS.length).setHorizontalAlignment('center');
  const markRange = m.getRange(2, 7, m.getMaxRows() - 1, SKILLS.length);
  const rule = (range, text, bg, fg) => SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(text).setBackground(bg).setFontColor(fg).setBold(true).setRanges([range]).build();
  m.setConditionalFormatRules([
    rule(markRange, 'Pass', '#d8f1e6', '#0f7a57'), rule(markRange, 'Hint', '#fbeccf', '#a35f00'),
    rule(markRange, 'Got there', '#fde2d3', '#c2410c'), rule(markRange, 'Failed', '#fbe0e0', '#bf2b2b'),
  ]);
  [t, st, m].forEach(s => s.setTabColor(INK));
}

// ─────────────────────────────────────────────────────────────── typing in the Teachers tab

/** Runs by itself on every edit: students typed in, the team password typed on Start here, new teachers. */
function onEdit(e) {
  try {
    const sh = e && e.range && e.range.getSheet(), name = sh && sh.getName(), r = e.range.getRow();
    if (name === SHEETS.STUDENTS.name && r >= 2) { studentsTyped_(sh, r, e.range.getNumRows()); gradesDirty_(); return; }
    if (name === SHEETS.START.name) { storeTeamPassword_(); return; }
    if (name === SHEETS.TEACHERS.name && r >= 2) {
      const last = Math.min(sh.getLastRow(), r + e.range.getNumRows() - 1);
      for (let i = r; i <= last; i++) if (sh.getRange(i, T.NAME + 1).getValue() && !sh.getRange(i, T.ROLE + 1).getValue()) sh.getRange(i, T.ROLE + 1).setValue('teacher');
    }
  } catch (err) {}
}

/** The team password: typed in plain on Start here, kept only as a salted hash in the script's own storage. */
function teamCell_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.START.name);
  if (!sh) return null;
  const col = sh.getRange(1, 1, Math.max(1, sh.getLastRow()), 1).getValues();
  const i = col.findIndex(r => String(r[0]) === TEAM_LABEL);
  return i < 0 ? null : sh.getRange(i + 1, 2);
}
function teamHash_() { const p = PropertiesService.getScriptProperties(); const h = p.getProperty('team_hash'); return h ? { hash: h, salt: p.getProperty('team_salt') } : null; }
function storeTeamPassword_() {
  const cell = teamCell_(); if (!cell) return;
  const typed = String(cell.getValue() || '').trim();
  if (!typed || typed === SET) return;
  if (typed.length < MIN_PASSWORD) { cell.setNote('Use ' + MIN_PASSWORD + ' characters or more.'); return; }
  setTeamPassword_(typed);
  cell.setValue(SET).setNote('');
}
function setTeamPassword_(pw) {
  const salt = Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperties({ team_hash: hash_(pw, salt), team_salt: salt });
}

function studentsTyped_(sh, top, count) {
  const first = Math.max(2, top), last = Math.min(sh.getLastRow(), top + count - 1);
  if (last < first) return;
  const above = first > 2 ? sh.getRange(first - 1, 1, 1, 3).getValues()[0] : ['', '', ''];
  const rows = sh.getRange(first, 1, last - first + 1, 7).getValues();
  const who = 'Sheet';   // typed in the sheet (reading the editor's email would need an extra Google permission)
  let prevDate = above[0], prevStudio = above[2], changed = false;
  rows.forEach(r => {
    const name = String(r[S_.NAME] || '').trim();
    if (name) {
      if (!r[S_.DAY] && prevDate) { r[S_.DAY] = prevDate; changed = true; }
      if (!r[S_.STUDIO] && prevStudio) { r[S_.STUDIO] = prevStudio; changed = true; }
      if (!r[S_.ID]) { r[S_.ID] = newId_(); r[S_.BY] = r[S_.BY] || who; r[S_.AT] = r[S_.AT] || new Date(); changed = true; }
    }
    if (r[S_.DAY]) prevDate = r[S_.DAY];
    if (r[S_.STUDIO]) prevStudio = r[S_.STUDIO];
  });
  if (changed) sh.getRange(first, 1, rows.length, 7).setValues(rows);
}

/** Catches rows that arrived without the edit trigger (an import, a copy from another sheet). */
function ensureStudentIds_() {
  const sh = sheet_(SHEETS.STUDENTS), last = sh.getLastRow();
  if (last < 2) return;
  const ids = sh.getRange(2, S_.ID + 1, last - 1, 1).getValues(), names = sh.getRange(2, S_.NAME + 1, last - 1, 1).getValues();
  if (names.some((n, i) => String(n[0] || '').trim() && !ids[i][0])) studentsTyped_(sh, 2, last - 1);
}

// ─────────────────────────────────────────────────────────────── sheet menu (still works, now optional)

function menuTeamPassword() {
  const ui = SpreadsheetApp.getUi();
  const p = ui.prompt('Team password', 'The one password every teacher uses (6 characters or more):', ui.ButtonSet.OK_CANCEL);
  if (p.getSelectedButton() !== ui.Button.OK) return;
  const pw = p.getResponseText().trim();
  if (pw.length < MIN_PASSWORD) { ui.alert('Use ' + MIN_PASSWORD + ' characters or more.'); return; }
  setTeamPassword_(pw); const cell = teamCell_(); if (cell) cell.setValue(SET);
  ui.alert('Team password set.');
}

function menuAddTeacher() {
  const ui = SpreadsheetApp.getUi();
  const n = ui.prompt('Add a teacher', 'Their name, as it should appear on the sign-in page:', ui.ButtonSet.OK_CANCEL);
  if (n.getSelectedButton() !== ui.Button.OK) return;
  const r = setTeacher_(n.getResponseText(), null);
  ui.alert(r.ok ? r.name + ' can now sign in with the team password.' : r.error);
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
  if (a === 'teachers') return { ok: true, teachers: teachers_().map(t => t.name) };   // the names on the sign-in page
  if (a === 'login') return withLock_(() => login_(req.name, req.password));
  const me = auth_(req.token);
  if (!me) return { ok: false, error: 'Please sign in again.', code: 'auth' };
  switch (a) {
    case 'state':          withLock_(() => { ensureStudentIds_(); maybeBuildGrades_(); return {}; }); return state_(me);
    case 'addStudents':    return withLock_(() => addStudents_(me, req));
    case 'updateStudent':  return withLock_(() => updateStudent_(me, req));
    case 'saveEntry':      return withLock_(() => saveEntry_(me, req));
    case 'publishResults': return withLock_(() => publishResults_(me, req));
    case 'logout':         PropertiesService.getScriptProperties().deleteProperty('s_' + req.token); return { ok: true };
    case 'listTeachers':   return admin_(me) || { ok: true, teachers: teachers_().map(t => ({ name: t.name, role: t.role })) };
    case 'addTeacher':     return admin_(me) || withLock_(() => setTeacher_(req.name, req.role === 'admin' ? 'admin' : 'teacher'));
    case 'removeTeacher':  return admin_(me) || (norm_(req.name) === norm_(me.name) ? { ok: false, error: 'You cannot remove yourself.' } : withLock_(() => removeTeacher_(req.name)));
  }
  return { ok: false, error: 'Unknown action.' };
}

// ─────────────────────────────────────────────────────────────── sign-in

/** A teacher picks their name from the list and types the team password. */
function login_(name, password) {
  name = clean_(name, 40); password = String(password || '').trim();
  if (!name) return { ok: false, error: 'Pick your name.' };
  if (!password) return { ok: false, error: 'Type the team password.' };
  const cache = CacheService.getScriptCache(), key = 'fail_' + norm_(name).replace(/[^a-z0-9]/g, '_');
  const fails = Number(cache.get(key) || 0);
  if (fails >= MAX_TRIES) return { ok: false, error: 'Too many wrong tries. Wait 10 minutes and try again.' };
  const t = teachers_().find(x => norm_(x.name) === norm_(name));
  if (!t) return { ok: false, error: 'That name is not in the Teachers tab any more.' };
  storeTeamPassword_();                                          // a password typed in the sheet but not yet hidden
  const team = teamHash_();
  if (!team) return { ok: false, error: 'No team password is set yet. Ask the test organiser to set it on the Start here tab.' };
  if (hash_(password, team.salt) !== team.hash) { cache.put(key, String(fails + 1), 600); return { ok: false, error: 'That password is not right.' }; }
  cache.remove(key);
  return startSession_(t);
}

function startSession_(t) {
  const token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('s_' + token, JSON.stringify({ name: t.name, exp: Date.now() + SESSION_DAYS * 864e5 }));
  try { sheet_(SHEETS.TEACHERS).getRange(t.row, T.LAST + 1).setValue(new Date()); } catch (e) {}
  return { ok: true, token, name: t.name, role: t.role, days: SESSION_DAYS };
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

/** Adds a teacher (or changes their role). */
function setTeacher_(name, role) {
  name = clean_(name, 40);
  if (!name) return { ok: false, error: 'Type a name.' };
  const sh = sheet_(SHEETS.TEACHERS), rows = teachers_();
  const found = rows.find(x => norm_(x.name) === norm_(name));
  const r = found ? found.row : Math.max(2, sh.getLastRow() + 1);
  sh.getRange(r, 1, 1, 3).setValues([[name, role || (found && found.role) || (rows.length ? 'teacher' : 'admin'), found ? found.last : '']]);
  return { ok: true, name, reset: !!found };
}

function removeTeacher_(name) {
  const t = teachers_().find(x => norm_(x.name) === norm_(name));
  if (!t) return { ok: false, error: 'No teacher with that name.' };
  sheet_(SHEETS.TEACHERS).deleteRow(t.row);
  endSessions_(t.name);
  return { ok: true };
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
  return rows_(SHEETS.TEACHERS).map((r, i) => ({ r, i })).filter(x => String(x.r[T.NAME] || '').trim()).map(({ r, i }) => ({
    row: i + 2, name: String(r[T.NAME]).trim(), role: String(r[T.ROLE] || 'teacher'), last: r[T.LAST] || '' }))
    .sort((a, b) => a.row - b.row);
}

// ─────────────────────────────────────────────────────────────── data

const S_ = { DAY: 0, NAME: 1, STUDIO: 2, BY: 3, AT: 4, REMOVED: 5, ID: 6 };

function students_() {
  return rows_(SHEETS.STUDENTS).map((r, i) => ({ row: i + 2, id: String(r[S_.ID] || ''), day: day_(r[S_.DAY]), name: String(r[S_.NAME] || '').trim(),
    studio: String(r[S_.STUDIO] || ''), addedBy: String(r[S_.BY] || ''), createdAt: +new Date(r[S_.AT]) || 0, removed: !!r[S_.REMOVED] }))
    .filter(s => s.id && s.name && s.day);
}
function marksRows_() {
  const dataCol = SHEETS.MARKS.header.length - 1, idCol = dataCol - 1;
  return rows_(SHEETS.MARKS).map((r, i) => {
    let data = {}; try { data = JSON.parse(r[dataCol] || '{}'); } catch (e) {}
    return { row: i + 2, sid: String(r[idCol] || ''), teacher: String(r[2] || ''), data };
  }).filter(x => x.sid);
}

function state_(me) {
  const students = students_().filter(s => !s.removed);
  const dayOf = {}; students.forEach(s => dayOf[s.id] = s.day);
  const entries = [];
  marksRows_().forEach(m => {
    if (!(m.sid in dayOf)) return;
    if (norm_(m.teacher) === norm_(me.name)) entries.push({ sid: m.sid, teacher: m.teacher, data: m.data });
    else entries.push({ sid: m.sid, teacher: m.teacher, progress: Object.keys(m.data.marks || {}).filter(k => m.data.marks[k] in POINTS).length });
  });
  return { ok: true, me: { name: me.name, role: me.role },
    students: students.map(s => ({ id: s.id, day: s.day, name: s.name, studio: s.studio, addedBy: s.addedBy, createdAt: s.createdAt })),
    entries, teachers: teachers_().map(t => t.name) };
}

function addStudents_(me, req) {
  gradesDirty_();
  const day = checkDay_(req.day); if (!day) return { ok: false, error: 'Pick a test date first.' };
  const names = (Array.isArray(req.names) ? req.names : []).map(n => clean_(n, 80)).filter(Boolean).slice(0, 60);
  if (!names.length) return { ok: false, error: 'Type at least one name.' };
  const studio = STUDIO_LIST.indexOf(clean_(req.studio, 10)) >= 0 ? clean_(req.studio, 10) : STUDIO_LIST[0], sh = sheet_(SHEETS.STUDENTS), now = new Date();
  const rows = names.map(n => [toDate_(day), n, studio, me.name, now, '', newId_()]);
  const start = sh.getLastRow() + 1;
  sh.getRange(start, 1, rows.length, 7).setValues(rows);
  return { ok: true, ids: rows.map(r => r[6]) };
}

function updateStudent_(me, req) {
  gradesDirty_();
  const s = students_().find(x => x.id === String(req.id));
  if (!s || s.removed) return { ok: false, error: 'That student is no longer on the list.' };
  const sh = sheet_(SHEETS.STUDENTS);
  if (req.deleted) {
    const others = marksRows_().filter(m => m.sid === s.id && norm_(m.teacher) !== norm_(me.name) && Object.keys(m.data.marks || {}).length);
    if (others.length && me.role !== 'admin') return { ok: false, error: others[0].teacher + ' has marked this student, so only an admin can remove them.' };
    sh.getRange(s.row, S_.REMOVED + 1).setValue('by ' + me.name + ', ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm'));
    return { ok: true };
  }
  if (req.day !== undefined) {
    const day = checkDay_(req.day); if (!day) return { ok: false, error: 'That date is not valid.' };
    sh.getRange(s.row, S_.DAY + 1).setValue(toDate_(day));
    marksRows_().filter(m => m.sid === s.id).forEach(m => sheet_(SHEETS.MARKS).getRange(m.row, 1).setValue(toDate_(day)));
  }
  if (req.studio !== undefined) sh.getRange(s.row, S_.STUDIO + 1).setValue(clean_(req.studio, 40));
  return { ok: true };
}

/** Merges a patch into this teacher's own row. null removes a key. Only known fields and mark values get in. */
function saveEntry_(me, req) {
  gradesDirty_();
  const sid = String(req.sid || '');
  const s = students_().find(x => x.id === sid && !x.removed);
  if (!s) return { ok: false, error: 'That student is no longer on the list.' };
  const p = req.patch || {}, clean = {};
  if (p.marks) { clean.marks = {}; Object.keys(p.marks).slice(0, 200).forEach(k => { const v = p.marks[k]; if (v === null || v in POINTS) clean.marks[clean_(k, 40)] = v; }); }
  if (p.notes) { clean.notes = {}; Object.keys(p.notes).slice(0, 200).forEach(k => { clean.notes[clean_(k, 40)] = p.notes[k] === null ? null : String(p.notes[k]).slice(0, 1000); }); }
  if (p.safety) { clean.safety = {}; Object.keys(p.safety).slice(0, 50).forEach(k => { const v = p.safety[k]; clean.safety[clean_(k, 40)] = v === null ? null : { t: Number(v.t) || Date.now(), text: clean_(v.text, 200) }; }); }
  if (p.summary !== undefined) clean.summary = String(p.summary || '').slice(0, 3000);
  const mine = marksRows_().find(m => m.sid === sid && norm_(m.teacher) === norm_(me.name));
  const data = merge_(mine ? mine.data : {}, clean);
  const sh = sheet_(SHEETS.MARKS), row = marksRow_(s, me.name, data);
  const r = mine ? mine.row : sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, row.length).setValues([row]);
  return { ok: true };
}

/** The readable row in Marks: every skill spelled out, plus the page's own copy in the last column. */
function marksRow_(s, teacher, data) {
  const m = data.marks || {}, ids = SKILLS.map(x => x[0]);
  const marked = ids.filter(id => m[id] in POINTS);
  const score = marked.length ? Math.round(marked.reduce((a, id) => a + POINTS[m[id]], 0) / marked.length * 100) : '';
  const errors = Object.keys(data.safety || {}).map(k => data.safety[k]).filter(Boolean).map(x => x.text);
  return [toDate_(s.day), s.name, teacher, marked.length + '/' + ids.length, score, errors.join('; ')]
    .concat(ids.map(id => MARK_LABEL[m[id]] || ''))
    .concat([data.summary || '', new Date(), s.id, JSON.stringify(data)]);
}

/** Kept for older pages that still ask: rebuilds the overview now. */
function publishResults_(me, req) {
  buildGrades_();
  return { ok: true };
}

// ─────────────────────────────────────────────────────────────── the Grades overview tab

function gradesDirty_() { try { CacheService.getScriptCache().put('grades_dirty', '1', 21600); } catch (e) {} }
/** Called on every page refresh: rebuilds at most every 20 seconds, and only after something changed. */
function maybeBuildGrades_() {
  const c = CacheService.getScriptCache();
  if (!c.get('grades_dirty') || c.get('grades_recent')) return;
  buildGrades_();
}

/** The same sum the page uses: each skill averaged over every teacher who marked it, the average of the skills,
 *  minus 10 per different critical error, a core skill under 50% means Resit. */
function gradeOf_(byTeacher) {
  const teachers = Object.keys(byTeacher).filter(u => Object.keys(byTeacher[u].marks || {}).some(k => byTeacher[u].marks[k] in POINTS));
  const per = SKILLS.map(([id]) => {
    const v = teachers.map(u => byTeacher[u].marks[id]).filter(x => x in POINTS).map(x => POINTS[x]);
    return { id, avg: v.length ? v.reduce((a, b) => a + b, 0) / v.length : null };
  });
  const marked = per.filter(x => x.avg !== null);
  const score = marked.length ? marked.reduce((a, x) => a + x.avg, 0) / marked.length * 100 : null;
  const errors = [];
  Object.keys(byTeacher).forEach(u => Object.keys(byTeacher[u].safety || {}).forEach(k => { const e = byTeacher[u].safety[k]; if (e && errors.indexOf(e.text) < 0) errors.push(e.text); }));
  const final = score === null ? null : Math.max(0, score - errors.length * ERR_POINTS);
  const coreFails = marked.filter(x => CORE[x.id] && x.avg < 0.5).map(x => SKILLS.find(s => s[0] === x.id)[1]);
  const complete = marked.length === SKILLS.length;
  const band = final === null ? null : coreFails.length ? BANDS[BANDS.length - 1] : BANDS.find(b => final >= b[1] - 1e-9);
  const stages = STAGES.map(([p]) => { const xs = marked.filter(x => x.id[0] === p); return xs.length ? Math.round(xs.reduce((a, x) => a + x.avg, 0) / xs.length * 100) : ''; });
  return { teachers, marked: marked.length, score, errors, final, coreFails, complete, band, stages };
}

/** Rewrites the Grades tab: a summary per grade at the top, then one row per student with bars and colours. */
function buildGrades_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEETS.GRADES.name);
  if (!sh) sh = ss.insertSheet(SHEETS.GRADES.name, 1);
  const cache = CacheService.getScriptCache();
  cache.remove('grades_dirty'); cache.put('grades_recent', '1', 20);

  const students = students_().filter(s => !s.removed).sort((a, b) => a.day < b.day ? -1 : a.day > b.day ? 1 : a.name.localeCompare(b.name));
  const marks = {}; marksRows_().forEach(m => { (marks[m.sid] = marks[m.sid] || {})[m.teacher] = m.data; });
  const N = SKILLS.length, top = 8;
  const head = ['Test date', 'Student', 'Studio', 'Teachers', 'Marked', 'Progress'].concat(STAGES.map(x => x[1]))
    .concat(['Score', 'Errors', 'Final', '', 'Grade', 'Core skill under 50%', 'Critical errors']);
  const W = head.length, C = { STAGE: 7, SCORE: 13, ERR: 14, FINAL: 15, BAR: 16, GRADE: 17 };

  const counts = { 'Grade A': 0, Pass: 0, Resit: 0, Unfinished: 0 };
  const values = [], bars = [], groupStarts = [];
  let prevDay = '';
  students.forEach((s, i) => {
    const g = gradeOf_(marks[s.id] || {});
    const label = !g.marked ? 'Not started' : g.complete ? g.band[0] : 'Unfinished';
    if (g.complete) counts[g.band[0]]++; else counts.Unfinished++;
    if (s.day !== prevDay) { groupStarts.push(top + i); prevDay = s.day; }
    values.push([toDate_(s.day), s.name, s.studio, g.teachers.join(', '), g.marked + '/' + N, '']
      .concat(g.stages)
      .concat([g.score === null ? '' : Math.round(g.score), g.errors.length ? -g.errors.length * ERR_POINTS : '', g.final === null ? '' : Math.round(g.final), '',
               label, g.coreFails.join(', '), g.errors.join('; ')]));
    const col = g.complete ? g.band[2] : '#9aa3b5';
    bars.push(['=SPARKLINE(' + g.marked + ',{"charttype","bar";"max",' + N + ';"color1","' + BLUE + '"})',
               g.final === null ? '' : '=SPARKLINE(' + Math.round(g.final) + ',{"charttype","bar";"max",100;"color1","' + col + '"})']);
  });

  sh.clear();
  try { sh.getBandings().forEach(b => b.remove()); } catch (e) {}
  sh.setConditionalFormatRules([]);
  const total = students.length || 1;
  sh.getRange(1, 1).setValue('Grades');
  sh.getRange('A1').setFontSize(22).setFontWeight('bold').setFontColor(BLUE);
  sh.setRowHeights(1, top - 1, 21);                    // reset the heights a previous layout may have left
  sh.setRowHeight(2, 8);                                // a thin gap under the title
  // the summary: one block per grade, with a bar showing its share of all students
  const names = ['Grade A', 'Pass', 'Resit', 'Unfinished'];
  const colours = [BANDS[0], BANDS[1], BANDS[2], ['', 0, '#9aa3b5', PAPER]];
  sh.getRange(3, 1, 2, 1).setValues([['Overview'], [students.length + (students.length === 1 ? ' student' : ' students')]]);
  sh.getRange('A3').setFontWeight('bold'); sh.getRange('A4').setFontColor(MUTE);
  names.forEach((n, i) => {
    const c = 2 + i;
    sh.getRange(3, c).setValue(n).setFontWeight('bold').setFontColor(colours[i][2]);
    sh.getRange(4, c).setValue(counts[n]).setFontSize(24).setFontWeight('bold').setFontColor(colours[i][2]).setHorizontalAlignment('left');
    sh.getRange(5, c).setFormula('=SPARKLINE(' + counts[n] + ',{"charttype","bar";"max",' + total + ';"color1","' + colours[i][2] + '"})');
    sh.getRange(3, c, 3, 1).setBackground(colours[i][3]);
  });
  sh.setRowHeight(4, 40);

  sh.getRange(top - 1, 1, 1, W).setValues([head]).setFontWeight('bold').setFontColor('#ffffff').setBackground(BLUE).setVerticalAlignment('middle').setWrap(true);
  sh.setRowHeight(top - 1, 36);
  if (values.length) {
    sh.getRange(top, 1, values.length, W).setValues(values).setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');
    sh.getRange(top, 6, bars.length, 1).setFormulas(bars.map(b => [b[0]]));
    sh.getRange(top, C.BAR, bars.length, 1).setFormulas(bars.map(b => [b[1]]));
    sh.getRange(top, 1, values.length, 1).setNumberFormat(DATE_FMT).setHorizontalAlignment('left');
    sh.getRange(top, 2, values.length, 1).setFontWeight('bold');
    sh.getRange(top, C.STAGE, values.length, 6).setHorizontalAlignment('center').setNumberFormat('0');
    sh.getRange(top, C.SCORE, values.length, 3).setHorizontalAlignment('center').setNumberFormat('0');
    sh.getRange(top, C.FINAL, values.length, 1).setFontWeight('bold').setFontSize(12);
    sh.getRange(top, C.GRADE, values.length, 1).setHorizontalAlignment('center').setFontWeight('bold');
    sh.getRange(top, C.GRADE + 1, values.length, 2).setFontColor(MUTE).setWrap(true);
    // a heavier line where a new test date starts
    groupStarts.forEach(r => { try { sh.getRange(r, 1, 1, W).setBorder(true, null, null, null, null, null, INK, SpreadsheetApp.BorderStyle.SOLID_MEDIUM); } catch (e) {} });
    for (let r = top; r < top + values.length; r++) sh.setRowHeight(r, 26);
    const heat = range => SpreadsheetApp.newConditionalFormatRule()
      .setGradientMinpointWithValue('#f4c7c3', SpreadsheetApp.InterpolationType.NUMBER, '0')
      .setGradientMidpointWithValue('#fce8b2', SpreadsheetApp.InterpolationType.NUMBER, '60')
      .setGradientMaxpointWithValue('#b7e1cd', SpreadsheetApp.InterpolationType.NUMBER, '100').setRanges([range]).build();
    const gradeCol = sh.getRange(top, C.GRADE, values.length, 1);
    const text = (t, bg, fg) => SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(t).setBackground(bg).setFontColor(fg).setBold(true).setRanges([gradeCol]).build();
    sh.setConditionalFormatRules([
      heat(sh.getRange(top, C.STAGE, values.length, 6)), heat(sh.getRange(top, C.SCORE, values.length, 1)), heat(sh.getRange(top, C.FINAL, values.length, 1)),
      text('Grade A', BANDS[0][3], BANDS[0][2]), text('Pass', BANDS[1][3], INK), text('Resit', BANDS[2][3], BANDS[2][2]),
      text('Unfinished', PAPER, MUTE), text('Not started', '#ffffff', '#9aa3b5'),
      SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setFontColor(BANDS[2][2]).setBold(true).setRanges([sh.getRange(top, C.ERR, values.length, 1)]).build(),
    ]);
  } else {
    sh.getRange(top, 1).setValue('No students yet. Add them in the Students tab or on the page.').setFontColor(MUTE);
  }
  [110, 170, 90, 150, 64, 110].concat(STAGES.map(() => 66)).concat([58, 58, 58, 110, 104, 200, 300]).forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setFrozenRows(top - 1); sh.setFrozenColumns(2);
  sh.setHiddenGridlines(true);
  sh.setTabColor(BLUE);
}

// ─────────────────────────────────────────────────────────────── moving older layouts over (run by setup)

function headerOf_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh || sh.getLastRow() < 1) return null;
  return { sh, head: sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String), rows: sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues() : [] };
}
/** Older layouts (with passwords or codes per teacher) become Name | Role | Last sign-in. */
function migrateTeachers_() {
  const o = headerOf_(SHEETS.TEACHERS.name);
  if (!o || o.head[1] === 'Role') return;
  const col = n => o.head.indexOf(n);
  const rows = o.rows.filter(r => r[0]).map(r => [r[0], (col('Role') >= 0 ? r[col('Role')] : '') || 'teacher', col('Last sign-in') >= 0 ? r[col('Last sign-in')] : '']);
  o.sh.clear(); o.sh.getRange(1, 1, 1, 3).setValues([SHEETS.TEACHERS.header]);
  if (rows.length) o.sh.getRange(2, 1, rows.length, 3).setValues(rows);
}

function migrateStudents_() {
  const o = headerOf_(SHEETS.STUDENTS.name);
  if (!o || o.head[0] === 'Test date') return;
  // v1 layout: ID | Test date | Student | Studio | Added by | Added at | Deleted
  const rows = o.rows.filter(r => r[0]).map(r => [toDate_(day_(r[1])), r[2], r[3], r[4], r[5], r[6] || '', r[0]]);
  o.sh.clear(); o.sh.getRange(1, 1, 1, 7).setValues([SHEETS.STUDENTS.header]);
  if (rows.length) o.sh.getRange(2, 1, rows.length, 7).setValues(rows);
}
/** Dates stored as text by the first version ('2026-11-04') become real dates, so the date picker and sorting work. */
function migrateDates_() {
  [SHEETS.STUDENTS, SHEETS.MARKS].forEach(def => {
    const sh = SpreadsheetApp.getActive().getSheetByName(def.name);
    if (!sh || sh.getLastRow() < 2) return;
    const col = sh.getRange(2, 1, sh.getLastRow() - 1, 1), v = col.getValues();
    let changed = false;
    const out = v.map(r => { if (r[0] && !(r[0] instanceof Date) && day_(r[0])) { changed = true; return [toDate_(day_(r[0]))]; } return r; });
    if (changed) col.setValues(out);
  });
}
function migrateMarks_() {
  const o = headerOf_(SHEETS.MARKS.name);
  if (!o || o.head[0] === 'Test date') return;
  // v1 layout: Student ID | Teacher | Marks (JSON) | Updated
  const studs = {}; students_().forEach(s => studs[s.id] = s);
  const rows = o.rows.filter(r => r[0] && studs[String(r[0])]).map(r => {
    let data = {}; try { data = JSON.parse(r[2] || '{}'); } catch (e) {}
    return marksRow_(studs[String(r[0])], String(r[1]), data);
  });
  o.sh.clear(); const h = SHEETS.MARKS.header;
  o.sh.getRange(1, 1, 1, h.length).setValues([h]);
  if (rows.length) o.sh.getRange(2, 1, rows.length, h.length).setValues(rows);
}

// ─────────────────────────────────────────────────────────────── helpers

function sheet_(def) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(def.name);
  if (!sh) {
    sh = ss.insertSheet(def.name);
    if (def.header.length) {
      sh.getRange(1, 1, 1, def.header.length).setValues([def.header]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  }
  return sh;
}
function rows_(def) { const sh = sheet_(def); if (sh.getLastRow() < 2) return []; return sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(def.header.length, 1)).getValues(); }
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
/** Any date the sheet may hold (a real date, 2026-11-04, or 4.11.2026 typed as text) → '2026-11-04'; '' if it is not a date. */
function day_(v) {
  if (v instanceof Date) return isNaN(v) ? '' : Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const t = String(v || '').trim();
  let m = t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  m = t.match(/^(\d{1,2})[./](\d{1,2})[./](\d{2,4})$/);
  if (m) return (m[3].length === 2 ? '20' + m[3] : m[3]) + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  return '';
}
/** '2026-11-04' → a real date at noon (noon keeps it on the same day in any time zone). */
function toDate_(d) { const p = String(d).split('-').map(Number); return new Date(p[0], p[1] - 1, p[2], 12); }
function newId_() { return Utilities.getUuid().slice(0, 13); }
function checkDay_(v) { const d = String(v || ''); return /^\d{4}-\d{2}-\d{2}$/.test(d) ? d : ''; }
