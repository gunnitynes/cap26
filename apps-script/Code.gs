/** @OnlyCurrentDoc  — this script can only open the spreadsheet it lives in, never your other files. */

/**
 * CATALYST OPERATIONS TEST — the private back end for the marking page.
 *
 *   The page (on GitHub Pages) talks only to this script, and this script keeps everything in this Sheet.
 *   Teachers need no Google account to mark: they sign in on the page with their name and a password.
 *
 *   Adding a teacher: type their name in the Teachers tab.
 *     · type a password next to it  → it is hidden the moment you press Enter, and they can sign in
 *     · or leave the password empty → they choose their own the first time they sign in
 *
 *   The page only shows a teacher their own marks until the grades for a date are revealed.
 *   (Anyone who can open this Sheet can of course see every tab.)
 *
 * Tabs (Ops Test ▸ Set up / tidy the sheet):
 *   Start here  what this is and how to add a teacher
 *   Teachers    Name | Password | Role | Status | Last sign-in    (+ two hidden columns holding the scrambled password)
 *   Students    one row per student and test date; type or paste names here (a date picker on the date column)
 *   Marks       one row per student per teacher, every skill readable; the last column is the page's own copy
 *   Days        tick "Grades revealed" to reveal a date (the page does this too)
 *   Results     the tally, written by the page when grades are revealed or downloaded
 */

const PAGE_URL = 'https://gunnitynes.github.io/ops-test/';
const SESSION_DAYS = 60;
const HASH_ROUNDS = 400;
const MAX_TRIES = 8;           // wrong passwords per name before a 10-minute pause
const MIN_PASSWORD = 6;
const SET = '●●●●●● set';      // what the Password column shows once a password is stored
const POINTS = { y: 1, h: 0.75, g: 0.5, n: 0 };
const MARK_LABEL = { y: 'Pass', h: 'Hint', g: 'Got there', n: 'Failed' };

// The skills, in test order. Keep the ids in step with CRIT in docs/index.html; the labels are only column headers here.
const SKILLS = [
  ['p1', 'Power-up order'], ['p2', 'CR level down for speakers/routing'], ['p3', 'Safe levels, methodical'],
  ['m1', 'Dynamic mic, placed'], ['m2', 'Wall box → right channel'], ['m3', 'Healthy desk signal'], ['m4', 'Phantom power'],
  ['r1', 'Mix bus → speakers'],
  ['t1', 'PT track in/out'], ['t2', 'PT record level'], ['t3', '1-2-3 test, other channel'], ['t4', 'Playback path'],
  ['c1', 'Usable cue mix'], ['c2', 'Pre vs post fader'], ['c3', 'Cue → headphone amp'], ['c4', 'Cue to a Live room'],
  ['s1', 'Shutdown order'], ['s2', 'Left as found'],
];

const DATE_FMT = 'dd.mm.yyyy';
const BLUE = '#0b43e0', INK = '#10131c', PAPER = '#f4f6fa', MUTE = '#5a6276';
const SHEETS = {
  START:    { name: 'Start here', header: [] },
  TEACHERS: { name: 'Teachers', header: ['Name', 'Password', 'Role', 'Status', 'Last sign-in', 'hash', 'salt'],
              widths: [180, 150, 100, 330, 150, 60, 60] },
  STUDENTS: { name: 'Students', header: ['Test date', 'Student', 'Studio', 'Added by', 'Added at', 'Removed', 'ID'],
              widths: [100, 200, 110, 120, 140, 220, 110] },
  MARKS:    { name: 'Marks', header: ['Test date', 'Student', 'Teacher', 'Skills marked', 'Teacher score %', 'Critical errors']
                .concat(SKILLS.map(s => s[1])).concat(['Feedback', 'Updated', 'Student ID', 'Page data (do not edit)']),
              widths: [100, 170, 120, 100, 110, 220].concat(SKILLS.map(() => 92)).concat([260, 140, 110, 120]) },
  DAYS:     { name: 'Days', header: ['Test date', 'Grades revealed', 'By', 'At'], widths: [110, 130, 120, 150] },
  RESULTS:  { name: 'Results', header: ['Test date', 'Student', 'Studio', 'Teachers', 'Skills marked', 'Score %', 'Critical errors',
                                        'Deduction', 'Final %', 'Grade', 'Core skills under 50%', 'Status', 'Feedback', 'Updated'],
              widths: [100, 170, 100, 170, 100, 80, 220, 90, 80, 110, 200, 100, 300, 140] },
};
const T = { NAME: 0, PASS: 1, ROLE: 2, STATUS: 3, LAST: 4, HASH: 5, SALT: 6 };

// ─────────────────────────────────────────────────────────────── sheet menu and layout

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Ops Test')
    .addItem('Set up / tidy the sheet', 'setup')
    .addItem('Add or reset a teacher…', 'menuAddTeacher')
    .addItem('Remove a teacher…', 'menuRemoveTeacher')
    .addToUi();
}

/** Builds every tab, moves older layouts over without losing anything, and styles it all. Safe to run any time. */
function setup() {
  const ss = SpreadsheetApp.getActive();
  migrateTeachers_(); migrateStudents_(); migrateMarks_(); migrateDates_();
  Object.keys(SHEETS).forEach(k => { if (k !== 'START') sheet_(SHEETS[k]); });
  buildStart_();
  style_();
  // tidy up the empty first sheet Google creates
  ['Sheet1', 'Tabelle1', 'Blatt1'].forEach(n => { const s = ss.getSheetByName(n); if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s); });
  refreshTeacherStatus_();
  try { SpreadsheetApp.getUi().alert('The sheet is set up. Add teachers in the Teachers tab: a name, and a password or leave it empty.'); } catch (e) {}
}

function buildStart_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEETS.START.name);
  if (!sh) sh = ss.insertSheet(SHEETS.START.name, 0);
  sh.clear();
  const rows = [
    ['Catalyst Operations Test', ''],
    ['Marking sheet for the music studio operations test · CAP 2026', ''],
    ['', ''],
    ['The page', PAGE_URL],
    ['', ''],
    ['Add a teacher', '1. Open the Teachers tab and type their name in a new row.'],
    ['', '2. Type a password next to it (6 characters or more). It is hidden as soon as you press Enter.'],
    ['', '   Or leave the password empty: they choose their own the first time they sign in on the page.'],
    ['', '3. Check the Status column. "Ready" means they can sign in.'],
    ['Add students', 'Open the Students tab. Type the test date, the student\'s name and the studio, one student per row.'],
    ['', 'Pasting a whole list? Paste the names into the Student column and type the date once at the top: rows below take the date and studio from the row above.'],
    ['', 'They appear on the page within a few seconds. Teachers can also add students on the page itself.'],
    ['Remove a student', 'Type anything (for example "absent") in their Removed column. Their row stays, so nothing is lost.'],
    ['Forgot a password?', 'Type a new one over the dots in the Password column, or clear it so they choose a new one.'],
    ['Remove a teacher', 'Delete their row in Teachers. Their marks stay in Marks.'],
    ['', ''],
    ['What the tabs hold', ''],
    ['Teachers', 'Who can sign in. Passwords are stored scrambled and never shown.'],
    ['Students', 'Everyone booked for a test: date, name, studio. Type or paste here, or add them on the page.'],
    ['Marks', 'One row per student per teacher, every skill readable. The last column is the page\'s own copy; leave it alone.'],
    ['Days', 'Tick "Grades revealed" to show the grades for a date on the page. The page can do this too.'],
    ['Results', 'The final tally per student. The page writes it when grades are revealed or downloaded.'],
    ['', ''],
    ['How the grade works', 'Pass 100 · Hint 75 · Got there 50 · Failed 0. Each skill is averaged over every teacher who marked it, every skill counts the same, each different critical error takes off 10, a Core skill under 50 means Resit. Distinction 90+, Merit 75–89, Pass 60–74, Resit under 60.'],
  ];
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.setColumnWidth(1, 190); sh.setColumnWidth(2, 640);
  sh.getRange(1, 1, rows.length, 2).setFontFamily('Arial').setFontSize(11).setVerticalAlignment('top').setWrap(true).setFontColor(INK);
  sh.getRange('A1').setFontSize(22).setFontWeight('bold').setFontColor(BLUE);
  sh.getRange('A2').setFontColor(MUTE);
  sh.getRange('A1:B1').merge(); sh.getRange('A2:B2').merge();
  sh.getRange('B4').setFormula('=HYPERLINK("' + PAGE_URL + '","' + PAGE_URL + '")').setFontWeight('bold');
  rows.forEach((r, i) => { if (r[0] && i > 1) sh.getRange(i + 1, 1).setFontWeight('bold'); });
  const tabsAt = rows.findIndex(r => r[0] === 'What the tabs hold') + 1;
  sh.getRange(tabsAt, 1, 1, 2).setBackground(PAPER);
  sh.getRange(tabsAt + 1, 1, 5, 1).setFontColor(BLUE);
  sh.setHiddenGridlines(true);
  sh.setTabColor(BLUE);
  sh.setFrozenRows(0);
}

function style_() {
  Object.keys(SHEETS).forEach(k => {
    if (k === 'START') return;
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
  t.hideColumns(T.HASH + 1, 2);
  t.getRange(2, T.ROLE + 1, t.getMaxRows() - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['teacher', 'admin'], true).setAllowInvalid(false).build());
  t.getRange(2, T.LAST + 1, t.getMaxRows() - 1, 1).setNumberFormat('dd.MM.yyyy HH:mm');
  t.getRange(2, T.STATUS + 1, t.getMaxRows() - 1, 1).setFontColor(MUTE);
  t.getRange(2, T.NAME + 1, t.getMaxRows() - 1, 1).setFontWeight('bold');
  const st = sheet_(SHEETS.STUDENTS);
  st.getRange(2, 1, st.getMaxRows() - 1, 1).setNumberFormat(DATE_FMT).setHorizontalAlignment('left').setDataValidation(
    SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false)
      .setHelpText('A test date, like 04.11.2026. Double-click for a calendar.').build());
  st.getRange(2, 5, st.getMaxRows() - 1, 1).setNumberFormat('dd.MM.yyyy HH:mm');
  st.getRange(2, 2, st.getMaxRows() - 1, 1).setFontWeight('bold');
  st.getRange(2, 4, st.getMaxRows() - 1, 3).setFontColor(MUTE);
  st.hideColumns(7);
  const m = sheet_(SHEETS.MARKS), mh = SHEETS.MARKS.header.length;
  m.getRange(2, 1, m.getMaxRows() - 1, 1).setNumberFormat(DATE_FMT).setHorizontalAlignment('left'); m.setFrozenColumns(3);
  m.getRange(2, mh - 2, m.getMaxRows() - 1, 1).setNumberFormat('dd.MM.yyyy HH:mm');
  m.hideColumns(mh - 1, 2);
  m.getRange(1, 7, 1, SKILLS.length).setTextRotation(0).setHorizontalAlignment('center');
  m.getRange(2, 7, m.getMaxRows() - 1, SKILLS.length).setHorizontalAlignment('center');
  const markRange = m.getRange(2, 7, m.getMaxRows() - 1, SKILLS.length);
  const d = sheet_(SHEETS.DAYS);
  d.getRange(2, 1, d.getMaxRows() - 1, 1).setNumberFormat(DATE_FMT).setHorizontalAlignment('left'); d.getRange(2, 2, d.getMaxRows() - 1, 1).insertCheckboxes();
  d.getRange(2, 4, d.getMaxRows() - 1, 1).setNumberFormat('dd.MM.yyyy HH:mm');
  const r = sheet_(SHEETS.RESULTS);
  r.getRange(2, 1, r.getMaxRows() - 1, 1).setNumberFormat(DATE_FMT).setHorizontalAlignment('left'); r.setFrozenColumns(2);
  r.getRange(2, 14, r.getMaxRows() - 1, 1).setNumberFormat('dd.MM.yyyy HH:mm');
  const gradeCol = r.getRange(2, 10, r.getMaxRows() - 1, 1);
  const rule = (range, text, bg, fg) => SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(text).setBackground(bg).setFontColor(fg).setBold(true).setRanges([range]).build();
  r.setConditionalFormatRules([
    rule(gradeCol, 'Distinction', '#dce6ff', BLUE), rule(gradeCol, 'Merit', '#d8f1e6', '#0f7a57'),
    rule(gradeCol, 'Pass', '#eef0f4', INK), rule(gradeCol, 'Resit', '#fbe0e0', '#bf2b2b'),
  ]);
  m.setConditionalFormatRules([
    rule(markRange, 'Pass', '#d8f1e6', '#0f7a57'), rule(markRange, 'Hint', '#fbeccf', '#a35f00'),
    rule(markRange, 'Got there', '#fde2d3', '#c2410c'), rule(markRange, 'Failed', '#fbe0e0', '#bf2b2b'),
  ]);
  [t, st, m, d, r].forEach(s => s.setTabColor(INK));
}

// ─────────────────────────────────────────────────────────────── typing in the Teachers tab

/** Runs by itself on every edit. A password typed in the Teachers tab is scrambled and hidden straight away. */
function onEdit(e) {
  try {
    const sh0 = e && e.range && e.range.getSheet();
    if (sh0 && sh0.getName() === SHEETS.STUDENTS.name && e.range.getRow() >= 2) { studentsTyped_(sh0, e.range.getRow(), e.range.getNumRows()); return; }
  } catch (err) {}
  try {
    const sh = e && e.range && e.range.getSheet();
    if (!sh || sh.getName() !== SHEETS.TEACHERS.name || e.range.getRow() < 2) return;
    const top = e.range.getRow(), rows = e.range.getNumRows();
    const col = e.range.getColumn ? e.range.getColumn() : T.PASS + 1, cols = e.range.getNumColumns ? e.range.getNumColumns() : 1;
    const touchedPassword = col <= T.PASS + 1 && col + cols - 1 >= T.PASS + 1;
    for (let r = top; r < top + rows; r++) {
      // an emptied Password cell means "forget the old password": they choose a new one at next sign-in
      if (touchedPassword && !String(sh.getRange(r, T.PASS + 1).getValue() || '') && sh.getRange(r, T.HASH + 1).getValue()) {
        sh.getRange(r, T.HASH + 1, 1, 2).setValues([['', '']]);
        try { endSessions_(sh.getRange(r, T.NAME + 1).getValue()); } catch (err) {}
      }
      storeTypedPassword_(sh, r);
    }
    refreshTeacherStatus_();
  } catch (err) {}
}

function storeTypedPassword_(sh, r) {
  const row = sh.getRange(r, 1, 1, 7).getValues()[0];
  const typed = String(row[T.PASS] || '');
  if (!typed || typed === SET) return;
  if (typed.length < MIN_PASSWORD) return;              // the Status column explains
  const salt = Utilities.getUuid().replace(/-/g, '');
  sh.getRange(r, T.HASH + 1, 1, 2).setValues([[hash_(typed, salt), salt]]);
  sh.getRange(r, T.PASS + 1).setValue(SET);
  if (!row[T.ROLE]) sh.getRange(r, T.ROLE + 1).setValue(isFirstTeacher_(sh, r) ? 'admin' : 'teacher');
  try { endSessions_(row[T.NAME]); } catch (e) {}   // a new password signs the old one out everywhere
}

function isFirstTeacher_(sh, r) { return sh.getRange(2, 1, Math.max(1, r - 2), 1).getValues().every(v => !v[0]) || r === 2; }

function refreshTeacherStatus_() {
  const sh = sheet_(SHEETS.TEACHERS), last = sh.getLastRow();
  if (last < 2) return;
  const v = sh.getRange(2, 1, last - 1, 7).getValues(), seen = {};
  const out = v.map(row => {
    const name = String(row[T.NAME] || '').trim(), pass = String(row[T.PASS] || ''), key = norm_(name);
    if (!name) return [pass ? 'Type a name in column A' : ''];
    if (seen[key]) return ['Same name as another row: give one of them a different name'];
    seen[key] = 1;
    if (pass && pass !== SET && pass.length < MIN_PASSWORD) return ['Password too short: use ' + MIN_PASSWORD + ' characters or more'];
    if (row[T.HASH]) return ['Ready: can sign in'];
    return ['Waiting: chooses a password at first sign-in'];
  });
  sh.getRange(2, T.STATUS + 1, out.length, 1).setValues(out);
}

/**
 * Students typed or pasted into the Students tab: a row with a name gets an ID so the page shows it.
 * An empty date or studio is taken from the row above, so a pasted list only needs the date once.
 */
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

function menuAddTeacher() {
  const ui = SpreadsheetApp.getUi();
  const n = ui.prompt('Add or reset a teacher', 'Teacher name, as they will type it to sign in:', ui.ButtonSet.OK_CANCEL);
  if (n.getSelectedButton() !== ui.Button.OK) return;
  const p = ui.prompt('Password for ' + n.getResponseText().trim(), 'At least 6 characters. Leave empty to let them choose it at first sign-in.', ui.ButtonSet.OK_CANCEL);
  if (p.getSelectedButton() !== ui.Button.OK) return;
  const r = setTeacher_(n.getResponseText(), p.getResponseText(), null, true);
  ui.alert(r.ok ? (r.name + (p.getResponseText() ? ' can now sign in.' : ' chooses a password at first sign-in.')) : r.error);
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
  if (a === 'login') return withLock_(() => login_(req.name, req.password, req.confirm));
  const me = auth_(req.token);
  if (!me) return { ok: false, error: 'Please sign in again.', code: 'auth' };
  switch (a) {
    case 'state':          withLock_(() => { ensureStudentIds_(); return {}; }); return state_(me);
    case 'addStudents':    return withLock_(() => addStudents_(me, req));
    case 'updateStudent':  return withLock_(() => updateStudent_(me, req));
    case 'saveEntry':      return withLock_(() => saveEntry_(me, req));
    case 'setReveal':      return withLock_(() => setReveal_(me, req));
    case 'publishResults': return withLock_(() => publishResults_(me, req));
    case 'changePassword': return withLock_(() => changePassword_(me, req));
    case 'logout':         PropertiesService.getScriptProperties().deleteProperty('s_' + req.token); return { ok: true };
    case 'listTeachers':   return admin_(me) || { ok: true, teachers: teachers_().map(t => ({ name: t.name, role: t.role })) };
    case 'addTeacher':     return admin_(me) || withLock_(() => setTeacher_(req.name, req.password, req.role === 'admin' ? 'admin' : 'teacher', true));
    case 'removeTeacher':  return admin_(me) || (norm_(req.name) === norm_(me.name) ? { ok: false, error: 'You cannot remove yourself.' } : withLock_(() => removeTeacher_(req.name)));
  }
  return { ok: false, error: 'Unknown action.' };
}

// ─────────────────────────────────────────────────────────────── sign-in

/**
 * Three cases, all from the same form on the page:
 *   the name has a password       → check it
 *   the name has no password yet  → first sign-in: the page asks for it twice (confirm) and it is stored
 *   the name is not in Teachers   → say so plainly
 */
function login_(name, password, confirm) {
  name = clean_(name, 40); password = String(password || '');
  if (!name) return { ok: false, error: 'Type your name.' };
  const cache = CacheService.getScriptCache(), key = 'fail_' + norm_(name).replace(/[^a-z0-9]/g, '_');
  const fails = Number(cache.get(key) || 0);
  if (fails >= MAX_TRIES) return { ok: false, error: 'Too many wrong passwords. Wait 10 minutes and try again.' };
  const sh = sheet_(SHEETS.TEACHERS);
  const t = teachers_().find(x => norm_(x.name) === norm_(name));
  if (!t) return { ok: false, error: 'There is no teacher called "' + name + '" yet. Ask the test organiser to add you in the Teachers tab, or check the spelling.', code: 'unknown' };
  if (t.typed && t.typed !== SET && !t.hash) {                // typed in the sheet and not yet scrambled
    if (t.typed.length >= MIN_PASSWORD) storeTypedPassword_(sh, t.row);
    t.hash = sh.getRange(t.row, T.HASH + 1).getValue(); t.salt = sh.getRange(t.row, T.SALT + 1).getValue();
  }
  if (!t.hash) {
    if (!password) return { ok: false, error: 'Choose a password for your first sign-in.', code: 'activate', name: t.name };
    if (password.length < MIN_PASSWORD) return { ok: false, error: 'Use ' + MIN_PASSWORD + ' characters or more.', code: 'activate', name: t.name };
    if (confirm === undefined || confirm === null) return { ok: false, error: 'First sign-in: type the same password again to confirm it.', code: 'activate', name: t.name };
    if (String(confirm) !== password) return { ok: false, error: 'The two passwords are different. Try again.', code: 'activate', name: t.name };
    const salt = Utilities.getUuid().replace(/-/g, '');
    sh.getRange(t.row, T.HASH + 1, 1, 2).setValues([[hash_(password, salt), salt]]);
    sh.getRange(t.row, T.PASS + 1).setValue(SET);
    if (!t.roleSet) sh.getRange(t.row, T.ROLE + 1).setValue(t.role);
    refreshTeacherStatus_();
    return startSession_(t, true);
  }
  if (!password) return { ok: false, error: 'Type your password.' };
  if (hash_(password, t.salt) !== t.hash) {
    cache.put(key, String(fails + 1), 600);
    return { ok: false, error: 'That password is not right. Forgot it? Ask the test organiser to reset it in the Teachers tab.' };
  }
  cache.remove(key);
  return startSession_(t, false);
}

function startSession_(t, activated) {
  const token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('s_' + token, JSON.stringify({ name: t.name, exp: Date.now() + SESSION_DAYS * 864e5 }));
  try { sheet_(SHEETS.TEACHERS).getRange(t.row, T.LAST + 1).setValue(new Date()); } catch (e) {}
  return { ok: true, token, name: t.name, role: t.role, activated };
}

function auth_(token) {
  if (!token || !/^[0-9a-f]{64}$/.test(String(token))) return null;
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty('s_' + token);
  if (!raw) return null;
  const s = JSON.parse(raw);
  if (s.exp < Date.now()) { props.deleteProperty('s_' + token); return null; }
  const t = teachers_().find(x => norm_(x.name) === norm_(s.name));
  if (!t || !t.hash) { props.deleteProperty('s_' + token); return null; }
  return { name: t.name, role: t.role, token };
}

function admin_(me) { return me.role === 'admin' ? null : { ok: false, error: 'Only an admin can do that.' }; }

/** Adds a teacher or resets one. An empty password means "they choose it at first sign-in". */
function setTeacher_(name, password, role, allowEmpty) {
  name = clean_(name, 40); password = String(password || '');
  if (!name) return { ok: false, error: 'Type a name.' };
  if (password && password.length < MIN_PASSWORD) return { ok: false, error: 'The password needs at least ' + MIN_PASSWORD + ' characters.' };
  if (!password && !allowEmpty) return { ok: false, error: 'Type a password.' };
  const sh = sheet_(SHEETS.TEACHERS), rows = teachers_();
  const found = rows.find(x => norm_(x.name) === norm_(name));
  const r = found ? found.row : Math.max(2, sh.getLastRow() + 1);
  const salt = password ? Utilities.getUuid().replace(/-/g, '') : '';
  const finalRole = role || (found && found.role) || (rows.length ? 'teacher' : 'admin');
  sh.getRange(r, 1, 1, 7).setValues([[name, password ? SET : '', finalRole, '', found ? found.last : '', password ? hash_(password, salt) : '', salt]]);
  endSessions_(name);
  refreshTeacherStatus_();
  return { ok: true, name, reset: !!found };
}

function removeTeacher_(name) {
  const t = teachers_().find(x => norm_(x.name) === norm_(name));
  if (!t) return { ok: false, error: 'No teacher with that name.' };
  sheet_(SHEETS.TEACHERS).deleteRow(t.row);
  endSessions_(t.name);
  return { ok: true };
}

function changePassword_(me, req) {
  const t = teachers_().find(x => norm_(x.name) === norm_(me.name));
  if (!t || hash_(String(req.current || ''), t.salt) !== t.hash) return { ok: false, error: 'Your current password is not right.' };
  const r = setTeacher_(me.name, req.next, t.role, false);
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
  const v = rows_(SHEETS.TEACHERS);
  let firstAdminGiven = v.some(r => String(r[T.ROLE]) === 'admin');
  return v.map((r, i) => ({ r, i })).filter(x => String(x.r[T.NAME] || '').trim()).map(({ r, i }) => {
    let role = String(r[T.ROLE] || '');
    const roleSet = !!role;
    if (!role) { role = firstAdminGiven ? 'teacher' : 'admin'; firstAdminGiven = true; }
    return { row: i + 2, name: String(r[T.NAME]).trim(), role, roleSet, typed: String(r[T.PASS] || ''),
             hash: String(r[T.HASH] || ''), salt: String(r[T.SALT] || ''), last: r[T.LAST] || '' };
  });
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
  const days = {}; rows_(SHEETS.DAYS).forEach(r => { if (r[0]) days[day_(r[0])] = r[1] === true || r[1] === 'yes' || r[1] === 'TRUE'; });
  const entries = [];
  marksRows_().forEach(m => {
    if (!(m.sid in dayOf)) return;
    if (norm_(m.teacher) === norm_(me.name) || days[dayOf[m.sid]]) entries.push({ sid: m.sid, teacher: m.teacher, data: m.data });
    else entries.push({ sid: m.sid, teacher: m.teacher, progress: Object.keys(m.data.marks || {}).filter(k => m.data.marks[k] in POINTS).length });
  });
  return { ok: true, me: { name: me.name, role: me.role },
    students: students.map(s => ({ id: s.id, day: s.day, name: s.name, studio: s.studio, addedBy: s.addedBy, createdAt: s.createdAt })),
    entries, days, teachers: teachers_().filter(t => t.hash).map(t => t.name) };
}

function addStudents_(me, req) {
  const day = checkDay_(req.day); if (!day) return { ok: false, error: 'Pick a test date first.' };
  const names = (Array.isArray(req.names) ? req.names : []).map(n => clean_(n, 80)).filter(Boolean).slice(0, 60);
  if (!names.length) return { ok: false, error: 'Type at least one name.' };
  const studio = clean_(req.studio, 40), sh = sheet_(SHEETS.STUDENTS), now = new Date();
  const rows = names.map(n => [toDate_(day), n, studio, me.name, now, '', newId_()]);
  const start = sh.getLastRow() + 1;
  sh.getRange(start, 1, rows.length, 7).setValues(rows);
  return { ok: true, ids: rows.map(r => r[6]) };
}

function updateStudent_(me, req) {
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

function setReveal_(me, req) {
  const day = checkDay_(req.day); if (!day) return { ok: false, error: 'That date is not valid.' };
  const sh = sheet_(SHEETS.DAYS), rows = sh.getDataRange().getValues();
  const row = [toDate_(day), !!req.on, me.name, new Date()];
  const i = rows.findIndex((r, k) => k > 0 && day_(r[0]) === day);
  const r = i > 0 ? i + 1 : Math.max(2, sh.getLastRow() + 1);
  sh.getRange(r, 1, 1, 4).setValues([row]);
  try { sh.getRange(r, 2).insertCheckboxes(); sh.getRange(r, 2).setValue(!!req.on); } catch (e) {}
  return { ok: true };
}

/** The page sends the tally rows for one date; they replace that date's rows in Results. */
function publishResults_(me, req) {
  const day = checkDay_(req.day); if (!day) return { ok: false, error: 'That date is not valid.' };
  const sh = sheet_(SHEETS.RESULTS), width = SHEETS.RESULTS.header.length - 1;
  const rows = (Array.isArray(req.rows) ? req.rows : []).slice(0, 300).map(r => {
    const a = (Array.isArray(r) ? r : []).slice(0, width).map(v => typeof v === 'number' ? v : String(v == null ? '' : v).slice(0, 3000));
    while (a.length < width) a.push('');
    a[0] = toDate_(day); a.push(new Date()); return a;
  });
  const all = sh.getDataRange().getValues();
  for (let i = all.length - 1; i >= 1; i--) if (day_(all[i][0]) === day) sh.deleteRow(i + 1);
  if (rows.length) {
    const start = sh.getLastRow() + 1;
    sh.getRange(start, 1, rows.length, width + 1).setValues(rows);
  }
  return { ok: true, rows: rows.length };
}

// ─────────────────────────────────────────────────────────────── moving older layouts over (run by setup)

function headerOf_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh || sh.getLastRow() < 1) return null;
  return { sh, head: sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String), rows: sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues() : [] };
}
function migrateTeachers_() {
  const o = headerOf_(SHEETS.TEACHERS.name);
  if (!o || o.head[1] === 'Password') return;
  // v1 layout: Name | Role | Password hash | Salt | Added. A value without a salt was typed by hand: it is a password.
  const rows = o.rows.filter(r => r[0]).map(r => {
    const hasSalt = !!r[3];
    return [r[0], hasSalt ? SET : String(r[2] || ''), r[1] || '', '', '', hasSalt ? r[2] : '', hasSalt ? r[3] : ''];
  });
  o.sh.clear(); o.sh.getRange(1, 1, 1, 7).setValues([SHEETS.TEACHERS.header]);
  if (rows.length) o.sh.getRange(2, 1, rows.length, 7).setValues(rows);
  for (let r = 2; r < rows.length + 2; r++) storeTypedPassword_(o.sh, r);
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
  [SHEETS.STUDENTS, SHEETS.MARKS, SHEETS.DAYS, SHEETS.RESULTS].forEach(def => {
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
