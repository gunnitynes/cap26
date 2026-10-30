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
 *   Students    one row per student: test date, test time, name, studio; type or paste here (a date picker on the date column)
 *   Questions   the test itself: one row per question, in order; the page reads it, so editing it changes the test
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

// The questions live in the Questions tab, where anyone with the Sheet can change them. This is only the first version,
// written there by setup when the tab does not exist yet. The ids keep marks attached to their question.
//   stage, stage heading, question, ask out loud, core, studio, column name, id
const DEFAULT_QUESTIONS = [
  ['Power', 'Power up & safe listening', 'Powers up in the right order', '', true, 'Both', 'Power-up order', 'p1'],
  ['', '', 'Control-room level down/mute before any console connections', '', true, 'Both', 'CR level down/mute for connections', 'p2'],
  ['', '', 'Keeps listening levels safe and works methodically', '', false, 'Both', 'Safe levels, methodical', 'p3'],
  ['Mic', 'Connecting a Mic', 'Chooses the correct mic and places it sensibly', '', false, 'Both', 'Correct mic, placed', 'm1'],
  ['', '', 'Connects the live-room wall box to the right desk channel', '', false, 'Both', 'Wall box → right channel', 'm2'],
  ['', '', 'Gets a healthy signal on the desk', '', true, 'Both', 'Healthy desk signal', 'm3'],
  ['', '', 'Understands phantom power', 'If I had asked you to use the condenser instead, what would you have needed to do?', true, 'Both', 'Phantom power', 'm4'],
  ['Route', 'Route to the speakers', 'Routes the channel to the console and hears it in the speakers', '', true, 'Both', 'Console → speakers', 'r1'],
  ['Record', 'Record & play back in Pro Tools', 'Creates a Pro Tools track with the correct input and output', '', true, 'Both', 'PT track in/out', 't1'],
  ['', '', 'Healthy record level in Pro Tools', '', false, 'Both', 'PT record level', 't2'],
  ['', '', 'Records a 1-2-3 test and returns it on a different desk channel', '', true, 'Both', '1-2-3 test, other channel', 't3'],
  ['', '', 'Plays it back through the path you asked for', '', true, 'Both', 'Playback path', 't4'],
  ['Cue', 'Headphones & cue mix', 'Builds a usable cue mix on the aux sends', '', true, 'Both', 'Usable cue mix', 'c1'],
  ['', '', 'Understands pre-fader vs. post-fader', 'Should this send be pre or post fader, and why?', true, 'Both', 'Pre vs post fader', 'c2'],
  ['', '', 'Routes the cue to the headphone amp and checks it', '', true, 'Both', 'Cue → headphone amp', 'c3'],
  ['', '', 'Explains sending the cue to a selected Live room', 'Without going there: how would the same cue reach this Live room?', false, 'Both', 'Cue to a Live room', 'c4'],
  ['Shut down', 'Shut down & pack down', 'Knows the shutdown order', 'If you had to shut down the studio now, what would you switch off first?', true, 'Both', 'Shutdown order', 's1'],
  ['', '', 'Leaves the studio as it was found', '', false, 'Both', 'Left as found', 's2'],
];
const Q_ = { STAGE: 0, HEADING: 1, TEXT: 2, ASK: 3, CORE: 4, STUDIO: 5, LABEL: 6, ID: 7 };
const MARKS_HEAD = ['Test date', 'Student', 'Teacher', 'Skills marked', 'Teacher score %', 'Critical errors'];
const MARKS_TAIL = ['Feedback', 'Updated', 'Student ID', 'Page data (do not edit)'];
const BANDS = [['Grade A', 75, '#0f7a57', '#d8f1e6'], ['Pass', 60, '#5a6276', '#eef0f4'], ['Resit', 0, '#bf2b2b', '#fbe0e0']];
const ERR_POINTS = 10;

const STUDIO_LIST = ['K3', 'K4'];   // the studios a test can be in
const DATE_FMT = 'dd.mm.yyyy';
const BLUE = '#0b43e0', INK = '#10131c', PAPER = '#f4f6fa', MUTE = '#5a6276';
const SHEETS = {
  START:    { name: 'Start here', header: [] },
  TEACHERS: { name: 'Teachers', header: ['Name', 'Role', 'Last sign-in'], widths: [220, 120, 170] },
  STUDENTS: { name: 'Students', header: ['Test date', 'Test time', 'Student', 'Studio', 'Removed', 'ID'],
              widths: [110, 90, 220, 90, 220, 110] },
  QUESTIONS:{ name: 'Questions', header: ['Stage', 'Stage heading', 'Question', 'Ask out loud (optional)', 'Core', 'Studio', 'Column name', 'ID'],
              widths: [110, 220, 380, 320, 60, 80, 220, 90] },
  MARKS:    { name: 'Marks', header: [] },   // its skill columns follow the questions: see syncMarks_
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
    .addItem('Use the current questions for a test date…', 'menuUseCurrentQuestions')
    .addToUi();
}

/** Builds every tab, moves older layouts over without losing anything, and styles it all. Safe to run any time. */
function setup() {
  const ss = SpreadsheetApp.getActive();
  migrateTeachers_(); migrateStudents_(); seedQuestions_(); ensureQuestionIds_(); migrateMarks_(); migrateDates_();
  Object.keys(SHEETS).forEach(k => { if (k !== 'START' && k !== 'GRADES') sheet_(SHEETS[k]); });
  freezeMarkedDays_();
  syncMarks_();
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
    ['Add students', 'Open the Students tab. Type the test date, the test time (like 10:30), the student\'s name and the studio, one student per row.'],
    ['', 'Pasting a whole list? Paste the names into the Student column and type the date once at the top: rows below take the date and studio from the row above. Each student keeps their own time.'],
    ['', 'They appear on the page within a few seconds. Teachers can also add students on the page itself.'],
    ['Remove a student', 'Type anything (for example "absent") in their Removed column. Their row stays, so nothing is lost.'],
    ['Remove a teacher', 'Delete their row in Teachers. Their marks stay in Marks.'],
    ['Change the questions', 'Edit the Questions tab: one row per question, in test order. Add a row, delete a row, reword, tick Core, or set a studio. The page picks it up within a few seconds.'],
    ['', 'A test date keeps the questions it was first marked with, so grades already given never move (rewording still shows). To give a marked date the new questions anyway: Ops Test ▸ Use the current questions for a test date.'],
    ['', ''],
    ['What the tabs hold', ''],
    ['Teachers', 'Who can sign in. Their names are the list on the sign-in page.'],
    ['Students', 'Everyone booked for a test: date, time, name, studio. Type or paste here, or add them on the page.'],
    ['Questions', 'The test itself: every question, its stage, whether it is Core, and which studio it is for.'],
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
  sh.getRange(tabsAt + 1, 1, 6, 1).setFontColor(BLUE);
  sh.setHiddenGridlines(true);
  sh.setTabColor(BLUE);
  sh.setFrozenRows(0);
}

function style_() {
  Object.keys(SHEETS).forEach(k => {
    if (k === 'START' || k === 'GRADES' || k === 'MARKS') return;
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
  const st = sheet_(SHEETS.STUDENTS), sr = st.getMaxRows() - 1;
  try { st.showColumns(1, st.getMaxColumns()); st.getRange(1, 1, st.getMaxRows(), st.getMaxColumns()).clearDataValidations(); } catch (e) {}
  st.getRange(2, S_.DAY + 1, sr, 1).setNumberFormat(DATE_FMT).setHorizontalAlignment('left').setDataValidation(
    SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false)
      .setHelpText('A test date, like 04.11.2026. Double-click for a calendar.').build());
  st.getRange(2, S_.TIME + 1, sr, 1).setNumberFormat('HH:mm').setHorizontalAlignment('center').setFontFamily('Courier New').setFontWeight('bold');
  st.getRange(1, S_.TIME + 1).setNote('The student\'s test time, like 10:30 (9.30 or 930 also work).');
  st.getRange(2, S_.NAME + 1, sr, 1).setFontWeight('bold');
  st.getRange(2, S_.STUDIO + 1, sr, 1).setHorizontalAlignment('center').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(STUDIO_LIST, true).setAllowInvalid(false).setHelpText('K3 or K4').build());
  st.getRange(2, S_.REMOVED + 1, sr, 1).setFontColor(MUTE);
  st.hideColumns(S_.ID + 1);
  const q = sheet_(SHEETS.QUESTIONS), qr = q.getMaxRows() - 1;
  q.getRange(2, Q_.STAGE + 1, qr, 1).setFontWeight('bold').setFontColor(BLUE);
  q.getRange(2, Q_.HEADING + 1, qr, 1).setFontColor(MUTE);
  q.getRange(2, Q_.TEXT + 1, qr, 2).setWrap(true);
  q.getRange(2, Q_.TEXT + 1, qr, 1).setFontWeight('bold');
  q.getRange(2, Q_.ASK + 1, qr, 1).setFontStyle('italic').setFontColor(MUTE);
  q.getRange(2, Q_.CORE + 1, Math.max(1, q.getLastRow() - 1) + 30, 1).insertCheckboxes();
  q.getRange(2, Q_.CORE + 1, qr, 2).setHorizontalAlignment('center');
  q.getRange(2, Q_.STUDIO + 1, qr, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Both'].concat(STUDIO_LIST), true)
    .setAllowInvalid(true).setHelpText('Both, or a studio: a K3 question is only asked of students tested in K3.').build());
  q.getRange(2, Q_.LABEL + 1, qr, 1).setFontColor(MUTE);
  [['Short stage name: the page\'s stage bar and a Grades column. Leave it blank to keep the stage of the row above.'],
   ['The heading over the stage on the page. Only needed on the first row of a stage.'],
   ['What the teacher marks. One row per question, in the order of the test. Delete the row to remove the question.'],
   ['Optional: a question the teacher reads out loud, shown under the question.'],
   ['Tick for a Core skill: a combined mark under 50% means Resit.'],
   ['Both, K3 or K4. A K3 question only shows for students tested in K3.'],
   ['Optional: a short name for the Marks and Grades tabs.'],
   ['Given by the sheet. Keeps marks attached to this question when you reword or move it.']]
    .forEach((n, i) => q.getRange(1, i + 1).setNote(n[0]));
  q.hideColumns(Q_.ID + 1);
  styleMarks_(markColumns_().length);
  [t, st].forEach(s => s.setTabColor(INK));
  q.setTabColor(BLUE);
}

/** Marks: dates, the skill columns in colour, the page's own copy hidden. Run again whenever the columns change. */
function styleMarks_(n) {
  const m = sheet_(SHEETS.MARKS), w = MARKS_HEAD.length + n + MARKS_TAIL.length, rows = m.getMaxRows() - 1;
  try { m.showColumns(1, m.getMaxColumns()); } catch (e) {}
  const head = m.getRange(1, 1, 1, w);
  head.setFontWeight('bold').setFontColor('#ffffff').setBackground(BLUE).setVerticalAlignment('middle').setWrap(true);
  m.setRowHeight(1, 40); m.setFrozenRows(1); m.setFrozenColumns(3);
  MARKS_HEAD.map((x, i) => [100, 170, 120, 100, 110, 220][i]).concat(Array(n).fill(92)).concat([260, 140, 110, 120]).forEach((x, i) => m.setColumnWidth(i + 1, x));
  m.getRange(2, 1, rows, w).setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');
  try { m.getBandings().forEach(b => b.remove()); m.getRange(1, 1, m.getMaxRows(), w).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false); } catch (e) {}
  head.setBackground(BLUE).setFontColor('#ffffff');
  m.getRange(2, 1, rows, 1).setNumberFormat(DATE_FMT).setHorizontalAlignment('left');
  m.getRange(2, w - 2, rows, 1).setNumberFormat('dd.MM.yyyy HH:mm');
  m.hideColumns(w - 1, 2);
  if (!n) { m.setConditionalFormatRules([]); m.setTabColor(INK); return; }
  m.getRange(1, 7, 1, n).setTextRotation(0).setHorizontalAlignment('center');
  const markRange = m.getRange(2, 7, rows, n);
  markRange.setHorizontalAlignment('center');
  const rule = (range, text, bg, fg) => SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(text).setBackground(bg).setFontColor(fg).setBold(true).setRanges([range]).build();
  m.setConditionalFormatRules([
    rule(markRange, 'Pass', '#d8f1e6', '#0f7a57'), rule(markRange, 'Hint', '#fbeccf', '#a35f00'),
    rule(markRange, 'Got there', '#fde2d3', '#c2410c'), rule(markRange, 'Failed', '#fbe0e0', '#bf2b2b'),
  ]);
  m.setTabColor(INK);
}

// ─────────────────────────────────────────────────────────────── typing in the Teachers tab

/** Runs by itself on every edit: students typed in, the team password typed on Start here, new teachers. */
function onEdit(e) {
  try {
    const sh = e && e.range && e.range.getSheet(), name = sh && sh.getName(), r = e.range.getRow();
    if (name === SHEETS.STUDENTS.name && r >= 2) { studentsTyped_(sh, r, e.range.getNumRows()); gradesDirty_(); return; }
    if (name === SHEETS.START.name) { storeTeamPassword_(); return; }
    if (name === SHEETS.QUESTIONS.name && r >= 2) { ensureQuestionIds_(); gradesDirty_(); return; }
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
  const W = SHEETS.STUDENTS.header.length;
  const above = first > 2 ? sh.getRange(first - 1, 1, 1, W).getValues()[0] : [];
  const range = sh.getRange(first, 1, last - first + 1, W), rows = range.getValues(), shown = range.getDisplayValues();
  let prevDate = above[S_.DAY] || '', prevStudio = above[S_.STUDIO] || '', changed = false;
  rows.forEach((r, i) => {
    const name = String(r[S_.NAME] || '').trim();
    if (name) {
      if (!r[S_.DAY] && prevDate) { r[S_.DAY] = prevDate; changed = true; }
      if (!r[S_.STUDIO] && prevStudio) { r[S_.STUDIO] = prevStudio; changed = true; }
      if (!r[S_.ID]) { r[S_.ID] = newId_(); changed = true; }
    }
    // a time typed as 9.30 or 930 would show as a wrong time: store it as 09:30
    const tm = time_(r[S_.TIME], shown[i][S_.TIME]);
    if (r[S_.TIME] !== '' && !(r[S_.TIME] instanceof Date) && tm) { r[S_.TIME] = tm; changed = true; }
    if (r[S_.DAY]) prevDate = r[S_.DAY];
    if (r[S_.STUDIO]) prevStudio = r[S_.STUDIO];
  });
  if (changed) range.setValues(rows);
}

/** A test time from the sheet → '09:30', or '' if there is none. Takes a real time, 9:30, 9.30, 930 or 0930. */
function time_(v, shown) {
  if (v instanceof Date) v = shown || '';
  if (typeof v === 'number') {
    if (v > 0 && v < 1) { const m = Math.round(v * 1440); return pad2_(Math.floor(m / 60) % 24) + ':' + pad2_(m % 60); }   // a fraction of a day
    if (v >= 0 && v < 24) { const h = Math.floor(v), m = Math.round((v - h) * 100); return m < 60 ? pad2_(h) + ':' + pad2_(m) : ''; }   // 9.3 = 9.30
    if (v >= 100 && v < 2400 && v % 100 < 60) return pad2_(Math.floor(v / 100)) + ':' + pad2_(v % 100);
    return '';
  }
  const t = String(v || '').trim().toLowerCase();
  let m = t.match(/^(\d{1,2})[:.h](\d{2})(?::\d{2})?\s*(am|pm)?$/) || t.match(/^(\d{1,2})(\d{2})$/) || t.match(/^(\d{1,2})\s*(am|pm)$/);
  if (!m) return '';
  let h = +m[1], mi = m[2] && /^\d+$/.test(m[2]) ? +m[2] : 0;
  const ap = m[3] || (m[2] && !/^\d+$/.test(m[2]) ? m[2] : '');
  if (ap === 'pm' && h < 12) h += 12; if (ap === 'am' && h === 12) h = 0;
  return h < 24 && mi < 60 ? pad2_(h) + ':' + pad2_(mi) : '';
}
function pad2_(n) { return ('0' + n).slice(-2); }

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

function menuUseCurrentQuestions() {
  const ui = SpreadsheetApp.getUi();
  const p = ui.prompt('Use the current questions for a test date',
    'A marked test date keeps the questions it was first marked with. Type a date (like 04.11.2026) to give it the questions in the Questions tab now. Marks on removed questions stop counting; new questions start unmarked.', ui.ButtonSet.OK_CANCEL);
  if (p.getSelectedButton() !== ui.Button.OK) return;
  const day = day_(p.getResponseText());
  if (!day) { ui.alert('That is not a date. Type it like 04.11.2026.'); return; }
  const r = withLock_(() => useCurrentQuestions_(day));
  ui.alert(r.ok ? 'Done: that date now uses the current questions.' : r.error);
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
    case 'state':          withLock_(() => { ensureStudentIds_(); ensureQuestionIds_(); syncMarks_(); maybeBuildGrades_(); return {}; }); return state_(me);
    case 'addStudents':    return withLock_(() => addStudents_(me, req));
    case 'updateStudent':  return withLock_(() => updateStudent_(me, req));
    case 'saveEntry':      return withLock_(() => saveEntry_(me, req));
    case 'publishResults': return withLock_(() => publishResults_(me, req));
    case 'grades':         return grades_();
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

const S_ = { DAY: 0, TIME: 1, NAME: 2, STUDIO: 3, REMOVED: 4, ID: 5 };

function students_() {
  const sh = sheet_(SHEETS.STUDENTS), n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const range = sh.getRange(2, 1, n, SHEETS.STUDENTS.header.length), rows = range.getValues(), shown = range.getDisplayValues();
  return rows.map((r, i) => ({ row: i + 2, id: String(r[S_.ID] || ''), day: day_(r[S_.DAY]), time: time_(r[S_.TIME], shown[i][S_.TIME]),
    name: String(r[S_.NAME] || '').trim(), studio: String(r[S_.STUDIO] || ''), removed: !!r[S_.REMOVED] }))
    .filter(s => s.id && s.name && s.day);
}
/** Test date, then test time (students without one last), then name. */
function byTestTime_(a, b) { return a.day < b.day ? -1 : a.day > b.day ? 1 : (a.time || '99') < (b.time || '99') ? -1 : (a.time || '99') > (b.time || '99') ? 1 : a.name.localeCompare(b.name); }
function marksRows_() {
  const sh = sheet_(SHEETS.MARKS), last = sh.getLastRow(), w = sh.getLastColumn();
  if (last < 2 || !w) return [];
  const head = sh.getRange(1, 1, 1, w).getValues()[0].map(String);
  let dataCol = head.indexOf(MARKS_TAIL[3]), idCol = head.indexOf(MARKS_TAIL[2]);
  if (dataCol < 0) { dataCol = w - 1; idCol = w - 2; }
  const upCol = head.indexOf(MARKS_TAIL[1]);
  return sh.getRange(2, 1, last - 1, w).getValues().map((r, i) => {
    let data = {}; try { data = JSON.parse(r[dataCol] || '{}'); } catch (e) {}
    return { row: i + 2, sid: String(r[idCol] || ''), teacher: String(r[2] || ''), data, day: day_(r[0]), name: String(r[1] || ''), updated: upCol >= 0 ? r[upCol] : '' };
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
  // the questions: today's, plus the frozen set of every listed date that has been marked
  const questions = questions_(), qdays = {};
  students.forEach(s => { if (!(s.day in qdays)) qdays[s.day] = frozen_(s.day) ? questionsFor_(s.day, questions) : null; });
  Object.keys(qdays).forEach(d => { if (!qdays[d]) delete qdays[d]; });
  return { ok: true, me: { name: me.name, role: me.role },
    students: students.map(s => ({ id: s.id, day: s.day, time: s.time, name: s.name, studio: s.studio })),
    entries, teachers: teachers_().map(t => t.name), questions, qdays };
}

function addStudents_(me, req) {
  gradesDirty_();
  const day = checkDay_(req.day); if (!day) return { ok: false, error: 'Pick a test date first.' };
  // a line may start with its test time: "10:30 Kim"
  const lines = (Array.isArray(req.names) ? req.names : []).map(n => clean_(n, 90)).filter(Boolean).slice(0, 60)
    .map(l => { const m = l.match(/^(\d{1,2}[:.]\d{2})\s+(.+)$/); return m && time_(m[1]) ? { time: time_(m[1]), name: clean_(m[2], 80) } : { time: '', name: clean_(l, 80) }; });
  if (!lines.length) return { ok: false, error: 'Type at least one name.' };
  const studio = STUDIO_LIST.indexOf(clean_(req.studio, 10)) >= 0 ? clean_(req.studio, 10) : STUDIO_LIST[0], sh = sheet_(SHEETS.STUDENTS);
  const rows = lines.map(l => [toDate_(day), l.time, l.name, studio, '', newId_()]);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  return { ok: true, ids: rows.map(r => r[S_.ID]) };
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
    if (marksRows_().some(m => m.sid === s.id && Object.keys(m.data.marks || {}).length)) freezeDay_(day);
    marksRows_().filter(m => m.sid === s.id).forEach(m => sheet_(SHEETS.MARKS).getRange(m.row, 1).setValue(toDate_(day)));
  }
  if (req.studio !== undefined) sh.getRange(s.row, S_.STUDIO + 1).setValue(clean_(req.studio, 40));
  if (req.time !== undefined) sh.getRange(s.row, S_.TIME + 1).setValue(time_(req.time));
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
  if (clean.marks && Object.keys(clean.marks).some(k => clean.marks[k])) freezeDay_(s.day);   // the first mark fixes this date's questions
  const ctx = syncMarks_();
  const mine = marksRows_().find(m => m.sid === sid && norm_(m.teacher) === norm_(me.name));
  const data = merge_(mine ? mine.data : {}, clean);
  const sh = sheet_(SHEETS.MARKS), row = marksRow_(s, me.name, data, ctx);
  const r = mine ? mine.row : sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, row.length).setValues([row]);
  return { ok: true };
}

/** The readable row in Marks: every skill spelled out, plus the page's own copy in the last column. */
function marksRow_(s, teacher, data, ctx, updated) {
  ctx = ctx || marksCtx_();
  const m = data.marks || {}, ids = forStudio_(questionsFor_(s.day, ctx.live), s.studio).map(q => q.id);
  const marked = ids.filter(id => m[id] in POINTS);
  const score = marked.length ? Math.round(marked.reduce((a, id) => a + POINTS[m[id]], 0) / marked.length * 100) : '';
  const errors = Object.keys(data.safety || {}).map(k => data.safety[k]).filter(Boolean).map(x => x.text);
  return [s.day ? toDate_(s.day) : '', s.name, teacher, marked.length + '/' + ids.length, score, errors.join('; ')]
    .concat(ctx.cols.map(c => MARK_LABEL[m[c.id]] || ''))
    .concat([data.summary || '', updated || new Date(), s.id, JSON.stringify(data)]);
}

/** Keeps the Marks columns in step with the questions. When they change, every row is rewritten from its own page copy. */
function syncMarks_() {
  const ctx = marksCtx_(), head = marksHeader_(ctx.cols), sh = sheet_(SHEETS.MARKS), w = sh.getLastColumn();
  const now = w ? sh.getRange(1, 1, 1, w).getValues()[0].map(String) : [];
  if (now.join('\u0001') === head.join('\u0001')) return ctx;
  const studs = {}; students_().forEach(s => studs[s.id] = s);
  const rows = marksRows_().map(m => marksRow_(studs[m.sid] || { id: m.sid, day: m.day, name: m.name, studio: '' }, m.teacher, m.data, ctx, m.updated));
  if (sh.getMaxColumns() < head.length) sh.insertColumnsAfter(sh.getMaxColumns(), head.length - sh.getMaxColumns());
  sh.clearContents();
  sh.getRange(1, 1, 1, head.length).setValues([head]);
  if (rows.length) sh.getRange(2, 1, rows.length, head.length).setValues(rows);
  styleMarks_(ctx.cols.length);
  return ctx;
}
function marksCtx_() { const live = questions_(); return { live, cols: markColumns_(live) }; }
function marksHeader_(cols) { return MARKS_HEAD.concat(cols.map(c => c.label)).concat(MARKS_TAIL); }
/** The skill columns in Marks: today's questions, then any that only a marked test date still uses. */
function markColumns_(live) {
  live = live || questions_();
  const cols = live.map(q => ({ id: q.id, label: q.label })), have = {};
  cols.forEach(c => have[c.id] = 1);
  frozenDays_().forEach(d => questionsFor_(d, live).forEach(q => { if (!have[q.id]) { have[q.id] = 1; cols.push({ id: q.id, label: q.label + ' (old)' }); } }));
  return cols;
}

// ─────────────────────────────────────────────────────────────── the questions (the Questions tab)

/** The questions as typed in the Questions tab, in test order. A blank Stage carries on the stage of the row above. */
function questions_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.QUESTIONS.name);
  if (!sh || sh.getLastRow() < 2) return [];
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, SHEETS.QUESTIONS.header.length).getValues();
  const out = [], heading = {}, seen = {};
  let stage = '';
  rows.forEach(r => {
    if (clean_(r[Q_.STAGE], 40)) stage = clean_(r[Q_.STAGE], 40);
    if (stage && clean_(r[Q_.HEADING], 80) && !heading[stage]) heading[stage] = clean_(r[Q_.HEADING], 80);
    const t = clean_(r[Q_.TEXT], 300), id = clean_(r[Q_.ID], 40);
    if (!t || !id || seen[id]) return;
    seen[id] = 1;
    out.push({ id, stage: stage || 'Test', t, ask: clean_(r[Q_.ASK], 300), core: yes_(r[Q_.CORE]) ? 1 : 0,
      studios: studios_(r[Q_.STUDIO]), label: clean_(r[Q_.LABEL], 40) || shorten_(t) });
  });
  out.forEach(q => q.heading = heading[q.stage] || q.stage);
  return out;
}
function yes_(v) { return v === true || /^(y|yes|x|✓|core|true|1)$/i.test(String(v == null ? '' : v).trim()); }
/** 'K3' → ['K3']; 'Both', blank or every studio → null (asked everywhere). */
function studios_(v) { const s = STUDIO_LIST.filter(k => new RegExp('\\b' + k + '\\b', 'i').test(String(v || ''))); return s.length && s.length < STUDIO_LIST.length ? s : null; }
function shorten_(t) { return t.length > 32 ? t.slice(0, 31).replace(/\s+\S*$/, '') + '…' : t; }
function forStudio_(qs, studio) { return qs.filter(q => !q.studios || !studio || q.studios.indexOf(studio) >= 0); }

/** Every typed question gets an ID in the hidden column; a row copied together with its ID gets a fresh one. */
function ensureQuestionIds_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.QUESTIONS.name);
  if (!sh || sh.getLastRow() < 2) return;
  const n = sh.getLastRow() - 1, texts = sh.getRange(2, Q_.TEXT + 1, n, 1).getValues(), idr = sh.getRange(2, Q_.ID + 1, n, 1), ids = idr.getValues();
  const seen = {}; let changed = false;
  ids.forEach((r, i) => {
    if (!String(texts[i][0] || '').trim()) return;
    let id = String(r[0] || '').trim();
    if (!id || seen[id]) { id = 'q' + newId_().replace(/-/g, '').slice(0, 8); r[0] = id; changed = true; }
    seen[id] = 1;
  });
  if (changed) idr.setValues(ids);
}

function seedQuestions_() {
  const ss = SpreadsheetApp.getActive();
  if (ss.getSheetByName(SHEETS.QUESTIONS.name)) return;
  const sh = sheet_(SHEETS.QUESTIONS);
  sh.getRange(2, 1, DEFAULT_QUESTIONS.length, DEFAULT_QUESTIONS[0].length).setValues(DEFAULT_QUESTIONS);
  try { const at = ss.getSheets().map(s => s.getName()).indexOf(SHEETS.STUDENTS.name); if (at >= 0) { ss.setActiveSheet(sh); ss.moveActiveSheet(at + 2); } } catch (e) {}
}

/* A test date keeps the questions it was first marked with, so changing the test never moves a grade already given.
   Rewording a question still shows on that date; adding, removing, Core and Studio changes do not.
   Stored as script property q_<date>; Ops Test ▸ Use the current questions for a test date removes it. */
function frozen_(day) { try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('q_' + day) || 'null'); } catch (e) { return null; } }
function frozenDays_() { return Object.keys(PropertiesService.getScriptProperties().getProperties()).filter(k => /^q_\d{4}-\d{2}-\d{2}$/.test(k)).map(k => k.slice(2)).sort(); }
function freezeDay_(day, qs) {
  if (!day || frozen_(day)) return;
  qs = qs || questions_();
  if (!qs.length) return;
  let s = JSON.stringify(qs);   // a script property holds 9 KB: drop the long text first, it comes from the tab anyway
  if (s.length > 8500) s = JSON.stringify(qs.map(q => ({ id: q.id, stage: q.stage, t: q.t.slice(0, 60), core: q.core, studios: q.studios, label: q.label })));
  if (s.length > 8500) s = JSON.stringify(qs.map(q => ({ id: q.id, core: q.core, studios: q.studios })));
  PropertiesService.getScriptProperties().setProperty('q_' + day, s);
}
/** One date's questions: its frozen set with today's wording, or the current questions if nothing is marked yet. */
function questionsFor_(day, live) {
  live = live || questions_();
  const f = frozen_(day);
  if (!f) return live;
  const byId = {}; live.forEach(q => byId[q.id] = q);
  return f.map(q => {
    const l = byId[q.id], o = Object.assign({ stage: 'Test', t: '(removed question)', ask: '', label: q.id }, q);
    if (l) Object.assign(o, { stage: l.stage, heading: l.heading, t: l.t, ask: l.ask, label: l.label });
    o.heading = o.heading || o.stage;
    return o;
  });
}
function freezeMarkedDays_() {
  const dayOf = {}; students_().forEach(s => dayOf[s.id] = s.day);
  const days = {}; marksRows_().forEach(m => { if (dayOf[m.sid] && Object.keys(m.data.marks || {}).length) days[dayOf[m.sid]] = 1; });
  const qs = questions_(); Object.keys(days).forEach(d => freezeDay_(d, qs));
}
function useCurrentQuestions_(day) {
  PropertiesService.getScriptProperties().deleteProperty('q_' + day);
  if (marksRows_().length) freezeMarkedDays_();   // still marked, so it is fixed again, now to the current questions
  syncMarks_(); buildGrades_();
  return { ok: true };
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
function gradeOf_(byTeacher, qs) {
  const teachers = Object.keys(byTeacher).filter(u => Object.keys(byTeacher[u].marks || {}).some(k => byTeacher[u].marks[k] in POINTS));
  const per = qs.map(q => {
    const v = teachers.map(u => byTeacher[u].marks[q.id]).filter(x => x in POINTS).map(x => POINTS[x]);
    return { q, avg: v.length ? v.reduce((a, b) => a + b, 0) / v.length : null };
  });
  const marked = per.filter(x => x.avg !== null);
  const score = marked.length ? marked.reduce((a, x) => a + x.avg, 0) / marked.length * 100 : null;
  const errors = [];
  Object.keys(byTeacher).forEach(u => Object.keys(byTeacher[u].safety || {}).forEach(k => { const e = byTeacher[u].safety[k]; if (e && errors.indexOf(e.text) < 0) errors.push(e.text); }));
  const final = score === null ? null : Math.max(0, score - errors.length * ERR_POINTS);
  const coreFails = marked.filter(x => x.q.core && x.avg < 0.5).map(x => x.q.label);
  const complete = qs.length > 0 && marked.length === qs.length;
  const band = final === null ? null : coreFails.length ? BANDS[BANDS.length - 1] : BANDS.find(b => final >= b[1] - 1e-9);
  const by = {}; marked.forEach(x => (by[x.q.stage] = by[x.q.stage] || []).push(x.avg));
  const stages = {}; Object.keys(by).forEach(k => stages[k] = Math.round(by[k].reduce((a, b) => a + b, 0) / by[k].length * 100));
  return { teachers, marked: marked.length, score, errors, final, coreFails, complete, band, stages };
}

/** Every student with their combined grade, each on their own date's questions. Used by the Grades tab and the page. */
function gradeList_() {
  const students = students_().filter(s => !s.removed).sort(byTestTime_);
  const marks = {}; marksRows_().forEach(m => { (marks[m.sid] = marks[m.sid] || {})[m.teacher] = m.data; });
  const live = questions_(), byDay = {};
  const list = students.map(s => {
    const qs = forStudio_(byDay[s.day] = byDay[s.day] || questionsFor_(s.day, live), s.studio);
    return { s, qs, g: gradeOf_(marks[s.id] || {}, qs) };
  });
  return { live, list };
}

/** The page's Grades screen: the combined result per student, never one teacher's marks. */
function grades_() {
  return { ok: true, students: gradeList_().list.map(({ s, qs, g }) => {
    const stages = [];
    qs.forEach(q => { if (!stages.some(x => x.name === q.stage)) stages.push({ name: q.stage, pct: q.stage in g.stages ? g.stages[q.stage] : null }); });
    return { id: s.id, day: s.day, time: s.time, name: s.name, studio: s.studio, marked: g.marked, total: qs.length,
      score: g.score === null ? null : Math.round(g.score), final: g.final === null ? null : Math.round(g.final),
      errors: g.errors, coreFails: g.coreFails, complete: g.complete, grade: g.band ? g.band[0] : null, teachers: g.teachers, stages };
  }) };
}

/** Rewrites the Grades tab: a summary per grade at the top, then one row per student with bars and colours. */
function buildGrades_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEETS.GRADES.name);
  if (!sh) sh = ss.insertSheet(SHEETS.GRADES.name, 1);
  const cache = CacheService.getScriptCache();
  cache.remove('grades_dirty'); cache.put('grades_recent', '1', 20);

  // each student is graded on their own date's questions; the stage columns follow today's stages, then any older ones
  const { live, list } = gradeList_(), students = list.map(x => x.s);
  const stageNames = [];
  live.concat(...list.map(x => x.qs)).forEach(q => { if (stageNames.indexOf(q.stage) < 0) stageNames.push(q.stage); });
  const nS = stageNames.length, top = 8;
  const head = ['Test date', 'Student', 'Studio', 'Teachers', 'Marked', 'Progress'].concat(stageNames)
    .concat(['Score', 'Errors', 'Final', '', 'Grade', 'Core skill under 50%', 'Critical errors']);
  const W = head.length, C = { STAGE: 7, SCORE: 7 + nS, ERR: 8 + nS, FINAL: 9 + nS, BAR: 10 + nS, GRADE: 11 + nS };

  const counts = { 'Grade A': 0, Pass: 0, Resit: 0, Unfinished: 0 };
  const values = [], bars = [], groupStarts = [];
  let prevDay = '';
  list.forEach(({ s, qs, g }, i) => {
    const N = qs.length;
    const label = !g.marked ? 'Not started' : g.complete ? g.band[0] : 'Unfinished';
    if (g.complete) counts[g.band[0]]++; else counts.Unfinished++;
    if (s.day !== prevDay) { groupStarts.push(top + i); prevDay = s.day; }
    values.push([toDate_(s.day), s.name, s.studio, g.teachers.join(', '), g.marked + '/' + N, '']
      .concat(stageNames.map(n => n in g.stages ? g.stages[n] : ''))
      .concat([g.score === null ? '' : Math.round(g.score), g.errors.length ? -g.errors.length * ERR_POINTS : '', g.final === null ? '' : Math.round(g.final), '',
               label, g.coreFails.join(', '), g.errors.join('; ')]));
    const col = g.complete ? g.band[2] : '#9aa3b5';
    bars.push(['=SPARKLINE(' + g.marked + ',{"charttype","bar";"max",' + Math.max(1, N) + ';"color1","' + BLUE + '"})',
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
    if (nS) sh.getRange(top, C.STAGE, values.length, nS).setHorizontalAlignment('center').setNumberFormat('0');
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
    sh.setConditionalFormatRules((nS ? [heat(sh.getRange(top, C.STAGE, values.length, nS))] : []).concat([
      heat(sh.getRange(top, C.SCORE, values.length, 1)), heat(sh.getRange(top, C.FINAL, values.length, 1)),
      text('Grade A', BANDS[0][3], BANDS[0][2]), text('Pass', BANDS[1][3], INK), text('Resit', BANDS[2][3], BANDS[2][2]),
      text('Unfinished', PAPER, MUTE), text('Not started', '#ffffff', '#9aa3b5'),
      SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setFontColor(BANDS[2][2]).setBold(true).setRanges([sh.getRange(top, C.ERR, values.length, 1)]).build(),
    ]));
  } else {
    sh.getRange(top, 1).setValue('No students yet. Add them in the Students tab or on the page.').setFontColor(MUTE);
  }
  [110, 170, 90, 150, 64, 110].concat(stageNames.map(() => 66)).concat([58, 58, 58, 110, 104, 200, 300]).forEach((w, i) => sh.setColumnWidth(i + 1, w));
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
  if (!o || o.head[1] === 'Test time') return;
  let rows;
  if (o.head[0] === 'Test date') {
    // v2 layout: Test date | Student | Studio | Added by | Added at | Removed | ID
    rows = o.rows.filter(r => r[1] || r[6]).map(r => [r[0], '', r[1], r[2], r[5] || '', r[6]]);
  } else {
    // v1 layout: ID | Test date | Student | Studio | Added by | Added at | Deleted
    rows = o.rows.filter(r => r[0]).map(r => [toDate_(day_(r[1])), '', r[2], r[3], r[6] || '', r[0]]);
  }
  try { o.sh.showColumns(1, o.sh.getMaxColumns()); o.sh.getRange(1, 1, o.sh.getMaxRows(), o.sh.getMaxColumns()).clearDataValidations(); } catch (e) {}
  o.sh.clear(); const W = SHEETS.STUDENTS.header.length;
  o.sh.getRange(1, 1, 1, W).setValues([SHEETS.STUDENTS.header]);
  if (rows.length) o.sh.getRange(2, 1, rows.length, W).setValues(rows);
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
  const ctx = marksCtx_();
  const rows = o.rows.filter(r => r[0] && studs[String(r[0])]).map(r => {
    let data = {}; try { data = JSON.parse(r[2] || '{}'); } catch (e) {}
    return marksRow_(studs[String(r[0])], String(r[1]), data, ctx);
  });
  o.sh.clear(); const h = marksHeader_(ctx.cols);
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
