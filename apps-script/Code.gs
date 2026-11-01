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
 *   Students    one row per student: test date, test time, name, CAP, studio, teacher; type or paste here (a date picker on the date column)
 *   Questions   the shared test: one row per question, in order (students without a teacher, and the start of each teacher's own tab)
 *   Questions · <teacher>   each teacher's own test, a copy of Questions they can change; their students are asked these
 *   Marks       one row per student per teacher, every skill readable; the last column is the page's own copy
 *   Grades      the overview: every student's combined grade, stages, progress and errors, rebuilt by the script
 */

const PAGE_URL = 'https://gunnitynes.github.io/cap26/';
const SESSION_DAYS = 30;        // a phone stays signed in this long, then signs in again with the team password
const OLD_SESSION_DAYS = 400;   // what sign-ins lasted before; used to work out when an older sign-in started
const HASH_ROUNDS = 400;
const MAX_TRIES = 8;            // wrong passwords per name before a 10-minute pause
const MIN_PASSWORD = 6;
const SET = '●●●●●● set';       // what Start here shows once the team password is stored
const TEAM_LABEL = 'Team password';
const PIN_LABEL = 'Admin PIN';
const RESET_ONE_LABEL = 'Reset a student';
const RESET_ALL_LABEL = 'Reset everything';
const STAGES_LABEL = 'Test stages';
const MAX_STAGES = 8;
const MIN_PIN = 4;
// Marks, best first. Tolerant on purpose: "Got there" on everything still passes.
const POINTS = { p: 1, y: 0.8, h: 0.65, g: 0.5, n: 0 };
const MARK_LABEL = { p: 'Perfect', y: 'Pass', h: 'Hint', g: 'Got there', n: 'Failed' };

// The questions live in the Questions tab, where anyone with the Sheet can change them. This is only the first version,
// written there by setup when the tab does not exist yet. The ids keep marks attached to their question.
//   stage, stage heading, question, ask out loud, studio, column name, id
const DEFAULT_QUESTIONS = [
  ['Power', 'Power up & safe listening', 'Powers up in the right order', '', 'Both', 'Power-up order', 'p1'],
  ['', '', 'Control-room level down/mute before any console connections', '', 'Both', 'CR level down/mute for connections', 'p2'],
  ['', '', 'Keeps listening levels safe and works methodically', '', 'Both', 'Safe levels, methodical', 'p3'],
  ['Mic', 'Connecting a Mic', 'Chooses the correct mic and places it sensibly', '', 'Both', 'Correct mic, placed', 'm1'],
  ['', '', 'Connects the live-room wall box to the right desk channel', '', 'Both', 'Wall box → right channel', 'm2'],
  ['', '', 'Gets a healthy signal on the desk', '', 'Both', 'Healthy desk signal', 'm3'],
  ['', '', 'Understands phantom power', 'If I had asked you to use the condenser instead, what would you have needed to do?', 'Both', 'Phantom power', 'm4'],
  ['Route', 'Route to the speakers', 'Routes the channel to the console and hears it in the speakers', '', 'Both', 'Console → speakers', 'r1'],
  ['Record', 'Record & play back in Pro Tools', 'Creates a Pro Tools track with the correct input and output', '', 'Both', 'PT track in/out', 't1'],
  ['', '', 'Healthy record level in Pro Tools', '', 'Both', 'PT record level', 't2'],
  ['', '', 'Records a 1-2-3 test and returns it on a different desk channel', '', 'Both', '1-2-3 test, other channel', 't3'],
  ['', '', 'Plays it back through the path you asked for', '', 'Both', 'Playback path', 't4'],
  ['Cue', 'Headphones & cue mix', 'Builds a usable cue mix on the aux sends', '', 'Both', 'Usable cue mix', 'c1'],
  ['', '', 'Understands pre-fader vs. post-fader', 'Should this send be pre or post fader, and why?', 'Both', 'Pre vs post fader', 'c2'],
  ['', '', 'Routes the cue to the headphone amp and checks it', '', 'Both', 'Cue → headphone amp', 'c3'],
  ['', '', 'Explains sending the cue to a selected Live room', 'Without going there: how would the same cue reach this Live room?', 'Both', 'Cue to a Live room', 'c4'],
  ['Shut down', 'Shut down & pack down', 'Knows the shutdown order', 'If you had to shut down the studio now, what would you switch off first?', 'Both', 'Shutdown order', 's1'],
  ['', '', 'Leaves the studio as it was found', '', 'Both', 'Left as found', 's2'],
];
// the standard column order of a questions tab; tabs are READ by their headings, so a moved or deleted column cannot scramble them
const Q_ = { STAGE: 0, HEADING: 1, TEXT: 2, ASK: 3, STUDIO: 4, LABEL: 5, ID: 6 };
/* Marks, one row per student per teacher:
     MARKS_HEAD | one % column per test stage | MARKS_MID | every skill (a folded column group) | MARKS_END
   The stage columns are the same for every teacher, so the rows line up whoever asked which questions. */
const MARKS_HEAD = ['Test date', 'Student', 'Teacher', 'Result', 'Score %', 'Marked'];
const MARKS_MID = ['Critical errors', 'Notes', 'Test ended', 'Updated'];
const MARKS_END = ['Student ID', 'Page data (do not edit)'];
const NOT_ASKED = '–';   // in Marks: this question is not part of that student's test (another teacher's own question)
// The score: each stage's marks averaged (a %), then the average of the stages, so every stage counts the same and a
// teacher with 3 questions in a stage weighs the same as one with 5.
// Grade A from 85 (needs Perfects), Pass from 40. Any critical error is a Resit whatever the score.
const BANDS = [['Grade A', 85, '#0f7a57', '#d8f1e6'], ['Pass', 40, '#5a6276', '#eef0f4'], ['Resit', 0, '#bf2b2b', '#fbe0e0']];

const STUDIO_LIST = ['K3', 'K4'];   // the studios a test can be in
const CAP_LIST = ['CAP 26A', 'CAP 26B', 'CAP 26C', 'TRANSFER'];   // the CAP groups, a dropdown in the Students tab; TRANSFER = a student from outside the class
const DATE_FMT = 'dd.mm.yyyy';
const BLUE = '#0b43e0', INK = '#10131c', PAPER = '#f4f6fa', MUTE = '#5a6276';
const SHEETS = {
  START:    { name: 'Start here', header: [] },
  TEACHERS: { name: 'Teachers', header: ['Name', 'Role', 'Last sign-in'], widths: [220, 120, 170] },
  STUDENTS: { name: 'Students', header: ['Test date', 'Test time', 'Student', 'CAP', 'Studio', 'Teacher', 'Removed', 'ID'],
              widths: [110, 110, 220, 110, 90, 140, 220, 110] },
  QUESTIONS:{ name: 'Questions', header: ['Stage', 'Stage heading', 'Question', 'Ask out loud (optional)', 'Studio', 'Column name', 'ID'],
              widths: [110, 220, 380, 320, 80, 220, 90] },
  MARKS:    { name: 'Marks', header: [] },   // its skill columns follow the questions: see syncMarks_
  GRADES:   { name: 'Grades', header: [] },
};
const T = { NAME: 0, ROLE: 1, LAST: 2 };

// ─────────────────────────────────────────────────────────────── sheet menu and layout

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Ops Test')
    .addItem('Set up / tidy the sheet', 'setup')
    .addItem('Set the team password…', 'menuTeamPassword')
    .addItem('Set the admin PIN…', 'menuAdminPin')
    .addItem('Add a teacher…', 'menuAddTeacher')
    .addItem('Remove a teacher…', 'menuRemoveTeacher')
    .addSeparator()
    .addItem('Refresh the Grades tab', 'buildGrades_')
    .addItem('Use the current questions for a test date…', 'menuUseCurrentQuestions')
    .addSeparator()
    .addItem('Reset a student\'s test…', 'menuResetStudent')
    .addItem('Reset ALL test results…', 'menuResetAll')
    .addToUi();
}

/** Builds every tab, moves older layouts over without losing anything, and styles it all. Safe to run any time. */
function setup() {
  const ss = SpreadsheetApp.getActive();
  migrateTeachers_(); migrateStudents_(); seedQuestions_(); questionSheets_().forEach(normalizeQuestionTab_); ensureTeacherQuestionTabs_(); ensureQuestionIds_(); migrateMarks_(); migrateDates_();
  splitDayFreezes_();
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
  storeTeamPassword_(); storeAdminPin_();     // keep a password or PIN typed there before the tab is rebuilt
  sh.clear();
  const rows = [
    ['Catalyst Operations Test', ''],
    ['Marking sheet for the music studio operations test · CAP 2026', ''],
    ['', ''],
    ['The page', PAGE_URL],
    ['', ''],
    [TEAM_LABEL, ''],
    ['', 'Everyone signs in with this one password. Type it in the cell above (6 characters or more); it is hidden as soon as you press Enter. Type a new one over the dots to change it.'],
    [PIN_LABEL, ''],
    ['', 'Only for admins (Role "admin" in the Teachers tab): asked after the team password, so the team password alone never opens an admin account. Type it in the cell above (4 characters or more); it is hidden at once. Setting it signs admins out once; teachers are not affected.'],
    [RESET_ONE_LABEL, ''],
    ['', 'Type a student\'s name in the cell above and press Enter: every teacher\'s marks for that student are removed, the student stays on the list. If the name is there twice, add the date, like "Kim 02.11.2026". (On the page: open the student, scroll down, tap Reset.)'],
    [RESET_ALL_LABEL, ''],
    ['', 'Type RESET in the cell above and press Enter to remove every mark for every student, for example after beta testing. Students, teachers and questions stay.'],
    ...stageSlots_().map((st, i) => [i ? '' : STAGES_LABEL, st]),
    ['', 'The stages of the test, in order, the same for every teacher: the Stage column of every questions tab is a dropdown of these, and Grades has one column per stage. Rename a stage here and it is renamed in every teacher\'s tab. Add one in an empty cell; clear one to take it off the list.'],
    ['Add a teacher', 'Type their name in the Teachers tab. It then appears on the sign-in page: they pick it and type the team password.'],
    ['Add students', 'Open the Students tab. Type the test date, the test time (like 10:30), the student\'s name and the studio, one student per row.'],
    ['', 'Pasting a whole list? Paste the names into the Student column and type the date once at the top: rows below take the date and studio from the row above. Each student keeps their own time.'],
    ['', 'They appear on the page within a few seconds. Teachers can also add students on the page itself.'],
    ['CAP', 'Pick the student\'s CAP in the CAP column: CAP 26A, CAP 26B, CAP 26C, or TRANSFER for a student from outside the class. The page shows it next to the student.'],
    ['Assign a teacher', 'Pick the teacher in the student\'s Teacher column. On the page each teacher sees only their own students; an admin (Role column in Teachers) sees everyone.'],
    ['Remove a student', 'Type anything (for example "absent") in their Removed column. Their row stays, so nothing is lost.'],
    ['Remove a teacher', 'Delete their row in Teachers. Their marks stay in Marks.'],
    ['Change the questions', 'Each teacher has their own tab, "Questions · their name". Edit yours: one row per question, in test order. Drag rows to reorder, type a new question over an old one, add or delete rows, or set a studio. Marks stay with their question when it moves. Every row shows its stage (repeats are faint), picked from a dropdown of the Test stages above: a moved row keeps its stage, a new row takes the stage above. Your students are asked your questions; the page picks it up within a few seconds.'],
    ['', 'The Questions tab is the shared starting point: a new teacher\'s tab starts as a copy of it, and students without a teacher are asked it.'],
    ['', 'A test date keeps the questions it was first marked with, so grades already given never move (rewording still shows). To give a marked date the new questions anyway: Ops Test ▸ Use the current questions for a test date.'],
    ['', ''],
    ['What the tabs hold', ''],
    ['Teachers', 'Who can sign in. Their names are the list on the sign-in page.'],
    ['Students', 'Everyone booked for a test: date, time, name, studio. Type or paste here, or add them on the page.'],
    ['Questions', 'The shared test: every question, its stage and which studio it is for. Each teacher\'s own tab starts as a copy of it.'],
    ['Marks', 'One row per student per teacher: result, score and a % per test stage, the same columns for every teacher. Every single skill is folded away on the right (press + above the columns). The last column is the page\'s own copy; leave it alone.'],
    ['Grades', 'The overview: every student\'s grade, progress, a colour per test stage and their critical errors. Kept up to date by itself.'],
    ['', ''],
    ['How the grade works', 'Each mark is worth: Perfect 100 · Pass 80 · Hint 65 · Got there 50 · Failed 0. Each test stage gets the average of its questions (each skill averaged over every teacher who marked it), and the score is the average of the stages, so every stage counts the same however many questions a teacher put in it. Grade A from 85%, Pass from 40%. Any critical error means Resit.'],
  ];
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.setColumnWidth(1, 190); sh.setColumnWidth(2, 640);
  sh.getRange(1, 1, rows.length, 2).setFontFamily('Arial').setFontSize(11).setVerticalAlignment('top').setWrap(true).setFontColor(INK);
  sh.getRange('A1').setFontSize(22).setFontWeight('bold').setFontColor(BLUE);
  sh.getRange('A2').setFontColor(MUTE);
  sh.getRange('A1:B1').merge(); sh.getRange('A2:B2').merge();
  sh.getRange('B4').setFormula('=HYPERLINK("' + PAGE_URL + '","' + PAGE_URL + '")').setFontWeight('bold');
  const tr = rows.findIndex(r => r[0] === TEAM_LABEL) + 1;
  sh.getRange(tr, 2).setNumberFormat('@').setNote(teamSetNote_()).setValue(teamHash_() ? SET : '').setBackground('#fff8d6').setFontWeight('bold').setFontFamily('Courier New')
    .setBorder(true, true, true, true, null, null, BLUE, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(tr + 1, 2).setFontColor(MUTE);
  const pr = rows.findIndex(r => r[0] === PIN_LABEL) + 1;
  sh.getRange(pr, 2).setNumberFormat('@').setValue(adminPinHash_() ? SET : '').setBackground('#fff8d6').setFontWeight('bold').setFontFamily('Courier New')
    .setBorder(true, true, true, true, null, null, BLUE, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(pr + 1, 2).setFontColor(MUTE);
  [RESET_ONE_LABEL, RESET_ALL_LABEL].forEach(label => {
    const at = rows.findIndex(r => r[0] === label) + 1;
    sh.getRange(at, 2).setNumberFormat('@').setBackground('#fdecec').setFontWeight('bold')
      .setBorder(true, true, true, true, null, null, '#c62828', SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(at, 1).setFontColor('#c62828');
    sh.getRange(at + 1, 2).setFontColor(MUTE);
  });
  const sa = rows.findIndex(r => r[0] === STAGES_LABEL) + 1;
  sh.getRange(sa, 2, MAX_STAGES, 1).setNumberFormat('@').setBackground('#eaf0ff').setFontWeight('bold').setFontColor(BLUE)
    .setBorder(true, true, true, true, null, true, BLUE, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(sa + MAX_STAGES, 2).setFontColor(MUTE);
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
  st.getRange(2, S_.CAP + 1, sr, 1).setHorizontalAlignment('center').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(CAP_LIST, true).setAllowInvalid(false).setHelpText(CAP_LIST.join(', ')).build());
  st.getRange(1, S_.CAP + 1).setNote('The student\'s CAP: ' + CAP_LIST.join(', ') + ' (TRANSFER = from outside the class). Shown next to their name on the page.');
  st.getRange(2, S_.STUDIO + 1, sr, 1).setHorizontalAlignment('center').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(STUDIO_LIST, true).setAllowInvalid(false).setHelpText('K3 or K4').build());
  st.getRange(2, S_.TEACHER + 1, sr, 1).setHorizontalAlignment('center').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInRange(sheet_(SHEETS.TEACHERS).getRange('A2:A'), true).setAllowInvalid(false)
      .setHelpText('The teacher who runs this student\'s test: a name from the Teachers tab.').build());
  st.getRange(2, S_.REMOVED + 1, sr, 1).setFontColor(MUTE);
  st.hideColumns(S_.ID + 1);
  questionSheets_().forEach(styleQuestions_);
  styleMarks_(marksCtx_());
  [t, st].forEach(s => s.setTabColor(INK));
}

/** Where each column of a questions tab is, found by its heading (-1 when a tab has no such column). */
function qcols_(sh) {
  const w = Math.max(1, sh.getLastColumn()), head = sh.getRange(1, 1, 1, w).getValues()[0].map(h => String(h || '').trim().toLowerCase());
  const at = (...names) => head.findIndex(h => names.some(n => h === n || h.indexOf(n) === 0));
  return { stage: head.indexOf('stage'), heading: at('stage heading'), text: at('question'),
           ask: at('ask'), studio: at('studio'), label: at('column name'), id: head.indexOf('id'), width: w };
}
/** A questions tab's rows in the standard column order, whatever order its columns are in now. */
function rowsStd_(sh) {
  if (!sh || sh.getLastRow() < 2) return [];
  const c = qcols_(sh), raw = sh.getRange(2, 1, sh.getLastRow() - 1, c.width).getValues();
  const get = (r, i) => i < 0 ? '' : (typeof r[i] === 'boolean' ? '' : r[i]);
  return raw.map(r => [get(r, c.stage), get(r, c.heading), get(r, c.text), get(r, c.ask), get(r, c.studio), get(r, c.label), get(r, c.id)]);
}
/** Puts a questions tab back into the standard columns (Stage … ID) when someone has moved, added or deleted a column.
 *  Every question, ask text, studio, column name and ID is kept; a Core column (no longer used) is dropped. */
function normalizeQuestionTab_(sh) {
  if (!sh) return;
  const want = SHEETS.QUESTIONS.header, w = Math.max(1, sh.getLastColumn());
  const head = sh.getRange(1, 1, 1, w).getValues()[0].map(h => String(h || '').trim());
  if (head.slice(0, want.length).join('|') === want.join('|') && head.slice(want.length).every(h => !h)) return;
  const rows = rowsStd_(sh);
  try { sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations(); sh.showColumns(1, sh.getMaxColumns()); } catch (e) {}
  sh.clear();
  sh.getRange(1, 1, 1, want.length).setValues([want]);
  if (rows.length) sh.getRange(2, 1, rows.length, want.length).setValues(rows);
}

/** One questions tab (the shared one or a teacher's): bold questions, Studio dropdown, notes, hidden IDs. */
/** Questions tab: every stage gets its own soft colour and a line above its first question, so it is clear which
 *  questions belong together (a blank Stage cell still means "the stage above"). Runs on setup and after every edit. */
const STAGE_TINTS = ['#eaf0ff', '#fff3df', '#e6f5ec', '#fdeaee', '#efe9fb', '#e4f4f6', '#f6f1e3'];
function colourStages_(q) {
  const last = q.getLastRow(), W = SHEETS.QUESTIONS.header.length;
  if (last < 2) return;
  const vals = rowsStd_(q), tintOf = {}, bg = [], starts = [];
  let prev = null;
  vals.forEach((r, i) => {
    const st = String(r[Q_.STAGE] || '').trim();
    if (st && !(st in tintOf)) tintOf[st] = STAGE_TINTS[Object.keys(tintOf).length % STAGE_TINTS.length];   // one colour per stage name
    if (st && st !== prev) starts.push(i + 2);
    if (st) prev = st;
    bg.push(Array(W).fill(st ? tintOf[st] : '#ffffff'));
  });
  try { q.getBandings().forEach(b => b.remove()); } catch (e) {}
  q.getRange(2, 1, vals.length, W).setBackgrounds(bg).setBorder(false, false, false, false, false, false);
  starts.forEach(r => q.getRange(r, 1, 1, W).setBorder(true, null, null, null, null, null, INK, SpreadsheetApp.BorderStyle.SOLID_MEDIUM));
  // the stage name is on every row (so a row keeps its stage when moved), but only the first of a block is shown strongly
  try {
    const col = q.getRange(3, Q_.STAGE + 1, Math.max(1, q.getMaxRows() - 2), 1);
    q.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(A3<>"",A3=A2)')
      .setFontColor('#aab3c5').setBold(false).setRanges([col]).build()]);
  } catch (e) {}
}
/* What happens when a teacher edits a questions tab:
   - a question without a stage takes the stage of the row above, written into its own cell, so it keeps it when moved;
   - IDs and colours are brought up to date. */
function questionsEdited_(sh, e) {
  try {
    fillStages_(sh);
  } catch (err) {}
  ensureQuestionIdsIn_(sh);
  try { colourStages_(sh); } catch (err) {}
}
/* The test's stages: one list for every teacher, kept in the script's storage and shown on Start here.
   Every questions tab's Stage column is a dropdown of it, and Grades has one column per stage in this order. */
let STAGES_CACHE_ = null;
function stages_() {
  if (STAGES_CACHE_) return STAGES_CACHE_;
  let list = null; try { list = JSON.parse(PropertiesService.getScriptProperties().getProperty('stages') || 'null'); } catch (e) {}
  if (!list || !list.length) {   // first time: the stages the tabs use now, in order
    list = [];
    questionSheets_().forEach(sh => rowsStd_(sh).forEach(r => { const st = clean_(r[Q_.STAGE], 40); if (st && !list.some(x => norm_(x) === norm_(st))) list.push(st); }));
    if (list.length) setStages_(list.slice(0, MAX_STAGES));
  }
  return (STAGES_CACHE_ = list);
}
function setStages_(list) { STAGES_CACHE_ = list; PropertiesService.getScriptProperties().setProperty('stages', JSON.stringify(list)); }
function stageSlots_() { const l = stages_(); return Array.from({ length: MAX_STAGES }, (_, i) => l[i] || ''); }
/** A stage name as the list spells it ("shut down " → "Shut down"); a name not on the list stays as typed. */
function stageName_(st) { st = String(st || '').trim(); const hit = stages_().find(x => norm_(x) === norm_(st)); return hit || st; }
/** Start here: a stage typed over another renames it in every teacher's tab; the list is saved and the dropdowns follow. */
function stagesTyped_(e) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.START.name), rg = e && e.range;
  if (!sh || !rg) return;
  const at = sh.getRange(1, 1, Math.max(1, sh.getLastRow()), 1).getValues().findIndex(r => String(r[0]) === STAGES_LABEL) + 1;
  if (!at || rg.getColumn() > 2 || rg.getLastColumn() < 2 || rg.getLastRow() < at || rg.getRow() > at + MAX_STAGES - 1) return;
  const from = String(e.oldValue || '').trim(), to = String(rg.getValue() || '').trim();
  if (rg.getNumRows() === 1 && from && to && from !== to) renameStageEverywhere_(from, to);   // a new name, or a corrected spelling
  const list = [];
  sh.getRange(at, 2, MAX_STAGES, 1).getValues().forEach(r => { const v = clean_(r[0], 40); if (v && !list.some(x => norm_(x) === norm_(v))) list.push(v); });
  setStages_(list);
  questionSheets_().forEach(applyStageList_);
  gradesDirty_();
}
function renameStageEverywhere_(from, to) {
  questionSheets_().forEach(sh => {
    const c = qcols_(sh), last = sh.getLastRow();
    if (c.stage < 0 || last < 2) return;
    const rg = sh.getRange(2, c.stage + 1, last - 1, 1), v = rg.getValues();
    let changed = false;
    v.forEach(r => { if (norm_(r[0]) === norm_(from)) { r[0] = to; changed = true; } });
    if (changed) rg.setValues(v);
  });
}
/** The Stage column of a questions tab: a dropdown of the stage list. */
function applyStageList_(sh) {
  const c = qcols_(sh), list = stages_();
  if (c.stage < 0) return;
  const col = sh.getRange(2, c.stage + 1, Math.max(1, sh.getMaxRows() - 1), 1);
  if (!list.length) { col.clearDataValidations(); return; }
  col.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(false)
    .setHelpText('A stage from the list on Start here (the same for every teacher). Ask the test organiser to add or rename one.').build());
}
/** Writes the stage into every question row that has none: the stage of the row above. */
function fillStages_(sh) {
  const c = qcols_(sh), last = sh.getLastRow();
  if (c.stage < 0 || c.text < 0 || last < 2) return;
  const sr = sh.getRange(2, c.stage + 1, last - 1, 1), st = sr.getValues(), tx = sh.getRange(2, c.text + 1, last - 1, 1).getValues();
  let cur = '', changed = false;
  st.forEach((r, i) => {
    const v = String(r[0] || '').trim();
    if (v) cur = v;
    else if (cur && String(tx[i][0] || '').trim()) { r[0] = cur; changed = true; }
  });
  if (changed) sr.setValues(st);
}

function styleQuestions_(q) {
  normalizeQuestionTab_(q);
  const qr = q.getMaxRows() - 1, own = q.getName() !== SHEETS.QUESTIONS.name;
  q.getRange(1, 1, 1, SHEETS.QUESTIONS.header.length).setValues([SHEETS.QUESTIONS.header]).setFontWeight('bold').setFontColor('#ffffff')
    .setBackground(own ? INK : BLUE).setVerticalAlignment('middle').setWrap(true);
  q.setRowHeight(1, 40); q.setFrozenRows(1);
  SHEETS.QUESTIONS.widths.forEach((w, i) => q.setColumnWidth(i + 1, w));
  q.getRange(2, 1, qr, SHEETS.QUESTIONS.header.length).setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');
  q.getRange(2, Q_.STAGE + 1, qr, 1).setFontWeight('bold').setFontColor(BLUE);
  q.getRange(2, Q_.HEADING + 1, qr, 1).setFontColor(MUTE);
  q.getRange(2, Q_.TEXT + 1, qr, 2).setWrap(true);
  q.getRange(2, Q_.TEXT + 1, qr, 1).setFontWeight('bold');
  q.getRange(2, Q_.ASK + 1, qr, 1).setFontStyle('italic').setFontColor(MUTE);
  q.getRange(2, Q_.STUDIO + 1, qr, 1).setHorizontalAlignment('center');
  q.getRange(2, Q_.STUDIO + 1, qr, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Both'].concat(STUDIO_LIST), true)
    .setAllowInvalid(true).setHelpText('Both, or a studio: a K3 question is only asked of students tested in K3.').build());
  q.getRange(2, Q_.LABEL + 1, qr, 1).setFontColor(MUTE);
  [['The stage this question belongs to: pick it from the dropdown (the stages are the same for every teacher; the test organiser keeps the list on Start here). Every row has one; repeats show faint. A new row left blank takes the stage of the row above.'],
   ['The heading over the stage on the page. Needed on one row of the stage.'],
   ['What the teacher marks. One row per question, in the order of the test. Delete the row to remove the question.'],
   ['Optional: a question the teacher reads out loud, shown under the question. Leave it empty for none.'],
   ['Both, K3 or K4. A K3 question only shows for students tested in K3.'],
   ['Optional: a short name, the column heading in Marks. Questions with the same name in different teachers\' tabs share a column. Cleared by itself when you type a different question over this row.'],
   ['Given by the sheet. Keeps marks attached to this question when you move it (drag, sort, copy or retype) or reword it a little.']]
    .forEach((n, i) => q.getRange(1, i + 1).setNote(n[0]));
  q.hideColumns(Q_.ID + 1);
  q.setTabColor(own ? '#7c8db5' : BLUE);
  try {   // the headings are what the script reads: changing them, or deleting a column, asks "are you sure?" first
    q.getProtections(SpreadsheetApp.ProtectionType.RANGE).filter(p => p.getDescription() === 'ops').forEach(p => p.remove());
    q.getRange(1, 1, 1, SHEETS.QUESTIONS.header.length).protect().setDescription('ops').setWarningOnly(true);
  } catch (e) {}
  fillStages_(q);
  applyStageList_(q);
  colourStages_(q);
}

/** Marks: result and stage columns in colour, every skill in a folded group, the page's own copy hidden. Run again whenever the columns change. */
function styleMarks_(ctx) {
  const m = sheet_(SHEETS.MARKS), nS = ctx.stages.length, nD = ctx.cols.length, rows = m.getMaxRows() - 1;
  const w = MARKS_HEAD.length + nS + MARKS_MID.length + nD + MARKS_END.length;
  const S0 = MARKS_HEAD.length + 1, M0 = S0 + nS, D0 = M0 + MARKS_MID.length;   // first stage, first middle, first skill column
  try { m.getRange(1, 1, 1, m.getMaxColumns()).shiftColumnGroupDepth(-8); } catch (e) {}   // drop the old fold
  try { m.showColumns(1, m.getMaxColumns()); } catch (e) {}
  try { m.getRange(1, 1, m.getMaxRows(), m.getMaxColumns()).setBackground(null).setFontColor(null).setFontWeight(null).setBorder(false, false, false, false, false, false); } catch (e) {}
  const head = m.getRange(1, 1, 1, w);
  head.setFontWeight('bold').setFontColor('#ffffff').setBackground(BLUE).setVerticalAlignment('middle').setWrap(true).setHorizontalAlignment('center');
  m.getRange(1, 1, 1, 3).setHorizontalAlignment('left');
  m.setRowHeight(1, 40); m.setFrozenRows(1); m.setFrozenColumns(3);
  [100, 170, 110, 104, 70, 70].concat(Array(nS).fill(76)).concat([220, 300, 120, 120]).concat(Array(nD).fill(92)).concat([110, 120])
    .forEach((x, i) => m.setColumnWidth(i + 1, x));
  m.getRange(2, 1, rows, w).setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');
  try { m.getBandings().forEach(b => b.remove()); m.getRange(1, 1, m.getMaxRows(), w).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false); } catch (e) {}
  head.setBackground(BLUE).setFontColor('#ffffff');
  if (nD) m.getRange(1, D0, 1, nD).setBackground('#5a6276');                     // the skill columns read as detail
  m.getRange(2, 1, rows, 1).setNumberFormat(DATE_FMT).setHorizontalAlignment('left');
  m.getRange(2, 2, rows, 1).setFontWeight('bold');
  m.getRange(2, 4, rows, 3 + nS).setHorizontalAlignment('center');
  m.getRange(2, 5, rows, 1 + nS).setNumberFormat('0');
  m.getRange(2, 5, rows, 1).setFontWeight('bold').setFontSize(12);
  m.getRange(2, M0, rows, 1).setFontColor(BANDS[2][2]).setWrap(true);
  m.getRange(2, M0 + 1, rows, 1).setWrap(true);
  m.getRange(2, M0 + 2, rows, 2).setNumberFormat('dd.MM.yyyy HH:mm').setFontColor(MUTE);
  // a thick line after the stages and after the summary, so the three parts read apart
  try {
    m.getRange(1, 3, m.getMaxRows(), 1).setBorder(null, null, null, true, null, null, INK, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    m.getRange(1, M0 - 1, m.getMaxRows(), 1).setBorder(null, null, null, true, null, null, INK, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  } catch (e) {}
  m.hideColumns(w - 1, 2);
  try { m.getRange(1, M0 - nS, 1, 1).setNote('Each stage: the average of its questions (Perfect 100 · Pass 80 · Hint 65 · Got there 50 · Failed 0). Score %: the average of the stages, so every stage counts the same however many questions a teacher asks in it.'); } catch (e) {}
  const result = m.getRange(2, 4, rows, 1), text = (t, bg, fg) => SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(t).setBackground(bg).setFontColor(fg).setBold(true).setRanges([result]).build();
  const heat = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpointWithValue('#f4c7c3', SpreadsheetApp.InterpolationType.NUMBER, '0')
    .setGradientMidpointWithValue('#fce8b2', SpreadsheetApp.InterpolationType.NUMBER, '60')
    .setGradientMaxpointWithValue('#b7e1cd', SpreadsheetApp.InterpolationType.NUMBER, '100').setRanges([m.getRange(2, 5, rows, 1 + nS)]).build();
  const rules = [heat, text('Grade A', BANDS[0][3], BANDS[0][2]), text('Pass', BANDS[1][3], INK), text('Resit', BANDS[2][3], BANDS[2][2]),
    text('In progress', PAPER, MUTE), text('Not started', '#ffffff', '#9aa3b5')];
  if (nD) {
    const det = m.getRange(2, D0, rows, nD), r = (t, bg, fg) => SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(t).setBackground(bg).setFontColor(fg).setRanges([det]).build();
    det.setHorizontalAlignment('center').setFontSize(9);
    rules.push(r('Perfect', '#d8f1e6', '#0f7a57'), r('Pass', '#dbe6ff', '#0b43e0'), r('Hint', '#fbeccf', '#a35f00'), r('Got there', '#fde2d3', '#c2410c'), r('Failed', '#fbe0e0', '#bf2b2b'),
      r(NOT_ASKED, '#ffffff', '#d5d9e1'));
    // every skill in one fold, closed: press + above the columns to open it
    try { m.getRange(1, D0, 1, nD).shiftColumnGroupDepth(1); m.getColumnGroup(D0, 1).collapse(); } catch (e) {}
    try { m.getRange(1, D0).setNote('Every skill, one column each, as each teacher marked it. A dash: not part of that student\'s test. Press − above the columns to fold them away again.'); } catch (e) {}
  }
  m.setConditionalFormatRules(rules);
  m.setTabColor(INK);
}

// ─────────────────────────────────────────────────────────────── typing in the Teachers tab

/** Runs by itself on every edit: students typed in, the team password typed on Start here, new teachers. */
function onEdit(e) {
  bump_();   // any edit in the Sheet: the phones' next refresh reads it again
  try {
    const sh = e && e.range && e.range.getSheet(), name = sh && sh.getName(), r = e.range.getRow();
    if (name === SHEETS.STUDENTS.name && r >= 2) { studentsTyped_(sh, r, e.range.getNumRows()); gradesDirty_(); return; }
    if (name === SHEETS.START.name) { storeTeamPassword_(); storeAdminPin_(); resetCells_(e.range); stagesTyped_(e); return; }
    if (isQuestionSheet_(name) && r >= 2) { questionsEdited_(sh, e); gradesDirty_(); return; }
    if (name === SHEETS.TEACHERS.name && r >= 2) {
      const last = Math.min(sh.getLastRow(), r + e.range.getNumRows() - 1);
      for (let i = r; i <= last; i++) if (sh.getRange(i, T.NAME + 1).getValue() && !sh.getRange(i, T.ROLE + 1).getValue()) sh.getRange(i, T.ROLE + 1).setValue('teacher');
      ensureTeacherQuestionTabs_();   // a new teacher gets their own questions tab straight away
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
/* The admin PIN: asked of admins after the team password, typed once on Start here, kept only as a salted hash.
   An admin's session remembers that the PIN was given; without it, an admin account works as a plain teacher. */
function pinCell_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.START.name);
  if (!sh) return null;
  const col = sh.getRange(1, 1, Math.max(1, sh.getLastRow()), 1).getValues();
  const i = col.findIndex(r => String(r[0]) === PIN_LABEL);
  return i < 0 ? null : sh.getRange(i + 1, 2);
}
function adminPinHash_() { const p = PropertiesService.getScriptProperties(); const h = p.getProperty('pin_hash'); return h ? { hash: h, salt: p.getProperty('pin_salt') } : null; }
function storeAdminPin_() {
  const cell = pinCell_(); if (!cell) return;
  const typed = typedTeamPassword_(cell);
  if (!typed || typed === SET) return;
  if (typed.length < MIN_PIN) { cell.setNote('Use ' + MIN_PIN + ' characters or more.'); return; }
  setAdminPin_(typed);
  cell.setNumberFormat('@').setValue(SET).setNote('');
}
/** Stores the PIN and signs admins out once, so every admin session from now on has given it. Teachers stay signed in. */
function setAdminPin_(pin) {
  const salt = Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperties({ pin_hash: hash_(pin, salt), pin_salt: salt });
  teachers_().filter(t => t.role === 'admin').forEach(t => endSessions_(t.name));
}
function menuAdminPin() {
  const ui = SpreadsheetApp.getUi();
  const p = ui.prompt('Admin PIN', 'The PIN admins type after the team password (' + MIN_PIN + ' characters or more):', ui.ButtonSet.OK_CANCEL);
  if (p.getSelectedButton() !== ui.Button.OK) return;
  const pin = p.getResponseText().trim();
  if (pin.length < MIN_PIN) { ui.alert('Use ' + MIN_PIN + ' characters or more.'); return; }
  setAdminPin_(pin); const cell = pinCell_(); if (cell) cell.setValue(SET);
  ui.alert('Admin PIN set. Admins sign in again once with it.');
}

function teamHash_() { const p = PropertiesService.getScriptProperties(); const h = p.getProperty('team_hash'); return h ? { hash: h, salt: p.getProperty('team_salt') } : null; }
function storeTeamPassword_() {
  const cell = teamCell_(); if (!cell) return;
  const typed = typedTeamPassword_(cell);
  if (!typed || typed === SET) return;
  if (typed.length < MIN_PASSWORD) { cell.setNote('Use ' + MIN_PASSWORD + ' characters or more.'); return; }
  setTeamPassword_(typed);
  cell.setNumberFormat('@').setValue(SET);
}
/** What is typed in the team password cell, exactly as it shows (the sheet may have turned 0123456 into a number). */
function typedTeamPassword_(cell) {
  let v = '';
  try { v = cell.getDisplayValue(); } catch (e) {}
  return String(v || cell.getValue() || '').trim();
}
function setTeamPassword_(pw) {
  const salt = Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperties({ team_hash: hash_(pw, salt), team_salt: salt, team_at: String(Date.now()) });
  try { const cell = teamCell_(); if (cell) cell.setNote(teamSetNote_()); } catch (e) {}
}
function teamSetNote_() {
  const at = Number(PropertiesService.getScriptProperties().getProperty('team_at'));
  return at ? 'Set ' + Utilities.formatDate(new Date(at), Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm') + '. Type a new one over the dots to change it.' : '';
}
/** Phones like to capitalise the first letter or curl a quote: those versions of the password count too. */
function teamPasswordMatches_(pw, team) {
  const straight = pw.replace(/[\u2018\u2019]/g, "'").replace(/[\u201c\u201d]/g, '"');
  const tries = [];
  [pw, straight].forEach(p => [p, p.charAt(0).toLowerCase() + p.slice(1), p.charAt(0).toUpperCase() + p.slice(1)]
    .forEach(x => { if (tries.indexOf(x) < 0) tries.push(x); }));
  return tries.some(x => hash_(x, team.salt) === team.hash);
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
    // a time typed as 9.30 or 930 becomes a number that would show as a wrong time: store it as 09:30
    // (a slot typed as text, like 11:15 - 11:30, stays exactly as typed)
    const tm = time_(r[S_.TIME], shown[i][S_.TIME]);
    if (typeof r[S_.TIME] === 'number' && tm) { r[S_.TIME] = tm; changed = true; }
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
/** A test slot from the sheet: '11:15 - 11:30' (hyphen, dash or "to") → {time: '11:15', end: '11:30'}; one time → {time, end: ''}. */
function slot_(v, shown) {
  const t = typeof v === 'string' ? v.trim() : '';
  const m = t.match(/^(.+?)\s*(?:-|\u2013|\u2014|to|bis)\s*(.+)$/i);
  if (m && time_(m[1]) && time_(m[2])) return { time: time_(m[1]), end: time_(m[2]) };
  return { time: time_(v, shown), end: '' };
}
function mins_(hm) { const p = String(hm).split(':'); return +p[0] * 60 + +p[1]; }
function hm_(m) { m = ((m % 1440) + 1440) % 1440; return pad2_(Math.floor(m / 60)) + ':' + pad2_(m % 60); }

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

/** Ops Test ▸ Reset a student's test: removes every teacher's marks for one student; the student stays on the list. */
function menuResetStudent() {
  const ui = SpreadsheetApp.getUi();
  const p = ui.prompt('Reset a student\'s test', 'The student\'s name as in the Students tab (add the date if the name is there twice, like "Kim 02.11.2026"):', ui.ButtonSet.OK_CANCEL);
  if (p.getSelectedButton() !== ui.Button.OK) return;
  const r = withLock_(() => resetStudent_(p.getResponseText()));
  ui.alert(r.ok ? 'Reset: ' + r.name + ' (' + r.rows + ' marking' + (r.rows === 1 ? '' : 's') + ' removed). The student stays on the list.' : r.error);
}
function resetStudent_(text) {
  const t = clean_(text, 120), m = t.match(/^(.*?)\s+(\d{1,2}[.\/]\d{1,2}[.\/]\d{2,4}|\d{4}-\d{2}-\d{2})$/);
  const name = norm_(m ? m[1] : t), day = m ? day_(m[2]) : '';
  const found = students_().filter(s => norm_(s.name) === name && (!day || s.day === day));
  if (!found.length) return { ok: false, error: 'No student called "' + t + '" in the Students tab.' };
  if (found.length > 1) return { ok: false, error: 'There are ' + found.length + ' students with that name. Add the date, like "' + found[0].name + ' ' + found[0].day.split('-').reverse().join('.') + '".' };
  const n = removeMarksRows_(row => row.sid === found[0].id);
  gradesDirty_(); buildGrades_();
  return { ok: true, name: found[0].name, rows: n };
}
/** Ops Test ▸ Reset ALL test results: every mark is removed and every test date takes the current questions again. */
function menuResetAll() {
  const ui = SpreadsheetApp.getUi();
  const p = ui.prompt('Reset ALL test results', 'This removes every mark, note and critical error for every student. Students, teachers and questions stay. Type RESET to go ahead:', ui.ButtonSet.OK_CANCEL);
  if (p.getSelectedButton() !== ui.Button.OK) return;
  if (p.getResponseText().trim() !== 'RESET') { ui.alert('Nothing was reset (type RESET in capitals to go ahead).'); return; }
  const n = withLock_(() => resetAll_());
  ui.alert('Done: ' + n + ' marking' + (n === 1 ? '' : 's') + ' removed.');
}
function resetAll_(quick) {
  const sh = sheet_(SHEETS.MARKS), n = Math.max(0, sh.getLastRow() - 1);
  if (n) sh.deleteRows(2, n);
  const props = PropertiesService.getScriptProperties();
  frozenKeys_().forEach(k => props.deleteProperty(k));
  gradesDirty_();
  if (!quick) syncMarks_();   // from a cell edit (30 s limit) the page's next tidy does this
  buildGrades_();
  return n;
}
/* The two reset cells on Start here: typed into, they reset, then say what happened in the cell itself.
   (The Ops Test menu does the same, but Sheets hides it when the window is narrow.) */
function startCell_(label) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.START.name);
  if (!sh) return null;
  const i = sh.getRange(1, 1, Math.max(1, sh.getLastRow()), 1).getValues().findIndex(r => String(r[0]) === label);
  return i < 0 ? null : sh.getRange(i + 1, 2);
}
function resetCells_(edited) {
  [[RESET_ONE_LABEL, t => {
     const r = withLock_(() => resetStudent_(t));
     return r.ok ? '✓ ' + r.name + ' reset: ' + r.rows + ' marking' + (r.rows === 1 ? '' : 's') + ' removed' : '✗ ' + r.error;
   }],
   [RESET_ALL_LABEL, t => {
     if (t !== 'RESET') return '✗ Nothing was reset. Type RESET in capitals.';
     const n = withLock_(() => resetAll_(true));
     if (typeof n !== 'number') return '✗ ' + n.error;
     return '✓ Everything reset: ' + n + ' marking' + (n === 1 ? '' : 's') + ' removed';
   }]].forEach(([label, run]) => {
    const cell = startCell_(label);
    if (!cell || (edited && (edited.getRow() > cell.getRow() || edited.getLastRow() < cell.getRow()))) return;
    const t = String(cell.getValue() || '').trim();
    if (!t || /^[✓✗]/.test(t)) return;
    let msg;
    try { msg = run(t); } catch (err) { msg = '✗ ' + (err && err.message || err); }
    cell.setValue(msg + ' · ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM HH:mm'));
  });
}
/** Removes the Marks rows that match, keeping the header. Returns how many went. */
function removeMarksRows_(match) {
  const sh = sheet_(SHEETS.MARKS), rows = marksRows_().filter(match).map(m => m.row).sort((a, b) => b - a);
  rows.forEach(r => sh.deleteRow(r));
  return rows.length;
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
  if (a === 'login') {
    // the students and questions come back with the sign-in, so the page needs no second trip before it shows them
    const r = withLock_(() => login_(req.name, req.password, req.pin));
    if (r.ok && req.withState) try { r.state = stateCached_({ name: r.name, role: r.role }, ''); } catch (e) {}
    return r;
  }
  const me = auth_(req.token);
  if (!me) return { ok: false, error: 'Please sign in again.', code: 'auth' };
  switch (a) {
    case 'state':          return stateCached_(me, req.h);   // only reads, so the page opens quickly
    case 'batch':          return withLock_(() => batch_(me, req));   // a whole test (or several) in one call
    case 'tidy':           return withLock_(() => { ensureStudentIds_(); ensureTeacherQuestionTabs_(); ensureQuestionIds_(); syncMarks_(); maybeBuildGrades_(); return { ok: true }; });
    case 'addStudents':    return withLock_(() => addStudents_(me, req));
    case 'updateStudent':  return withLock_(() => updateStudent_(me, req));
    case 'saveEntry':      return withLock_(() => saveEntry_(me, req));
    case 'publishResults': return withLock_(() => publishResults_(me, req));
    case 'grades':         return grades_(me);
    case 'logout':         PropertiesService.getScriptProperties().deleteProperty('s_' + req.token); return { ok: true };
    case 'listTeachers':   return admin_(me) || { ok: true, teachers: teachers_().map(t => ({ name: t.name, role: t.role })) };
    case 'addTeacher':     return admin_(me) || withLock_(() => setTeacher_(req.name, req.role === 'admin' ? 'admin' : 'teacher'));
    case 'removeTeacher':  return admin_(me) || (norm_(req.name) === norm_(me.name) ? { ok: false, error: 'You cannot remove yourself.' } : withLock_(() => removeTeacher_(req.name)));
  }
  return { ok: false, error: 'Unknown action.' };
}

// ─────────────────────────────────────────────────────────────── sign-in

/** A teacher picks their name from the list and types the team password. */
function login_(name, password, pin) {
  name = clean_(name, 40); password = String(password || '').trim();
  if (!name) return { ok: false, error: 'Pick your name.' };
  if (!password) return { ok: false, error: 'Type the team password.' };
  // wrong tries are counted per name and per team password: setting the password again clears every pause at once
  const stamp = PropertiesService.getScriptProperties().getProperty('team_at') || '0';
  const cache = CacheService.getScriptCache(), key = 'fail_' + stamp + '_' + norm_(name).replace(/[^a-z0-9]/g, '_');
  const fails = Number(cache.get(key) || 0);
  if (fails >= MAX_TRIES) return { ok: false, error: 'Too many wrong tries for this name. Wait 10 minutes, or ask the organiser to type the team password again on Start here (that clears it at once).' };
  const t = teachers_().find(x => norm_(x.name) === norm_(name));
  if (!t) return { ok: false, error: 'That name is not in the Teachers tab any more.' };
  storeTeamPassword_();                                          // a password typed in the sheet but not yet hidden
  const team = teamHash_();
  if (!team) return { ok: false, error: 'No team password is set yet. Ask the test organiser to set it on the Start here tab.' };
  if (!teamPasswordMatches_(password, team)) {
    cache.put(key, String(fails + 1), 600);
    // a new password typed on Start here that was too short was not taken, so the old one still applies: say so
    const cell = teamCell_(), pending = cell ? typedTeamPassword_(cell) : '';
    if (pending && pending !== SET) return { ok: false, error: 'That password is not right. A new password was typed on the Start here tab but it is too short, so the old one still applies. Type one with ' + MIN_PASSWORD + ' characters or more there.' };
    return { ok: false, error: 'That password is not right.' };
  }
  // an admin also gives the admin PIN (once one is set); a wrong PIN counts as a wrong try
  const adminPin = adminPinHash_();
  if (t.role === 'admin' && adminPin) {
    pin = String(pin || '').trim();
    if (!pin) return { ok: false, needPin: true, error: 'Type your admin PIN.' };
    if (hash_(pin, adminPin.salt) !== adminPin.hash) { cache.put(key, String(fails + 1), 600); return { ok: false, needPin: true, error: 'That admin PIN is not right.' }; }
  }
  cache.remove(key);
  return startSession_(t, t.role === 'admin' && !!adminPin);
}

function startSession_(t, pin) {
  const token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('s_' + token, JSON.stringify({ name: t.name, pin: !!pin, start: Date.now(), exp: Date.now() + SESSION_DAYS * 864e5 }));
  try { sheet_(SHEETS.TEACHERS).getRange(t.row, T.LAST + 1).setValue(new Date()); } catch (e) {}
  return { ok: true, token, name: t.name, role: t.role, days: SESSION_DAYS };
}

function auth_(token) {
  if (!token || !/^[0-9a-f]{64}$/.test(String(token))) return null;
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty('s_' + token);
  if (!raw) return null;
  const s = JSON.parse(raw);
  // a sign-in lasts SESSION_DAYS from when it started (older sign-ins were made for 400 days: they end 30 days after they began)
  const start = s.start || (s.exp - OLD_SESSION_DAYS * 864e5);
  if (s.exp < Date.now() || start + SESSION_DAYS * 864e5 < Date.now()) { props.deleteProperty('s_' + token); return null; }
  const t = teachers_().find(x => norm_(x.name) === norm_(s.name));
  if (!t) { props.deleteProperty('s_' + token); return null; }
  // admin rights only for a session that gave the admin PIN (when one is set); otherwise the account works as a teacher
  const role = t.role === 'admin' && (s.pin || !adminPinHash_()) ? 'admin' : 'teacher';
  return { name: t.name, role, token };
}

/** A teacher only ever gets the students whose Teacher column has their name; an admin gets everyone. */
function sees_(me, s) { return me.role === 'admin' || (!!s.teacher && norm_(s.teacher) === norm_(me.name)); }
const NOT_YOURS = { ok: false, error: 'This student is not on your list any more. Ask the organiser if they should be.' };

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

/** A name from the Teachers tab, spelled as it is there; '' for anything else. */
function teacherName_(name) { const t = teachers_().find(x => norm_(x.name) === norm_(name)); return t ? t.name : ''; }

function teachers_() {
  return rows_(SHEETS.TEACHERS).map((r, i) => ({ r, i })).filter(x => String(x.r[T.NAME] || '').trim()).map(({ r, i }) => ({
    row: i + 2, name: String(r[T.NAME]).trim(), role: String(r[T.ROLE] || 'teacher'), last: r[T.LAST] || '' }))
    .sort((a, b) => a.row - b.row);
}

// ─────────────────────────────────────────────────────────────── data

const S_ = { DAY: 0, TIME: 1, NAME: 2, CAP: 3, STUDIO: 4, TEACHER: 5, REMOVED: 6, ID: 7 };

function students_() {
  const sh = sheet_(SHEETS.STUDENTS), n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const range = sh.getRange(2, 1, n, SHEETS.STUDENTS.header.length), rows = range.getValues(), shown = range.getDisplayValues();
  return rows.map((r, i) => ({ row: i + 2, id: String(r[S_.ID] || ''), day: day_(r[S_.DAY]), ...slot_(r[S_.TIME], shown[i][S_.TIME]),
    name: String(r[S_.NAME] || '').trim(), cap: String(r[S_.CAP] || '').trim(), studio: String(r[S_.STUDIO] || ''), teacher: String(r[S_.TEACHER] || '').trim(), removed: !!r[S_.REMOVED] }))
    .filter(s => s.id && s.name && s.day);
}
/** Test date, then test time (students without one last), then name. */
function byTestTime_(a, b) { return a.day < b.day ? -1 : a.day > b.day ? 1 : (a.time || '99') < (b.time || '99') ? -1 : (a.time || '99') > (b.time || '99') ? 1 : a.name.localeCompare(b.name); }
function marksRows_() {
  const sh = sheet_(SHEETS.MARKS), last = sh.getLastRow(), w = sh.getLastColumn();
  if (last < 2 || !w) return [];
  const head = sh.getRange(1, 1, 1, w).getValues()[0].map(String);
  let dataCol = head.indexOf('Page data (do not edit)'), idCol = head.indexOf('Student ID');
  if (dataCol < 0) { dataCol = w - 1; idCol = w - 2; }
  const upCol = head.indexOf('Updated');
  return sh.getRange(2, 1, last - 1, w).getValues().map((r, i) => {
    let data = {}; try { data = JSON.parse(r[dataCol] || '{}'); } catch (e) {}
    return { row: i + 2, sid: String(r[idCol] || ''), teacher: String(r[2] || ''), data, day: day_(r[0]), name: String(r[1] || ''), updated: upCol >= 0 ? r[upCol] : '' };
  }).filter(x => x.sid);
}

function state_(me) {
  const students = students_().filter(s => !s.removed && sees_(me, s));
  const dayOf = {}; students.forEach(s => dayOf[s.id] = s.day);
  const entries = [];
  marksRows_().forEach(m => {
    if (!(m.sid in dayOf) || !m.teacher.trim()) return;   // rows without a teacher do not count
    if (norm_(m.teacher) === norm_(me.name)) entries.push({ sid: m.sid, teacher: m.teacher, data: m.data });
    else entries.push({ sid: m.sid, teacher: m.teacher, progress: Object.keys(m.data.marks || {}).filter(k => m.data.marks[k] in POINTS).length });
  });
  // the questions: each teacher's current set (by teacher, '' = shared), and the fixed set of every marked date and teacher
  // only the question tabs this teacher's students use (an admin: everyone's)
  const sets = liveSets_(me.role === 'admin' ? null : [me.name].concat(students.map(s => s.teacher))), qlive = { '': sets[''] }, qsets = {};
  [me.name].concat(students.map(s => s.teacher)).forEach(t => { const k = norm_(t); if (k && sets[k]) qlive[k] = sets[k]; });
  students.forEach(s => { const k = s.day + '|' + norm_(s.teacher); if (!(k in qsets) && frozenFor_(s.day, s.teacher)) qsets[k] = questionsFor_(s.day, s.teacher, sets); });
  const questions = sets[norm_(me.name)] || sets[''];
  return { ok: true, me: { name: me.name, role: me.role },
    students: students.map(s => ({ id: s.id, day: s.day, time: s.time, end: s.end, name: s.name, cap: s.cap, studio: s.studio, teacher: s.teacher })),
    entries, teachers: teachers_().map(t => t.name), questions, qlive, qsets };
}

/* Every write (and every edit in the Sheet) bumps a version number. A teacher's state is kept for up to a minute and
   reused while the version is the same, so most refreshes answer without reading the Sheet. The page sends the hash
   of the copy it already has, and an unchanged state comes back as just {same: true}.
   (Inserting or deleting rows does not fire onEdit; the one-minute limit covers that.) */
const STATE_KEEP_MS = 60000;
function bump_() { try { CacheService.getScriptCache().put('ver', String(Date.now()) + Math.random().toString(36).slice(2, 6), 21600); } catch (e) {} }
function stateCached_(me, have) {
  const c = CacheService.getScriptCache(), key = 'st_' + norm_(me.name).replace(/[^a-z0-9]/g, '_');
  let ver = null, hit = null;
  try { ver = c.get('ver'); hit = JSON.parse(c.get(key) || 'null'); } catch (e) {}
  if (!(ver && hit && hit.v === ver && Date.now() - hit.at < STATE_KEEP_MS)) {
    if (!ver) { bump_(); ver = c.get('ver'); }
    const body = JSON.stringify(state_(me));
    hit = { v: ver, at: Date.now(), h: Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, body, Utilities.Charset.UTF_8)).slice(0, 16), body };
    try { if (body.length < 90000) c.put(key, JSON.stringify(hit), 120); } catch (e) {}
  }
  if (have && have === hit.h) return { ok: true, same: true, h: hit.h };
  const out = JSON.parse(hit.body); out.h = hit.h; return out;
}
/** Several saves in one call and one lock: [['saveEntry', {sid, patch}], ['updateStudent', {id, ...}], ...]. */
function batch_(me, req) {
  const steps = Array.isArray(req.steps) ? req.steps.slice(0, 40) : [];
  return { ok: true, results: steps.map(st => {
    const a = st && st[0], b = (st && st[1]) || {};
    try {
      if (a === 'saveEntry') return saveEntry_(me, b);
      if (a === 'updateStudent') return updateStudent_(me, b);
    } catch (err) { return { ok: false, error: 'Something went wrong on the server: ' + (err && err.message || err) }; }
    return { ok: false, error: 'Unknown action.' };
  }) };
}

function addStudents_(me, req) {
  gradesDirty_();
  const day = checkDay_(req.day); if (!day) return { ok: false, error: 'Pick a test date first.' };
  // a line may start with its test time: "10:30 Kim"
  const lines = (Array.isArray(req.names) ? req.names : []).map(n => clean_(n, 90)).filter(Boolean).slice(0, 60)
    .map(l => {
      const m = l.match(/^(\d{1,2}[:.]\d{2})(?:\s*[-\u2013]\s*(\d{1,2}[:.]\d{2}))?\s+(.+)$/);
      if (!m || !time_(m[1])) return { time: '', name: clean_(l, 80) };
      return { time: m[2] && time_(m[2]) ? time_(m[1]) + ' - ' + time_(m[2]) : time_(m[1]), name: clean_(m[3], 80) };
    });
  if (!lines.length) return { ok: false, error: 'Type at least one name.' };
  const studio = STUDIO_LIST.indexOf(clean_(req.studio, 10)) >= 0 ? clean_(req.studio, 10) : STUDIO_LIST[0], sh = sheet_(SHEETS.STUDENTS);
  const t = me.role === 'admin' ? teacherName_(req.teacher) : me.name;   // a teacher adds students to their own list
  const rows = lines.map(l => [toDate_(day), l.time, l.name, '', studio, t, '', newId_()]);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  return { ok: true, ids: rows.map(r => r[S_.ID]) };
}

function updateStudent_(me, req) {
  gradesDirty_();
  const s = students_().find(x => x.id === String(req.id));
  if (!s || s.removed) return { ok: false, error: 'That student is no longer on the list.' };
  if (!sees_(me, s)) return NOT_YOURS;
  if (req.teacher !== undefined && me.role !== 'admin') return { ok: false, error: 'Only an admin can move a student to another teacher.' };
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
    if (marksRows_().some(m => m.sid === s.id && Object.keys(m.data.marks || {}).length)) freezeSet_(day, s.teacher);
    marksRows_().filter(m => m.sid === s.id).forEach(m => sheet_(SHEETS.MARKS).getRange(m.row, 1).setValue(toDate_(day)));
  }
  if (req.studio !== undefined) sh.getRange(s.row, S_.STUDIO + 1).setValue(clean_(req.studio, 40));
  if (req.time !== undefined) {
    // a new start time keeps the slot's length: 10:00 - 10:15 moved to 11:00 becomes 11:00 - 11:15
    const start = time_(req.time), len = s.time && s.end ? mins_(s.end) - mins_(s.time) : 0;
    sh.getRange(s.row, S_.TIME + 1).setValue(start && len > 0 ? start + ' - ' + hm_(mins_(start) + len) : start);
  }
  if (req.teacher !== undefined) {
    const t = teacherName_(req.teacher);
    sh.getRange(s.row, S_.TEACHER + 1).setValue(t);
    if (marksRows_().some(m => m.sid === s.id && Object.keys(m.data.marks || {}).length)) freezeSet_(s.day, t);   // marked: the new teacher's set is fixed too
  }
  return { ok: true };
}

/** Merges a patch into this teacher's own row. null removes a key. Only known fields and mark values get in. */
function saveEntry_(me, req) {
  gradesDirty_();
  const sid = String(req.sid || '');
  const s = students_().find(x => x.id === sid && !x.removed);
  if (!s) return { ok: false, error: 'That student is no longer on the list.' };
  if (!sees_(me, s)) return NOT_YOURS;
  const p = req.patch || {}, clean = {};
  if (p.marks) { clean.marks = {}; Object.keys(p.marks).slice(0, 200).forEach(k => { const v = p.marks[k]; if (v === null || v in POINTS) clean.marks[clean_(k, 40)] = v; }); }
  if (p.notes) { clean.notes = {}; Object.keys(p.notes).slice(0, 200).forEach(k => { clean.notes[clean_(k, 40)] = p.notes[k] === null ? null : String(p.notes[k]).slice(0, 1000); }); }
  if (p.safety) { clean.safety = {}; Object.keys(p.safety).slice(0, 50).forEach(k => { const v = p.safety[k]; clean.safety[clean_(k, 40)] = v === null ? null : { t: Number(v.t) || Date.now(), text: clean_(v.text, 200) }; }); }
  if (p.summary !== undefined) clean.summary = String(p.summary || '').slice(0, 3000);
  if (p.ended !== undefined) clean.ended = p.ended === null ? null : Number(p.ended) || Date.now();   // "End test and grade" was pressed
  if (clean.marks && Object.keys(clean.marks).some(k => clean.marks[k])) freezeSet_(s.day, s.teacher);   // the first mark fixes this date's questions for this teacher
  const ctx = syncMarks_();
  const mine = marksRows_().find(m => m.sid === sid && norm_(m.teacher) === norm_(me.name));
  const data = merge_(mine ? mine.data : {}, clean);
  const sh = sheet_(SHEETS.MARKS), row = marksRow_(s, me.name, data, ctx);
  const r = mine ? mine.row : sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, row.length).setValues([row]);
  return { ok: true };
}

/** The readable row in Marks: result, score and a % per stage first, every skill spelled out further right, plus the page's own copy. */
function marksRow_(s, teacher, data, ctx, updated) {
  ctx = ctx || marksCtx_();
  const qs = forStudio_(questionsFor_(s.day, s.teacher, ctx.sets), s.studio), m = data.marks || {}, ids = qs.map(q => q.id);
  const byName = {}; colNames_(qs).forEach((c, i) => byName[c.key] = ids[i]);
  const marked = ids.filter(id => m[id] in POINTS);
  const sc = stageScores_(qs, q => m[q.id] in POINTS ? POINTS[m[q.id]] : null);
  const errors = Object.keys(data.safety || {}).map(k => data.safety[k]).filter(Boolean).map(x => x.text);
  const result = resultOf_(sc.score, errors, marked.length, ids.length, !!data.ended);
  // the notes: the teacher's own, then any note on a single skill
  const notes = [data.summary || ''].concat(qs.filter(q => data.notes && data.notes[q.id]).map(q => q.label + ': ' + data.notes[q.id])).filter(x => String(x).trim()).join('\n');
  return [s.day ? toDate_(s.day) : '', s.name, teacher, result, sc.score === null ? '' : Math.round(sc.score), marked.length + '/' + ids.length]
    .concat(ctx.stages.map(st => st in sc.stages ? Math.round(sc.stages[st]) : ''))
    .concat([errors.join('; '), notes, data.ended ? new Date(Number(data.ended)) : '', updated || new Date()])
    .concat(ctx.cols.map(c => !(c.key in byName) ? NOT_ASKED : MARK_LABEL[m[byName[c.key]]] || ''))   // a dash: not part of this student's test
    .concat([s.id, JSON.stringify(data)]);
}
/** One stage list for every teacher's scoring. val(q) is a question's points 0..1, or null when it is not marked.
 *  Each stage: the average of its marked questions, as a %. Score: the average of those stages. */
function stageScores_(qs, val) {
  const by = {};
  qs.forEach(q => { const v = val(q); if (v !== null && v !== undefined) (by[q.stage] = by[q.stage] || []).push(v); });
  const stages = {}; Object.keys(by).forEach(k => stages[k] = by[k].reduce((a, b) => a + b, 0) / by[k].length * 100);
  const ks = Object.keys(stages);
  return { stages, score: ks.length ? ks.reduce((a, k) => a + stages[k], 0) / ks.length : null };
}
/** In words: Grade A / Pass / Resit once every skill is marked or the test was ended; before that In progress or Not started. */
function resultOf_(score, errors, marked, total, ended) {
  if (!marked && !errors.length) return ended ? 'Not marked' : 'Not started';
  if (errors.length) return 'Resit';
  if (!ended && marked < total) return 'In progress';
  return BANDS.find(b => score >= b[1] - 1e-9)[0];
}

/** Keeps the Marks columns in step with the questions. When they change, every row is rewritten from its own page copy. */
function syncMarks_() {
  const ctx = marksCtx_(), head = marksHeader_(ctx), sh = sheet_(SHEETS.MARKS), w = sh.getLastColumn();
  const now = w ? sh.getRange(1, 1, 1, w).getValues()[0].map(String) : [];
  if (now.join('\u0001') === head.join('\u0001')) return ctx;
  const studs = {}; students_().forEach(s => studs[s.id] = s);
  const rows = marksRows_().map(m => marksRow_(studs[m.sid] || { id: m.sid, day: m.day, name: m.name, studio: '', teacher: '' }, m.teacher, m.data, ctx, m.updated));
  if (sh.getMaxColumns() < head.length) sh.insertColumnsAfter(sh.getMaxColumns(), head.length - sh.getMaxColumns());
  sh.clearContents();
  sh.getRange(1, 1, 1, head.length).setValues([head]);
  if (rows.length) sh.getRange(2, 1, rows.length, head.length).setValues(rows);
  styleMarks_(ctx);
  return ctx;
}
function marksCtx_() { const sets = liveSets_(); return { sets, live: sets[''], cols: markColumns_(sets), stages: stageColumns_(sets) }; }
/** The stage columns: the shared list on Start here, then any other stage a question set still uses. */
function stageColumns_(sets) {
  const out = stages_().slice();
  Object.keys(sets).forEach(k => sets[k].forEach(q => { if (out.indexOf(q.stage) < 0) out.push(q.stage); }));
  return out;
}
function marksHeader_(ctx) { return MARKS_HEAD.concat(ctx.stages).concat(MARKS_MID).concat(ctx.cols.map(c => c.label)).concat(MARKS_END); }
/* The skill columns in Marks go by NAME (the Column name, or the question shortened), not by question:
   questions with the same name in different teachers' tabs share one column, and renaming a Column name renames it.
   Every teacher's questions in the Teachers order; the shared tab only while a student has no teacher. */
function markColumns_(sets) {
  sets = sets || liveSets_();
  const cols = [], at = {};
  const keys = teachers_().map(t => norm_(t.name)).filter(k => sets[k]);
  if (!keys.length || students_().some(s => !s.removed && !s.teacher)) keys.unshift('');
  const all = keys.map(k => sets[k] || []);
  frozenKeys_().forEach(k => all.push(questionsFor_(k.slice(2, 12), k.slice(13), sets)));   // a fixed date: its questions, named as now
  all.forEach(qs => {
    let prev = -1;
    colNames_(qs).forEach(c => {
      if (!(c.key in at)) { cols.splice(prev + 1, 0, c); cols.forEach((x, i) => at[x.key] = i); }   // a new name goes after the one before it
      prev = at[c.key];
    });
  });
  return cols;
}
/** Each question's column name in one set; a name used twice in the same set gets "(2)". */
function colNames_(qs) {
  const seen = {};
  return qs.map(q => {
    const base = String(q.label || q.id).trim(); let label = base, n = 1;
    while (seen[norm_(label)]) label = base + ' (' + (++n) + ')';
    seen[norm_(label)] = 1;
    return { key: norm_(label), label };
  });
}

// ─────────────────────────────────────────────────────────────── the questions (the Questions tab)

/* Each teacher has their own tab, "Questions · <name>", a copy of the shared Questions tab they can change.
   A student is asked their teacher's questions; a student without a teacher (or a teacher whose tab is empty) the shared ones. */
const OWN_PREFIX = 'Questions · ';
function teacherTabName_(name) { return OWN_PREFIX + String(name).replace(/[:\/?*\[\]]/g, ' ').trim().slice(0, 80); }
function isQuestionSheet_(name) { return name === SHEETS.QUESTIONS.name || String(name).indexOf(OWN_PREFIX) === 0; }
function questionSheets_() { return SpreadsheetApp.getActive().getSheets().filter(sh => isQuestionSheet_(sh.getName())).map(sh => SpreadsheetApp.getActive().getSheetByName(sh.getName())); }
/** The questions a teacher's students are asked: their own tab, or the shared one. No teacher: the shared one. */
function questions_(teacher) {
  if (teacher) { const own = SpreadsheetApp.getActive().getSheetByName(teacherTabName_(teacher)); const qs = readQuestions_(own); if (qs.length) return qs; }
  return readQuestions_(SpreadsheetApp.getActive().getSheetByName(SHEETS.QUESTIONS.name));
}
/** Every current set, by teacher ('' = shared), read once per request. */
function liveSets_(only) {
  const sets = { '': questions_() };
  const names = only ? only.filter(Boolean) : teachers_().map(t => t.name);
  names.forEach(n => { if (!(norm_(n) in sets)) sets[norm_(n)] = questions_(n); });
  return sets;
}
/** Gives every teacher in the Teachers tab their own questions tab, starting as a copy of the shared one. */
function ensureTeacherQuestionTabs_() {
  const ss = SpreadsheetApp.getActive(), base = ss.getSheetByName(SHEETS.QUESTIONS.name);
  if (!base) return;
  const W = SHEETS.QUESTIONS.header.length, rows = rowsStd_(base);
  teachers_().forEach(t => {
    const name = teacherTabName_(t.name);
    if (ss.getSheetByName(name)) return;
    const sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, W).setValues([SHEETS.QUESTIONS.header]);
    if (rows.length) sh.getRange(2, 1, rows.length, W).setValues(rows);
    try { styleQuestions_(sh); } catch (e) {}
  });
}

/** The questions as typed in one questions tab, in test order. A blank Stage carries on the stage of the row above. */
function readQuestions_(sh) {
  if (!sh || sh.getLastRow() < 2) return [];
  const rows = rowsStd_(sh);
  const out = [], heading = {}, seen = {};
  let stage = '';
  rows.forEach(r => {
    if (clean_(r[Q_.STAGE], 40)) stage = stageName_(clean_(r[Q_.STAGE], 40));
    if (stage && clean_(r[Q_.HEADING], 80) && !heading[stage]) heading[stage] = clean_(r[Q_.HEADING], 80);
    const t = clean_(r[Q_.TEXT], 300), id = clean_(r[Q_.ID], 40);
    if (!t || !id || seen[id]) return;
    seen[id] = 1;
    out.push({ id, stage: stage || 'Test', t, ask: clean_(r[Q_.ASK], 300), core: 0,
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

/** Every typed question gets an ID in the hidden column; a row copied together with its ID (in the same tab) gets a fresh one.
 *  The same ID in two teachers' tabs is the same question, so marks line up. */
function ensureQuestionIds_() { questionSheets_().forEach(ensureQuestionIdsIn_); }
/* Every question keeps an ID (hidden ID column) that its marks hang on. Each tab remembers which text each ID had, so:
   - a question MOVED by any means (dragging rows, sorting, copying or retyping cells) gets its own ID back;
   - a question REPLACED by a different one gets a new ID (the old marks do not slide onto it);
   - a question REWORDED a little keeps its ID. */
function ensureQuestionIdsIn_(sh) {
  if (!sh || sh.getLastRow() < 2) return;
  let c = qcols_(sh);
  if (c.id < 0) { sh.getRange(1, c.width + 1).setValue('ID'); c = qcols_(sh); }
  if (c.text < 0) return;
  const n = sh.getLastRow() - 1, texts = sh.getRange(2, c.text + 1, n, 1).getValues(), idr = sh.getRange(2, c.id + 1, n, 1), ids = idr.getValues();
  const before = ids.map(r => String(r[0] || '').trim());
  const props = PropertiesService.getScriptProperties(), key = 'qreg_' + sh.getName(), lkey = 'qlab_' + sh.getName();
  let reg = {}; try { reg = JSON.parse(props.getProperty(key) || '{}'); } catch (e) {}
  let labs = {}; try { labs = JSON.parse(props.getProperty(lkey) || '{}'); } catch (e) {}
  const byText = {}; Object.keys(reg).forEach(id => byText[normQ_(reg[id])] = id);
  const taken = {}, next = {}; let changed = false;
  // first pass: rows whose ID still holds its own text keep it
  ids.forEach((r, i) => { const id = String(r[0] || '').trim(), t = String(texts[i][0] || '').trim();
    if (t && id && reg[id] !== undefined && normQ_(reg[id]) === normQ_(t) && !taken[id]) taken[id] = i + 1; });
  ids.forEach((r, i) => {
    const t = String(texts[i][0] || '').trim();
    if (!t) return;
    let id = String(r[0] || '').trim();
    if (id && taken[id] === i + 1) { next[id] = t; return; }
    const moved = byText[normQ_(t)];
    if (moved && !taken[moved]) id = moved;                                              // this text had an ID elsewhere: it moved here
    else if (id && !taken[id] && (reg[id] === undefined || similar_(reg[id], t))) {}     // first time seen, or a small rewording
    else id = 'q' + newId_().replace(/-/g, '').slice(0, 8);                             // a new or replaced question
    taken[id] = i + 1; next[id] = t;
    if (String(r[0] || '').trim() !== id) { r[0] = id; changed = true; }
  });
  if (changed) idr.setValues(ids);
  // a Column name belongs to its question: when a row now holds a different question, the old name goes (or the moved one comes along)
  const nextLabs = {};
  if (c.label >= 0) {
    const lr = sh.getRange(2, c.label + 1, n, 1), lv = lr.getValues();
    let lchanged = false;
    ids.forEach((r, i) => {
      const id = String(r[0] || '').trim(), was = before[i], lab = String(lv[i][0] || '').trim();
      if (id && was && id !== was && labs[was] !== undefined && lab === labs[was]) { lv[i][0] = labs[id] || ''; lchanged = true; }
      if (id) nextLabs[id] = String(lv[i][0] || '').trim();
    });
    if (lchanged) lr.setValues(lv);
  }
  Object.keys(next).forEach(k => next[k] = next[k].slice(0, 160));
  try { props.setProperty(key, JSON.stringify(next)); props.setProperty(lkey, JSON.stringify(nextLabs)); } catch (e) {}
}
function normQ_(t) { return String(t || '').toLowerCase().replace(/[^a-z0-9äöüßæøåéèáíóúð]+/g, ' ').trim(); }
/** Two wordings of the same question: most of their words are the same. */
function similar_(a, b) {
  const words = t => normQ_(t).split(' ').filter(w => w.length > 2);
  const A = words(a), B = words(b); if (!A.length || !B.length) return false;
  const inB = A.filter(w => B.indexOf(w) >= 0).length;
  return inB / Math.max(A.length, B.length) >= 0.5;
}

function seedQuestions_() {
  const ss = SpreadsheetApp.getActive();
  if (ss.getSheetByName(SHEETS.QUESTIONS.name)) return;
  const sh = sheet_(SHEETS.QUESTIONS);
  sh.getRange(2, 1, DEFAULT_QUESTIONS.length, DEFAULT_QUESTIONS[0].length).setValues(DEFAULT_QUESTIONS);
  try { const at = ss.getSheets().map(s => s.getName()).indexOf(SHEETS.STUDENTS.name); if (at >= 0) { ss.setActiveSheet(sh); ss.moveActiveSheet(at + 2); } } catch (e) {}
}

/* A test date keeps the questions it was first marked with, per teacher, so changing the test never moves a grade already given.
   Rewording a question still shows on that date; adding, removing, Core and Studio changes do not.
   Stored as script property q_<date>|<teacher> (older dates: q_<date>, for every teacher that day);
   Ops Test ▸ Use the current questions for a test date removes them. */
function frozenRaw_(key) { try { return JSON.parse(PropertiesService.getScriptProperties().getProperty(key) || 'null'); } catch (e) { return null; } }
function frozenFor_(day, teacher) { return frozenRaw_('q_' + day + '|' + norm_(teacher)) || frozenRaw_('q_' + day); }
function frozenKeys_() { return Object.keys(PropertiesService.getScriptProperties().getProperties()).filter(k => /^q_\d{4}-\d{2}-\d{2}(\|.*)?$/.test(k)).sort(); }
function frozenDays_() { const d = {}; frozenKeys_().forEach(k => d[k.slice(2, 12)] = 1); return Object.keys(d).sort(); }
function freezeSet_(day, teacher, qs) {
  if (!day || frozenFor_(day, teacher)) return;
  qs = qs || questions_(teacher);
  if (!qs.length) return;
  let s = JSON.stringify(qs);   // a script property holds 9 KB: drop the long text first, it comes from the tab anyway
  if (s.length > 8500) s = JSON.stringify(qs.map(q => ({ id: q.id, stage: q.stage, t: q.t.slice(0, 60), core: q.core, studios: q.studios, label: q.label })));
  if (s.length > 8500) s = JSON.stringify(qs.map(q => ({ id: q.id, core: q.core, studios: q.studios })));
  PropertiesService.getScriptProperties().setProperty('q_' + day + '|' + norm_(teacher), s);
}
/** A student's questions: the set fixed for their date and teacher (with today's wording), or their teacher's current set. */
function questionsFor_(day, teacher, sets) {
  const live = sets ? (sets[norm_(teacher)] && sets[norm_(teacher)].length ? sets[norm_(teacher)] : sets['']) : questions_(teacher);
  const f = frozenFor_(day, teacher);
  if (!f) return live;
  const byId = {}; live.forEach(q => byId[q.id] = q);
  return f.map(q => {
    const l = byId[q.id], o = Object.assign({ stage: 'Test', t: '(removed question)', ask: '', label: q.id }, q);
    if (l) Object.assign(o, { stage: l.stage, heading: l.heading, t: l.t, ask: l.ask, label: l.label });
    o.stage = stageName_(o.stage); o.heading = o.heading || o.stage;
    return o;
  });
}
/** Dates fixed before teachers had their own tabs (q_<date>, one set for everyone) become one fixed set per teacher.
 *  A teacher whose tab still holds the same questions keeps that set exactly; a teacher who has since changed which
 *  questions are in their tab gets their current tab, since the old shared set was never theirs to choose. */
function splitDayFreezes_() {
  const props = PropertiesService.getScriptProperties(), studs = students_().filter(s => !s.removed);
  Object.keys(props.getProperties()).filter(k => /^q_\d{4}-\d{2}-\d{2}$/.test(k)).forEach(k => {
    const day = k.slice(2), old = frozenRaw_(k) || [], oldIds = old.map(q => q.id).sort().join('|');
    const teachersThatDay = {};
    studs.filter(x => x.day === day).forEach(x => teachersThatDay[norm_(x.teacher)] = x.teacher);
    Object.keys(teachersThatDay).forEach(t => {
      const key = 'q_' + day + '|' + t;
      if (props.getProperty(key)) return;
      const live = questions_(teachersThatDay[t]);
      const same = live.map(q => q.id).sort().join('|') === oldIds;
      props.setProperty(key, same ? JSON.stringify(old) : JSON.stringify(live));
    });
    props.deleteProperty(k);
  });
}
function freezeMarkedDays_() {
  const of = {}; students_().forEach(s => of[s.id] = s);
  marksRows_().forEach(m => { const s = of[m.sid]; if (s && Object.keys(m.data.marks || {}).length) freezeSet_(s.day, s.teacher); });
}
function useCurrentQuestions_(day) {
  const props = PropertiesService.getScriptProperties();
  frozenKeys_().filter(k => k.slice(2, 12) === day).forEach(k => props.deleteProperty(k));
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

/** The combined grade: each skill averaged over every teacher who marked it, each stage the average of its skills,
 *  the score the average of the stages. Any critical error means Resit. */
function gradeOf_(byTeacher, qs) {
  const teachers = Object.keys(byTeacher).filter(u => Object.keys(byTeacher[u].marks || {}).some(k => byTeacher[u].marks[k] in POINTS));
  const per = qs.map(q => {
    const v = teachers.map(u => byTeacher[u].marks[q.id]).filter(x => x in POINTS).map(x => POINTS[x]);
    return { q, avg: v.length ? v.reduce((a, b) => a + b, 0) / v.length : null };
  });
  const marked = per.filter(x => x.avg !== null);
  const sc = stageScores_(qs, q => per.find(x => x.q === q).avg), score = sc.score;   // every stage counts the same
  const errors = [];
  Object.keys(byTeacher).forEach(u => Object.keys(byTeacher[u].safety || {}).forEach(k => { const e = byTeacher[u].safety[k]; if (e && errors.indexOf(e.text) < 0) errors.push(e.text); }));
  const final = score;   // no points off: a critical error decides the grade on its own (below)
  const coreFails = [];   // (Core skills are no longer used: every question counts the same)
  const complete = qs.length > 0 && marked.length === qs.length;
  const band = final === null ? null : errors.length ? BANDS[BANDS.length - 1] : BANDS.find(b => final >= b[1] - 1e-9);
  const stages = {}; Object.keys(sc.stages).forEach(k => stages[k] = Math.round(sc.stages[k]));
  return { teachers, marked: marked.length, score, errors, final, coreFails, complete, band, stages };
}

/** Every student with their combined grade, each on their own date's questions. Used by the Grades tab and the page. */
function gradeList_() {
  const students = students_().filter(s => !s.removed).sort(byTestTime_);
  // a Marks row whose Teacher cell is empty cannot be anyone's marking (e.g. cleared by hand): it does not count
  const marks = {}; marksRows_().forEach(m => { if (m.teacher.trim()) (marks[m.sid] = marks[m.sid] || {})[m.teacher] = m.data; });
  const sets = liveSets_(), byKey = {};
  const list = students.map(s => {
    const k = s.day + '|' + norm_(s.teacher);
    const qs = forStudio_(byKey[k] = byKey[k] || questionsFor_(s.day, s.teacher, sets), s.studio);
    return { s, qs, g: gradeOf_(marks[s.id] || {}, qs) };
  });
  return { live: sets[''], list };
}

/** The page's Grades screen: the combined result per student, never one teacher's marks. */
function grades_(me) {
  return { ok: true, students: gradeList_().list.filter(x => sees_(me, x.s)).map(({ s, qs, g }) => {
    const stages = [];
    qs.forEach(q => { if (!stages.some(x => x.name === q.stage)) stages.push({ name: q.stage, pct: q.stage in g.stages ? g.stages[q.stage] : null }); });
    return { id: s.id, day: s.day, time: s.time, end: s.end, name: s.name, cap: s.cap, studio: s.studio, teacher: s.teacher, marked: g.marked, total: qs.length,
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
  // grouped by teacher (in the order of the Teachers tab, students without a teacher last), then date and time
  const g0 = gradeList_(), live = g0.live, order = teachers_().map(t => norm_(t.name));
  const rank = s => { const i = order.indexOf(norm_(s.teacher)); return i < 0 ? order.length : i; };
  const list = g0.list.slice().sort((a, b) => rank(a.s) - rank(b.s) || byTestTime_(a.s, b.s)), students = list.map(x => x.s);
  const stageNames = [];
  stages_().concat(live.concat(...list.map(x => x.qs)).map(q => q.stage)).forEach(n => { if (stageNames.indexOf(n) < 0) stageNames.push(n); });
  const nS = stageNames.length, top = 8;
  const head = ['Test date', 'Student', 'Studio', 'Teachers', 'Marked', 'Progress'].concat(stageNames)
    .concat(['Score %', 'Errors', '', 'Grade', 'Critical errors']);
  const W = head.length, C = { STAGE: 7, SCORE: 7 + nS, ERR: 8 + nS, BAR: 9 + nS, GRADE: 10 + nS };

  const counts = { 'Grade A': 0, Pass: 0, Resit: 0, Unfinished: 0 };
  const values = [], bars = [], groupStarts = [], bandRows = [];
  let prevDay = '', prevTeacher = null;
  list.forEach(({ s, qs, g }) => {
    const N = qs.length;
    const label = !g.marked ? 'Not started' : g.complete ? g.band[0] : 'Unfinished';
    if (g.complete) counts[g.band[0]]++; else counts.Unfinished++;
    const tk = rank(s);
    if (tk !== prevTeacher) {
      // a band naming the teacher and how far their students are
      const mine = list.filter(x => rank(x.s) === tk), done = mine.filter(x => x.g.complete).length;
      bandRows.push(top + values.length);
      values.push([(s.teacher || 'No teacher yet') + '  ·  ' + mine.length + ' student' + (mine.length === 1 ? '' : 's') + '  ·  ' + done + ' graded'].concat(Array(W - 1).fill('')));
      bars.push(['', '']);
      prevTeacher = tk; prevDay = '';
    }
    if (s.day !== prevDay) { if (prevDay) groupStarts.push(top + values.length); prevDay = s.day; }
    values.push([toDate_(s.day), s.name, s.studio, g.teachers.join(', '), g.marked + '/' + N, '']
      .concat(stageNames.map(n => n in g.stages ? g.stages[n] : ''))
      .concat([g.score === null ? '' : Math.round(g.score), g.errors.length || '', '', label, g.errors.join('; ')]));
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
    sh.getRange(top, C.SCORE, values.length, 2).setHorizontalAlignment('center').setNumberFormat('0');
    sh.getRange(top, C.SCORE, values.length, 1).setFontWeight('bold').setFontSize(12);
    sh.getRange(top, C.GRADE, values.length, 1).setHorizontalAlignment('center').setFontWeight('bold');
    sh.getRange(top, C.GRADE + 1, values.length, 1).setFontColor(MUTE).setWrap(true);
    // a heavier line where a new test date starts
    groupStarts.forEach(r => { try { sh.getRange(r, 1, 1, W).setBorder(true, null, null, null, null, null, INK, SpreadsheetApp.BorderStyle.SOLID_MEDIUM); } catch (e) {} });
    for (let r = top; r < top + values.length; r++) sh.setRowHeight(r, 26);
    bandRows.forEach(r => { try {
      sh.getRange(r, 1, 1, W).setBackground(INK).setFontColor('#ffffff').setFontWeight('bold').setFontSize(11).setNumberFormat('@').setBorder(null, null, null, null, null, null);
      sh.setRowHeight(r, 30);
    } catch (e) {} });
    const heat = range => SpreadsheetApp.newConditionalFormatRule()
      .setGradientMinpointWithValue('#f4c7c3', SpreadsheetApp.InterpolationType.NUMBER, '0')
      .setGradientMidpointWithValue('#fce8b2', SpreadsheetApp.InterpolationType.NUMBER, '60')
      .setGradientMaxpointWithValue('#b7e1cd', SpreadsheetApp.InterpolationType.NUMBER, '100').setRanges([range]).build();
    const gradeCol = sh.getRange(top, C.GRADE, values.length, 1);
    const text = (t, bg, fg) => SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(t).setBackground(bg).setFontColor(fg).setBold(true).setRanges([gradeCol]).build();
    sh.setConditionalFormatRules((nS ? [heat(sh.getRange(top, C.STAGE, values.length, nS))] : []).concat([
      heat(sh.getRange(top, C.SCORE, values.length, 1)),
      text('Grade A', BANDS[0][3], BANDS[0][2]), text('Pass', BANDS[1][3], INK), text('Resit', BANDS[2][3], BANDS[2][2]),
      text('Unfinished', PAPER, MUTE), text('Not started', '#ffffff', '#9aa3b5'),
      SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setFontColor(BANDS[2][2]).setBold(true).setRanges([sh.getRange(top, C.ERR, values.length, 1)]).build(),
    ]));
  } else {
    sh.getRange(top, 1).setValue('No students yet. Add them in the Students tab or on the page.').setFontColor(MUTE);
  }
  [110, 170, 90, 150, 64, 110].concat(stageNames.map(() => 66)).concat([70, 58, 110, 104, 300]).forEach((w, i) => sh.setColumnWidth(i + 1, w));
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
  if (!o || o.head.join('|') === SHEETS.STUDENTS.header.join('|')) return;
  if (o.head[1] === 'Test time') {
    // layouts with a time: columns added later go in where they belong, nothing else moves
    const add = (name, after) => {
      const head = o.sh.getRange(1, 1, 1, o.sh.getLastColumn()).getValues()[0].map(String);
      if (head.indexOf(name) >= 0 || head.indexOf(after) < 0) return;
      const at = head.indexOf(after) + 1;
      o.sh.insertColumnAfter(at);
      o.sh.getRange(1, at + 1).setValue(name);
    };
    add('Teacher', 'Studio');
    add('CAP', 'Student');
    return;
  }
  let rows;
  if (o.head[0] === 'Test date') {
    // v2 layout: Test date | Student | Studio | Added by | Added at | Removed | ID
    rows = o.rows.filter(r => r[1] || r[6]).map(r => [r[0], '', r[1], '', r[2], '', r[5] || '', r[6]]);
  } else {
    // v1 layout: ID | Test date | Student | Studio | Added by | Added at | Deleted
    rows = o.rows.filter(r => r[0]).map(r => [toDate_(day_(r[1])), '', r[2], '', r[3], '', r[6] || '', r[0]]);
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
  o.sh.clear(); const h = marksHeader_(ctx);
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
  try { return fn(); } finally { bump_(); lock.releaseLock(); }
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
