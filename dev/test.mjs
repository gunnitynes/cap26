// Checks the sign-in paths of Code.gs on this computer:  node dev/test.mjs
import assert from 'node:assert/strict';
import { loadScript } from './google-mock.mjs';

// the checks edit the mock sheet directly, which in the real Sheet fires onEdit (and so a new version); raw() does not
const raw = (g, body) => JSON.parse(g.doPost({ postData: { contents: JSON.stringify(body) } }).s);
const call = (g, body) => { if (body.action === 'state') g.bump_(); return raw(g, body); };
const tab = (book, name) => book.sheets.find(s => s.name === name);
let n = 0; const ok = msg => console.log('  ✓', msg, ++n && '');

// 1. Older layouts: personal passwords or codes per teacher become a plain list of names.
{
  const { g, book } = loadScript();
  const sh = g.SpreadsheetApp.getActive().insertSheet('Teachers');
  sh.getRange(1, 1, 1, 7).setValues([['Name', 'Sign-in code', 'Role', 'Status', 'Last sign-in', 'hash', 'salt']]);
  sh.getRange(2, 1, 2, 7).setValues([['Gunni', '●●●●●● set', 'admin', 'Admin', '', 'xx', 'yy'], ['Isuru', '', '', 'No code yet', '', '', '']]);
  g.setup();
  const t = tab(book, 'Teachers').data;
  assert.deepEqual(t[0], ['Name', 'Role', 'Last sign-in']);
  assert.deepEqual(t.slice(1).map(r => r.slice(0, 2)), [['Gunni', 'admin'], ['Isuru', 'teacher']]);
  ok('older Teachers layouts become Name | Role | Last sign-in, keeping roles');
  assert.equal(book.sheets[0].name, 'Start here');
  ok('"Start here" is the first tab');
  assert.deepEqual(call(g, { action: 'teachers' }).teachers, ['Gunni', 'Isuru']);
  ok('the sign-in page gets the list of names without signing in');
  assert.match(call(g, { action: 'login', name: 'Isuru', password: 'anything' }).error, /No team password/);
  ok('before a team password is set, sign-in says so plainly');
}

// 2. The team password, typed on Start here (what onEdit sees).
{
  const { g, book } = loadScript();
  g.setup();
  const ss = g.SpreadsheetApp.getActive(), start = ss.getSheetByName('Start here'), te = ss.getSheetByName('Teachers');
  te.getRange(2, 1, 2, 1).setValues([['Gunni'], ['Anna']]); g.onEdit({ range: te.getRange(2, 1, 2, 1) });
  assert.deepEqual(tab(book, 'Teachers').data.slice(1).map(r => r[1]), ['teacher', 'teacher']);
  ok('a name typed in Teachers gets a role by itself');
  const cell = g.teamCell_();
  cell.setValue('abc'); g.onEdit({ range: cell });
  assert.equal(g.teamHash_(), null);
  ok('a too-short team password is not taken');
  cell.setValue('test-pass-1'); g.onEdit({ range: cell });
  assert.equal(cell.getValue(), '●●●●●● set');
  assert.ok(!JSON.stringify(tab(book, 'Start here').data).includes('test-pass-1'));
  ok('the team password is hidden the moment it is typed');
  let r = call(g, { action: 'login', name: 'anna', password: 'test-pass-1' });
  assert.equal(r.ok, true); assert.equal(r.days, 30);
  const annaTok = r.token;
  assert.equal(call(g, { action: 'login', name: 'Gunni', password: 'test-pass-1' }).ok, true);
  ok('every teacher signs in with their name and the same password, and stays signed in');
  assert.equal(call(g, { action: 'login', name: 'Anna', password: 'wrong-one' }).ok, false);
  assert.match(call(g, { action: 'login', name: 'Zed', password: 'test-pass-1' }).error, /not in the Teachers tab/);
  ok('a wrong password or a name not in Teachers is refused');
  cell.setValue('new-pass-2'); g.onEdit({ range: cell });
  assert.equal(call(g, { action: 'state', token: annaTok }).ok, true);
  assert.equal(call(g, { action: 'login', name: 'Anna', password: 'test-pass-1' }).ok, false);
  assert.equal(call(g, { action: 'login', name: 'Anna', password: 'new-pass-2' }).ok, true);
  ok('changing the team password keeps signed-in phones signed in, and only the new one works');
  // a password typed while the trigger did not run is picked up at sign-in
  cell.setValue('third-pass-3');
  assert.equal(call(g, { action: 'login', name: 'Anna', password: 'third-pass-3' }).ok, true);
  ok('a password typed without the edit trigger still works at the next sign-in');
  assert.equal(call(g, { action: 'login', name: 'Anna', password: 'Third-pass-3' }).ok, true);
  assert.equal(call(g, { action: 'login', name: 'Anna', password: ' third-pass-3 ' }).ok, true);
  ok('a first letter the phone capitalised, or spaces around it, still sign in');
  cell.setValue('short'); g.onEdit({ range: cell });
  assert.equal(call(g, { action: 'login', name: 'Anna', password: 'third-pass-3' }).ok, true);
  assert.match(call(g, { action: 'login', name: 'Anna', password: 'short' }).error, /too short, so the old one still applies/);
  ok('a too-short new password is not taken: the old one keeps working and a wrong try says why');
  cell.setValue('third-pass-3'); g.onEdit({ range: cell });
  te.deleteRow(3);
  assert.equal(call(g, { action: 'state', token: annaTok }).code, 'auth');
  ok('removing a teacher from the tab signs them out');
  // 5. A teacher marks, the readable Marks row and the hidden copy agree.
  const gt = call(g, { action: 'login', name: 'gunni', password: 'third-pass-3' }).token;
  call(g, { action: 'addStudents', token: gt, names: ['Polina'], studio: 'K3', day: '2026-11-04' });
  const sid = call(g, { action: 'state', token: gt }).students[0].id;
  call(g, { action: 'saveEntry', token: gt, sid, patch: { marks: { p1: 'y', m4: 'h', c2: 'n', zz: 'bogus' }, safety: { a: { t: 1, text: 'Speakers switched with level up' } } } });
  const m = tab(book, 'Marks').data;
  const head = m[0], row = m[1];
  assert.equal(row[head.indexOf('Student')], 'Polina'); assert.equal(row[head.indexOf('Teacher')], 'Gunni');
  assert.equal(row[head.indexOf('Phantom power')], 'Hint'); assert.equal(row[head.indexOf('Pre vs post fader')], 'Failed');
  assert.equal(row[head.indexOf('Skills marked')], '3/18');
  assert.equal(row[head.indexOf('Critical errors')], 'Speakers switched with level up');
  assert.equal(head.indexOf('Critical errors'), head.indexOf('Feedback') - 1);   // last before Feedback
  assert.deepEqual(JSON.parse(row[head.length - 1]).marks, { p1: 'y', m4: 'h', c2: 'n' });
  ok('Marks shows every skill in words; unknown values are dropped');
  const st = call(g, { action: 'state', token: gt });
  assert.equal(st.entries[0].data.marks.m4, 'h');
  ok('the page reads its marks back from the sheet');
}
// 6. Students typed or pasted straight into the Students tab: date | time | name | CAP | studio | teacher | removed | id.
{
  const { g, book } = loadScript();
  g.setup();
  const ss = g.SpreadsheetApp.getActive(), te = ss.getSheetByName('Teachers'), st = ss.getSheetByName('Students');
  assert.deepEqual(tab(book, 'Students').data[0], ['Test date', 'Test time', 'Student', 'CAP', 'Studio', 'Teacher', 'Removed', 'ID']);
  te.getRange(2, 1, 1, 2).setValues([['Gunni', 'admin']]); g.setTeamPassword_('test-pass-1');
  const tok = call(g, { action: 'login', name: 'gunni', password: 'test-pass-1' }).token;
  // dates must come from inside the sandbox, like the ones Google hands the script
  const pick = d => { const x = g.toDate_(d); x.setHours(0); return x; };   // what the date picker stores: midnight
  // 9.30 typed in Sheets arrives as the number 9.3
  st.getRange(2, 1, 1, 5).setValues([[pick('2026-11-04'), 9.3, 'Polina', 'CAP 1', 'K3']]);
  g.onEdit({ range: st.getRange(2, 1, 1, 5) });
  st.getRange(3, 2, 3, 2).setValues([['10:30', 'Paul'], [1100, 'Faddei'], ['', '  Yutaro ']]);
  g.onEdit({ range: st.getRange(3, 2, 3, 2) });
  st.getRange(6, 1, 1, 5).setValues([['5.11.2026', '8:05', 'Tales', 'CAP 2', 'K4']]);
  g.onEdit({ range: st.getRange(6, 1, 1, 5) });
  const rows = tab(book, 'Students').data.slice(1);
  assert.ok(rows.every(r => r[7]), 'every named row gets an ID');
  assert.deepEqual(rows.map(r => r[4]), ['K3', 'K3', 'K3', 'K3', 'K4']);
  ok('a pasted list takes the date and studio from the row above and gets IDs');
  assert.deepEqual(rows.map(r => r[1]), ['09:30', '10:30', '11:00', '', '8:05']);
  ok('test times typed as 9.30 or 1100 (numbers) are stored as a time; text stays as typed');
  const s1 = call(g, { action: 'state', token: tok }).students;
  assert.deepEqual(s1.map(x => x.name + ' ' + x.day + ' ' + x.time), ['Polina 2026-11-04 09:30', 'Paul 2026-11-04 10:30', 'Faddei 2026-11-04 11:00', 'Yutaro 2026-11-04 ', 'Tales 2026-11-05 08:05']);
  ok('the page sees them on the right dates and times, including a date typed as text');
  assert.deepEqual(s1.map(x => x.cap), ['CAP 1', '', '', '', 'CAP 2']);
  ok('each student\'s CAP from the sheet reaches the page');
  st.getRange(7, 1, 1, 5).setValues([[pick('2026-11-05'), '', 'Ida', '', 'K4']]);
  call(g, { action: 'tidy', token: tok });
  assert.ok(call(g, { action: 'state', token: tok }).students.some(x => x.name === 'Ida'));
  ok('rows added without the edit trigger get an ID on the next refresh');
  st.getRange(3, 7).setValue('absent');
  assert.ok(!call(g, { action: 'state', token: tok }).students.some(x => x.name === 'Paul'));
  ok('typing in Removed takes a student off the page');
  call(g, { action: 'addStudents', token: tok, names: ['10:45 Mira', 'Noor'], studio: 'K3', day: '2026-11-04' });
  const last = tab(book, 'Students').data.slice(-2);
  assert.equal(g.day_(last[0][0]), '2026-11-04'); assert.equal(typeof last[0][0], 'object');
  assert.deepEqual(last.map(r => [r[1], r[2]]), [['10:45', 'Mira'], ['', 'Noor']]);
  ok('students added on the page land in the sheet as real dates, with a time if the line starts with one');
  const mira = call(g, { action: 'state', token: tok }).students.find(x => x.name === 'Mira');
  call(g, { action: 'updateStudent', token: tok, id: mira.id, time: '13:15' });
  assert.equal(call(g, { action: 'state', token: tok }).students.find(x => x.name === 'Mira').time, '13:15');
  ok('a test time changed on the page is saved in the sheet');
  // slots typed as text, as the organiser does: 11:15 - 11:30 (hyphen or dash)
  st.getRange(20, 1, 2, 5).setValues([[pick('2026-11-06'), '11:15 - 11:30', 'Andreas', '', 'K3'], [pick('2026-11-06'), '11:30 – 11:45', 'Connor', '', 'K3']]);
  g.onEdit({ range: st.getRange(20, 1, 2, 5) });
  assert.equal(tab(book, 'Students').data[19][1], '11:15 - 11:30');
  let sl = call(g, { action: 'state', token: tok }).students;
  const andreas = sl.find(x => x.name === 'Andreas'), connor = sl.find(x => x.name === 'Connor');
  assert.deepEqual([andreas.time, andreas.end, connor.time, connor.end], ['11:15', '11:30', '11:30', '11:45']);
  ok('a slot typed as 11:15 - 11:30 stays as typed and reaches the page as start and end');
  call(g, { action: 'updateStudent', token: tok, id: andreas.id, time: '12:00' });
  assert.equal(tab(book, 'Students').data[19][1], '12:00 - 12:15');
  call(g, { action: 'addStudents', token: tok, names: ['10:00-10:15 Efe'], studio: 'K4', day: '2026-11-06' });
  sl = call(g, { action: 'state', token: tok }).students;
  assert.deepEqual([sl.find(x => x.name === 'Efe').time, sl.find(x => x.name === 'Efe').end], ['10:00', '10:15']);
  ok('a slot moved on the page keeps its length, and "10:00-10:15 Efe" adds a slot');
  te.getRange(3, 1).setValue('Isuru');
  call(g, { action: 'addStudents', token: tok, names: ['Lior'], studio: 'K4', day: '2026-11-04', teacher: 'isuru' });
  call(g, { action: 'addStudents', token: tok, names: ['Sam'], studio: 'K4', day: '2026-11-04', teacher: 'Nobody' });
  let ss2 = call(g, { action: 'state', token: tok }).students;
  assert.equal(ss2.find(x => x.name === 'Lior').teacher, 'Isuru'); assert.equal(ss2.find(x => x.name === 'Sam').teacher, '');
  call(g, { action: 'updateStudent', token: tok, id: ss2.find(x => x.name === 'Sam').id, teacher: 'Gunni' });
  assert.equal(call(g, { action: 'state', token: tok }).students.find(x => x.name === 'Sam').teacher, 'Gunni');
  assert.equal(tab(book, 'Students').data.find(r => r[2] === 'Sam')[5], 'Gunni');
  ok('each student can be given a teacher from the Teachers tab, on the page or in the sheet; other names are refused');
}
// 6b. The older Students layout (with Added by / Added at) moves over without losing anyone.
{
  const { g, book } = loadScript();
  const sh = g.SpreadsheetApp.getActive().insertSheet('Students');
  sh.getRange(1, 1, 1, 7).setValues([['Test date', 'Student', 'Studio', 'Added by', 'Added at', 'Removed', 'ID']]);
  sh.getRange(2, 1, 2, 7).setValues([[g.toDate_('2026-11-02'), 'Faddei', 'K3', 'Gunni', new Date(), '', 'abc-1'], [g.toDate_('2026-11-02'), 'Paul', 'K4', 'Sheet', new Date(), 'absent', 'abc-2']]);
  g.setup();
  const d = tab(book, 'Students').data;
  assert.deepEqual(d[0], ['Test date', 'Test time', 'Student', 'CAP', 'Studio', 'Teacher', 'Removed', 'ID']);
  assert.deepEqual(d.slice(1).map(r => r.slice(1)), [['', 'Faddei', '', 'K3', '', '', 'abc-1'], ['', 'Paul', '', 'K4', '', 'absent', 'abc-2']]);
  assert.equal(g.day_(d[1][0]), '2026-11-02');
  ok('the oldest Students layout becomes date | time | name | CAP | studio | teacher | removed | id, keeping IDs');
}
// 6c. The layout with a time but no teacher gets a Teacher column after Studio; everything else stays where it was.
{
  const { g, book } = loadScript();
  const sh = g.SpreadsheetApp.getActive().insertSheet('Students');
  sh.getRange(1, 1, 1, 6).setValues([['Test date', 'Test time', 'Student', 'Studio', 'Removed', 'ID']]);
  sh.getRange(2, 1, 1, 6).setValues([[g.toDate_('2026-11-02'), '09:30', 'Faddei', 'K3', '', 'abc-1']]);
  g.setup();
  const d = tab(book, 'Students').data;
  assert.deepEqual(d[0], ['Test date', 'Test time', 'Student', 'CAP', 'Studio', 'Teacher', 'Removed', 'ID']);
  assert.deepEqual(d[1].slice(1), ['09:30', 'Faddei', '', 'K3', '', '', 'abc-1']);
  ok('the layout without Teacher and CAP gets both where they belong, nothing else moves');
}
// 6d. Today's layout (with Teacher, without CAP) gets a CAP column after Student, keeping every value.
{
  const { g, book } = loadScript();
  const sh = g.SpreadsheetApp.getActive().insertSheet('Students');
  sh.getRange(1, 1, 1, 7).setValues([['Test date', 'Test time', 'Student', 'Studio', 'Teacher', 'Removed', 'ID']]);
  sh.getRange(2, 1, 1, 7).setValues([[g.toDate_('2026-11-02'), '10:00 - 10:15', 'Alice', 'K4', 'Gunni', '', 'abc-1']]);
  g.setup();
  const d = tab(book, 'Students').data;
  assert.deepEqual(d[0], ['Test date', 'Test time', 'Student', 'CAP', 'Studio', 'Teacher', 'Removed', 'ID']);
  assert.deepEqual(d[1].slice(1), ['10:00 - 10:15', 'Alice', '', 'K4', 'Gunni', '', 'abc-1']);
  ok('the current layout gets a CAP column after Student; times, studios and teachers stay put');
}
// 7. The Grades tab: same sum as the page, a row per student, a summary per grade.
{
  const { g, book } = loadScript();
  g.setup();
  const te = g.SpreadsheetApp.getActive().getSheetByName('Teachers');
  te.getRange(2, 1, 2, 2).setValues([['Gunni', 'admin'], ['Anna', 'admin']]); g.setTeamPassword_('test-pass-1');
  const gt = call(g, { action: 'login', name: 'gunni', password: 'test-pass-1' }).token;
  const at = call(g, { action: 'login', name: 'anna', password: 'test-pass-1' }).token;
  call(g, { action: 'addStudents', token: gt, names: ['Polina', 'Paul', 'Kim', 'Lior', 'Ida'], studio: 'K3', day: '2026-11-04' });
  const [pol, paul, kim, lior, ida] = call(g, { action: 'state', token: gt }).students.map(s => s.id);
  const IDS = ['p1','p2','p3','m1','m2','m3','m4','r1','t1','t2','t3','t4','c1','c2','c3','c4','s1','s2'];
  const all = v => Object.fromEntries(IDS.map(id => [id, v(id)]));
  const save = (t, sid, patch) => call(g, { action: 'saveEntry', token: t, sid, patch });
  save(gt, pol, { marks: all(id => id === 'm4' ? 'h' : 'y') }); save(at, pol, { marks: all(() => 'y') });   // Pass everywhere, one Hint
  save(gt, paul, { marks: { p1: 'y', r1: 'n' } });                                                           // unfinished, a core skill failed
  save(gt, kim, { marks: all(() => 'p'), safety: { a: { t: 1, text: 'Speakers switched with level up' } } });   // perfect, but a critical error
  save(at, lior, { marks: all(id => id === 't2' || id === 'c4' ? 'y' : 'p') });                              // nearly all Perfect
  save(gt, ida, { marks: all(() => 'g') });                                                                   // "got there" on everything
  call(g, { action: 'publishResults', token: gt, day: '2026-11-04' });
  const tabs = book.sheets.map(s => s.name);
  assert.ok(tabs.includes('Grades') && !tabs.includes('Results') && !tabs.includes('Days'));
  assert.equal(tabs[1], 'Grades');
  const d = tab(book, 'Grades').data, head = d[6];
  const row = name => d.find(r => r[1] === name), col = (name, c) => row(name)[head.indexOf(c)];
  // Polina: Pass 80 on 17 skills, m4 (Hint 65 + Pass 80)/2 → 79.6 → Pass
  assert.equal(col('Polina', 'Score %'), 80); assert.equal(col('Polina', 'Grade'), 'Pass'); assert.equal(col('Polina', 'Teachers'), 'Gunni, Anna');
  assert.equal(col('Polina', 'Mic'), 78);   // (80+80+80+72.5)/4
  assert.equal(col('Kim', 'Score %'), 100); assert.equal(col('Kim', 'Errors'), 1); assert.equal(col('Kim', 'Grade'), 'Resit');
  assert.equal(col('Lior', 'Score %'), 98); assert.equal(col('Lior', 'Grade'), 'Grade A');
  assert.equal(col('Ida', 'Score %'), 50); assert.equal(col('Ida', 'Grade'), 'Pass');
  assert.equal(col('Paul', 'Grade'), 'Unfinished'); assert.equal(head.indexOf('Core skill under 50%'), -1);
  assert.equal(head.indexOf('In the app'), -1);
  ok('Perfect 100 · Pass 80 · Hint 65 · Got there 50: Grade A from 85, "got there" on everything still passes');
  ok('any critical error is a Resit, even with every skill Perfect');
  assert.deepEqual([d[3][1], d[3][2], d[3][3], d[3][4]], [1, 2, 1, 1]);   // Grade A, Pass, Resit, Unfinished
  assert.equal((d[1] || []).join(''), '');   // row 2 stays empty: no subtitle line
  ok('the summary counts each grade');
  const mk = tab(book, 'Marks').data, mh = mk[0];
  assert.equal(mk.find(r => r[1] === 'Kim')[mh.indexOf('Power-up order')], 'Perfect');
  ok('Marks shows Perfect in words');
  const gr = call(g, { action: 'grades', token: at }).students, gp = gr.find(x => x.name === 'Polina'), gq = gr.find(x => x.name === 'Paul'), gk = gr.find(x => x.name === 'Kim');
  assert.equal(gp.grade, 'Pass'); assert.equal(gp.final, 80); assert.equal(gp.score, 80); assert.equal(gp.errors.length, 0); assert.equal(gp.complete, true);
  assert.equal(gk.grade, 'Resit'); assert.equal(gk.errors.length, 1);
  assert.equal(gp.stages.find(x => x.name === 'Mic').pct, 78); assert.deepEqual(gp.teachers, ['Gunni', 'Anna']);
  assert.equal(gq.complete, false); assert.equal(gq.marked, 2); assert.deepEqual(gq.coreFails, []);
  assert.ok(!JSON.stringify(gr).includes('"marks"'));
  ok('the page\'s Grades screen gets each student\'s combined result, and no teacher\'s single marks');
  const seen = call(g, { action: 'state', token: at }).entries.filter(e => e.sid === pol);
  assert.ok(seen.find(e => e.teacher === 'Anna').data);
  const gun = seen.find(e => e.teacher === 'Gunni'); assert.equal(gun.data, undefined); assert.equal(gun.progress, 18);
  ok('in the app a teacher only ever sees their own marks, and how far the others have got');
}
// 8. The questions come from the Questions tab, and a marked test date keeps its own.
{
  const { g, book } = loadScript();
  g.setup();
  const ss = g.SpreadsheetApp.getActive(), q = ss.getSheetByName('Questions'), te = ss.getSheetByName('Teachers');
  te.getRange(2, 1, 1, 2).setValues([['Gunni', 'admin']]); g.setTeamPassword_('test-pass-1');
  const tok = call(g, { action: 'login', name: 'gunni', password: 'test-pass-1' }).token;
  let st = call(g, { action: 'state', token: tok });
  assert.equal(st.qlive[''].length, 18);
  assert.deepEqual([...new Set(st.qlive[''].map(x => x.stage))], ['Power', 'Mic', 'Route', 'Record', 'Cue', 'Shut down']);
  const m4 = st.qlive[''].find(x => x.id === 'm4');
  assert.equal(m4.heading, 'Connecting a Mic'); assert.match(m4.ask, /condenser/); assert.equal(m4.studios, null);
  ok('setup writes the 18 questions into the Questions tab and the page reads them from there');

  call(g, { action: 'addStudents', token: tok, names: ['Polina'], studio: 'K3', day: '2026-11-04' });
  call(g, { action: 'addStudents', token: tok, names: ['Tales'], studio: 'K4', day: '2026-11-11' });
  const [pol, tal] = call(g, { action: 'state', token: tok }).students.map(s => s.id);
  call(g, { action: 'saveEntry', token: tok, sid: pol, patch: { marks: { p1: 'y' } } });

  // a new question for K4 only, typed under the last row with no stage: it joins "Shut down"
  const last = q.getLastRow() + 1;
  q.getRange(last, 3, 1, 3).setValues([['Patches the headphone line to the K4 booth', '', 'K4']]);
  g.onEdit({ range: q.getRange(last, 3, 1, 3) });
  const newId = q.getRange(last, 7).getValue();
  assert.match(String(newId), /^q[0-9a-f]{8}$/);
  // reword an existing one
  q.getRange(2, 3).setValue('Powers up in the right order, speakers last');
  st = call(g, { action: 'state', token: tok });
  const nq = st.qlive[''].find(x => x.id === newId);
  assert.equal(st.qlive[''].length, 19); assert.equal(nq.stage, 'Shut down'); assert.deepEqual(nq.studios, ['K4']);
  ok('a question typed in the tab gets an ID, its stage from the row above, and its studio');
  assert.equal(st.qsets['2026-11-04|'].length, 18);
  assert.equal(st.qsets['2026-11-04|'][0].t, 'Powers up in the right order, speakers last');
  assert.equal(st.qsets['2026-11-11|'], undefined);
  ok('a marked date keeps its 18 questions (with the new wording); an unmarked date gets the new set');
  call(g, { action: 'tidy', token: tok });   // the page sends this after every refresh
  const head = tab(book, 'Marks').data[0];
  assert.ok(head.includes('Patches the headphone line…') || head.some(h => /^Patches the headphone/.test(h)));
  assert.equal(head.at(-1), 'Page data (do not edit)');
  call(g, { action: 'saveEntry', token: tok, sid: tal, patch: { marks: { [newId]: 'n', p1: 'y' } } });
  const mk = tab(book, 'Marks').data, mh = mk[0], trow = mk.find(r => r[1] === 'Tales'), prow = mk.find(r => r[1] === 'Polina');
  assert.equal(trow[mh.indexOf('Skills marked')], '2/19'); assert.equal(prow[mh.indexOf('Skills marked')], '1/18');
  assert.equal(JSON.parse(prow.at(-1)).marks.p1, 'y');
  ok('Marks gains a column for it, and every row is rewritten without losing a mark');

  // remove a question from the tab: gone from today's set, still counted on the date that was marked with it
  const rows = tab(book, 'Questions').data; const at = rows.findIndex(r => r[6] === 's2') + 1;
  q.deleteRow(at);
  call(g, { action: 'tidy', token: tok });
  st = call(g, { action: 'state', token: tok });
  assert.ok(!st.qlive[''].some(x => x.id === 's2'));
  assert.ok(st.qsets['2026-11-04|'].some(x => x.id === 's2'));
  assert.ok(tab(book, 'Marks').data[0].some(h => /^Left as found/.test(h)));   // still a column: marked dates (and Gunni's own tab) use it
  ok('a removed question disappears from new tests but stays on dates already marked with it');

  // the Grades tab: each date's own set. Tales's date was first marked while there were 19 questions, Polina's with 18
  const IDS18 = ['p1','p2','p3','m1','m2','m3','m4','r1','t1','t2','t3','t4','c1','c2','c3','c4','s1','s2'];
  call(g, { action: 'saveEntry', token: tok, sid: pol, patch: { marks: Object.fromEntries(IDS18.map(id => [id, 'p'])) } });
  assert.equal(st.qsets['2026-11-11|'].length, 19);
  const tq = st.qsets['2026-11-11|'].map(x => x.id);
  call(g, { action: 'saveEntry', token: tok, sid: tal, patch: { marks: Object.fromEntries(tq.map(id => [id, 'p'])) } });
  g.buildGrades_();
  const gd = tab(book, 'Grades').data, gh = gd[6], gr = n => gd.find(r => r[1] === n);
  assert.equal(gr('Polina')[gh.indexOf('Grade')], 'Grade A'); assert.equal(gr('Polina')[gh.indexOf('Marked')], '18/18');
  assert.equal(gr('Tales')[gh.indexOf('Grade')], 'Grade A'); assert.equal(gr('Tales')[gh.indexOf('Marked')], '19/19');
  ok('Grades marks each student out of their own date\'s questions');

  g.useCurrentQuestions_('2026-11-04');
  st = call(g, { action: 'state', token: tok });
  assert.equal(st.qsets['2026-11-04|'].length, 18);   // today's 18; the page and Grades leave out the K4 one for her
  const gd2 = tab(book, 'Grades').data, gh2 = gd2[6];
  assert.equal(gd2.find(r => r[1] === 'Polina')[gh2.indexOf('Marked')], '17/17');
  ok('"Use the current questions for a test date" moves a marked date onto today\'s questions');
}
// 9. Each teacher gets only the students they are testing; an admin gets everyone.
{
  const { g, book } = loadScript();
  g.setup();
  const te = g.SpreadsheetApp.getActive().getSheetByName('Teachers');
  te.getRange(2, 1, 3, 2).setValues([['Gunni', 'admin'], ['Isuru', 'teacher'], ['Ainsley', 'teacher']]); g.setTeamPassword_('test-pass-1');
  const tok = n => call(g, { action: 'login', name: n, password: 'test-pass-1' }).token, gt = tok('Gunni'), it = tok('Isuru'), at = tok('Ainsley');
  call(g, { action: 'addStudents', token: gt, names: ['Kim', 'Lior'], studio: 'K3', day: '2026-11-04', teacher: 'Isuru' });
  call(g, { action: 'addStudents', token: gt, names: ['Noor'], studio: 'K4', day: '2026-11-04', teacher: 'Ainsley' });
  call(g, { action: 'addStudents', token: gt, names: ['Sam'], studio: 'K4', day: '2026-11-04' });
  const names = t => call(g, { action: 'state', token: t }).students.map(x => x.name).sort().join(',');
  assert.equal(names(it), 'Kim,Lior'); assert.equal(names(at), 'Noor'); assert.equal(names(gt), 'Kim,Lior,Noor,Sam');
  ok('a teacher only gets the students with their name in the Teacher column; the admin gets everyone');
  assert.deepEqual(call(g, { action: 'grades', token: it }).students.map(x => x.name).sort(), ['Kim', 'Lior']);
  ok('the Grades screen shows a teacher only their own students');
  const noor = call(g, { action: 'state', token: gt }).students.find(x => x.name === 'Noor').id;
  assert.match(call(g, { action: 'saveEntry', token: it, sid: noor, patch: { marks: { p1: 'y' } } }).error, /not on your list/);
  assert.match(call(g, { action: 'updateStudent', token: it, id: noor, time: '10:00' }).error, /not on your list/);
  const kim = call(g, { action: 'state', token: it }).students.find(x => x.name === 'Kim').id;
  assert.match(call(g, { action: 'updateStudent', token: it, id: kim, teacher: 'Ainsley' }).error, /Only an admin/);
  ok('a teacher cannot mark, move or hand over another teacher\'s student');
  call(g, { action: 'addStudents', token: it, names: ['Pia'], studio: 'K3', day: '2026-11-04', teacher: 'Ainsley' });
  assert.equal(names(it), 'Kim,Lior,Pia');
  ok('a student a teacher adds on the page goes on their own list');
  call(g, { action: 'updateStudent', token: gt, id: noor, teacher: 'Isuru' });
  assert.equal(names(it), 'Kim,Lior,Noor,Pia'); assert.equal(names(at), '');
  ok('when the admin moves a student to another teacher, it moves between their lists');
}
// 10. Every teacher signs in with the one password, whatever the organiser tried first.
{
  const { g } = loadScript();
  g.setup();
  const te = g.SpreadsheetApp.getActive().getSheetByName('Teachers');
  te.getRange(2, 1, 3, 2).setValues([['Gunni', 'admin'], ['Isuru', 'teacher'], ['Ainsley', 'teacher']]); g.setTeamPassword_('team-pass-1');
  const login = (n, p) => call(g, { action: 'login', name: n, password: p });
  // the organiser tries each name on their own phone first, then signs out
  ['Gunni', 'Isuru', 'Ainsley'].forEach(n => { const t = login(n, 'team-pass-1').token; assert.ok(t); call(g, { action: 'logout', token: t }); });
  const phones = ['Gunni', 'Isuru', 'Ainsley'].map(n => login(n, 'team-pass-1'));
  assert.ok(phones.every(r => r.ok));
  ok('every teacher signs in with the same password, also after the organiser tried their name first');
  const isuru2 = login('isuru', 'team-pass-1').token;
  call(g, { action: 'logout', token: isuru2 });
  assert.equal(call(g, { action: 'state', token: phones[1].token }).ok, true);
  ok('the same name signed in on two phones: signing out on one leaves the other signed in');
  for (let i = 0; i < 8; i++) login('Ainsley', 'typo');
  assert.match(login('Ainsley', 'team-pass-1').error, /Too many wrong tries/);
  assert.equal(login('Isuru', 'team-pass-1').ok, true);
  ok('8 wrong tries pause only that name, nobody else');
  g.setTeamPassword_('team-pass-2');
  assert.equal(login('Ainsley', 'team-pass-2').ok, true);
  ok('typing the team password again on Start here clears a pause at once');
}
// 11. Each teacher has their own questions tab; their students are asked those questions.
{
  const { g, book } = loadScript();
  g.setup();
  const ss = g.SpreadsheetApp.getActive(), te = ss.getSheetByName('Teachers');
  te.getRange(2, 1, 3, 2).setValues([['Gunni', 'admin'], ['Isuru', 'teacher'], ['Ainsley', 'teacher']]);
  g.onEdit({ range: te.getRange(2, 1, 3, 2) });
  g.setTeamPassword_('test-pass-1');
  const names = book.sheets.map(x => x.name);
  assert.ok(['Questions · Gunni', 'Questions · Isuru', 'Questions · Ainsley'].every(n => names.includes(n)));
  assert.equal(tab(book, 'Questions · Isuru').data.length, 19);   // header + the 18 shared questions
  ok('every teacher gets their own questions tab, a copy of the shared one');
  const tok = n => call(g, { action: 'login', name: n, password: 'test-pass-1' }).token, gt = tok('Gunni'), it = tok('Isuru'), at = tok('Ainsley');
  call(g, { action: 'addStudents', token: gt, names: ['Kim'], studio: 'K3', day: '2026-11-04', teacher: 'Isuru' });
  call(g, { action: 'addStudents', token: gt, names: ['Noor'], studio: 'K3', day: '2026-11-04', teacher: 'Ainsley' });
  // Isuru adds a question of his own and rewords one; Ainsley's tab is untouched
  const iq = ss.getSheetByName('Questions · Isuru'), last = iq.getLastRow() + 1;
  iq.getRange(last, 1, 1, 5).setValues([['Shut down', '', 'Coils the cables over-under', '', 'Both']]);
  g.onEdit({ range: iq.getRange(last, 1, 1, 5) });
  iq.getRange(2, 3).setValue('Powers up in Isuru\'s order');
  let si = call(g, { action: 'state', token: it }), sa = call(g, { action: 'state', token: at });
  assert.equal(si.questions.length, 19); assert.equal(si.questions[0].t, 'Powers up in Isuru\'s order');
  assert.equal(sa.questions.length, 18); assert.equal(sa.questions[0].t, 'Powers up in the right order');
  assert.equal(si.qlive['isuru'].length, 19); assert.equal(sa.qlive['ainsley'].length, 18);
  ok('a teacher\'s own questions reach only their phone; the other teacher keeps theirs');
  const kim = si.students[0].id, noor = sa.students[0].id;
  call(g, { action: 'saveEntry', token: it, sid: kim, patch: { marks: { p1: 'p' } } });
  call(g, { action: 'tidy', token: gt });
  const mh = tab(book, 'Marks').data[0];
  assert.ok(mh.includes('Coils the cables over-under'));
  const kimRow = tab(book, 'Marks').data.find(r => r[1] === 'Kim');
  assert.equal(kimRow[mh.indexOf('Coils the cables over-under')], '');   // in Kim's test, not marked yet
  const noorRow = g.marksRow_({ id: noor, day: '2026-11-04', name: 'Noor', studio: 'K3', teacher: 'Ainsley' }, 'Ainsley', { marks: { p1: 'y' } });
  assert.equal(noorRow[mh.indexOf('Coils the cables over-under')], '–');   // not in Ainsley's test: a dash
  assert.equal(noorRow[mh.indexOf('Power-up order')], 'Pass');
  ok('a teacher\'s own question gets its own Marks column, a dash for other teachers\' students');
  const gr = call(g, { action: 'grades', token: gt }).students;
  assert.equal(gr.find(x => x.name === 'Kim').total, 19); assert.equal(gr.find(x => x.name === 'Noor').total, 18);
  ok('Marks gets a column for a teacher\'s own question, and each student is graded on their teacher\'s set');
  // Kim is marked, so Isuru's set is fixed for that date; Noor (unmarked) still follows Ainsley's tab
  iq.deleteRow(last);
  const aq = ss.getSheetByName('Questions · Ainsley');
  aq.getRange(aq.getLastRow() + 1, 3, 1, 1).setValues([['Labels the patch bay']]); g.onEdit({ range: aq.getRange(aq.getLastRow(), 3) });
  si = call(g, { action: 'state', token: it }); sa = call(g, { action: 'state', token: at });
  assert.equal(si.qsets['2026-11-04|isuru'].length, 19); assert.equal(si.qlive['isuru'].length, 18);
  assert.equal(sa.qsets['2026-11-04|ainsley'], undefined); assert.equal(sa.qlive['ainsley'].length, 19);
  ok('a marked date keeps its teacher\'s questions; an unmarked one follows the teacher\'s tab');
  // an older date fixed for everyone (before teachers had their own tabs) still applies to every teacher
  g.PropertiesService.getScriptProperties().setProperty('q_2026-11-02', JSON.stringify([{ id: 'p1', core: 1, studios: null }]));
  assert.equal(g.questionsFor_('2026-11-02', 'Ainsley').length, 1); assert.equal(g.questionsFor_('2026-11-02', 'Isuru').length, 1);
  ok('dates fixed before this change keep their questions for every teacher');
  te.getRange(5, 1).setValue('Mila'); g.onEdit({ range: te.getRange(5, 1) });
  assert.ok(book.sheets.some(x => x.name === 'Questions · Mila'));
  ok('a teacher typed into the Teachers tab gets their own questions tab at once');
}
// 12. The admin PIN: admins give it after the team password; teachers are not affected.
{
  const { g, book } = loadScript();
  g.setup();
  const te = g.SpreadsheetApp.getActive().getSheetByName('Teachers');
  te.getRange(2, 1, 2, 2).setValues([['Gunni', 'admin'], ['Isuru', 'teacher']]); g.setTeamPassword_('team-pass-1');
  const login = (n, p, pin) => call(g, { action: 'login', name: n, password: p, pin });
  const oldAdmin = login('Gunni', 'team-pass-1').token, oldTeacher = login('Isuru', 'team-pass-1').token;
  assert.ok(oldAdmin && oldTeacher);
  ok('before a PIN is set, admins sign in as before');
  const cell = g.pinCell_(); cell.setValue('4711'); g.onEdit({ range: cell });
  assert.equal(cell.getValue(), '●●●●●● set'); assert.ok(!JSON.stringify(tab(book, 'Start here').data).includes('4711'));
  ok('the admin PIN typed on Start here is hidden at once');
  assert.equal(call(g, { action: 'state', token: oldTeacher }).ok, true);
  assert.equal(call(g, { action: 'state', token: oldAdmin }).code, 'auth');
  ok('setting the PIN signs admins out once; teachers stay signed in');
  assert.equal(login('Isuru', 'team-pass-1').ok, true);
  let r = login('Gunni', 'team-pass-1'); assert.equal(r.ok, false); assert.equal(r.needPin, true);
  r = login('Gunni', 'team-pass-1', '0000'); assert.equal(r.ok, false); assert.match(r.error, /PIN is not right/);
  assert.equal(login('Gunni', 'wrong-pass', '4711').ok, false);
  r = login('Gunni', 'team-pass-1', '4711'); assert.equal(r.ok, true); assert.equal(r.role, 'admin');
  assert.equal(call(g, { action: 'state', token: r.token }).me.role, 'admin');
  ok('an admin needs the team password and the PIN; a teacher still needs only the team password');
  // a teacher's session does not become an admin's when someone changes the role in the sheet
  const it = login('Isuru', 'team-pass-1').token;
  te.getRange(3, 2).setValue('admin');
  assert.equal(call(g, { action: 'state', token: it }).me.role, 'teacher');
  ok('admin rights need a session that gave the PIN, even if a role is changed in the sheet');
}
// 13. Sign-ins last 30 days, older 400-day ones included.
{
  const { g } = loadScript();
  g.setup();
  g.SpreadsheetApp.getActive().getSheetByName('Teachers').getRange(2, 1).setValue('Isuru'); g.setTeamPassword_('team-pass-1');
  const r = call(g, { action: 'login', name: 'Isuru', password: 'team-pass-1' });
  assert.equal(r.days, 30); assert.equal(call(g, { action: 'state', token: r.token }).ok, true);
  const props = g.PropertiesService.getScriptProperties(), key = 's_' + r.token, sess = JSON.parse(props.getProperty(key));
  props.setProperty(key, JSON.stringify({ ...sess, start: Date.now() - 31 * 864e5 }));
  assert.equal(call(g, { action: 'state', token: r.token }).code, 'auth');
  ok('a sign-in ends after 30 days');
  const old = (Date.now() + 400 * 864e5) - 10 * 864e5, recent = 'a'.repeat(64), stale = 'b'.repeat(64);
  props.setProperty('s_' + recent, JSON.stringify({ name: 'Isuru', exp: old }));                       // made 10 days ago for 400 days
  props.setProperty('s_' + stale, JSON.stringify({ name: 'Isuru', exp: old - 40 * 864e5 }));           // made 50 days ago
  assert.equal(call(g, { action: 'state', token: recent }).ok, true); assert.equal(call(g, { action: 'state', token: stale }).code, 'auth');
  ok('older 400-day sign-ins now also end 30 days after they began');
}
// 14. Isuru's case: a date fixed for everyone before teachers had their own tabs, then a teacher changes their tab.
{
  const { g, book } = loadScript();
  g.setup();
  const ss = g.SpreadsheetApp.getActive(), te = ss.getSheetByName('Teachers');
  te.getRange(2, 1, 3, 2).setValues([['Gunni', 'admin'], ['Isuru', 'teacher'], ['Ainsley', 'teacher']]); g.onEdit({ range: te.getRange(2, 1, 3, 2) });
  g.setTeamPassword_('team-pass-1');
  const gt = call(g, { action: 'login', name: 'Gunni', password: 'team-pass-1' }).token;
  call(g, { action: 'addStudents', token: gt, names: ['Emiliano'], studio: 'K3', day: '2026-11-02', teacher: 'Isuru' });
  call(g, { action: 'addStudents', token: gt, names: ['Faddei'], studio: 'K4', day: '2026-11-02', teacher: 'Ainsley' });
  // the old, day-wide fixed set: the 18 shared questions
  g.PropertiesService.getScriptProperties().setProperty('q_2026-11-02', JSON.stringify(g.questions_()));
  // Isuru writes his own questions over the rows (IDs stay) and adds a Shut down question at the end
  const iq = ss.getSheetByName('Questions · Isuru'), rows = tab(book, 'Questions · Isuru').data;
  const s1 = rows.findIndex(r => r[6] === 's1') + 1;
  iq.getRange(s1, 1, 2, 1).setValues([[''], ['']]);                   // s1 and s2 now fall under the stage above (Cue)
  const last = iq.getLastRow() + 1;
  iq.getRange(last, 1, 1, 3).setValues([['Shut down', 'Shut down & pack down', 'Demonstrated correct shutdown & reset procedures']]);
  g.onEdit({ range: iq.getRange(last, 1, 1, 3) });
  const before = g.questionsFor_('2026-11-02', 'Isuru');
  assert.ok(!before.some(q => q.stage === 'Shut down'));
  g.setup();   // the update runs setup
  const isu = g.questionsFor_('2026-11-02', 'Isuru'), ain = g.questionsFor_('2026-11-02', 'Ainsley');
  assert.ok(isu.some(q => q.stage === 'Shut down' && /shutdown & reset/.test(q.t))); assert.equal(isu.length, 19);
  assert.equal(ain.length, 18);
  assert.equal(g.PropertiesService.getScriptProperties().getProperty('q_2026-11-02'), null);
  ok('an old day-wide set follows a teacher\'s changed tab (Isuru gets his Shut down question); unchanged teachers keep theirs');
  // the Grades tab: one band per teacher, then their students
  g.buildGrades_();
  const d = tab(book, 'Grades').data.slice(7).map(r => String(r[0]).startsWith('Isuru') || String(r[0]).startsWith('Ainsley') ? 'BAND ' + r[0].split('  ')[0] : r[1]);
  assert.deepEqual(d.filter(Boolean), ['BAND Isuru', 'Emiliano', 'BAND Ainsley', 'Faddei']);
  ok('the Grades tab lists each teacher\'s students together under a band with the teacher\'s name');
}
// 15. The pass bar is 40: mostly "got there" with a few failed (non-core) skills still passes; below 40 is a Resit.
{
  const { g } = loadScript();
  g.setup();
  const qs = g.questions_(), nonCore = qs.filter(q => !q.core).map(q => q.id);
  const marks = f => ({ T: { marks: Object.fromEntries(qs.map(q => [q.id, f(q)])) } });
  const three = nonCore.slice(0, 3), five = nonCore.slice(0, 5);
  const a = g.gradeOf_(marks(q => three.includes(q.id) ? 'n' : 'g'), qs), b = g.gradeOf_(marks(q => five.includes(q.id) ? 'n' : 'g'), qs);
  assert.equal(Math.round(a.score), 42); assert.equal(a.band[0], 'Pass');
  assert.equal(Math.round(b.score), 36); assert.equal(b.band[0], 'Resit');
  ok('Pass from 40: "got there" with three failed skills passes (42), with five it is a Resit (36)');
}
// 16. A Marks row whose Teacher cell was cleared by hand does not count towards a grade.
{
  const { g, book } = loadScript();
  g.setup();
  const te = g.SpreadsheetApp.getActive().getSheetByName('Teachers');
  te.getRange(2, 1, 1, 2).setValues([['Ainsley', 'teacher']]); g.setTeamPassword_('team-pass-1');
  const at = call(g, { action: 'login', name: 'Ainsley', password: 'team-pass-1' }).token;
  call(g, { action: 'addStudents', token: at, names: ['Ayham'], studio: 'K3', day: '2026-11-03' });
  const sid = call(g, { action: 'state', token: at }).students[0].id, all = v => Object.fromEntries(g.questions_().map(q => [q.id, v]));
  call(g, { action: 'saveEntry', token: at, sid, patch: { marks: all('n') } });
  const mk = g.SpreadsheetApp.getActive().getSheetByName('Marks');
  mk.getRange(2, 2, 1, 2).setValues([['', '']]);   // someone clears Student and Teacher on that row
  call(g, { action: 'saveEntry', token: at, sid, patch: { marks: all('p') } });   // Ainsley marks again: a new row
  const gr = call(g, { action: 'grades', token: at }).students[0];
  assert.equal(gr.score, 100); assert.deepEqual(gr.teachers, ['Ainsley']);
  ok('a Marks row with its Teacher cell cleared is ignored: only the teacher\'s own row counts');
}
// 17. A questions tab with a column deleted (Gunni's case) is read by its headings and put back in order by setup.
{
  const { g, book } = loadScript();
  g.setup();
  const ss = g.SpreadsheetApp.getActive(), te = ss.getSheetByName('Teachers');
  te.getRange(2, 1, 1, 2).setValues([['Gunni', 'admin']]); g.onEdit({ range: te.getRange(2, 1, 1, 2) });
  // Gunni deleted "Ask out loud": Stage | Stage heading | Question | Core | Studio | Column name | (empty) | ID
  const gq = ss.getSheetByName('Questions · Gunni');
  gq.clear();
  gq.getRange(1, 1, 3, 8).setValues([
    ['Stage', 'Stage heading', 'Question', 'Core', 'Studio', 'Column name', '', 'ID'],
    ['Power', 'Power up', 'Power in the right order', true, 'Both', 'Turning on', '', 'p1'],
    ['', '', 'Mute before console connections', false, 'K3', 'CR mute', '', 'p2']]);
  let qs = g.questions_('Gunni');
  assert.deepEqual(JSON.parse(JSON.stringify(qs.map(q => [q.id, q.t, q.ask, q.label]))), [['p1', 'Power in the right order', '', 'Turning on'], ['p2', 'Mute before console connections', '', 'CR mute']]);
  assert.deepEqual(JSON.parse(JSON.stringify(qs[1].studios)), ['K3']);
  ok('a tab with a deleted column is read by its headings: an empty Ask stays empty (no "true")');
  g.setup();
  assert.deepEqual(tab(book, 'Questions · Gunni').data[0], ['Stage', 'Stage heading', 'Question', 'Ask out loud (optional)', 'Studio', 'Column name', 'ID']);
  assert.deepEqual(tab(book, 'Questions · Gunni').data[2], ['Power', '', 'Mute before console connections', '', 'K3', 'CR mute', 'p2']);
  ok('setup puts the tab back in the standard columns, keeping every question, studio, column name and ID (Core is gone)');
}
// 18. Moving, replacing and rewording questions: marks stay with their question.
{
  const { g, book } = loadScript();
  g.setup();
  const q = g.SpreadsheetApp.getActive().getSheetByName('Questions');
  g.ensureQuestionIds_();   // the tab remembers which text each ID has
  const textOf = id => g.questions_().find(x => x.id === id)?.t;
  const t1 = textOf('p1'), t2 = textOf('p2');
  // retype p1's and p2's texts the other way round (the hidden IDs stay in their rows)
  q.getRange(2, 3).setValue(t2); q.getRange(3, 3).setValue(t1);
  g.onEdit({ range: q.getRange(2, 3, 2, 1) });
  assert.equal(textOf('p1'), t1); assert.equal(textOf('p2'), t2);
  assert.deepEqual(JSON.parse(JSON.stringify(g.questions_().slice(0, 2).map(x => x.id))), ['p2', 'p1']);
  ok('questions swapped by retyping keep their own IDs, so their marks follow them');
  q.getRange(4, 3).setValue('Labels every cable on the patch bay'); g.onEdit({ range: q.getRange(4, 3) });
  assert.equal(textOf('p3'), undefined); assert.match(g.questions_()[2].id, /^q[0-9a-f]{8}$/);
  ok('a question replaced by a different one gets a new ID (old marks do not slide onto it)');
  q.getRange(5, 3).setValue('Chooses the correct mic and places it well'); g.onEdit({ range: q.getRange(5, 3) });
  assert.equal(textOf('m1'), 'Chooses the correct mic and places it well');
  ok('a small rewording keeps the ID');
}
// 19. Resetting from the Sheet: one student, or everything.
{
  const { g, book } = loadScript();
  g.setup();
  const te = g.SpreadsheetApp.getActive().getSheetByName('Teachers');
  te.getRange(2, 1, 1, 2).setValues([['Gunni', 'admin']]); g.setTeamPassword_('team-pass-1');
  const tok = call(g, { action: 'login', name: 'Gunni', password: 'team-pass-1' }).token;
  call(g, { action: 'addStudents', token: tok, names: ['Kim', 'Lior'], studio: 'K3', day: '2026-11-02' });
  call(g, { action: 'addStudents', token: tok, names: ['Kim'], studio: 'K3', day: '2026-11-03' });
  const st = call(g, { action: 'state', token: tok }).students;
  st.forEach(x => call(g, { action: 'saveEntry', token: tok, sid: x.id, patch: { marks: { p1: 'y' } } }));
  assert.match(g.resetStudent_('Kim').error, /2 students with that name/);
  const r = g.resetStudent_('Kim 02.11.2026');
  assert.equal(r.ok, true); assert.equal(r.rows, 1);
  assert.equal(tab(book, 'Marks').data.filter(x => x[1] === 'Kim').length, 1);
  assert.equal(call(g, { action: 'state', token: tok }).students.length, 3);
  ok('"Reset a student\'s test" removes that student\'s marks (a date picks between two with one name); students stay');
  g.resetAll_();
  assert.equal(tab(book, 'Marks').data.filter(x => x[1]).length, 1);   // only the header
  assert.equal(g.frozenKeys_().length, 0);
  ok('"Reset ALL test results" removes every mark and unfixes every test date');
  // the same from the two cells on Start here (the menu hides on narrow windows)
  g.setup();
  call(g, { action: 'saveEntry', token: tok, sid: st[1].id, patch: { marks: { p1: 'y' } } });
  call(g, { action: 'saveEntry', token: tok, sid: st[0].id, patch: { marks: { p1: 'y' } } });
  const one = g.startCell_('Reset a student'), all = g.startCell_('Reset everything');
  one.setValue('Lior'); g.onEdit({ range: one });
  assert.match(String(one.getValue()), /^✓ Lior reset: 1 marking removed/);
  assert.equal(tab(book, 'Marks').data.filter(x => x[1] === 'Lior').length, 0);
  one.setValue('Nobody'); g.onEdit({ range: one });
  assert.match(String(one.getValue()), /^✗ No student called "Nobody"/);
  all.setValue('reset'); g.onEdit({ range: all });
  assert.match(String(all.getValue()), /^✗ Nothing was reset/);
  assert.equal(tab(book, 'Marks').data.filter(x => x[1]).length, 2);
  all.setValue('RESET'); g.onEdit({ range: all });
  assert.match(String(all.getValue()), /^✓ Everything reset: 1 marking removed/);
  assert.equal(tab(book, 'Marks').data.filter(x => x[1]).length, 1);
  g.onEdit({ range: all });   // editing the message again does nothing
  ok('typing a name, or RESET, in the Start here reset cells resets and says so in the cell');
}
// 21. Marks columns go by name: the same name in two teachers' tabs is one column; a shared-only question has none.
{
  const { g, book } = loadScript();
  g.setup();
  const ss = g.SpreadsheetApp.getActive(), te = ss.getSheetByName('Teachers');
  te.getRange(2, 1, 2, 2).setValues([['Gunni', 'admin'], ['Isuru', 'teacher']]); g.onEdit({ range: te.getRange(2, 1, 2, 2) });
  g.setTeamPassword_('team-pass-1');
  const gt = call(g, { action: 'login', name: 'Gunni', password: 'team-pass-1', pin: '' }).token;
  call(g, { action: 'addStudents', token: gt, names: ['Kim'], studio: 'K3', day: '2026-11-04' });
  ss.getSheetByName('Students').getRange(2, 6).setValue('Isuru');   // every student has a teacher: the shared tab is not asked
  const gq = ss.getSheetByName('Questions · Gunni'), iq = ss.getSheetByName('Questions · Isuru');
  const lastG = gq.getLastRow(), lastI = iq.getLastRow();
  gq.getRange(lastG, 3).setValue('Leaves the studio as it was found'); gq.getRange(lastG, 6).setValue('Leaves the studio');
  iq.getRange(lastI, 3).setValue('Puts every cable back where it belongs'); iq.getRange(lastI, 6).setValue('Leaves the studio');
  g.onEdit({ range: gq.getRange(lastG, 3) }); g.onEdit({ range: iq.getRange(lastI, 3) });
  const sh = ss.getSheetByName('Questions'); sh.getRange(sh.getLastRow(), 6).setValue('Shared only name');
  call(g, { action: 'tidy', token: gt });
  const mh = tab(book, 'Marks').data[0];
  assert.equal(mh.filter(h => h === 'Leaves the studio').length, 1);
  assert.ok(!mh.some(h => /·/.test(h)));
  assert.ok(!mh.includes('Shared only name'));
  assert.ok(mh.includes('%') && !mh.includes('Teacher score %'));
  ok('Marks: one column per name across teachers, no teacher tags, no column for a question nobody asks, "%" header');
}
// 22. Teachers editing their tab: moving, renaming and adding stages; Column names stay with their question.
{
  const { g, book } = loadScript();
  g.setup();
  const ss = g.SpreadsheetApp.getActive(), te = ss.getSheetByName('Teachers');
  te.getRange(2, 1, 1, 2).setValues([['Gunni', 'admin']]); g.onEdit({ range: te.getRange(2, 1, 1, 2) });
  const q = ss.getSheetByName('Questions · Gunni'), rows = () => tab(book, 'Questions · Gunni').data;
  const stageOf = id => g.questions_('Gunni').find(x => x.id === id).stage;
  const rowOf = id => rows().findIndex(r => r[6] === id) + 1;
  assert.ok(rows().slice(1).filter(r => r[2]).every(r => r[0]), 'every question row shows its stage after setup');
  // move the FIRST row of the Mic stage to the very end: the other Mic questions stay Mic, the moved one too
  const micFirst = rows().findIndex((r, i) => i > 0 && r[0] === 'Mic') + 1, moved = rows()[micFirst - 1].slice();
  const micRest = rows().filter(r => r[0] === 'Mic').map(r => r[6]).filter(id => id !== moved[6]);
  q.deleteRow(micFirst); const end = q.getLastRow() + 1;
  q.getRange(end, 1, 1, 7).setValues([moved]); g.onEdit({ range: q.getRange(end, 1, 1, 7) });
  assert.ok(micRest.length && micRest.every(id => stageOf(id) === 'Mic'));
  assert.equal(stageOf(moved[6]), 'Mic');
  ok('moving the first row of a stage keeps every question in its stage');
  // the stages are one list on Start here: renaming one there renames it in every teacher's tab
  assert.deepEqual(JSON.parse(JSON.stringify(g.stages_())), ['Power', 'Mic', 'Route', 'Record', 'Cue', 'Shut down']);
  const routeIds = rows().filter(r => r[0] === 'Record').map(r => r[6]);
  const start = ss.getSheetByName('Start here'), sa = tab(book, 'Start here').data.findIndex(r => r[0] === 'Test stages') + 1;
  const recCell = start.getRange(sa + 3, 2);
  assert.equal(recCell.getValue(), 'Record');
  recCell.setValue('Recording'); g.onEdit({ range: recCell, oldValue: 'Record' });
  assert.ok(routeIds.length === 4 && routeIds.every(id => stageOf(id) === 'Recording'));
  assert.ok(!tab(book, 'Questions').data.some(r => r[0] === 'Record') && tab(book, 'Questions').data.some(r => r[0] === 'Recording'));
  assert.equal(g.stages_()[3], 'Recording');
  // picking another stage for one question moves just that question
  const second = rowOf(routeIds[1]);
  q.getRange(second, 1).setValue('Cue'); g.onEdit({ range: q.getRange(second, 1), oldValue: 'Recording' });
  assert.equal(stageOf(routeIds[1]), 'Cue'); assert.equal(stageOf(routeIds[2]), 'Recording');
  ok('renaming a stage on Start here renames it in every teacher\'s tab; picking a stage moves one question');
  // a stage typed with other capitals or spaces still counts as the listed stage
  q.getRange(rowOf(routeIds[2]), 1).setValue('  recording '); g.onEdit({ range: q.getRange(rowOf(routeIds[2]), 1) });
  assert.equal(stageOf(routeIds[2]), 'Recording');
  g.buildGrades_();
  const gh = tab(book, 'Grades').data[6];
  assert.equal(gh.filter(h => /^recording$/i.test(String(h).trim())).length, 1);
  assert.deepEqual(JSON.parse(JSON.stringify(gh.slice(6, 12))), ['Power', 'Mic', 'Route', 'Recording', 'Cue', 'Shut down']);
  ok('Grades has one column per listed stage, in list order, whatever capitals or spaces a teacher typed');
  // a new row left without a stage takes the stage above
  const after = rowOf(routeIds[2]);
  q.insertRowAfter(after); q.getRange(after + 1, 3).setValue('Checks the talkback'); g.onEdit({ range: q.getRange(after + 1, 3) });
  assert.equal(rows()[after][0], 'Recording');
  ok('a new question row takes the stage of the row above, written into its own cell');
  // typing a different question over a row clears its old Column name
  const r1 = rowOf('p1');
  assert.ok(rows()[r1 - 1][5]);
  q.getRange(r1, 3).setValue('Labels every cable on the patch bay'); g.onEdit({ range: q.getRange(r1, 3) });
  assert.equal(rows()[r1 - 1][5], '');
  assert.equal(g.questions_('Gunni').find(x => x.t === 'Labels every cable on the patch bay').label, 'Labels every cable on the…');
  ok('a question typed over another loses the old Column name (it shows its own words until named)');
}
// 23. Landmark saving: refreshes reuse a kept copy, a whole test is sent in one call, End test and grade is stored.
{
  const { g, book } = loadScript();
  g.setup();
  const ss = g.SpreadsheetApp.getActive(), te = ss.getSheetByName('Teachers');
  te.getRange(2, 1, 2, 2).setValues([['Gunni', 'admin'], ['Isuru', 'teacher']]); g.onEdit({ range: te.getRange(2, 1, 2, 2) });
  g.setTeamPassword_('team-pass-1');
  const gt = raw(g, { action: 'login', name: 'Gunni', password: 'team-pass-1', pin: '' }).token;
  const ids = raw(g, { action: 'addStudents', token: gt, names: ['Kim', 'Lou'], studio: 'K3', day: '2026-11-04' }).ids;
  const s1 = raw(g, { action: 'state', token: gt });
  assert.ok(s1.h && s1.students.length === 2);
  assert.deepEqual(raw(g, { action: 'state', token: gt, h: s1.h }), { ok: true, same: true, h: s1.h });
  ok('a refresh with nothing new answers just "same" (the page keeps its copy)');
  const q = s1.questions.map(x => x.id);
  const b = raw(g, { action: 'batch', token: gt, steps: [
    ['saveEntry', { sid: ids[0], patch: { marks: Object.fromEntries(q.map(id => [id, 'y'])), summary: 'Clean run', ended: 1793444955000 } }],
    ['saveEntry', { sid: ids[1], patch: { marks: { [q[0]]: 'h' } } }],
    ['saveEntry', { sid: 'nobody', patch: { marks: { [q[0]]: 'y' } } }]] });
  assert.deepEqual(b.results.map(r => r.ok), [true, true, false]);
  ok('one batch call saves two students and refuses an unknown one on its own');
  const s2 = raw(g, { action: 'state', token: gt, h: s1.h });
  assert.ok(!s2.same && s2.h !== s1.h);
  const e = s2.entries.find(x => x.sid === ids[0]).data;
  assert.equal(e.ended, 1793444955000); assert.equal(e.summary, 'Clean run');
  assert.equal(s2.entries.find(x => x.sid === ids[1]).data.ended, undefined);
  ok('after a save the next refresh is new, and the ended test says so');
  const m = tab(book, 'Marks').data, h = m[0];
  const kim = m.find(r => r[1] === 'Kim');
  assert.equal(kim[h.indexOf('Skills marked')], q.length + '/' + q.length);
  assert.equal(kim[h.indexOf('Feedback')], 'Clean run');
  ok('the ended test is in Marks with every skill and the feedback');
  raw(g, { action: 'saveEntry', token: gt, sid: ids[0], patch: { ended: null } });
  assert.equal(raw(g, { action: 'state', token: gt }).entries.find(x => x.sid === ids[0]).data.ended, undefined);
  ok('Reset takes the ended mark off again');
}

console.log(`\n${n} checks passed`);
