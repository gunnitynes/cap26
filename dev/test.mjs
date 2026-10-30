// Checks the sign-in paths of Code.gs on this computer:  node dev/test.mjs
import assert from 'node:assert/strict';
import { loadScript } from './google-mock.mjs';

const call = (g, body) => JSON.parse(g.doPost({ postData: { contents: JSON.stringify(body) } }).s);
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
  assert.equal(r.ok, true); assert.ok(r.days > 300);
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
  assert.deepEqual(JSON.parse(row[head.length - 1]).marks, { p1: 'y', m4: 'h', c2: 'n' });
  ok('Marks shows every skill in words; unknown values are dropped');
  const st = call(g, { action: 'state', token: gt });
  assert.equal(st.entries[0].data.marks.m4, 'h');
  ok('the page reads its marks back from the sheet');
}
// 6. Students typed or pasted straight into the Students tab: date | time | name | studio | teacher | removed | id.
{
  const { g, book } = loadScript();
  g.setup();
  const ss = g.SpreadsheetApp.getActive(), te = ss.getSheetByName('Teachers'), st = ss.getSheetByName('Students');
  assert.deepEqual(tab(book, 'Students').data[0], ['Test date', 'Test time', 'Student', 'Studio', 'Teacher', 'Removed', 'ID']);
  te.getRange(2, 1, 1, 2).setValues([['Gunni', 'admin']]); g.setTeamPassword_('test-pass-1');
  const tok = call(g, { action: 'login', name: 'gunni', password: 'test-pass-1' }).token;
  // dates must come from inside the sandbox, like the ones Google hands the script
  const pick = d => { const x = g.toDate_(d); x.setHours(0); return x; };   // what the date picker stores: midnight
  // 9.30 typed in Sheets arrives as the number 9.3
  st.getRange(2, 1, 1, 4).setValues([[pick('2026-11-04'), 9.3, 'Polina', 'K3']]);
  g.onEdit({ range: st.getRange(2, 1, 1, 4) });
  st.getRange(3, 2, 3, 2).setValues([['10:30', 'Paul'], [1100, 'Faddei'], ['', '  Yutaro ']]);
  g.onEdit({ range: st.getRange(3, 2, 3, 2) });
  st.getRange(6, 1, 1, 4).setValues([['5.11.2026', '8:05', 'Tales', 'K4']]);
  g.onEdit({ range: st.getRange(6, 1, 1, 4) });
  const rows = tab(book, 'Students').data.slice(1);
  assert.ok(rows.every(r => r[6]), 'every named row gets an ID');
  assert.deepEqual(rows.map(r => r[3]), ['K3', 'K3', 'K3', 'K3', 'K4']);
  ok('a pasted list takes the date and studio from the row above and gets IDs');
  assert.deepEqual(rows.map(r => r[1]), ['09:30', '10:30', '11:00', '', '08:05']);
  ok('test times typed as 9.30, 10:30 or 1100 are all stored as a time');
  const s1 = call(g, { action: 'state', token: tok }).students;
  assert.deepEqual(s1.map(x => x.name + ' ' + x.day + ' ' + x.time), ['Polina 2026-11-04 09:30', 'Paul 2026-11-04 10:30', 'Faddei 2026-11-04 11:00', 'Yutaro 2026-11-04 ', 'Tales 2026-11-05 08:05']);
  ok('the page sees them on the right dates and times, including a date typed as text');
  st.getRange(7, 1, 1, 4).setValues([[pick('2026-11-05'), '', 'Ida', 'K4']]);
  call(g, { action: 'tidy', token: tok });
  assert.ok(call(g, { action: 'state', token: tok }).students.some(x => x.name === 'Ida'));
  ok('rows added without the edit trigger get an ID on the next refresh');
  st.getRange(3, 6).setValue('absent');
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
  te.getRange(3, 1).setValue('Isuru');
  call(g, { action: 'addStudents', token: tok, names: ['Lior'], studio: 'K4', day: '2026-11-04', teacher: 'isuru' });
  call(g, { action: 'addStudents', token: tok, names: ['Sam'], studio: 'K4', day: '2026-11-04', teacher: 'Nobody' });
  let ss2 = call(g, { action: 'state', token: tok }).students;
  assert.equal(ss2.find(x => x.name === 'Lior').teacher, 'Isuru'); assert.equal(ss2.find(x => x.name === 'Sam').teacher, '');
  call(g, { action: 'updateStudent', token: tok, id: ss2.find(x => x.name === 'Sam').id, teacher: 'Gunni' });
  assert.equal(call(g, { action: 'state', token: tok }).students.find(x => x.name === 'Sam').teacher, 'Gunni');
  assert.equal(tab(book, 'Students').data.find(r => r[2] === 'Sam')[4], 'Gunni');
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
  assert.deepEqual(d[0], ['Test date', 'Test time', 'Student', 'Studio', 'Teacher', 'Removed', 'ID']);
  assert.deepEqual(d.slice(1).map(r => r.slice(1)), [['', 'Faddei', 'K3', '', '', 'abc-1'], ['', 'Paul', 'K4', '', 'absent', 'abc-2']]);
  assert.equal(g.day_(d[1][0]), '2026-11-02');
  ok('the oldest Students layout becomes date | time | name | studio | teacher | removed | id, keeping IDs');
}
// 6c. The layout with a time but no teacher gets a Teacher column after Studio; everything else stays where it was.
{
  const { g, book } = loadScript();
  const sh = g.SpreadsheetApp.getActive().insertSheet('Students');
  sh.getRange(1, 1, 1, 6).setValues([['Test date', 'Test time', 'Student', 'Studio', 'Removed', 'ID']]);
  sh.getRange(2, 1, 1, 6).setValues([[g.toDate_('2026-11-02'), '09:30', 'Faddei', 'K3', '', 'abc-1']]);
  g.setup();
  const d = tab(book, 'Students').data;
  assert.deepEqual(d[0], ['Test date', 'Test time', 'Student', 'Studio', 'Teacher', 'Removed', 'ID']);
  assert.deepEqual(d[1].slice(1), ['09:30', 'Faddei', 'K3', '', '', 'abc-1']);
  ok('the layout without a Teacher column gets one after Studio, nothing else moves');
}
// 7. The Grades tab: same sum as the page, a row per student, a summary per grade.
{
  const { g, book } = loadScript();
  g.setup();
  const te = g.SpreadsheetApp.getActive().getSheetByName('Teachers');
  te.getRange(2, 1, 2, 2).setValues([['Gunni', 'admin'], ['Anna', 'admin']]); g.setTeamPassword_('test-pass-1');
  const gt = call(g, { action: 'login', name: 'gunni', password: 'test-pass-1' }).token;
  const at = call(g, { action: 'login', name: 'anna', password: 'test-pass-1' }).token;
  call(g, { action: 'addStudents', token: gt, names: ['Polina', 'Paul'], studio: 'K3', day: '2026-11-04' });
  const [pol, paul] = call(g, { action: 'state', token: gt }).students.map(s => s.id);
  const IDS = ['p1','p2','p3','m1','m2','m3','m4','r1','t1','t2','t3','t4','c1','c2','c3','c4','s1','s2'];
  const all = v => Object.fromEntries(IDS.map(id => [id, v(id)]));
  call(g, { action: 'saveEntry', token: gt, sid: pol, patch: { marks: all(id => id === 'm4' ? 'h' : 'y') } });
  call(g, { action: 'saveEntry', token: at, sid: pol, patch: { marks: all(() => 'y'), safety: { a: { t: 1, text: 'Speakers switched with level up' } } } });
  call(g, { action: 'saveEntry', token: gt, sid: paul, patch: { marks: { p1: 'y', r1: 'n' } } });
  call(g, { action: 'publishResults', token: gt, day: '2026-11-04' });
  const tabs = book.sheets.map(s => s.name);
  assert.ok(tabs.includes('Grades') && !tabs.includes('Results') && !tabs.includes('Days'));
  assert.equal(tabs[1], 'Grades');
  const d = tab(book, 'Grades').data, head = d[6];
  const row = name => d.find(r => r[1] === name);
  const P = row('Polina'), Q = row('Paul');
  // Polina: m4 averages (75+100)/2 → score 99, one critical error → final 89 → Grade A
  assert.equal(P[head.indexOf('Score')], 99); assert.equal(P[head.indexOf('Errors')], -10); assert.equal(P[head.indexOf('Final')], 89);
  assert.equal(P[head.indexOf('Grade')], 'Grade A'); assert.equal(P[head.indexOf('Teachers')], 'Gunni, Anna');
  assert.equal(P[head.indexOf('Mic')], 97);   // (100+100+100+87.5)/4
  assert.equal(Q[head.indexOf('Grade')], 'Unfinished'); assert.equal(Q[head.indexOf('Core skill under 50%')], 'Console → speakers');
  assert.equal(head.indexOf('In the app'), -1);
  ok('the Grades tab works out the same grade as the page, per stage, with errors and core fails');
  assert.equal(d[3][1], 1); assert.equal(d[3][4], 1);   // one Grade A, one Unfinished in the summary
  assert.equal((d[1] || []).join(''), '');   // row 2 stays empty: no subtitle line
  ok('the summary counts each grade');
  const gr = call(g, { action: 'grades', token: at }).students, gp = gr.find(x => x.name === 'Polina'), gq = gr.find(x => x.name === 'Paul');
  assert.equal(gp.grade, 'Grade A'); assert.equal(gp.final, 89); assert.equal(gp.score, 99); assert.equal(gp.errors.length, 1); assert.equal(gp.complete, true);
  assert.equal(gp.stages.find(x => x.name === 'Mic').pct, 97); assert.deepEqual(gp.teachers, ['Gunni', 'Anna']);
  assert.equal(gq.complete, false); assert.equal(gq.marked, 2); assert.deepEqual(gq.coreFails, ['Console → speakers']);
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
  assert.equal(st.questions.length, 18);
  assert.deepEqual([...new Set(st.questions.map(x => x.stage))], ['Power', 'Mic', 'Route', 'Record', 'Cue', 'Shut down']);
  const m4 = st.questions.find(x => x.id === 'm4');
  assert.equal(m4.heading, 'Connecting a Mic'); assert.equal(m4.core, 1); assert.match(m4.ask, /condenser/); assert.equal(m4.studios, null);
  assert.equal(st.questions.find(x => x.id === 'p3').core, 0);
  ok('setup writes the 18 questions into the Questions tab and the page reads them from there');

  call(g, { action: 'addStudents', token: tok, names: ['Polina'], studio: 'K3', day: '2026-11-04' });
  call(g, { action: 'addStudents', token: tok, names: ['Tales'], studio: 'K4', day: '2026-11-11' });
  const [pol, tal] = call(g, { action: 'state', token: tok }).students.map(s => s.id);
  call(g, { action: 'saveEntry', token: tok, sid: pol, patch: { marks: { p1: 'y' } } });

  // a new core question for K4 only, typed under the last row with no stage: it joins "Shut down"
  const last = q.getLastRow() + 1;
  q.getRange(last, 3, 1, 4).setValues([['Patches the headphone line to the K4 booth', '', true, 'K4']]);
  g.onEdit({ range: q.getRange(last, 3, 1, 4) });
  const newId = q.getRange(last, 8).getValue();
  assert.match(String(newId), /^q[0-9a-f]{8}$/);
  // reword an existing one
  q.getRange(2, 3).setValue('Powers up in the right order, speakers last');
  st = call(g, { action: 'state', token: tok });
  const nq = st.questions.find(x => x.id === newId);
  assert.equal(st.questions.length, 19); assert.equal(nq.stage, 'Shut down'); assert.deepEqual(nq.studios, ['K4']); assert.equal(nq.core, 1);
  ok('a question typed in the tab gets an ID, its stage from the row above, and its studio');
  assert.equal(st.qdays['2026-11-04'].length, 18);
  assert.equal(st.qdays['2026-11-04'][0].t, 'Powers up in the right order, speakers last');
  assert.equal(st.qdays['2026-11-11'], undefined);
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
  const rows = tab(book, 'Questions').data; const at = rows.findIndex(r => r[7] === 's2') + 1;
  q.deleteRow(at);
  call(g, { action: 'tidy', token: tok });
  st = call(g, { action: 'state', token: tok });
  assert.ok(!st.questions.some(x => x.id === 's2'));
  assert.ok(st.qdays['2026-11-04'].some(x => x.id === 's2'));
  assert.ok(tab(book, 'Marks').data[0].includes('Left as found (old)'));
  ok('a removed question disappears from new tests but stays on dates already marked with it');

  // the Grades tab: each date's own set. Tales's date was first marked while there were 19 questions, Polina's with 18
  const IDS18 = ['p1','p2','p3','m1','m2','m3','m4','r1','t1','t2','t3','t4','c1','c2','c3','c4','s1','s2'];
  call(g, { action: 'saveEntry', token: tok, sid: pol, patch: { marks: Object.fromEntries(IDS18.map(id => [id, 'y'])) } });
  assert.equal(st.qdays['2026-11-11'].length, 19);
  const tq = st.qdays['2026-11-11'].map(x => x.id);
  call(g, { action: 'saveEntry', token: tok, sid: tal, patch: { marks: Object.fromEntries(tq.map(id => [id, 'y'])) } });
  g.buildGrades_();
  const gd = tab(book, 'Grades').data, gh = gd[6], gr = n => gd.find(r => r[1] === n);
  assert.equal(gr('Polina')[gh.indexOf('Grade')], 'Grade A'); assert.equal(gr('Polina')[gh.indexOf('Marked')], '18/18');
  assert.equal(gr('Tales')[gh.indexOf('Grade')], 'Grade A'); assert.equal(gr('Tales')[gh.indexOf('Marked')], '19/19');
  ok('Grades marks each student out of their own date\'s questions');

  g.useCurrentQuestions_('2026-11-04');
  st = call(g, { action: 'state', token: tok });
  assert.equal(st.qdays['2026-11-04'].length, 18);   // today's 18; the page and Grades leave out the K4 one for her
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
console.log(`\n${n} checks passed`);
