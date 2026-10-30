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
// 6. Students typed or pasted straight into the Students tab.
{
  const { g, book } = loadScript();
  g.setup();
  const ss = g.SpreadsheetApp.getActive(), te = ss.getSheetByName('Teachers'), st = ss.getSheetByName('Students');
  te.getRange(2, 1).setValue('Gunni'); g.setTeamPassword_('test-pass-1');
  const tok = call(g, { action: 'login', name: 'gunni', password: 'test-pass-1' }).token;
  // one typed row with a real date (what the date picker gives), then a pasted list with no dates or studios
  // dates must come from inside the sandbox, like the ones Google hands the script
  const pick = d => { const x = g.toDate_(d); x.setHours(0); return x; };   // what the date picker stores: midnight
  st.getRange(2, 1, 1, 3).setValues([[pick('2026-11-04'), 'Polina', 'K3']]);
  g.onEdit({ range: st.getRange(2, 1, 1, 3) });
  st.getRange(3, 2, 3, 1).setValues([['Paul'], ['Faddei'], ['  Yutaro ']]);
  g.onEdit({ range: st.getRange(3, 2, 3, 1) });
  // someone types a date as text in a later row
  st.getRange(6, 1, 1, 3).setValues([['5.11.2026', 'Tales', 'K4']]);
  g.onEdit({ range: st.getRange(6, 1, 1, 3) });
  const rows = tab(book, 'Students').data.slice(1);
  assert.ok(rows.every(r => r[6]), 'every named row gets an ID');
  assert.deepEqual(rows.map(r => r[2]), ['K3', 'K3', 'K3', 'K3', 'K4']);
  ok('a pasted list takes the date and studio from the row above and gets IDs');
  const s1 = call(g, { action: 'state', token: tok }).students;
  assert.deepEqual(s1.map(x => x.name + ' ' + x.day), ['Polina 2026-11-04', 'Paul 2026-11-04', 'Faddei 2026-11-04', 'Yutaro 2026-11-04', 'Tales 2026-11-05']);
  ok('the page sees them on the right dates, including a date typed as text');
  // a row that arrived without the edit trigger (an import) still shows up
  st.getRange(7, 1, 1, 3).setValues([[pick('2026-11-05'), 'Ida', 'K4']]);
  assert.ok(call(g, { action: 'state', token: tok }).students.some(x => x.name === 'Ida'));
  ok('rows added without the edit trigger get an ID on the next refresh');
  st.getRange(3, 6).setValue('absent');
  assert.ok(!call(g, { action: 'state', token: tok }).students.some(x => x.name === 'Paul'));
  ok('typing in Removed takes a student off the page');
  call(g, { action: 'addStudents', token: tok, names: ['Mira'], studio: 'K3', day: '2026-11-04' });
  const last = tab(book, 'Students').data.at(-1);
  assert.equal(g.day_(last[0]), '2026-11-04'); assert.equal(typeof last[0], 'object'); assert.equal(last[1], 'Mira');
  ok('students added on the page land in the sheet as real dates');
}
// 7. The Grades tab: same sum as the page, a row per student, a summary per grade.
{
  const { g, book } = loadScript();
  g.setup();
  const te = g.SpreadsheetApp.getActive().getSheetByName('Teachers');
  te.getRange(2, 1, 2, 1).setValues([['Gunni'], ['Anna']]); g.setTeamPassword_('test-pass-1');
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
  const seen = call(g, { action: 'state', token: at }).entries.filter(e => e.sid === pol);
  assert.ok(seen.find(e => e.teacher === 'Anna').data);
  const gun = seen.find(e => e.teacher === 'Gunni'); assert.equal(gun.data, undefined); assert.equal(gun.progress, 18);
  ok('in the app a teacher only ever sees their own marks, and how far the others have got');
}
console.log(`\n${n} checks passed`);
