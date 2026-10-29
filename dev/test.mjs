// Checks the sign-in paths of Code.gs on this computer:  node dev/test.mjs
import assert from 'node:assert/strict';
import { loadScript } from './google-mock.mjs';

const call = (g, body) => JSON.parse(g.doPost({ postData: { contents: JSON.stringify(body) } }).s);
const tab = (book, name) => book.sheets.find(s => s.name === name);
let n = 0; const ok = msg => console.log('  ✓', msg, ++n && '');

// 1. The old layout, exactly as typed by hand: a name, no role, the password in "Password hash", no salt.
{
  const { g, book } = loadScript();
  const ss = g.SpreadsheetApp.getActive();
  const sh = ss.insertSheet('Teachers');
  sh.getRange(1, 1, 1, 5).setValues([['Name', 'Role', 'Password hash', 'Salt', 'Added']]);
  sh.getRange(2, 1, 1, 5).setValues([['Gunni', '', 'test-pass-1', '', '']]);
  g.setup();
  const t = tab(book, 'Teachers').data;
  assert.deepEqual(t[0].slice(0, 5), ['Name', 'Sign-in code', 'Role', 'Status', 'Last sign-in']);
  assert.equal(t[1][0], 'Gunni'); assert.equal(t[1][1], '●●●●●● set'); assert.equal(t[1][2], 'admin');
  assert.match(String(t[1][3]), /^Admin: signs in/);
  assert.ok(!JSON.stringify(t).includes('test-pass-1'), 'the typed password must not stay in the sheet');
  ok('old layout moved over; the hand-typed password is hidden and Gunni is admin');
  const r = call(g, { action: 'login', name: ' gunni ', password: 'test-pass-1' });
  assert.equal(r.ok, true); assert.equal(r.role, 'admin');
  ok('Gunni signs in with the password typed in the sheet');
  assert.equal(call(g, { action: 'login', name: 'Gunni', password: 'wrong123' }).ok, false);
  ok('a wrong password is refused');
  assert.equal(book.sheets[0].name, 'Start here');
  ok('"Start here" is the first tab');
}

// 2. Typing in the Teachers tab: an admin password, and single-use codes for teachers (what onEdit sees).
{
  const { g, book } = loadScript();
  g.setup();
  const sh = g.SpreadsheetApp.getActive().getSheetByName('Teachers');
  sh.getRange(2, 1, 1, 2).setValues([['Gunni', 'test-pass-1']]);
  g.onEdit({ range: sh.getRange(2, 2) });
  sh.getRange(3, 1, 1, 2).setValues([['Mira', 'abc']]);
  g.onEdit({ range: sh.getRange(3, 2) });
  sh.getRange(4, 1, 1, 2).setValues([['Anna', '4821']]);
  g.onEdit({ range: sh.getRange(4, 2) });
  sh.getRange(5, 1).setValue('Jon');
  g.onEdit({ range: sh.getRange(5, 1) });
  const t = tab(book, 'Teachers').data;
  assert.equal(t[1][1], '●●●●●● set'); assert.equal(t[1][2], 'admin'); assert.match(String(t[1][3]), /^Admin/);
  assert.equal(t[2][2], 'teacher'); assert.match(String(t[2][3]), /too short/);
  assert.equal(t[3][1], '●●●● code ready'); assert.match(String(t[3][3]), /^Code ready/);
  assert.match(String(t[4][3]), /^No code yet/);
  assert.ok(!JSON.stringify(t).includes('4821'));
  ok('the first row is admin with a password; teachers get codes, hidden at once; Status explains each row');
  assert.equal(call(g, { action: 'login', name: 'gunni', password: 'test-pass-1' }).ok, true);
  ok('the admin signs in with the password');

  // 3. A teacher's code works once and the phone stays signed in.
  let r = call(g, { action: 'login', name: 'anna', password: '4822' });
  assert.equal(r.ok, false);
  r = call(g, { action: 'login', name: ' Anna ', password: ' 4821 ' });
  assert.equal(r.ok, true); assert.equal(r.role, 'teacher'); assert.ok(r.days > 300);
  const annaTok = r.token;
  assert.match(String(tab(book, 'Teachers').data[3][1]), /^used /); assert.match(String(tab(book, 'Teachers').data[3][3]), /^Signed in on a phone/);
  ok('a teacher signs in with name + code; the code is marked used and the phone stays signed in for the year');
  assert.equal(call(g, { action: 'login', name: 'anna', password: '4821' }).ok, false);
  ok('the same code cannot be used a second time');
  assert.equal(call(g, { action: 'state', token: annaTok }).ok, true);
  ok('the first phone keeps working after its code is used up');
  assert.equal(call(g, { action: 'login', name: 'jon', password: '1234' }).ok, false);
  assert.match(call(g, { action: 'login', name: 'jon', password: '' }).error, /no sign-in code/);
  ok('a teacher without a code is told to ask for one (no password ever)');
  sh.getRange(4, 2).setValue('7777'); g.onEdit({ range: sh.getRange(4, 2) });
  const r2 = call(g, { action: 'login', name: 'anna', password: '7777' });
  assert.equal(r2.ok, true); assert.equal(call(g, { action: 'state', token: annaTok }).ok, true);
  ok('a new code adds a second phone without signing out the first');
  assert.equal(call(g, { action: 'changePassword', token: r2.token, current: '7777', next: 'whatever1' }).ok, false);
  ok('teachers cannot set a password');

  // 4. Clearing the admin password signs the admin out and lets them choose a new one.
  const tok = call(g, { action: 'login', name: 'gunni', password: 'test-pass-1' }).token;
  sh.getRange(2, 2).setValue(''); g.onEdit({ range: sh.getRange(2, 2) });
  assert.equal(call(g, { action: 'state', token: tok }).code, 'auth');
  assert.equal(call(g, { action: 'login', name: 'gunni', password: '' }).code, 'activate');
  assert.equal(call(g, { action: 'login', name: 'gunni', password: 'test-pass-1', confirm: 'test-pass-1' }).ok, true);
  ok('clearing the admin password signs them out; they choose a new one at next sign-in');

  // 5. A teacher marks, the readable Marks row and the hidden copy agree.
  const gt = call(g, { action: 'login', name: 'gunni', password: 'test-pass-1' }).token;
  call(g, { action: 'addStudents', token: gt, names: ['Polina'], studio: 'Studio A', day: '2026-11-04' });
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
  te.getRange(2, 1, 1, 2).setValues([['Gunni', 'test-pass-1']]); g.onEdit({ range: te.getRange(2, 2) });
  const tok = call(g, { action: 'login', name: 'gunni', password: 'test-pass-1' }).token;
  // one typed row with a real date (what the date picker gives), then a pasted list with no dates or studios
  // dates must come from inside the sandbox, like the ones Google hands the script
  const pick = d => { const x = g.toDate_(d); x.setHours(0); return x; };   // what the date picker stores: midnight
  st.getRange(2, 1, 1, 3).setValues([[pick('2026-11-04'), 'Polina', 'Studio A']]);
  g.onEdit({ range: st.getRange(2, 1, 1, 3) });
  st.getRange(3, 2, 3, 1).setValues([['Paul'], ['Faddei'], ['  Yutaro ']]);
  g.onEdit({ range: st.getRange(3, 2, 3, 1) });
  // someone types a date as text in a later row
  st.getRange(6, 1, 1, 3).setValues([['5.11.2026', 'Tales', 'Studio B']]);
  g.onEdit({ range: st.getRange(6, 1, 1, 3) });
  const rows = tab(book, 'Students').data.slice(1);
  assert.ok(rows.every(r => r[6]), 'every named row gets an ID');
  assert.deepEqual(rows.map(r => r[2]), ['Studio A', 'Studio A', 'Studio A', 'Studio A', 'Studio B']);
  ok('a pasted list takes the date and studio from the row above and gets IDs');
  const s1 = call(g, { action: 'state', token: tok }).students;
  assert.deepEqual(s1.map(x => x.name + ' ' + x.day), ['Polina 2026-11-04', 'Paul 2026-11-04', 'Faddei 2026-11-04', 'Yutaro 2026-11-04', 'Tales 2026-11-05']);
  ok('the page sees them on the right dates, including a date typed as text');
  // a row that arrived without the edit trigger (an import) still shows up
  st.getRange(7, 1, 1, 3).setValues([[pick('2026-11-05'), 'Ida', 'Studio B']]);
  assert.ok(call(g, { action: 'state', token: tok }).students.some(x => x.name === 'Ida'));
  ok('rows added without the edit trigger get an ID on the next refresh');
  st.getRange(3, 6).setValue('absent');
  assert.ok(!call(g, { action: 'state', token: tok }).students.some(x => x.name === 'Paul'));
  ok('typing in Removed takes a student off the page');
  call(g, { action: 'addStudents', token: tok, names: ['Mira'], studio: 'Studio A', day: '2026-11-04' });
  const last = tab(book, 'Students').data.at(-1);
  assert.equal(g.day_(last[0]), '2026-11-04'); assert.equal(typeof last[0], 'object'); assert.equal(last[1], 'Mira');
  ok('students added on the page land in the sheet as real dates');
}
// 7. The Grades tab: same sum as the page, a row per student, a summary per grade.
{
  const { g, book } = loadScript();
  g.setup();
  const te = g.SpreadsheetApp.getActive().getSheetByName('Teachers');
  te.getRange(2, 1, 1, 2).setValues([['Gunni', 'test-pass-1']]); g.onEdit({ range: te.getRange(2, 2) });
  te.getRange(3, 1, 1, 2).setValues([['Anna', '4821']]); g.onEdit({ range: te.getRange(3, 2) });
  const gt = call(g, { action: 'login', name: 'gunni', password: 'test-pass-1' }).token;
  const at = call(g, { action: 'login', name: 'anna', password: '4821' }).token;
  call(g, { action: 'addStudents', token: gt, names: ['Polina', 'Paul'], studio: 'Studio A', day: '2026-11-04' });
  const [pol, paul] = call(g, { action: 'state', token: gt }).students.map(s => s.id);
  const IDS = ['p1','p2','p3','m1','m2','m3','m4','r1','t1','t2','t3','t4','c1','c2','c3','c4','s1','s2'];
  const all = v => Object.fromEntries(IDS.map(id => [id, v(id)]));
  call(g, { action: 'saveEntry', token: gt, sid: pol, patch: { marks: all(id => id === 'm4' ? 'h' : 'y') } });
  call(g, { action: 'saveEntry', token: at, sid: pol, patch: { marks: all(() => 'y'), safety: { a: { t: 1, text: 'Speakers switched with level up' } } } });
  call(g, { action: 'saveEntry', token: gt, sid: paul, patch: { marks: { p1: 'y', r1: 'n' } } });
  call(g, { action: 'publishResults', token: gt, day: '2026-11-04' });
  const tabs = book.sheets.map(s => s.name);
  assert.ok(tabs.includes('Grades') && !tabs.includes('Results'));
  assert.equal(tabs[1], 'Grades');
  const d = tab(book, 'Grades').data, head = d[7];
  const row = name => d.find(r => r[1] === name);
  const P = row('Polina'), Q = row('Paul');
  // Polina: m4 averages (75+100)/2 → score 99, one critical error → final 89 → Merit
  assert.equal(P[head.indexOf('Score')], 99); assert.equal(P[head.indexOf('Errors')], -10); assert.equal(P[head.indexOf('Final')], 89);
  assert.equal(P[head.indexOf('Grade')], 'Merit'); assert.equal(P[head.indexOf('Teachers')], 'Gunni, Anna');
  assert.equal(P[head.indexOf('Mic')], 97);   // (100+100+100+87.5)/4
  assert.equal(Q[head.indexOf('Grade')], 'Unfinished'); assert.equal(Q[head.indexOf('Core skill under 50%')], 'Mix bus → speakers');
  assert.equal(P[head.indexOf('In the app')], 'Hidden');
  ok('the Grades tab works out the same grade as the page, per stage, with errors and core fails');
  assert.equal(d[4][2], 1); assert.equal(d[4][5], 1);   // one Merit, one Unfinished in the summary
  ok('the summary counts each grade');
}
console.log(`\n${n} checks passed`);
