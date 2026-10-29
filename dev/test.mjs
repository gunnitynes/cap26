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
  assert.deepEqual(t[0].slice(0, 5), ['Name', 'Password', 'Role', 'Status', 'Last sign-in']);
  assert.equal(t[1][0], 'Gunni'); assert.equal(t[1][1], '●●●●●● set'); assert.equal(t[1][2], 'admin');
  assert.match(String(t[1][3]), /^Ready/);
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

// 2. Typing a password straight into the new layout (what onEdit sees).
{
  const { g, book } = loadScript();
  g.setup();
  const sh = g.SpreadsheetApp.getActive().getSheetByName('Teachers');
  sh.getRange(2, 1, 1, 2).setValues([['Gunni', 'test-pass-1']]);
  g.onEdit({ range: sh.getRange(2, 2) });
  sh.getRange(3, 1, 1, 2).setValues([['Mira', 'abc']]);
  g.onEdit({ range: sh.getRange(3, 2) });
  sh.getRange(4, 1).setValue('Anna');
  g.onEdit({ range: sh.getRange(4, 1) });
  const t = tab(book, 'Teachers').data;
  assert.equal(t[1][1], '●●●●●● set'); assert.equal(t[1][2], 'admin'); assert.match(String(t[1][3]), /^Ready/);
  assert.match(String(t[2][3]), /too short/);
  assert.match(String(t[3][3]), /^Waiting/);
  ok('typed passwords are hidden at once; too short and empty ones are explained in Status');
  assert.equal(call(g, { action: 'login', name: 'gunni', password: 'test-pass-1' }).ok, true);
  ok('sign-in works right after typing it in the sheet');

  // 3. First sign-in for a name with no password.
  let r = call(g, { action: 'login', name: 'anna', password: '' });
  assert.equal(r.code, 'activate'); assert.equal(r.name, 'Anna');
  r = call(g, { action: 'login', name: 'anna', password: 'anna2026', confirm: 'anna2027' });
  assert.equal(r.ok, false); assert.equal(r.code, 'activate');
  r = call(g, { action: 'login', name: 'anna', password: 'anna2026', confirm: 'anna2026' });
  assert.equal(r.ok, true); assert.equal(r.activated, true); assert.equal(r.role, 'teacher');
  assert.match(String(tab(book, 'Teachers').data[3][3]), /^Ready/);
  ok('a teacher with no password chooses one (typed twice) and is signed in');
  assert.equal(call(g, { action: 'login', name: 'anna', password: 'anna2026' }).ok, true);
  assert.equal(call(g, { action: 'login', name: 'anna', password: 'other123', confirm: 'other123' }).ok, false);
  ok('once chosen, nobody else can choose it again');

  // 4. Clearing a password lets the teacher choose a new one and signs them out.
  const tok = call(g, { action: 'login', name: 'anna', password: 'anna2026' }).token;
  sh.getRange(4, 2).setValue('');                 // only the visible cell, as a teacher would
  g.onEdit({ range: sh.getRange(4, 2) });
  assert.equal(call(g, { action: 'state', token: tok }).code, 'auth');
  assert.equal(call(g, { action: 'login', name: 'anna', password: '' }).code, 'activate');
  ok('clearing a password in the sheet signs them out and lets them choose a new one');

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
console.log(`\n${n} checks passed`);
