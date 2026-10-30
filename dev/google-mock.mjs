// Stand-ins for the Google services Code.gs uses, so it can run on this computer (dev server and tests).
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import vm from 'node:vm';
const root = path.join(path.dirname(new URL(import.meta.url).pathname), '..');

export function loadScript() {
  // ---- a small in-memory spreadsheet. Formatting calls (colours, widths, validation…) are accepted and ignored.
  const chain = () => new Proxy(function () {}, { get: (t, k) => k === 'build' ? () => ({}) : () => chain() });
  const lenient = obj => new Proxy(obj, { get: (t, k) => k in t ? t[k] : (typeof k === 'string' ? () => lenient(t) : undefined) });
  const colNum = s => s.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  class Range {
    constructor(sh, r, c, nr, nc) { Object.assign(this, { sh, r, c, nr, nc }); }
    getValues() { const out = []; for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) row.push(this.sh.get(this.r + i, this.c + j)); out.push(row); } return out; }
    getValue() { return this.sh.get(this.r, this.c); }
    getDisplayValues() { return this.getValues().map(r => r.map(v => v instanceof Date ? v.toISOString() : String(v ?? ''))); }
    setValues(v) { v.forEach((row, i) => row.forEach((x, j) => this.sh.set(this.r + i, this.c + j, x))); return lenient(this); }
    setValue(x) { this.sh.set(this.r, this.c, x); return lenient(this); }
    setFormula(f) { this.sh.set(this.r, this.c, f); return lenient(this); }
    getRow() { return this.r; } getNumRows() { return this.nr; } getColumn() { return this.c; } getNumColumns() { return this.nc; } getSheet() { return this.sh; }
  }
  class Sheet {
    constructor(name) { this.name = name; this.data = []; }
    getName() { return this.name; }
    get(r, c) { return (this.data[r - 1] || [])[c - 1] ?? ''; }
    set(r, c, x) { while (this.data.length < r) this.data.push([]); const row = this.data[r - 1]; while (row.length < c) row.push(''); row[c - 1] = x; }
    getLastRow() { let n = this.data.length; while (n && !this.data[n - 1].some(v => v !== '' && v != null)) n--; return n; }
    getLastColumn() { return Math.max(0, ...this.data.map(r => r.length)); }
    getMaxRows() { return 1000; }
    getDataRange() { return lenient(new Range(this, 1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn()))); }
    getRange(a, b, c, d) {
      if (typeof a === 'string') {
        const [p, q] = a.split(':'), m1 = p.match(/^([A-Z]+)(\d*)$/), m2 = (q || p).match(/^([A-Z]+)(\d*)$/);
        const r1 = m1[2] ? +m1[2] : 1, r2 = m2[2] ? +m2[2] : 1000;
        return lenient(new Range(this, r1, colNum(m1[1]), r2 - r1 + 1, colNum(m2[1]) - colNum(m1[1]) + 1));
      }
      return lenient(new Range(this, a, b, c || 1, d || 1));
    }
    appendRow(row) { this.data.splice(this.getLastRow(), 0, row.slice()); }
    deleteRow(r) { this.data.splice(r - 1, 1); }
    insertColumnAfter(c) { this.data.forEach(row => { if (row.length > c) row.splice(c, 0, ''); }); }
    clear() { this.data = []; return this; }
    clearContents() { this.data = []; return this; }
    getMaxColumns() { return Math.max(26, this.getLastColumn()); }
    getBandings() { return []; }
  }
  const book = { sheets: [new Sheet('Sheet1')] };
  const SpreadsheetApp = {
    getActive: () => ({
      getSheetByName: n => { const s = book.sheets.find(s => s.name === n); return s ? lenient(s) : null; },
      insertSheet: (n, i) => { const s = new Sheet(n); typeof i === 'number' ? book.sheets.splice(i, 0, s) : book.sheets.push(s); return lenient(s); },
      getSheets: () => book.sheets,
      deleteSheet: s => { book.sheets = book.sheets.filter(x => x.name !== s.getName()); },
    }),
    getUi: () => { throw new Error('no ui here'); },
    newDataValidation: chain, newConditionalFormatRule: chain, BandingTheme: {}, InterpolationType: { NUMBER: 'NUMBER' }, BorderStyle: {},
  };
  const props = new Map(), cache = new Map();
  const g = {
    SpreadsheetApp,
    PropertiesService: { getScriptProperties: () => ({
      getProperty: k => props.has(k) ? props.get(k) : null, setProperty: (k, v) => props.set(k, v), setProperties: o => Object.entries(o).forEach(([k, v]) => props.set(k, v)),
      deleteProperty: k => props.delete(k), getProperties: () => Object.fromEntries(props) }) },
    CacheService: { getScriptCache: () => ({ get: k => cache.get(k) ?? null, put: (k, v) => cache.set(k, v), remove: k => cache.delete(k) }) },
    LockService: { getScriptLock: () => ({ tryLock: () => true, waitLock: () => {}, releaseLock: () => {} }) },
    ContentService: { MimeType: { JSON: 'json' }, createTextOutput: s => ({ s, setMimeType() { return this; } }) },
    Session: { getScriptTimeZone: () => 'Europe/Berlin' },
    Utilities: {
      getUuid: () => crypto.randomUUID(),
      DigestAlgorithm: { SHA_256: 'sha256' }, Charset: { UTF_8: 'utf8' },
      computeDigest: (alg, s) => [...crypto.createHash('sha256').update(s, 'utf8').digest()].map(b => b > 127 ? b - 256 : b),
      base64Encode: bytes => Buffer.from(bytes.map(b => b & 255)).toString('base64'),
      formatDate: (d, tz, f) => {
        // enough of Google's patterns for Code.gs: yyyy MM dd HH mm, in the given time zone
        const p = Object.fromEntries(new Intl.DateTimeFormat('en-GB', { timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', hour12: false })
          .formatToParts(d).map(x => [x.type, x.value]));
        return f.replace('yyyy', p.year).replace('MM', p.month).replace('dd', p.day).replace('HH', p.hour).replace('mm', p.minute);
      },
    },
    console,
  };
  vm.createContext(g);
  vm.runInContext(fs.readFileSync(path.join(root, 'apps-script/Code.gs'), 'utf8'), g);
  return { g, book };
  g.setup();
  g.setTeacher_('Gunnar', 'catalyst', 'admin');
  g.setTeacher_('Anna', '', null, true);            // no password yet: she chooses one at first sign-in

}
