// Runs ../apps-script/Code.gs on this computer with stand-ins for Google's services, and serves ../docs/
// with the API at /api, so the whole thing can be tried before anything is deployed.
//   node dev/server.mjs [port]      then open http://localhost:8787/?api=/api
// Starts with one admin teacher: name "Gunnar", password "catalyst".
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import vm from 'node:vm';

const here = path.dirname(new URL(import.meta.url).pathname);
const root = path.join(here, '..');
const port = Number(process.argv[2] || 8787);

// ---- a small in-memory spreadsheet
class Range {
  constructor(sh, r, c, nr, nc) { Object.assign(this, { sh, r, c, nr, nc }); }
  getValues() { const out = []; for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) row.push(this.sh.get(this.r + i, this.c + j)); out.push(row); } return out; }
  setValues(v) { v.forEach((row, i) => row.forEach((x, j) => this.sh.set(this.r + i, this.c + j, x))); return this; }
  setValue(x) { this.sh.set(this.r, this.c, x); return this; }
  setNumberFormat() { return this; } setFontWeight() { return this; }
}
class Sheet {
  constructor(name) { this.name = name; this.data = []; }
  getName() { return this.name; }
  get(r, c) { return (this.data[r - 1] || [])[c - 1] ?? ''; }
  set(r, c, x) { while (this.data.length < r) this.data.push([]); const row = this.data[r - 1]; while (row.length < c) row.push(''); row[c - 1] = x; }
  getLastRow() { return this.data.length; }
  getLastColumn() { return Math.max(0, ...this.data.map(r => r.length)); }
  getDataRange() { return new Range(this, 1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn())); }
  getRange(a, b, c, d) {
    if (typeof a === 'string') { const col = a.charCodeAt(0) - 64; return new Range(this, 1, col, Math.max(1, this.getLastRow()), 1); }
    return new Range(this, a, b, c || 1, d || 1);
  }
  appendRow(row) { this.data.push(row.slice()); }
  deleteRow(r) { this.data.splice(r - 1, 1); }
  setFrozenRows() {}
}
const book = { sheets: [new Sheet('Sheet1')] };
const SpreadsheetApp = {
  getActive: () => ({
    getSheetByName: n => book.sheets.find(s => s.name === n) || null,
    insertSheet: n => { const s = new Sheet(n); book.sheets.push(s); return s; },
    getSheets: () => book.sheets,
    deleteSheet: s => { book.sheets = book.sheets.filter(x => x !== s); },
  }),
  getUi: () => { throw new Error('no ui here'); },
};
const props = new Map(), cache = new Map();
const g = {
  SpreadsheetApp,
  PropertiesService: { getScriptProperties: () => ({
    getProperty: k => props.has(k) ? props.get(k) : null, setProperty: (k, v) => props.set(k, v),
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
    formatDate: (d, tz, f) => d.toISOString().slice(0, f.length > 10 ? 16 : 10).replace('T', ' '),
  },
  console,
};
vm.createContext(g);
vm.runInContext(fs.readFileSync(path.join(root, 'apps-script/Code.gs'), 'utf8'), g);
g.setup();
g.setTeacher_('Gunnar', 'catalyst', 'admin');
const call = body => JSON.parse(g.doPost({ postData: { contents: body } }).s);

const types = { '.html': 'text/html', '.js': 'text/javascript', '.json': 'application/json', '.png': 'image/png', '.svg': 'image/svg+xml', '.webmanifest': 'application/manifest+json' };
http.createServer((req, res) => {
  if (req.url.startsWith('/api')) {
    let body = ''; req.on('data', d => body += d);
    req.on('end', () => {
      // a little delay, like the real script, so loading states get exercised
      setTimeout(() => { res.writeHead(200, { 'content-type': 'application/json', 'access-control-allow-origin': '*' }); res.end(JSON.stringify(call(body))); }, 250);
    });
    return;
  }
  if (req.url.startsWith('/__sheets')) { res.writeHead(200, { 'content-type': 'application/json' }); res.end(JSON.stringify(book.sheets.map(s => ({ name: s.name, rows: s.data })))); return; }
  let f = decodeURIComponent(req.url.split('?')[0]); if (f.endsWith('/')) f += 'index.html';
  const file = path.join(root, 'docs', f);
  if (!file.startsWith(path.join(root, 'docs')) || !fs.existsSync(file)) { res.writeHead(404); res.end('not found'); return; }
  res.writeHead(200, { 'content-type': types[path.extname(file)] || 'application/octet-stream' });
  fs.createReadStream(file).pipe(res);
}).listen(port, () => console.log(`http://localhost:${port}/?api=/api   (sign in: Gunnar / catalyst)`));
