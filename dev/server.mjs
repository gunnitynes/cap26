// Runs ../apps-script/Code.gs on this computer with stand-ins for Google's services, and serves ../docs/
// with the API at /api, so the whole thing can be tried before anything is deployed.
//   node dev/server.mjs [port]      then open http://localhost:8787/?api=/api
// Starts with two teachers, Gunnar (admin) and Anna; the team password is "catalyst".
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import vm from 'node:vm';

const here = path.dirname(new URL(import.meta.url).pathname);
const root = path.join(here, '..');
const port = Number(process.argv[2] || 8787);

import { loadScript } from './google-mock.mjs';
const { g, book } = loadScript();
g.setup();
g.setTeacher_('Gunnar', 'admin');
g.setTeacher_('Anna', 'teacher');
g.setTeamPassword_('catalyst');                    // everyone signs in with this
const call = body => JSON.parse(g.doPost({ postData: { contents: body } }).s);
// OPS_TRY_QUESTIONS=1: 45 s after start, someone "edits the Questions tab" (rewords the first question, adds a K3 one
// to Mic), so you can watch an open page pick it up.
if (process.env.OPS_TRY_QUESTIONS) setTimeout(() => {
  const q = g.SpreadsheetApp.getActive().getSheetByName('Questions');
  q.getRange(2, 3).setValue('Powers up in the right order (speakers last)');
  const at = q.getLastRow() + 1;   // a later row with Stage "Mic" still lands in the Mic stage
  q.getRange(at, 1, 1, 7).setValues([['Mic', '', 'Checks the K3 mic splitter is set to the desk', 'Where does the second feed go?', false, 'K3', '']]);
  g.onEdit({ range: q.getRange(at, 1, 1, 7) });
  console.log('Questions tab edited');
}, 45000);

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
}).listen(port, () => console.log(`http://localhost:${port}/?api=/api   (team password: catalyst)`));
