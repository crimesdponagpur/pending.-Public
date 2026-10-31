// शोध: Mac वरील फोल्डरमधील माहिती शोधून Word / Excel / PDF बनवणारा स्थानिक सर्व्हर
// चालवा: npm start   (मग ब्राउझरमध्ये http://localhost:3939)
const http = require('http');
const fs = require('fs');
const os = require('os');
const path = require('path');
const crypto = require('crypto');
const { execFile } = require('child_process');
const { Index } = require('./lib/indexer');
const { ollamaModels, ollamaAnswer, planRequest, summarize, webSearch } = require('./lib/ai');
const { buildTables, catalog, runPlan } = require('./lib/tables');
const { toDocx, toXlsx, toPrintHtml } = require('./lib/output');

const PORT = Number(process.env.PORT) || 3939;
// Mac ॲपमध्ये SHODH_DATA = ~/Library/Application Support/शोध
const DATA = process.env.SHODH_DATA || __dirname;
const CONFIG = path.join(DATA, 'config.json');
const expand = p => p.replace(/^~(?=$|\/)/, os.homedir());

function loadConfig() {
  let c = {};
  try { c = JSON.parse(fs.readFileSync(CONFIG, 'utf8')); } catch {}
  const folder = (require.main === module && process.argv[2]) || process.env.SHODH_FOLDER || c.folder || '~/Shodh-Mahiti';
  return { ...c, folder: path.resolve(expand(folder)) };
}

let config = loadConfig();
fs.mkdirSync(config.folder, { recursive: true });
const index = new Index();
const reports = new Map();

function reindex() {
  console.log('निर्देशांक बनवत आहे:', config.folder);
  return index.build(config.folder).then(s => {
    tables = buildTables(index.chunks);
    console.log(`तयार: ${s.files} फाइल्स, ${s.chunks} तुकडे` + (s.errors.length ? `, ${s.errors.length} फाइल्स वाचता आल्या नाहीत` : ''));
    return s;
  }).catch(err => { index.status = { state: 'error', errors: [err.message] }; });
}

const send = (res, code, body, type = 'application/json; charset=utf-8', extra = {}) => {
  res.writeHead(code, { 'Content-Type': type, ...extra });
  res.end(typeof body === 'string' || Buffer.isBuffer(body) ? body : JSON.stringify(body));
};

const readBody = req => new Promise((ok, bad) => {
  let s = '';
  req.on('data', d => { s += d; if (s.length > 15e6) req.destroy(); });
  req.on('end', () => { try { ok(s ? JSON.parse(s) : {}); } catch (e) { bad(e); } });
});

let tables = [];

// Google Drive for desktop चा फोल्डर: ~/Library/CloudStorage/GoogleDrive-<email>/My Drive
function driveFolders() {
  const base = path.join(os.homedir(), 'Library', 'CloudStorage');
  const out = [];
  try {
    for (const d of fs.readdirSync(base)) {
      if (!d.startsWith('GoogleDrive')) continue;
      for (const sub of fs.readdirSync(path.join(base, d))) {
        if (!sub.startsWith('.')) out.push(path.join(base, d, sub));
      }
    }
  } catch {}
  return out;
}

// Drive मधून नवीन फाइल्स येत राहतात, म्हणून दर 10 मिनिटांनी बदल तपासा (फक्त बदललेल्या फाइल्स वाचल्या जातात)
setInterval(() => { if (index.status.state !== 'indexing') reindex(); }, 10 * 60 * 1000).unref();

async function ask({ request = '', useAI = false, model, useWeb = false, limit = 30 }) {
  request = String(request).trim();
  if (!request) throw new Error('काय हवे ते लिहा');
  const hits = index.query(request, limit);
  const web = useWeb ? await webSearch(request) : [];
  let answer = null, aiError = null, analysis = null, plan = null;
  if (useAI) {
    const models = await ollamaModels();
    const m = model && models.includes(model) ? model : models[0];
    if (!m) aiError = 'Ollama चालू नाही किंवा मॉडेल नाही (README पहा). फक्त शोध-निकाल दिले आहेत.';
    else {
      try {
        // ज्या तक्त्यांत शोध-निकाल आले ते AI ला आधी दाखवायचे
        if (tables.length) {
          const hitCount = new Map();
          for (const h of hits) if (h.headers) { const sig = h.headers.join('\u0001'); hitCount.set(sig, (hitCount.get(sig) || 0) + 1); }
          const ranked = [...tables].sort((a, b) => (hitCount.get(b.headers.slice(0, -1).join('\u0001')) || 0) - (hitCount.get(a.headers.slice(0, -1).join('\u0001')) || 0));
          plan = await planRequest({ model: m, request, catalog: catalog(ranked) });
        }
        if (plan && plan.mode === 'analysis') {
          try {
            analysis = runPlan(tables, plan);
            answer = await summarize({ model: m, request, result: analysis });
          } catch (e) {
            aiError = 'मोजणी जमली नाही (' + e.message + '). त्याऐवजी लिखित उत्तर दिले आहे.';
            analysis = null;
          }
        }
        if (!analysis) {
          const kind = /excel|एक्सेल|तक्ता|table|यादी/i.test(request) ? 'table' : 'doc';
          answer = await ollamaAnswer({ model: m, request, hits: hits.slice(0, 15), web, kind });
        }
      } catch (e) { aiError = e.message; }
    }
  }
  const title = request.length > 80 ? request.slice(0, 77) + '...' : request;
  const id = crypto.randomBytes(6).toString('hex');
  const rep = { id, title, request, answer, hits, web, aiError, analysis, plan, at: new Date().toISOString() };
  reports.set(id, rep);
  if (reports.size > 50) reports.delete(reports.keys().next().value);
  // पानावर मूळ ओळी पाठवायची गरज नाही
  return { ...rep, analysis: analysis && { ...analysis, rows: undefined } };
}

const safeName = s => s.replace(/[\\/:*?"<>|\n]+/g, ' ').trim().slice(0, 60) || 'shodh';

const server = http.createServer(async (req, res) => {
  const url = new URL(req.url, 'http://localhost');
  // दुसऱ्या वेबसाइटने या सर्व्हरला गुपचूप विनंती पाठवू नये
  const origin = req.headers.origin;
  if (req.method !== 'GET' && origin && !/^http:\/\/(localhost|127\.0\.0\.1):\d+$/.test(origin)) return send(res, 403, { error: 'forbidden' });
  try {
    if (req.method === 'GET' && url.pathname === '/') {
      return send(res, 200, fs.readFileSync(path.join(__dirname, 'public', 'index.html')), 'text/html; charset=utf-8');
    }
    if (req.method === 'GET' && url.pathname === '/chart.js') {
      return send(res, 200, fs.readFileSync(path.join(__dirname, 'node_modules', 'chart.js', 'dist', 'chart.umd.min.js')), 'text/javascript; charset=utf-8');
    }
    const ch = url.pathname.match(/^\/api\/chart\/([0-9a-f]+)$/);
    if (req.method === 'POST' && ch) {
      const rep = reports.get(ch[1]);
      const { png } = await readBody(req);
      if (rep && typeof png === 'string' && png.startsWith('data:image/png;base64,')) rep.chartPng = Buffer.from(png.split(',')[1], 'base64');
      return send(res, 200, { ok: !!rep });
    }
    if (req.method === 'GET' && url.pathname === '/api/status') {
      return send(res, 200, { folder: config.folder, index: index.status, models: await ollamaModels(), drives: driveFolders() });
    }
    if (req.method === 'POST' && url.pathname === '/api/folder') {
      const { folder } = await readBody(req);
      const f = path.resolve(expand(String(folder || '')));
      if (!folder || !fs.existsSync(f) || !fs.statSync(f).isDirectory()) return send(res, 400, { error: 'हा फोल्डर सापडला नाही: ' + f });
      config.folder = f;
      fs.writeFileSync(CONFIG, JSON.stringify({ folder: f }, null, 2));
      reindex();
      return send(res, 200, { ok: true, folder: f });
    }
    if (req.method === 'POST' && url.pathname === '/api/reindex') {
      reindex();
      return send(res, 200, { ok: true });
    }
    if (req.method === 'POST' && url.pathname === '/api/ask') {
      return send(res, 200, await ask(await readBody(req)));
    }
    const dl = url.pathname.match(/^\/api\/download\/([0-9a-f]+)\.(docx|xlsx|pdf)$/);
    if (req.method === 'GET' && dl) {
      const rep = reports.get(dl[1]);
      if (!rep) return send(res, 404, { error: 'निकाल जुना झाला, पुन्हा शोधा' });
      const name = encodeURIComponent(safeName(rep.title));
      if (dl[2] === 'pdf') return send(res, 200, toPrintHtml(rep, !url.searchParams.has('app')), 'text/html; charset=utf-8');
      const buf = dl[2] === 'docx' ? await toDocx(rep) : await toXlsx(rep);
      const type = dl[2] === 'docx'
        ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
      return send(res, 200, buf, type, { 'Content-Disposition': `attachment; filename*=UTF-8''${name}.${dl[2]}` });
    }
    const file = url.pathname.match(/^\/api\/open$/) && url.searchParams.get('f');
    if (req.method === 'POST' && file) {
      const p = path.resolve(config.folder, file);
      if (!p.startsWith(config.folder + path.sep) || !fs.existsSync(p)) return send(res, 404, { error: 'फाइल सापडली नाही' });
      if (process.platform === 'darwin') execFile('open', [p]);
      return send(res, 200, { ok: true });
    }
    send(res, 404, { error: 'not found' });
  } catch (err) {
    console.error(err);
    send(res, 500, { error: err.message });
  }
});

// फक्त याच Mac वरून उघडता यावे म्हणून 127.0.0.1
function start(port = PORT) {
  return new Promise((ok, bad) => {
    server.once('error', bad);
    server.listen(port, '127.0.0.1', () => {
      port = server.address().port;
      console.log(`\nशोध चालू आहे: http://localhost:${port}\nफोल्डर: ${config.folder}\nबंद करण्यासाठी: Ctrl + C\n`);
      reindex();
      ok(port);
    });
  });
}

module.exports = { start };

// Terminal मधून (npm run web) चालवले तर ब्राउझरमध्ये उघडा
if (require.main === module) {
  start().then(port => {
    if (process.platform === 'darwin' && !process.env.NO_OPEN) execFile('open', [`http://localhost:${port}`]);
  });
}
