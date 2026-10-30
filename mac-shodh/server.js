// शोध: Mac वरील फोल्डरमधील माहिती शोधून Word / Excel / PDF बनवणारा स्थानिक सर्व्हर
// चालवा: npm start   (मग ब्राउझरमध्ये http://localhost:3939)
const http = require('http');
const fs = require('fs');
const os = require('os');
const path = require('path');
const crypto = require('crypto');
const { execFile } = require('child_process');
const { Index } = require('./lib/indexer');
const { ollamaModels, ollamaAnswer, webSearch } = require('./lib/ai');
const { toDocx, toXlsx, toPrintHtml } = require('./lib/output');

const PORT = Number(process.env.PORT) || 3939;
const CONFIG = path.join(__dirname, 'config.json');
const expand = p => p.replace(/^~(?=$|\/)/, os.homedir());

function loadConfig() {
  let c = {};
  try { c = JSON.parse(fs.readFileSync(CONFIG, 'utf8')); } catch {}
  const folder = process.argv[2] || process.env.SHODH_FOLDER || c.folder || '~/Shodh-Mahiti';
  return { ...c, folder: path.resolve(expand(folder)) };
}

let config = loadConfig();
fs.mkdirSync(config.folder, { recursive: true });
const index = new Index();
const reports = new Map();

function reindex() {
  console.log('निर्देशांक बनवत आहे:', config.folder);
  return index.build(config.folder).then(s => {
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
  req.on('data', d => { s += d; if (s.length > 1e6) req.destroy(); });
  req.on('end', () => { try { ok(s ? JSON.parse(s) : {}); } catch (e) { bad(e); } });
});

async function ask({ request = '', useAI = false, model, useWeb = false, limit = 30 }) {
  request = String(request).trim();
  if (!request) throw new Error('काय हवे ते लिहा');
  const hits = index.query(request, useAI ? Math.min(limit, 15) : limit);
  const web = useWeb ? await webSearch(request) : [];
  let answer = null, aiError = null;
  if (useAI) {
    const models = await ollamaModels();
    const m = model && models.includes(model) ? model : models[0];
    if (!m) aiError = 'Ollama चालू नाही किंवा मॉडेल नाही (README पहा). फक्त शोध-निकाल दिले आहेत.';
    else {
      try {
        const kind = /excel|एक्सेल|तक्ता|table|यादी/i.test(request) ? 'table' : 'doc';
        answer = await ollamaAnswer({ model: m, request, hits, web, kind });
      } catch (e) { aiError = e.message; }
    }
  }
  const title = request.length > 80 ? request.slice(0, 77) + '...' : request;
  const id = crypto.randomBytes(6).toString('hex');
  const rep = { id, title, request, answer, hits, web, aiError, at: new Date().toISOString() };
  reports.set(id, rep);
  if (reports.size > 50) reports.delete(reports.keys().next().value);
  return rep;
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
    if (req.method === 'GET' && url.pathname === '/api/status') {
      return send(res, 200, { folder: config.folder, index: index.status, models: await ollamaModels() });
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
      if (dl[2] === 'pdf') return send(res, 200, toPrintHtml(rep), 'text/html; charset=utf-8');
      const buf = dl[2] === 'docx' ? await toDocx(rep) : toXlsx(rep);
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
server.listen(PORT, '127.0.0.1', () => {
  const link = `http://localhost:${PORT}`;
  console.log(`\nशोध ॲप चालू आहे: ${link}\nफोल्डर: ${config.folder}\nबंद करण्यासाठी: Ctrl + C\n`);
  if (process.platform === 'darwin' && !process.env.NO_OPEN) execFile('open', [link]);
  reindex();
});
