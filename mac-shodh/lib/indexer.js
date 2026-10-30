// फोल्डरमधील सर्व फाइल्सचा शोध-निर्देशांक (index). बदललेल्या फाइल्सच पुन्हा वाचल्या जातात.
const fs = require('fs');
const path = require('path');
const MiniSearch = require('minisearch');
const { extractFile, SUPPORTED } = require('./extract');

const CACHE = path.join(process.env.SHODH_DATA || path.join(__dirname, '..'), '.cache', 'files.json');

// मराठी/इंग्रजी शब्द तोडणे: अक्षरे, मात्रा (\p{M}) व अंक एकत्र ठेवायचे
const tokenize = s => s.toLowerCase().match(/[\p{L}\p{M}\p{N}]+/gu) || [];

// विनंतीतील सामान्य शब्द (शोधात उपयोग नाही)
const STOP = new Set(('आणि व या ही हा हे ते ती तो की का च ची चा चे ला ना ने नी त मध्ये मधील मधे वर खाली साठी करिता बद्दल संबंधी विषयी सर्व माहिती ' +
  'द्या दे दाखवा काढा शोधा बनवा करा पाहिजे हवी हवे हवा आहे आहेत होते होता एक कोणते कोण काय कसे किती ' +
  'word excel pdf वर्ड एक्सेल पीडीएफ फाइल फाईल format फॉरमॅट तक्ता अहवाल यादी ' +
  'the a an of in on for to and or is are was give show me find all about with from list make please').split(/\s+/));

function walk(dir, out = []) {
  let entries = [];
  try { entries = fs.readdirSync(dir, { withFileTypes: true }); } catch { return out; }
  for (const e of entries) {
    if (e.name.startsWith('.') || e.name.startsWith('~$') || e.name === 'node_modules') continue;
    const p = path.join(dir, e.name);
    if (e.isDirectory()) walk(p, out);
    else if (SUPPORTED.includes(path.extname(e.name).toLowerCase())) out.push(p);
  }
  return out;
}

class Index {
  constructor() {
    this.folder = '';
    this.chunks = [];
    this.search = null;
    this.status = { state: 'idle', files: 0, chunks: 0, errors: [] };
  }

  async build(folder, onProgress = () => {}) {
    this.folder = folder;
    this.status = { state: 'indexing', files: 0, chunks: 0, done: 0, errors: [] };
    let cache = {};
    try { cache = JSON.parse(fs.readFileSync(CACHE, 'utf8')); } catch {}
    const files = walk(folder);
    this.status.files = files.length;
    const fresh = {};
    for (const f of files) {
      const st = fs.statSync(f);
      const key = st.mtimeMs + ':' + st.size;
      if (cache[f] && cache[f].key === key) {
        fresh[f] = cache[f];
      } else {
        try {
          fresh[f] = { key, chunks: await extractFile(f) };
        } catch (err) {
          this.status.errors.push(path.relative(folder, f) + ': ' + err.message);
          fresh[f] = { key, chunks: [] };
        }
      }
      this.status.done++;
      onProgress(this.status);
    }
    fs.mkdirSync(path.dirname(CACHE), { recursive: true });
    fs.writeFileSync(CACHE, JSON.stringify(fresh));

    this.chunks = [];
    for (const [f, v] of Object.entries(fresh)) {
      for (const c of v.chunks) this.chunks.push({ id: this.chunks.length, file: path.relative(folder, f), ...c });
    }
    this.search = new MiniSearch({
      fields: ['text', 'file'],
      storeFields: [],
      tokenize,
      processTerm: t => t,
      searchOptions: { processTerm: t => (STOP.has(t) ? null : t) }
    });
    this.search.addAll(this.chunks);
    this.status = { ...this.status, state: 'ready', chunks: this.chunks.length };
    return this.status;
  }

  query(q, limit = 30) {
    if (!this.search || !q.trim()) return [];
    return this.search
      .search(q, { prefix: true, fuzzy: 0.15, combineWith: 'OR', boost: { file: 2 } })
      .slice(0, limit)
      .map(r => ({ ...this.chunks[r.id], score: r.score, terms: r.terms }));
  }
}

module.exports = { Index, tokenize };
