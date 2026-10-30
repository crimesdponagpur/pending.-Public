// ऐच्छिक: Ollama (Mac वरच चालणारे फुकट AI) आणि फुकट वेब शोध (DuckDuckGo)
const OLLAMA = process.env.OLLAMA_URL || 'http://127.0.0.1:11434';

async function ollamaModels() {
  try {
    const r = await fetch(OLLAMA + '/api/tags', { signal: AbortSignal.timeout(1500) });
    const j = await r.json();
    return (j.models || []).map(m => m.name);
  } catch {
    return [];
  }
}

const FORMAT_HINT = {
  table: 'उत्तर फक्त Markdown तक्त्याच्या (| स्तंभ | स्तंभ |) स्वरूपात द्या. पहिली ओळ स्तंभांची नावे.',
  doc: 'उत्तर व्यवस्थित अहवालासारखे द्या: शीर्षक (# ने), उपशीर्षके (## ने), मुद्दे (- ने). गरज असेल तर Markdown तक्ता वापरा.'
};

async function ollamaAnswer({ model, request, hits, web, kind }) {
  const sources = hits.map((h, i) => `[${i + 1}] (${h.file}, ${h.where})\n${h.text}`).join('\n\n');
  const webText = (web || []).map((w, i) => `[W${i + 1}] ${w.title}: ${w.snippet} (${w.url})`).join('\n');
  const prompt =
`तुम्ही एक मदतनीस आहात. खाली दिलेल्या माहितीच्या आधारेच उत्तर द्या. माहितीत नसलेली गोष्ट स्वतःहून बनवू नका; नसेल तर "माहिती उपलब्ध नाही" असे लिहा.
वापरकर्ता ज्या भाषेत विचारतो (मराठी/इंग्रजी) त्याच भाषेत उत्तर द्या. प्रत्येक मुद्द्यानंतर स्रोत क्रमांक [1], [2] असा लिहा.
${FORMAT_HINT[kind] || FORMAT_HINT.doc}

### फोल्डरमधील माहिती
${sources || '(काही सापडले नाही)'}
${webText ? '\n### वेबवरील माहिती\n' + webText : ''}

### वापरकर्त्याची विनंती
${request}`;
  const r = await fetch(OLLAMA + '/api/generate', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ model, prompt, stream: false, options: { temperature: 0.2, num_ctx: 8192 } })
  });
  if (!r.ok) throw new Error('Ollama: ' + r.status + ' ' + (await r.text()));
  return (await r.json()).response.trim();
}

// DuckDuckGo चे HTML पान (API key लागत नाही). जमले नाही तर रिकामी यादी.
async function webSearch(q, n = 6) {
  try {
    const r = await fetch('https://html.duckduckgo.com/html/?q=' + encodeURIComponent(q), {
      headers: { 'User-Agent': 'Mozilla/5.0 (Macintosh)' },
      signal: AbortSignal.timeout(8000)
    });
    const html = await r.text();
    const out = [];
    const re = /<a[^>]+class="result__a"[^>]+href="([^"]+)"[^>]*>([\s\S]*?)<\/a>[\s\S]*?class="result__snippet"[^>]*>([\s\S]*?)<\/a>/g;
    const strip = s => s.replace(/<[^>]+>/g, '').replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&#x27;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').trim();
    let m;
    while ((m = re.exec(html)) && out.length < n) {
      let url = m[1];
      const u = url.match(/uddg=([^&]+)/);
      if (u) url = decodeURIComponent(u[1]);
      out.push({ title: strip(m[2]), snippet: strip(m[3]), url });
    }
    return out;
  } catch {
    return [];
  }
}

module.exports = { ollamaModels, ollamaAnswer, webSearch };
