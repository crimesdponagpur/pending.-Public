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
  return generate(model, prompt);
}

async function generate(model, prompt, json = false) {
  const r = await fetch(OLLAMA + '/api/generate', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ model, prompt, stream: false, ...(json ? { format: 'json' } : {}), options: { temperature: json ? 0 : 0.2, num_ctx: 8192 } })
  });
  if (!r.ok) throw new Error('Ollama: ' + r.status + ' ' + (await r.text()));
  return (await r.json()).response.trim();
}

// AI ला Excel तक्त्यांची ओळख देऊन "काय मोजायचे" ते JSON मध्ये विचारणे
async function planRequest({ model, request, catalog }) {
  const prompt =
`You help a Marathi-speaking police office analyse their own Excel data. Decide how to answer the user's request.

Available tables (columns with example values):
${catalog || '(no tables)'}

Reply with ONLY one JSON object, no other text:
{
  "mode": "analysis" | "write",
  "table": "T1",
  "filters": [{"column": "<exact column name>", "op": "contains|equals|not_contains|not_equals|gte|lte|not_empty", "value": "<text or [list]>", "derive": "value|year|month"}],
  "groupBy": {"column": "<column>", "derive": "value|year|month"} or null,
  "seriesBy": {"column": "<column>", "derive": "value|year|month"} or null,
  "metric": {"type": "count|sum|avg", "column": null},
  "chart": "bar|line|pie|none",
  "sort": "desc|asc|label",
  "top": 25,
  "title": "<short title in the user's language>"
}

Rules:
- "analysis" when the user wants counts, comparison, chart, statistics, table or list from the tables. "write" when they want a letter, summary, note, explanation or the data is not in the tables.
- Use column names exactly as listed. Use example values to choose filter values (e.g. section/act text such as "NDPS" goes in a contains filter on the section column).
- "A vs B" comparison (e.g. 2024 vs 2025): put that dimension in seriesBy (derive "year" for a date or crime-number column, or the "फाइल" column if each year is a separate file) and filter to only those values.
- groupBy is the x-axis category (e.g. police station, month). For a comparison by year only, use groupBy with derive "year" and seriesBy null.
- derive "year"/"month" extracts the year/month from dates like 2025-03-14 or crime numbers like 123/2025.
- Digits may be Marathi (२०२५) or English (2025); use English digits.

User request: ${request}`;
  const txt = await generate(model, prompt, true);
  try { return JSON.parse(txt); } catch { throw new Error('AI चे उत्तर समजले नाही: ' + txt.slice(0, 200)); }
}

// मोजलेल्या आकड्यांवरून थोडक्यात निरीक्षण (आकडे बदलू नयेत)
async function summarize({ model, request, result }) {
  const tbl = result.table.map(r => r.join(' | ')).join('\n');
  const prompt =
`खालील तक्ता वापरकर्त्याच्या Excel माहितीवरून अचूक मोजलेला आहे. त्यावरून 3 ते 5 मुद्द्यांत थोडक्यात निरीक्षण लिहा (वाढ/घट, सर्वाधिक, टक्केवारी). आकडे फक्त तक्त्यातील वापरा, स्वतःहून बनवू नका. वापरकर्त्याच्या भाषेत (मराठी असल्यास मराठीत) लिहा. प्रत्येक मुद्दा "- " ने सुरू करा.

विनंती: ${request}
मोजणी कशी केली: ${result.desc}
तक्ता:
${tbl}`;
  return generate(model, prompt);
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

module.exports = { ollamaModels, ollamaAnswer, planRequest, summarize, webSearch };
