// Excel मधील ओळींवर मोजणी/तुलना. AI फक्त "काय मोजायचे" ठरवते; आकडे इथे खऱ्या ओळींवरून मोजले जातात.

const DEV = '०१२३४५६७८९';
const norm = s => String(s ?? '').replace(/[०-९]/g, d => DEV.indexOf(d)).toLowerCase().replace(/\s+/g, ' ').trim();
const SOURCE = 'फाइल';

// एकाच स्तंभांच्या सर्व शीट्स (उदा. दर महिन्याच्या फाइल्स) एकाच तक्त्यात; "फाइल" हा जादा स्तंभ
function buildTables(chunks) {
  const groups = new Map();
  for (const c of chunks) {
    if (!c.headers) continue;
    const sig = c.headers.join('\u0001');
    if (!groups.has(sig)) groups.set(sig, { headers: [...c.headers, SOURCE], rows: [], sources: new Set() });
    const g = groups.get(sig);
    const src = c.file + ' › ' + c.sheet;
    g.sources.add(src);
    g.rows.push([...c.cells, src]);
  }
  return [...groups.values()]
    .sort((a, b) => b.rows.length - a.rows.length)
    .map((g, i) => ({ id: 'T' + (i + 1), headers: g.headers, rows: g.rows, sources: [...g.sources] }));
}

function profile(t, maxCols = 40) {
  const cols = t.headers.slice(0, maxCols).map((h, ci) => {
    const counts = new Map();
    let filled = 0;
    for (const r of t.rows) {
      const v = String(r[ci] ?? '').trim();
      if (!v) continue;
      filled++;
      counts.set(v, (counts.get(v) || 0) + 1);
    }
    const top = [...counts.entries()].sort((a, b) => b[1] - a[1]).slice(0, 6).map(([v]) => v.length > 40 ? v.slice(0, 40) + '…' : v);
    return `  - "${h}" (${counts.size} वेगवेगळ्या किमती; उदा: ${top.map(v => JSON.stringify(v)).join(', ')})`;
  });
  const src = t.sources.length > 8 ? t.sources.slice(0, 8).join('; ') + ` … (एकूण ${t.sources.length})` : t.sources.join('; ');
  return `${t.id}: ${t.rows.length} ओळी. स्रोत: ${src}\n  स्तंभ:\n${cols.join('\n')}`;
}

function catalog(tables, max = 12) {
  return tables.slice(0, max).map(t => profile(t)).join('\n\n');
}

// स्तंभाच्या किमतीतून वर्ष / महिना / किंमत
function derive(v, how) {
  const s = norm(v);
  if (!how || how === 'value') return String(v ?? '').trim();
  let m = s.match(/\b((?:19|20)\d{2})-(\d{1,2})-(\d{1,2})\b/);
  let y, mo;
  if (m) { y = m[1]; mo = m[2]; }
  else if ((m = s.match(/\b(\d{1,2})[./-](\d{1,2})[./-]((?:19|20)\d{2}|\d{2})\b/))) {
    y = m[3].length === 2 ? '20' + m[3] : m[3]; mo = m[2];
  } else if ((m = s.match(/\b((?:19|20)\d{2})\b/))) {
    y = m[1];
  } else if ((m = s.match(/\/(\d{2})$/))) {
    y = '20' + m[1];
  }
  if (how === 'year') return y || '';
  if (how === 'month') return y && mo ? `${y}-${String(mo).padStart(2, '0')}` : '';
  return String(v ?? '').trim();
}

const num = v => {
  const n = parseFloat(norm(v).replace(/,/g, ''));
  return Number.isFinite(n) ? n : null;
};

// AI ने दिलेले स्तंभाचे नाव थोडे वेगळे असले तरी शोधणे
function resolveCol(t, name) {
  if (name == null || name === '') return -1;
  const n = norm(name);
  let i = t.headers.findIndex(h => norm(h) === n);
  if (i < 0) i = t.headers.findIndex(h => norm(h).includes(n) || n.includes(norm(h)));
  if (i < 0) throw new Error(`"${name}" हा स्तंभ ${t.id} मध्ये नाही`);
  return i;
}

function testFilter(r, f) {
  const a = norm(derive(r[f.ci], f.derive));
  const vals = (Array.isArray(f.value) ? f.value : [f.value]).map(norm).filter(v => v !== '');
  switch (f.op) {
    case 'equals': return vals.some(v => a === v);
    case 'not_contains': return !vals.some(v => a.includes(v));
    case 'not_equals': return !vals.some(v => a === v);
    case 'gte': return num(a) !== null && num(a) >= num(vals[0]);
    case 'lte': return num(a) !== null && num(a) <= num(vals[0]);
    case 'not_empty': return a !== '';
    default: return vals.some(v => a.includes(v));
  }
}

const OPS = { contains: 'मध्ये', equals: '=', not_contains: 'मध्ये नाही', not_equals: '≠', gte: '≥', lte: '≤', not_empty: 'रिकामे नाही' };
const DER = { year: 'वर्ष', month: 'महिना' };
const colLabel = (t, c) => `"${t.headers[c.ci]}"` + (DER[c.derive] ? ` (${DER[c.derive]})` : '');

function runPlan(tables, plan) {
  const t = tables.find(x => x.id === plan.table) || tables[0];
  if (!t) throw new Error('फोल्डरमध्ये Excel तक्ता सापडला नाही');
  const filters = (plan.filters || []).filter(f => f && f.column).map(f => ({ ...f, ci: resolveCol(t, f.column) }));
  const g = plan.groupBy && plan.groupBy.column ? { ...plan.groupBy, ci: resolveCol(t, plan.groupBy.column) } : null;
  const s = plan.seriesBy && plan.seriesBy.column ? { ...plan.seriesBy, ci: resolveCol(t, plan.seriesBy.column) } : null;
  const metric = plan.metric || { type: 'count' };
  const mi = metric.type !== 'count' && metric.column ? resolveCol(t, metric.column) : -1;

  const rows = t.rows.filter(r => filters.every(f => testFilter(r, f)));
  const cell = new Map();
  const cats = new Map(), sers = new Map();
  for (const r of rows) {
    const c = g ? derive(r[g.ci], g.derive) || '(रिकामे)' : 'एकूण';
    const se = s ? derive(r[s.ci], s.derive) || '(रिकामे)' : (metric.type === 'count' ? 'संख्या' : metric.type === 'sum' ? 'बेरीज' : 'सरासरी');
    cats.set(c, (cats.get(c) || 0) + 1);
    sers.set(se, true);
    const k = c + '\u0001' + se;
    const cur = cell.get(k) || { n: 0, sum: 0 };
    cur.n++;
    if (mi >= 0) cur.sum += num(r[mi]) || 0;
    cell.set(k, cur);
  }
  const val = (c, se) => {
    const x = cell.get(c + '\u0001' + se);
    if (!x) return 0;
    if (metric.type === 'sum') return Math.round(x.sum * 100) / 100;
    if (metric.type === 'avg') return Math.round((x.sum / x.n) * 100) / 100;
    return x.n;
  };
  const byLabel = (a, b) => a.localeCompare(b, 'mr', { numeric: true });
  let series = [...sers.keys()].sort(byLabel);
  let categories = [...cats.keys()];
  const total = c => series.reduce((a, se) => a + val(c, se), 0);
  if (plan.sort === 'label' || (g && (g.derive === 'year' || g.derive === 'month'))) categories.sort(byLabel);
  else if (plan.sort === 'asc') categories.sort((a, b) => total(a) - total(b));
  else categories.sort((a, b) => total(b) - total(a));
  const top = Math.max(1, Math.min(Number(plan.top) || 25, 200));
  if (categories.length > top) categories = categories.slice(0, top);

  const table = [[g ? t.headers[g.ci] : '', ...series, ...(series.length > 1 ? ['एकूण'] : [])]];
  for (const c of categories) {
    const vals = series.map(se => val(c, se));
    table.push([c, ...vals, ...(series.length > 1 ? [vals.reduce((a, b) => a + b, 0)] : [])]);
  }

  const desc = [
    `तक्ता ${t.id} (${t.rows.length} ओळी)`,
    filters.length ? 'फिल्टर: ' + filters.map(f => `${colLabel(t, f)} ${OPS[f.op] || 'मध्ये'} ${f.op === 'not_empty' ? '' : JSON.stringify(f.value)}`).join(', ') : 'फिल्टर नाही',
    `जुळलेल्या ओळी: ${rows.length}`,
    g ? 'गट: ' + colLabel(t, g) : null,
    s ? 'तुलना: ' + colLabel(t, s) : null,
    'मोजणी: ' + (metric.type === 'count' ? 'ओळींची संख्या' : (metric.type === 'sum' ? 'बेरीज ' : 'सरासरी ') + `"${t.headers[mi]}"`)
  ].filter(Boolean).join(' · ');

  return {
    title: plan.title || 'तुलना',
    chart: ['bar', 'line', 'pie', 'none'].includes(plan.chart) ? plan.chart : 'bar',
    categories,
    series: series.map(se => ({ name: se, values: categories.map(c => val(c, se)) })),
    table,
    desc,
    matched: rows.length,
    rowsHeaders: t.headers,
    rows: rows.slice(0, 5000)
  };
}

module.exports = { buildTables, catalog, runPlan, derive };
