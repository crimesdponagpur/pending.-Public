// निकालावरून Word / Excel / PDF (छपाई) बनवणे
const XLSX = require('xlsx');
const D = require('docx');

const FONT = { ascii: 'Calibri', hAnsi: 'Calibri', cs: 'Kohinoor Devanagari', eastAsia: 'Kohinoor Devanagari' };

// सोपा Markdown वाचक: # शीर्षक, - मुद्दा, | तक्ता |, बाकी परिच्छेद
function parseMarkdown(md) {
  const blocks = [];
  const lines = (md || '').replace(/\r/g, '').split('\n');
  for (let i = 0; i < lines.length; i++) {
    const l = lines[i];
    if (/^\s*\|/.test(l)) {
      const rows = [];
      while (i < lines.length && /^\s*\|/.test(lines[i])) {
        const cells = lines[i].trim().replace(/^\||\|$/g, '').split('|').map(c => c.trim());
        if (!cells.every(c => /^:?-{2,}:?$/.test(c))) rows.push(cells);
        i++;
      }
      i--;
      blocks.push({ type: 'table', rows });
    } else if (/^#{1,6}\s/.test(l)) {
      blocks.push({ type: 'h', level: l.match(/^#+/)[0].length, text: l.replace(/^#+\s*/, '') });
    } else if (/^\s*([-*•]|\d+[.)])\s+/.test(l)) {
      blocks.push({ type: 'li', text: l.replace(/^\s*([-*•]|\d+[.)])\s+/, '') });
    } else if (l.trim()) {
      blocks.push({ type: 'p', text: l.trim() });
    }
  }
  return blocks;
}

const clean = s => String(s).replace(/\*\*(.+?)\*\*/g, '$1').replace(/`/g, '');

// Excel ओळींचे निकाल: एकाच शीटचे (एकाच स्तंभांचे) निकाल एकत्र
function groupRows(hits) {
  const groups = new Map();
  for (const h of hits) {
    if (!h.headers) continue;
    const key = h.file + ' › ' + h.sheet;
    if (!groups.has(key)) groups.set(key, { key, headers: h.headers, rows: [] });
    groups.get(key).rows.push(h);
  }
  return [...groups.values()];
}

// AI नसेल तेव्हा निकालांवरूनच अहवाल
function reportBlocks(rep) {
  if (rep.answer) return parseMarkdown(rep.answer);
  const blocks = [{ type: 'h', level: 1, text: rep.title }];
  if (!rep.hits.length) blocks.push({ type: 'p', text: 'फोल्डरमध्ये काही सापडले नाही.' });
  for (const g of groupRows(rep.hits)) {
    blocks.push({ type: 'h', level: 2, text: g.key });
    blocks.push({ type: 'table', rows: [g.headers, ...g.rows.map(r => r.cells)] });
  }
  const byFile = new Map();
  for (const h of rep.hits.filter(h => !h.headers)) {
    if (!byFile.has(h.file)) byFile.set(h.file, []);
    byFile.get(h.file).push(h);
  }
  for (const [file, hs] of byFile) {
    blocks.push({ type: 'h', level: 2, text: file });
    for (const h of hs) {
      blocks.push({ type: 'p', text: '(' + h.where + ')', small: true });
      for (const line of h.text.split('\n')) blocks.push({ type: 'p', text: line });
    }
  }
  return blocks;
}

function sourceLines(rep) {
  const out = rep.hits.map((h, i) => `[${i + 1}] ${h.file} (${h.where})`);
  (rep.web || []).forEach((w, i) => out.push(`[W${i + 1}] ${w.title}: ${w.url}`));
  return out;
}

async function toDocx(rep) {
  const children = [];
  const run = (text, o = {}) => new D.TextRun({ text: clean(text), font: FONT, ...o });
  const HL = [D.HeadingLevel.HEADING_1, D.HeadingLevel.HEADING_2, D.HeadingLevel.HEADING_3];
  const blocks = reportBlocks(rep);
  if (rep.answer) children.push(new D.Paragraph({ heading: D.HeadingLevel.TITLE, children: [run(rep.title)] }));
  for (const b of blocks) {
    if (b.type === 'h') children.push(new D.Paragraph({ heading: HL[Math.min(b.level, 3) - 1], children: [run(b.text)] }));
    else if (b.type === 'li') children.push(new D.Paragraph({ bullet: { level: 0 }, children: [run(b.text)] }));
    else if (b.type === 'p') children.push(new D.Paragraph({ children: [run(b.text, b.small ? { italics: true, color: '666666', size: 18 } : {})] }));
    else if (b.type === 'table' && b.rows.length) {
      const n = Math.max(...b.rows.map(r => r.length));
      children.push(new D.Table({
        width: { size: 100, type: D.WidthType.PERCENTAGE },
        rows: b.rows.map((r, ri) => new D.TableRow({
          tableHeader: ri === 0,
          children: Array.from({ length: n }, (_, ci) => new D.TableCell({
            shading: ri === 0 ? { fill: 'E8EEF7', type: D.ShadingType.CLEAR, color: 'auto' } : undefined,
            children: [new D.Paragraph({ children: [run(r[ci] ?? '', { bold: ri === 0, size: 18 })] })]
          }))
        }))
      }));
      children.push(new D.Paragraph(''));
    }
  }
  children.push(new D.Paragraph({ heading: D.HeadingLevel.HEADING_2, children: [run('स्रोत')] }));
  for (const s of sourceLines(rep)) children.push(new D.Paragraph({ children: [run(s, { size: 16, color: '666666' })] }));
  const doc = new D.Document({ sections: [{ children }] });
  return D.Packer.toBuffer(doc);
}

function toXlsx(rep) {
  const wb = XLSX.utils.book_new();
  const used = new Set();
  const add = (name, aoa) => {
    let n = name.replace(/[\\/?*[\]:]/g, ' ').slice(0, 28) || 'शीट';
    let k = n, i = 2;
    while (used.has(k)) k = n.slice(0, 25) + ' ' + i++;
    used.add(k);
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = (aoa[0] || []).map((_, c) => ({ wch: Math.min(50, Math.max(10, ...aoa.map(r => String(r[c] ?? '').length))) }));
    XLSX.utils.book_append_sheet(wb, ws, k);
  };
  if (rep.answer) {
    const tables = parseMarkdown(rep.answer).filter(b => b.type === 'table');
    if (tables.length) tables.forEach((t, i) => add('उत्तर ' + (i + 1), t.rows.map(r => r.map(clean))));
    else add('उत्तर', rep.answer.split('\n').map(l => [clean(l)]));
  }
  for (const g of groupRows(rep.hits)) {
    add(g.rows[0].sheet, [[...g.headers, 'फाइल', 'ठिकाण'], ...g.rows.map(r => [...r.cells, r.file, r.where])]);
  }
  add('सर्व निकाल', [['क्र.', 'फाइल', 'ठिकाण', 'मजकूर'], ...rep.hits.map((h, i) => [i + 1, h.file, h.where, h.text])]);
  if (rep.web && rep.web.length) add('वेब', [['शीर्षक', 'सारांश', 'लिंक'], ...rep.web.map(w => [w.title, w.snippet, w.url])]);
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

// PDF: ब्राउझर स्वतः मराठी अक्षरे बरोबर जोडतो, म्हणून छपाई-योग्य पान बनवून "Save as PDF"
function toPrintHtml(rep) {
  const blocks = reportBlocks(rep);
  let body = rep.answer ? `<h1>${esc(rep.title)}</h1>` : '';
  let inList = false;
  for (const b of blocks) {
    if (b.type !== 'li' && inList) { body += '</ul>'; inList = false; }
    if (b.type === 'h') body += `<h${Math.min(b.level + (rep.answer ? 1 : 0), 4)}>${esc(clean(b.text))}</h${Math.min(b.level + (rep.answer ? 1 : 0), 4)}>`;
    else if (b.type === 'li') { if (!inList) { body += '<ul>'; inList = true; } body += `<li>${esc(clean(b.text))}</li>`; }
    else if (b.type === 'p') body += `<p${b.small ? ' class="sm"' : ''}>${esc(clean(b.text))}</p>`;
    else if (b.type === 'table') body += '<table>' + b.rows.map((r, i) => '<tr>' + r.map(c => i ? `<td>${esc(clean(c))}</td>` : `<th>${esc(clean(c))}</th>`).join('') + '</tr>').join('') + '</table>';
  }
  if (inList) body += '</ul>';
  body += '<h2>स्रोत</h2>' + sourceLines(rep).map(s => `<p class="sm">${esc(s)}</p>`).join('');
  const wide = blocks.some(b => b.type === 'table' && b.rows[0] && b.rows[0].length > 6);
  return `<!doctype html><html lang="mr"><head><meta charset="utf-8"><title>${esc(rep.title)}</title>
<style>
body{font-family:"Kohinoor Devanagari","Noto Sans Devanagari",system-ui,sans-serif;max-width:800px;margin:24px auto;padding:0 16px;color:#111;line-height:1.6}
h1{font-size:22px;border-bottom:2px solid #333;padding-bottom:6px}h2{font-size:18px;margin-top:22px}h3{font-size:16px}
table{border-collapse:collapse;width:100%;margin:8px 0;font-size:12px}th,td{border:1px solid #999;padding:4px 6px;text-align:left;vertical-align:top}th{background:#eef}
.sm{font-size:11px;color:#555;margin:2px 0}.bar{position:sticky;top:0;background:#fff;padding:8px 0;text-align:right}
@media print{.bar{display:none}body{margin:0;max-width:none}}${wide ? '@page{size:A4 landscape}' : ''}
</style></head><body><div class="bar"><button onclick="print()">PDF म्हणून जतन करा / छापा</button></div>${body}
<script>setTimeout(()=>print(),400)</script></body></html>`;
}

module.exports = { toDocx, toXlsx, toPrintHtml, parseMarkdown };
