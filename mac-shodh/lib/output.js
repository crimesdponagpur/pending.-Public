// निकालावरून Word / Excel / PDF (छपाई) बनवणे
const ExcelJS = require('exceljs');
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

// अहवालाचे भाग: शीर्षक, (chart व मोजणी तक्ता), AI चे लिखाण, शोध-निकाल
function reportBlocks(rep) {
  const blocks = [{ type: 'h', level: 1, text: rep.analysis ? rep.analysis.title : rep.title }];
  const a = rep.analysis;
  if (a) {
    blocks.push({ type: 'p', text: a.desc, small: true });
    if (rep.chartPng) blocks.push({ type: 'img', png: rep.chartPng });
    blocks.push({ type: 'table', rows: a.table.map(r => r.map(String)) });
    if (rep.answer) blocks.push({ type: 'h', level: 2, text: 'निरीक्षण' });
  }
  if (rep.answer) {
    const md = parseMarkdown(rep.answer);
    // AI ने स्वतःचे मुख्य शीर्षक दिले असेल तर दोनदा नको
    if (!a && md[0] && md[0].type === 'h' && md[0].level === 1) blocks.shift();
    blocks.push(...md);
    return blocks;
  }
  if (a) return blocks;
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
  if (rep.analysis) return [...new Set(rep.analysis.rows.map(r => r[r.length - 1]))].map(s => '• ' + s);
  const out = rep.hits.map((h, i) => `[${i + 1}] ${h.file} (${h.where})`);
  (rep.web || []).forEach((w, i) => out.push(`[W${i + 1}] ${w.title}: ${w.url}`));
  return out;
}

async function toDocx(rep) {
  const children = [];
  const run = (text, o = {}) => new D.TextRun({ text: clean(text), font: FONT, ...o });
  const HL = [D.HeadingLevel.TITLE, D.HeadingLevel.HEADING_1, D.HeadingLevel.HEADING_2, D.HeadingLevel.HEADING_3];
  for (const b of reportBlocks(rep)) {
    if (b.type === 'h') children.push(new D.Paragraph({ heading: HL[Math.min(b.level, 4) - 1], children: [run(b.text)] }));
    else if (b.type === 'li') children.push(new D.Paragraph({ bullet: { level: 0 }, children: [run(b.text)] }));
    else if (b.type === 'p') children.push(new D.Paragraph({ children: [run(b.text, b.small ? { italics: true, color: '666666', size: 18 } : {})] }));
    else if (b.type === 'img') children.push(new D.Paragraph({ children: [new D.ImageRun({ type: 'png', data: b.png, transformation: { width: 600, height: 330 } })] }));
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
  return D.Packer.toBuffer(new D.Document({ sections: [{ children }] }));
}

async function toXlsx(rep) {
  const wb = new ExcelJS.Workbook();
  const used = new Set();
  const add = (name, aoa) => {
    const n = name.replace(/[\\/?*[\]:]/g, ' ').slice(0, 28) || 'शीट';
    let k = n, i = 2;
    while (used.has(k)) k = n.slice(0, 25) + ' ' + i++;
    used.add(k);
    const ws = wb.addWorksheet(k);
    aoa.forEach(r => ws.addRow(r.map(v => (typeof v === 'number' ? v : clean(v ?? '')))));
    if (aoa.length) {
      ws.getRow(1).font = { bold: true };
      ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8EEF7' } };
      ws.views = [{ state: 'frozen', ySplit: 1 }];
      ws.columns.forEach((col, c) => { col.width = Math.min(50, Math.max(10, ...aoa.map(r => String(r[c] ?? '').length + 2))); });
    }
    return ws;
  };
  const a = rep.analysis;
  if (a) {
    const ws = add('तुलना', a.table);
    if (rep.chartPng) {
      const img = wb.addImage({ buffer: rep.chartPng, extension: 'png' });
      ws.addImage(img, { tl: { col: a.table[0].length + 1, row: 0 }, ext: { width: 640, height: 352 } });
    }
    ws.addRow([]);
    ws.addRow([a.desc]).font = { italic: true, color: { argb: 'FF666666' } };
    if (rep.answer) add('निरीक्षण', rep.answer.split('\n').filter(l => l.trim()).map(l => [clean(l)]));
    add('मूळ ओळी', [a.rowsHeaders, ...a.rows]);
  } else if (rep.answer) {
    const tables = parseMarkdown(rep.answer).filter(b => b.type === 'table');
    if (tables.length) tables.forEach((t, i) => add('उत्तर ' + (i + 1), t.rows));
    else add('उत्तर', rep.answer.split('\n').map(l => [l]));
  }
  if (!a) {
    for (const g of groupRows(rep.hits)) {
      add(g.rows[0].sheet, [[...g.headers, 'फाइल', 'ठिकाण'], ...g.rows.map(r => [...r.cells, r.file, r.where])]);
    }
    add('सर्व निकाल', [['क्र.', 'फाइल', 'ठिकाण', 'मजकूर'], ...rep.hits.map((h, i) => [i + 1, h.file, h.where, h.text])]);
    if (rep.web && rep.web.length) add('वेब', [['शीर्षक', 'सारांश', 'लिंक'], ...rep.web.map(w => [w.title, w.snippet, w.url])]);
  }
  return Buffer.from(await wb.xlsx.writeBuffer());
}

const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

// PDF: Chromium मराठी अक्षरे बरोबर जोडतो, म्हणून हे पान PDF मध्ये छापले जाते (ॲपमध्ये थेट PDF, ब्राउझरमध्ये "Save as PDF")
function toPrintHtml(rep, autoPrint = true) {
  const blocks = reportBlocks(rep);
  let body = '';
  let inList = false;
  for (const b of blocks) {
    if (b.type !== 'li' && inList) { body += '</ul>'; inList = false; }
    if (b.type === 'h') body += `<h${b.level}>${esc(clean(b.text))}</h${b.level}>`;
    else if (b.type === 'li') { if (!inList) { body += '<ul>'; inList = true; } body += `<li>${esc(clean(b.text))}</li>`; }
    else if (b.type === 'p') body += `<p${b.small ? ' class="sm"' : ''}>${esc(clean(b.text))}</p>`;
    else if (b.type === 'img') body += `<img src="data:image/png;base64,${b.png.toString('base64')}">`;
    else if (b.type === 'table') body += '<table>' + b.rows.map((r, i) => '<tr>' + r.map(c => i ? `<td>${esc(clean(c))}</td>` : `<th>${esc(clean(c))}</th>`).join('') + '</tr>').join('') + '</table>';
  }
  if (inList) body += '</ul>';
  body += '<h2>स्रोत</h2>' + sourceLines(rep).map(s => `<p class="sm">${esc(s)}</p>`).join('');
  const wide = blocks.some(b => b.type === 'table' && b.rows[0] && b.rows[0].length > 6);
  return `<!doctype html><html lang="mr"><head><meta charset="utf-8"><title>${esc(rep.title)}</title>
<style>
body{font-family:"Kohinoor Devanagari","Noto Sans Devanagari",system-ui,sans-serif;max-width:800px;margin:24px auto;padding:0 16px;color:#111;background:#fff;line-height:1.6}
h1{font-size:22px;border-bottom:2px solid #333;padding-bottom:6px}h2{font-size:18px;margin-top:22px}h3,h4{font-size:16px}
table{border-collapse:collapse;width:100%;margin:8px 0;font-size:12px}th,td{border:1px solid #999;padding:4px 6px;text-align:left;vertical-align:top}th{background:#eef}
img{max-width:100%;margin:8px 0}
.sm{font-size:11px;color:#555;margin:2px 0}.bar{position:sticky;top:0;background:#fff;padding:8px 0;text-align:right}
@media print{.bar{display:none}body{margin:0;max-width:none}}${wide ? '@page{size:A4 landscape}' : ''}
</style></head><body>${autoPrint ? '<div class="bar"><button onclick="print()">PDF म्हणून जतन करा / छापा</button></div>' : ''}${body}
${autoPrint ? '<script>setTimeout(()=>print(),400)</script>' : ''}</body></html>`;
}

module.exports = { toDocx, toXlsx, toPrintHtml, parseMarkdown };
