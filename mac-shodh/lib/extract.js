// फाइलमधून मजकूर काढणे: txt, md, csv, pdf, docx, xlsx/xls
const fs = require('fs');
const path = require('path');
const pdfParse = require('pdf-parse/lib/pdf-parse.js');
const mammoth = require('mammoth');
const XLSX = require('xlsx');

const SUPPORTED = ['.txt', '.md', '.csv', '.pdf', '.docx', '.xlsx', '.xls', '.json', '.html', '.htm'];
const CHUNK = 900;

// मोठा मजकूर छोट्या तुकड्यांत (परिच्छेद तुटू नयेत असा प्रयत्न)
function chunkText(text, where) {
  const out = [];
  const paras = text.replace(/\r/g, '').split(/\n\s*\n/);
  let buf = '';
  for (const p of paras) {
    const t = p.replace(/[ \t]+/g, ' ').trim();
    if (!t) continue;
    if (buf && buf.length + t.length > CHUNK) { out.push({ where, text: buf }); buf = ''; }
    if (t.length > CHUNK * 2) {
      for (let i = 0; i < t.length; i += CHUNK) out.push({ where, text: t.slice(i, i + CHUNK) });
    } else {
      buf = buf ? buf + '\n' + t : t;
    }
  }
  if (buf) out.push({ where, text: buf });
  return out;
}

async function extractPdf(file) {
  const pages = [];
  await pdfParse(fs.readFileSync(file), {
    pagerender: async (pageData) => {
      const tc = await pageData.getTextContent();
      let last, text = '';
      for (const item of tc.items) {
        if (last !== undefined && last !== item.transform[5]) text += '\n';
        text += item.str;
        last = item.transform[5];
      }
      pages.push(text);
      return text;
    }
  });
  return pages.flatMap((t, i) => chunkText(t, 'पान ' + (i + 1)));
}

// तारीख नेहमी YYYY-MM-DD, म्हणजे वर्ष/महिना काढता येतो
const pad = n => String(n).padStart(2, '0');
const cellText = v => v instanceof Date
  ? `${v.getFullYear()}-${pad(v.getMonth() + 1)}-${pad(v.getDate())}`
  : typeof v === 'number' ? String(Math.round(v * 1e6) / 1e6) : String(v ?? '');

// Excel: प्रत्येक ओळ स्वतंत्र तुकडा, स्तंभांच्या नावांसह (म्हणजे Excel आउटपुटमध्ये मूळ स्तंभ परत मिळतात)
function extractSheet(file) {
  const wb = XLSX.readFile(file, { cellDates: true });
  const out = [];
  for (const name of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '', raw: true })
      .map(r => r.map(cellText));
    // शीर्षक-ओळ: सुरुवातीच्या ओळींपैकी जवळपास सर्व स्तंभ भरलेली पहिली ओळ
    const filled = rows.slice(0, 20).map(r => r.filter(c => String(c).trim()).length);
    const most = Math.max(0, ...filled);
    const hi = most < 2 ? -1 : filled.findIndex(n => n >= Math.max(2, most * 0.6));
    if (hi < 0) continue;
    const headers = rows[hi].map((h, i) => String(h).trim() || 'स्तंभ ' + (i + 1));
    for (let r = hi + 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row.some(c => String(c).trim())) continue;
      const cells = headers.map((_, i) => String(row[i] ?? '').trim());
      out.push({
        where: name + ' / ओळ ' + (r + 1),
        text: headers.map((h, i) => cells[i] ? h + ': ' + cells[i] : '').filter(Boolean).join(' | '),
        sheet: name,
        headers,
        cells
      });
    }
  }
  return out;
}

async function extractFile(file) {
  const ext = path.extname(file).toLowerCase();
  if (ext === '.pdf') return extractPdf(file);
  if (ext === '.docx') {
    const { value } = await mammoth.extractRawText({ path: file });
    return chunkText(value, 'मजकूर');
  }
  if (ext === '.xlsx' || ext === '.xls' || ext === '.csv') return extractSheet(file);
  let text = fs.readFileSync(file, 'utf8');
  if (ext === '.html' || ext === '.htm') text = text.replace(/<script[\s\S]*?<\/script>|<style[\s\S]*?<\/style>/gi, ' ').replace(/<[^>]+>/g, ' ');
  return chunkText(text, 'मजकूर');
}

module.exports = { extractFile, SUPPORTED };
