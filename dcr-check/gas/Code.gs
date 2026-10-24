/**
 * जिल्हा DCR चेक व B भाग ॲप (नागपूर ग्रामीण) • स्वतंत्र Google Apps Script Web app.
 * रोजचा DCR (Word) → गुन्हे नोंदवही + त्रुटी, गुन्ह्यावर SP शेरा → DCR चेक (तपास सूचना) पत्र Word + PDF,
 * महिन्याचे सर्व DCR → हेडवाइज B भाग Word + PDF. सर्व फाइल्स "DCR_चेक_ॲप" या Drive फोल्डरमध्ये आपोआप जतन.
 * MCR / प्रलंबित ॲपशी याचा काही संबंध नाही: स्वतःचा फोल्डर, स्वतःची Sheet.
 */
var VERSION = '1.0.0';
var ROOT_NAME = 'DCR_चेक_ॲप';
var SS_NAME = 'DCR_नोंदवही';
var CASE_COLS = ['key', 'kind', 'sec', 'ps', 'div', 'no', 'yr', 'kalam', 'label', 'sid', 'head', 'headSet', 'comp', 'acc', 'victim', 'deceased', 'place',
  'occ', 'reg', 'regDate', 'occDate', 'prop', 'io', 'incharge', 'regOff', 'visit', 'facts', 'prior', 'priorAction', 'arrest', 'arrestReason', 'cause',
  'nCases', 'nAcc', 'psCr', 'dcrDate', 'dcrIssue', 'dcrList', 'shera', 'sheraAt', 'letters', 'updated'];
var SHEETS = {
  CASES: CASE_COLS,
  DCRS: ['reportDate', 'issueDate', 'jaKr', 'fileName', 'fileUrl', 'pdfUrl', 'cases', 'bad', 'warn', 'info', 'uploaded'],
  ISSUES: ['reportDate', 'level', 'msg', 'key'],
  LETTERS: ['id', 'jaKr', 'date', 'key', 'ps', 'div', 'crime', 'head', 'shera', 'keys', 'docxUrl', 'krutiUrl', 'pdfUrl', 'created', 'status', 'replyDate', 'note'],
  BPART: ['month', 'scope', 'cases', 'docxUrl', 'pdfUrl', 'created'],
  PS: ['ps', 'div'],
  SETTINGS: ['key', 'value']
};
var DEF = {
  spName: '', spDesig: 'पोलीस अधीक्षक,\nनागपूर जिल्हा (ग्रामीण), नागपूर.',
  office: 'कार्यालय पोलीस अधीक्षक, नागपूर जिल्हा (ग्रामीण)\nसिव्हील लाईन, नागपूर - 440001\nई-मेल: spreader-ngp.r@mahapolice.gov.in • दूरध्वनी क्र. 0712-2566650 • फॅक्स क्र. 0712-2560779',
  jaKrPrefix: 'नाजिग्रा/डी-32/वाशा/दैगुअ तपास सूचना', jaKrFile: '26-1962', autoNo: 'होय', letterSeq: '0', rules: '[]', logoL: '', logoC: '', logoR: '', days: '3', font: 'Mukta',
  divOrder: 'काटोल,सावनेर,नागपूर,उमरेड,रामटेक,कन्हान', distName: 'नागपूर ग्रामीण जिल्हा',
  quick: 'आरोपीस तात्काळ अटक करावी\nउपविभागीय पोलीस अधिकारी यांनी घटनास्थळी भेट द्यावी\nआरोपीविरुद्ध प्रतिबंधक कार्यवाही करावी\nसीसीटीव्ही फुटेज तपासावे\nमुद्देमाल हस्तगत करावा\nगुन्हा लवकर उघडकीस आणावा\nपीडितेचा तात्काळ शोध घ्यावा\nतपास दुसऱ्या अधिकाऱ्याकडे वर्ग करावा',
  paras: ''
};

/* ---------- Web app ---------- */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('App').setTitle('DCR चेक • B भाग')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
/** ॲपमधील सर्व विनंत्या इथून जातात: उत्तर नेहमी { ok, d } किंवा { ok:false, e } */
function api(fn, args) {
  var F = { init: init_, saveDcr: saveDcr_, savePdf: savePdf_, search: search_, saveShera: saveShera_, setHead: setHead_, makeLetter: makeLetter_, makeLetters: makeLetters_, saveLogo: saveLogo_,
    letters: letters_, setLetter: setLetter_, monthCheck: monthCheck_, buildBPart: buildBPart_, saveSettings: saveSettings_, savePs: savePs_ };
  try {
    if (!F[fn]) throw new Error('अज्ञात क्रिया: ' + fn);
    return JSON.stringify({ ok: true, d: F[fn].apply(null, args || []) });
  } catch (e) {
    console.error(fn, e && e.stack);
    return JSON.stringify({ ok: false, e: String(e && e.message || e) });
  }
}

/* ---------- Drive / Sheet ---------- */
function props_() { return PropertiesService.getScriptProperties(); }
function root_() {
  var id = props_().getProperty('ROOT');
  if (id) try { var f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) {}
  var it = DriveApp.getRootFolder().getFoldersByName(ROOT_NAME), r = it.hasNext() ? it.next() : DriveApp.getRootFolder().createFolder(ROOT_NAME);
  props_().setProperty('ROOT', r.getId());
  return r;
}
function folder_(path) {
  var f = root_();
  path.forEach(function (n) { var it = f.getFoldersByName(n); f = it.hasNext() ? it.next() : f.createFolder(n); });
  return f;
}
function ss_() {
  var id = props_().getProperty('SS');
  if (id) try { return SpreadsheetApp.openById(id); } catch (e) {}
  var r = root_(), it = r.getFilesByName(SS_NAME), ss;
  if (it.hasNext()) ss = SpreadsheetApp.open(it.next());
  else { ss = SpreadsheetApp.create(SS_NAME); DriveApp.getFileById(ss.getId()).moveTo(r); }
  props_().setProperty('SS', ss.getId());
  return ss;
}
function sh_(name) {
  var ss = ss_(), s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.getRange(1, 1, 1, SHEETS[name].length).setValues([SHEETS[name]]).setFontWeight('bold');
    s.setFrozenRows(1);
    if (name === 'SETTINGS') s.getRange(2, 1, Object.keys(DEF).length, 2).setValues(Object.keys(DEF).map(function (k) { return [k, DEF[k]]; }));
    var d = ss.getSheetByName('Sheet1') || ss.getSheetByName('पत्रक1');
    if (d && ss.getSheets().length > 1 && d.getLastRow() === 0) ss.deleteSheet(d);
  }
  return s;
}
function rows_(name) {
  var s = sh_(name), n = s.getLastRow();
  if (n < 2) return [];
  var cols = SHEETS[name], v = s.getRange(2, 1, n - 1, cols.length).getDisplayValues();
  return v.map(function (r, i) { var o = { _row: i + 2 }; cols.forEach(function (c, j) { o[c] = r[j]; }); return o; });
}
function toRow_(name, o) { return SHEETS[name].map(function (c) { var v = o[c]; return v === undefined || v === null ? '' : typeof v === 'object' ? JSON.stringify(v) : String(v); }); }
function settings_() {
  var o = {}; Object.keys(DEF).forEach(function (k) { o[k] = DEF[k]; });
  rows_('SETTINGS').forEach(function (r) { if (r.key) o[r.key] = r.value; });
  return o;
}
function setSetting_(k, v) {
  var s = sh_('SETTINGS'), rs = rows_('SETTINGS');
  for (var i = 0; i < rs.length; i++) if (rs[i].key === k) { s.getRange(rs[i]._row, 2).setValue(String(v)); return; }
  s.appendRow([k, String(v)]);
}
function now_() { return Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd HH:mm'); }
function today_() { return Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd'); }
function dmy_(iso) { return iso ? iso.slice(8, 10) + '/' + iso.slice(5, 7) + '/' + iso.slice(0, 4) : ''; }
function saveBlob_(folder, blob) {
  var it = folder.getFilesByName(blob.getName());
  while (it.hasNext()) it.next().setTrashed(true);
  return folder.createFile(blob);
}
function caseObj_(r) { var o = {}; CASE_COLS.forEach(function (c) { o[c] = r[c]; }); o.no = +o.no || o.no; o.yr = +o.yr || o.yr; o.letters = +o.letters || 0; return o; }

/* ---------- सुरुवात ---------- */
function init_() {
  var S = settings_();
  var dcrs = rows_('DCRS').sort(function (a, b) { return a.reportDate < b.reportDate ? 1 : -1; }).slice(0, 62);
  return { version: VERSION, settings: S, dcrs: dcrs, ps: rows_('PS'), rootUrl: root_().getUrl(), ssUrl: ss_().getUrl(), today: today_() };
}

/* ---------- DCR जतन ---------- */
function saveDcr_(p) {
  if (!p || !p.reportDate) throw new Error('DCR ची तारीख नाही');
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var month = p.reportDate.slice(0, 7), file = null;
    if (p.b64) {
      var name = p.saveName || ('CRO_DCR_' + dmy_(p.reportDate).replace(/\//g, '-') + '.docx');
      file = saveBlob_(folder_(['DCR', month]), Utilities.newBlob(Utilities.base64Decode(p.b64), p.mime || 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', name));
    }
    var sh = sh_('CASES'), all = rows_('CASES'), idx = {}, add = [], upd = 0, ts = now_();
    all.forEach(function (r) { idx[r.key] = r; });
    (p.cases || []).forEach(function (c) {
      c.dcrDate = p.reportDate; c.dcrIssue = p.issueDate || ''; c.updated = ts;
      var old = idx[c.key];
      if (old) {
        var list = (old.dcrList || '').split(',').filter(String);
        if (list.indexOf(p.reportDate) < 0) list.push(p.reportDate);
        c.dcrList = list.sort().join(',');
        c.dcrDate = list[0]; c.dcrIssue = old.dcrDate && old.dcrDate < p.reportDate ? old.dcrIssue : c.dcrIssue;
        ['shera', 'sheraAt', 'letters'].forEach(function (k) { c[k] = old[k]; });
        if (old.headSet) { c.head = old.head; c.headSet = old.headSet; }
        sh.getRange(old._row, 1, 1, CASE_COLS.length).setValues([toRow_('CASES', c)]);
        upd++;
      } else { c.dcrList = p.reportDate; add.push(toRow_('CASES', c)); idx[c.key] = c; }
    });
    if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, CASE_COLS.length).setValues(add);
    /* पो.स्टे. → विभाग (DCR वरून शिकते) */
    var psS = sh_('PS'), psR = rows_('PS'), known = {};
    psR.forEach(function (r) { known[r.ps] = r; });
    var newPs = [];
    (p.cases || []).forEach(function (c) { if (c.ps && c.div && !known[c.ps]) { known[c.ps] = { ps: c.ps, div: c.div }; newPs.push([c.ps, c.div]); } });
    if (newPs.length) psS.getRange(psS.getLastRow() + 1, 1, newPs.length, 2).setValues(newPs);
    /* त्रुटी */
    var isS = sh_('ISSUES'), old = rows_('ISSUES').filter(function (r) { return r.reportDate === p.reportDate; }).map(function (r) { return r._row; }).sort(function (a, b) { return b - a; });
    old.forEach(function (r) { isS.deleteRow(r); });
    var iss = (p.issues || []).map(function (i) { return [p.reportDate, i.level, i.msg, i.key || '']; });
    if (iss.length) isS.getRange(isS.getLastRow() + 1, 1, iss.length, 4).setValues(iss);
    /* DCR यादी */
    var cnt = function (l) { return (p.issues || []).filter(function (i) { return i.level === l; }).length; };
    var dS = sh_('DCRS'), dR = rows_('DCRS').filter(function (r) { return r.reportDate === p.reportDate; })[0];
    var row = { reportDate: p.reportDate, issueDate: p.issueDate, jaKr: p.jaKr, fileName: p.fileName, fileUrl: file ? file.getUrl() : (dR ? dR.fileUrl : ''),
      pdfUrl: dR ? dR.pdfUrl : '', cases: (p.cases || []).length, bad: cnt('bad'), warn: cnt('warn'), info: cnt('info'), uploaded: ts };
    if (dR) dS.getRange(dR._row, 1, 1, SHEETS.DCRS.length).setValues([toRow_('DCRS', row)]); else dS.appendRow(toRow_('DCRS', row));
    return { added: add.length, updated: upd, fileUrl: row.fileUrl, month: month };
  } finally { lock.releaseLock(); }
}
function savePdf_(p) {
  var date = p.reportDate || today_(), month = date.slice(0, 7);
  var f = saveBlob_(folder_(['DCR', month]), Utilities.newBlob(Utilities.base64Decode(p.b64), 'application/pdf', p.fileName || ('CRO_DCR_' + dmy_(date).replace(/\//g, '-') + '.pdf')));
  if (p.reportDate) {
    var dS = sh_('DCRS'), dR = rows_('DCRS').filter(function (r) { return r.reportDate === p.reportDate; })[0];
    if (dR) dS.getRange(dR._row, SHEETS.DCRS.indexOf('pdfUrl') + 1).setValue(f.getUrl());
    else dS.appendRow(toRow_('DCRS', { reportDate: p.reportDate, fileName: p.fileName, pdfUrl: f.getUrl(), cases: 0, uploaded: now_() }));
  }
  return { url: f.getUrl() };
}

/* ---------- शोध / शेरा ---------- */
function search_(q) {
  q = q || {};
  var t = String(q.text || '').trim(), toks = t ? t.split(/\s+/) : [];
  var rs = rows_('CASES').filter(function (r) {
    if (q.date && (r.dcrList || r.dcrDate).split(',').indexOf(q.date) < 0) return false;
    if (q.month && String(r.dcrDate || r.regDate).slice(0, 7) !== q.month) return false;
    if (q.ps && r.ps !== q.ps) return false;
    if (q.div && r.div !== q.div) return false;
    if (q.head && r.head !== q.head) return false;
    if (q.shera === 'yes' && !r.shera) return false;
    if (!toks.length) return true;
    var hay = [r.ps, r.no + '/' + r.yr, r.no + '/' + String(r.yr).slice(2), r.kalam, r.label, r.comp, r.acc, r.place, r.io, r.facts, r.shera].join(' ');
    return toks.every(function (k) { return hay.indexOf(k) >= 0; });
  });
  rs.sort(function (a, b) { return (b.regDate || b.dcrDate) > (a.regDate || a.dcrDate) ? 1 : (b.regDate || b.dcrDate) < (a.regDate || a.dcrDate) ? -1 : 0; });
  var more = rs.length > 400;
  return { cases: rs.slice(0, 400).map(caseObj_), more: more, total: rs.length };
}
function findCase_(key) {
  var r = rows_('CASES').filter(function (x) { return x.key === key; })[0];
  if (!r) throw new Error('गुन्हा सापडला नाही: ' + key);
  return r;
}
function saveShera_(key, shera) {
  var r = findCase_(key), s = sh_('CASES');
  s.getRange(r._row, CASE_COLS.indexOf('shera') + 1, 1, 2).setValues([[shera || '', shera ? now_() : '']]);
  return { ok: 1 };
}
function setHead_(key, head) {
  var r = findCase_(key), s = sh_('CASES');
  s.getRange(r._row, CASE_COLS.indexOf('head') + 1, 1, 2).setValues([[head, head ? 'हाताने' : '']]);
  return { ok: 1 };
}

/* ---------- Google Doc → Word / PDF ---------- */
function exportDocx_(docId, folder, base) {
  var res = UrlFetchApp.fetch('https://docs.google.com/document/d/' + docId + '/export?format=docx', { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('Word फाइल बनवता आली नाही (' + res.getResponseCode() + ')');
  var docx = saveBlob_(folder, res.getBlob().setName(base + '.docx'));
  DriveApp.getFileById(docId).setTrashed(true);
  return { docxUrl: docx.getUrl(), docxId: docx.getId() };
}
function exportDoc_(docId, folder, base) {
  var out = exportDocx_(docId, folder, base);
  var pdf = saveBlob_(folder, DriveApp.getFileById(docId).getAs('application/pdf').setName(base + '.pdf'));
  return { docxUrl: out.docxUrl, docxId: out.docxId, pdfUrl: pdf.getUrl(), pdfId: pdf.getId() };
}
function newDoc_(name, folder) {
  var d = DocumentApp.create(name);
  DriveApp.getFileById(d.getId()).moveTo(folder);
  return d;
}
function para_(body, text, o) {
  o = o || {};
  var p = body.appendParagraph('');
  p.setAlignment(o.align || DocumentApp.HorizontalAlignment.JUSTIFY).setSpacingAfter(o.after == null ? 6 : o.after).setLineSpacing(o.line || 1.15);
  if (o.indent) p.setIndentFirstLine(o.indent);
  if (o.start) p.setIndentStart(o.start);
  rich_(p, text, o);
  return p;
}
/* मजकूर लिहितो. कृतिदेव मोडमध्ये देवनागरी → Kruti Dev; इंग्रजी अक्षरे (E-mail, SP, CDR) Arial मध्ये राहतात. */
function rich_(p, text, o) {
  var kr = !!o.kr, segs = [], re = /[A-Za-z][A-Za-z0-9@._\-\/]*/g, last = 0, m;
  text = String(text || '');
  while ((m = re.exec(text))) { if (m.index > last) segs.push({ t: text.slice(last, m.index), lat: 0 }); segs.push({ t: m[0], lat: 1 }); last = re.lastIndex; }
  if (last < text.length) segs.push({ t: text.slice(last), lat: 0 });
  var full = '', ranges = [];
  segs.forEach(function (s) {
    var conv = kr && !s.lat ? uni2kruti(s.t) : s.t, st = full.length;
    full += conv; if (conv.length) ranges.push({ s: st, e: full.length - 1, lat: s.lat });
  });
  p.setText(full);
  var t = p.editAsText(), mainF = o.font || 'Mukta';
  if (o.size) t.setFontSize(o.size);
  if (full.length) {
    t.setFontFamily(mainF);
    if (kr) ranges.forEach(function (r) { if (r.lat) t.setFontFamily(r.s, r.e, 'Arial'); });
    if (o.bold) t.setBold(true);
    if (o.boldPre) { var n = Math.min((kr ? uni2kruti(o.boldPre) : o.boldPre).length, full.length); t.setBold(0, n - 1, true); }
  }
  return p;
}
function jaKrText_(S, date, seq) {
  return S.jaKrPrefix + '/' + S.jaKrFile + '/ ' + (seq || '______') + ' /' + date.slice(0, 4);
}
function logoOf_(S, slot) {
  var id = S['logo' + slot];
  try { if (id) return DriveApp.getFileById(id).getBlob(); } catch (e) { /* अपलोड केलेला लोगो मिळाला नाही: तयार लोगो वापरा */ }
  var g = LOGOS[slot];
  return g ? Utilities.newBlob(Utilities.base64Decode(g.b64), g.mime, 'logo' + slot) : null;
}
/* एक पत्र (शीर्षक, जा.क्र., मजकूर, सही, प्रति) b मध्ये लिहितो. सर्व पत्रांची मांडणी एकसारखी. */
function renderLetter_(b, L, S, kr) {
  var A = DocumentApp.HorizontalAlignment, F = kr ? 'Kruti Dev 010' : (S.font || 'Mukta');
  var sz = function (u, k) { return kr ? k : u; };
  var o = { kr: kr, font: F };
  var off = String(S.office || '').split('\n').filter(String);
  /* शीर्षक: डावा लोगो | मध्यभागी कार्यालय (वर मध्यभागी चिन्ह) | उजवा लोगो */
  var ht = b.appendTable(); ht.setBorderWidth(0);
  var row = ht.appendTableRow();
  var cells = { L: row.appendTableCell(), C: row.appendTableCell(), R: row.appendTableCell() };
  ['L', 'C', 'R'].forEach(function (k) {
    var bl = logoOf_(S, k); if (!bl) return;
    var p = cells[k].getChild(0).asParagraph(), w = k === 'C' ? 25 : 89;
    p.setAlignment(A.CENTER);
    var im = p.appendInlineImage(bl); im.setWidth(w); im.setHeight(w * im.getHeight() / im.getWidth());
  });
  var hasC = !!logoOf_(S, 'C');
  off.forEach(function (l, i) {
    var p = (i === 0 && !hasC) ? cells.C.getChild(0).asParagraph() : cells.C.appendParagraph('');
    p.setAlignment(A.CENTER).setSpacingAfter(0);
    rich_(p, l, { kr: kr, font: F, bold: i === 0, size: i === 0 ? sz(14, 20) : (i === off.length - 1 ? sz(9, 13) : sz(11, 16)) });
  });
  ht.setColumnWidth(0, 100); ht.setColumnWidth(2, 100); ht.setColumnWidth(1, 612 - 115 - 100 - 100);
  b.appendHorizontalRule();
  var t = b.appendTable([['', '']]); t.setBorderWidth(0);
  var jt = 'जा.क्र.: ' + L.jaKr, dt = 'दिनांक: ' + (L.dateText != null ? L.dateText : dmy_(L.date || today_()));
  var c0 = t.getCell(0, 0), c1 = t.getCell(0, 1);
  rich_(c0.getChild(0).asParagraph(), jt, { kr: kr, font: F, size: sz(11, 17), bold: true });
  rich_(c1.getChild(0).asParagraph(), dt, { kr: kr, font: F, size: sz(11, 17), bold: true });
  c1.getChild(0).asParagraph().setAlignment(A.RIGHT);
  para_(b, 'संदर्भ: ' + (L.ref || ''), Object.assign({ start: 0, after: 4, align: A.LEFT, size: sz(11, 17), boldPre: 'संदर्भ:' }, o));
  para_(b, 'विषय: ' + (L.subject || ''), Object.assign({ bold: true, after: 6, align: A.LEFT, size: sz(11, 17) }, o));
  para_(b, L.intro || '', Object.assign({ indent: 36, size: sz(11, 17) }, o));
  (L.body || []).forEach(function (x) { if (x && x.trim()) para_(b, x.trim(), Object.assign({ indent: 36, size: sz(11, 17) }, o)); });
  if (L.close) para_(b, L.close, Object.assign({ indent: 36, size: sz(11, 17) }, o));
  para_(b, '', { after: 14 });
  var sig = ['(' + (S.spName || '                    ') + ')'].concat(String(S.spDesig || '').split('\n'));
  sig.forEach(function (l, i) { para_(b, l, Object.assign({ align: A.LEFT, start: 300, after: 0, bold: i === 0, size: sz(11, 17) }, o)); });
  para_(b, '', { after: 6 });
  var pr = (L.prati || []).filter(String), pt = (L.prat || []).filter(String);
  if (pr.length) {
    para_(b, 'प्रति,', Object.assign({ align: A.LEFT, after: 0, size: sz(11, 17) }, o));
    pr.forEach(function (x, i) { para_(b, (pr.length > 1 ? (i + 1) + ') ' : '') + x, Object.assign({ align: A.LEFT, start: 18, after: 0, size: sz(11, 17) }, o)); });
  }
  if (pt.length) {
    para_(b, 'प्रत: माहिती व पुढील कार्यवाहीस,', Object.assign({ align: A.LEFT, after: 0, size: sz(11, 17) }, o));
    pt.forEach(function (x, i) { para_(b, (pt.length > 1 ? (i + 1) + ') ' : '') + x, Object.assign({ align: A.LEFT, start: 18, after: 0, size: sz(11, 17) }, o)); });
  }
}
function pageSetup_(b) {
  b.setMarginTop(57.6).setMarginBottom(43.2).setMarginLeft(57.6).setMarginRight(57.6);
  try { b.setPageWidth(612); b.setPageHeight(792); } catch (e) { Logger.log('पेज आकार: ' + e); }
}
/* जा.क्र. क्रमांक: सेटिंग्जच्या आधारे; autoNo बंद असल्यास रिकामी जागा */
function nextJaKr_(S, date) {
  if (S.autoNo !== 'होय') return jaKrText_(S, date, '');
  var seq = (+S.letterSeq || 0) + 1;
  setSetting_('letterSeq', seq);
  S.letterSeq = String(seq);
  return jaKrText_(S, date, String(seq));
}
/* पत्र तयार करताना वापरकर्त्याने केलेले बदल (मूळ व बदललेला मजकूर) → पुढील पत्रांसाठी नियम */
function learn_(S, orig, L) {
  if (!orig || !L) return;
  var pairs = [['ref', orig.ref, L.ref], ['subject', orig.subject, L.subject], ['intro', orig.intro, L.intro], ['close', orig.close, L.close]];
  var okeys = orig.keys || [], lkeys = L.keys || [];
  okeys.forEach(function (k, i) { var j = lkeys.indexOf(k); if (j >= 0) pairs.push(['body' + k, (orig.body || [])[i], (L.body || [])[j]]); });
  var rules = []; try { rules = JSON.parse(S.rules || '[]'); } catch (e) { rules = []; }
  pairs.forEach(function (p) {
    var r = LETTER.diffRule(p[1], p[2]);
    if (!r) return;
    rules = rules.filter(function (x) { return x.f !== r.f; });
    rules.push(r);
  });
  setSetting_('rules', JSON.stringify(rules.slice(-300)));
}
/* items: [{ key, L, orig }] → सर्व पत्रे एका Word फाईलमध्ये (Unicode व कृतिदेव) + Unicode PDF */
function makeLetters_(items) {
  items = (items || []).filter(function (it) { return it && it.key && it.L; });
  if (!items.length) throw new Error('पत्रासाठी गुन्हा निवडलेला नाही');
  var S = settings_(), date = items[0].L.date || today_(), month = date.slice(0, 7), folder = folder_(['DCR_चेक', month]);
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  var built = [];
  try {
    items.forEach(function (it) {
      var L = it.L, jaKr = L.jaKr || nextJaKr_(S, date);
      built.push({ key: it.key, c: caseObj_(findCase_(it.key)), L: Object.assign({}, L, { jaKr: jaKr, date: date }), orig: it.orig });
    });
  } finally { lock.releaseLock(); }
  built.forEach(function (x) { learn_(S, x.orig, x.L); });
  var one = built.length === 1, c0 = built[0].c;
  var base = one ? c0.ps + '_' + (c0.kind === 'AD' ? 'मर्ग_' : '') + c0.no + '-' + String(c0.yr).slice(2) + '_तपास_सूचना' : 'DCR_चेक_पत्रे_' + built.length + '_' + date;
  var docU = newDoc_(base, folder), bu = docU.getBody(); pageSetup_(bu);
  built.forEach(function (x, i) { if (i) bu.appendPageBreak(); renderLetter_(bu, x.L, S, false); });
  bu.editAsText().setFontFamily(S.font || 'Mukta');
  docU.saveAndClose();
  var outU = exportDoc_(docU.getId(), folder, base);
  var docK = newDoc_(base + '_कृतिदेव', folder), bk = docK.getBody(); pageSetup_(bk);
  built.forEach(function (x, i) { if (i) bk.appendPageBreak(); renderLetter_(bk, x.L, S, true); });
  docK.saveAndClose();
  var outK = exportDocx_(docK.getId(), folder, base + '_कृतिदेव');
  var sL = sh_('LETTERS'), sC = sh_('CASES'), now = now_(), ids = [];
  built.forEach(function (x) {
    var id = Utilities.getUuid().slice(0, 8), r = findCase_(x.key);
    ids.push(id);
    sL.appendRow(toRow_('LETTERS', { id: id, jaKr: x.L.jaKr, date: date, key: x.key, ps: x.c.ps, div: x.c.div, crime: (x.c.kind === 'AD' ? 'मर्ग ' : '') + x.c.no + '/' + x.c.yr,
      head: x.c.head, shera: x.L.shera || '', keys: (x.L.keys || []).join(','), docxUrl: outU.docxUrl, krutiUrl: outK.docxUrl, pdfUrl: outU.pdfUrl, created: now, status: '', replyDate: '', note: '' }));
    sC.getRange(r._row, CASE_COLS.indexOf('shera') + 1, 1, 2).setValues([[x.L.shera || r.shera || '', now]]);
    sC.getRange(r._row, CASE_COLS.indexOf('letters') + 1).setValue((+r.letters || 0) + 1);
  });
  return { docxUrl: outU.docxUrl, krutiUrl: outK.docxUrl, pdfUrl: outU.pdfUrl, count: built.length, jaKrs: built.map(function (x) { return x.L.jaKr; }), ids: ids };
}
/* जुनी एक-पत्र क्रिया: makeLetters_ चा एकच पत्र */
function makeLetter_(key, L) {
  var r = makeLetters_([{ key: key, L: L, orig: L.orig }]);
  return { docxUrl: r.docxUrl, krutiUrl: r.krutiUrl, pdfUrl: r.pdfUrl, jaKr: r.jaKrs[0], id: r.ids[0] };
}
/* लोगो (डावा, मध्य, उजवा) Drive मध्ये जतन करून सेटिंगमध्ये ठेवतो */
function saveLogo_(slot, b64, mime) {
  if (['L', 'C', 'R'].indexOf(slot) < 0) throw new Error('अवैध लोगो');
  var blob = Utilities.newBlob(Utilities.base64Decode(b64), mime || 'image/png', 'लोगो_' + slot);
  var f = saveBlob_(root_(), blob);
  setSetting_('logo' + slot, f.getId());
  return { id: f.getId() };
}
function letters_(q) {
  q = q || {};
  return rows_('LETTERS').filter(function (r) { return !q.month || String(r.date).slice(0, 7) === q.month; })
    .sort(function (a, b) { return a.created < b.created ? 1 : -1; });
}
function setLetter_(id, o) {
  var r = rows_('LETTERS').filter(function (x) { return x.id === id; })[0];
  if (!r) throw new Error('पत्र सापडले नाही');
  var s = sh_('LETTERS'), C = SHEETS.LETTERS;
  ['status', 'replyDate', 'note'].forEach(function (k) { if (k in o) s.getRange(r._row, C.indexOf(k) + 1).setValue(o[k] || ''); });
  return { ok: 1 };
}

/* ---------- महिन्याची तपासणी ---------- */
function monthCases_(month, scope) {
  return rows_('CASES').filter(function (r) {
    return r.kind !== 'NC' && String(r.dcrDate || r.regDate).slice(0, 7) === month && (!scope || scope === 'ALL' || r.div === scope);
  }).map(caseObj_);
}
function monthCheck_(month, scope) {
  var cs = monthCases_(month, scope), S = settings_(), I = [];
  /* DCR न आलेल्या तारखा */
  var have = {}; rows_('DCRS').forEach(function (r) { if (+r.cases > 0) have[r.reportDate] = 1; });
  var y = +month.slice(0, 4), m = +month.slice(5, 7), last = new Date(y, m, 0).getDate(), td = today_(), miss = [];
  for (var d = 1; d <= last; d++) { var iso = month + '-' + (d < 10 ? '0' : '') + d; if (iso >= td) break; if (!have[iso]) miss.push(d); }
  if (miss.length) I.push({ level: 'bad', msg: 'या तारखांचे DCR अपलोड झाले नाहीत: ' + miss.join(', ') + ' (' + month.slice(5) + '/' + y + '). ते दिल्याशिवाय B भाग अपूर्ण राहील, आणि खालील अप.क्र. गॅप त्यामुळेही दिसू शकतात.' });
  /* अप.क्र. मधील गॅप (FIR क्रमांक एकाच मालिकेत असतात) */
  var by = {};
  cs.forEach(function (c) { if (c.kind === 'FIR' && c.no) { var k = c.ps + '|' + c.yr; (by[k] = by[k] || []).push(+c.no); } });
  Object.keys(by).sort().forEach(function (k) {
    var a = by[k].sort(function (x, z) { return x - z; }), gaps = [];
    for (var i = 1; i < a.length; i++) { var g = a[i] - a[i - 1]; if (g > 1 && g <= 15) for (var n = a[i - 1] + 1; n < a[i]; n++) gaps.push(n); }
    if (gaps.length) I.push({ level: miss.length ? 'info' : 'warn', msg: k.split('|')[0] + ': अप.क्र. ' + gaps.join(', ') + '/' + k.split('|')[1].slice(2) + ' हे गुन्हे या महिन्याच्या DCR मध्ये आले नाहीत (NC, इतर कायदा किंवा DCR मधून राहिले?).' });
  });
  var oth = cs.filter(function (c) { return c.head === 'OTH5' || c.head === 'OTH6'; });
  if (oth.length) I.push({ level: 'info', msg: oth.length + ' गुन्ह्यांचे B भाग शीर्षक ठरले नाही ("इतर" मध्ये). खाली यादीत शीर्षक निवडा.' });
  var nod = cs.filter(function (c) { return !c.div; });
  if (nod.length) I.push({ level: 'warn', msg: nod.length + ' गुन्ह्यांना उपविभाग नाही: ' + nod.slice(0, 10).map(function (c) { return c.ps + ' ' + c.no; }).join(', ') });
  var cnt = {}, divs = String(S.divOrder).split(',');
  cs.forEach(function (c) { var h = c.head || 'OTH5'; cnt[h] = cnt[h] || {}; cnt[h][c.div || '-'] = (cnt[h][c.div || '-'] || 0) + 1; });
  return { issues: I, counts: cnt, divs: divs, total: cs.length, others: oth.map(function (c) { return { key: c.key, ps: c.ps, no: c.no, yr: c.yr, kalam: c.kalam, label: c.label, head: c.head }; }) };
}

/* ---------- हेडवाइज B भाग ---------- */
function bRow_(c, n) {
  var cr = c.no + '/' + String(c.yr).slice(2) + (c.kalam ? ' कलम ' + c.kalam : '') + (c.label ? '\n(' + c.label + ')' : '');
  var arrest = [c.arrest, c.arrestReason].filter(String).join('\n');
  if (!arrest) {
    var seg = String(c.acc || '').split(' / '), ai = -1;
    seg.forEach(function (x, i) { if (ai < 0 && /अटक|सु[ूु]?चना\s*पत्र|नोटीस/.test(x)) ai = i; });
    arrest = ai < 0 ? '-' : seg[ai] + (seg[ai + 1] && /^(दि|\d|चे)/.test(seg[ai + 1]) ? ' ' + seg[ai + 1] : '');
  }
  var sl = function (x) { return String(x || '').replace(/ \/ /g, ' ').trim(); };
  var facts = sl(c.facts); if (facts.length > 900) facts = facts.slice(0, 900) + '...';
  if (c.head === 'AD' || c.kind === 'AD')
    return [String(n), c.ps, 'मर्ग क्र. ' + c.no + '/' + String(c.yr).slice(2) + (c.kalam ? ' कलम ' + c.kalam : ''), c.place, sl(c.occ), sl(c.reg),
      c.comp, c.deceased || c.victim, [c.regOff ? 'दाखल: ' + c.regOff : '', c.io ? 'तपास: ' + c.io : ''].filter(String).join('\n'), c.cause || facts];
  return [String(n), c.ps, cr, c.place, sl(c.occ), sl(c.reg),
    'फिर्यादी: ' + (c.comp || '-') + '\nआरोपी: ' + (c.acc || '-'), arrest, c.io || '-', facts];
}
function buildBPart_(month, scope) {
  var S = settings_(), cs = monthCases_(month, scope), font = S.font || 'Mukta';
  if (!cs.length) throw new Error(month + ' या महिन्याचे गुन्हे नोंदवहीत नाहीत. आधी DCR अपलोड करा.');
  var divs = String(S.divOrder).split(','), dIdx = {}; divs.forEach(function (d, i) { dIdx[d] = i; });
  var HD = DCR.HEADS, groups = {};
  cs.forEach(function (c) { var h = c.head || 'OTH5'; (groups[h] = groups[h] || []).push(c); });
  var mm = month.slice(5) + '/' + month.slice(0, 4), scopeTxt = !scope || scope === 'ALL' ? S.distName : scope + ' विभाग';
  var folder = folder_(['B_भाग', month]), base = 'B_भाग_' + (!scope || scope === 'ALL' ? 'जिल्हा' : scope) + '_' + month.slice(5) + '-' + month.slice(0, 4);
  var doc = newDoc_(base, folder), b = doc.getBody(), A = DocumentApp.HorizontalAlignment;
  b.setPageWidth(842).setPageHeight(595).setMarginTop(24).setMarginBottom(24).setMarginLeft(24).setMarginRight(24);
  para_(b, scopeTxt, { align: A.CENTER, bold: true, size: 15, after: 0 });
  para_(b, 'बी पार्ट (हेडवाइज)', { align: A.CENTER, bold: true, size: 13, after: 0 });
  para_(b, 'माहे ' + mm, { align: A.CENTER, bold: true, size: 12, after: 8 });
  /* गोषवारा */
  var useDivs = !scope || scope === 'ALL' ? divs : [scope];
  var sum = [['अ.क्र.', 'शीर्षक'].concat(useDivs).concat(['एकूण'])], k = 0, tot = useDivs.map(function () { return 0; }), gt = 0;
  HD.forEach(function (h) {
    var g = groups[h[0]]; if (!g) return;
    var row = [String(++k), h[1]], n = 0;
    useDivs.forEach(function (d, i) { var x = g.filter(function (c) { return c.div === d; }).length; tot[i] += x; n += x; row.push(String(x || '-')); });
    var other = g.length - n; gt += g.length;
    row.push(String(g.length) + (other ? ' (' + other + ' विभाग नाही)' : ''));
    sum.push(row);
  });
  sum.push(['', 'एकूण'].concat(tot.map(String)).concat([String(gt)]));
  var st = b.appendTable(sum); st.editAsText().setFontSize(9);
  st.getRow(0).editAsText().setBold(true); st.getRow(sum.length - 1).editAsText().setBold(true);
  for (var ci = 0; ci < sum[0].length; ci++) st.getRow(0).getCell(ci).setBackgroundColor('#e8eef9');
  b.appendPageBreak();
  /* शीर्षकवार तक्ते */
  var COLS = ['अ.क्र.', 'पोस्टे', 'अप.क्र. व कलम', 'घटनास्थळ', 'घ. ता. वेळ', 'दा. ता. वेळ', 'फिर्यादी व आरोपीचे नाव व पत्ता', 'अटक / प्रतिबंधक', 'तपास अधिकारी', 'घटनेचे थोडक्यात विवरण'];
  var ADC = ['अ.क्र.', 'पोस्टे', 'मर्ग क्र. व कलम', 'घटनास्थळ', 'घ. ता. वेळ', 'दा. ता. वेळ', 'फिर्यादीचे नाव', 'मृतकाचे नाव', 'दाखल / तपास', 'मरणाचे कारण'];
  var W = [22, 50, 80, 68, 56, 56, 120, 60, 70, 212];
  k = 0;
  HD.forEach(function (h) {
    var g = groups[h[0]]; if (!g) return;
    g.sort(function (a, z) { var da = dIdx[a.div] == null ? 99 : dIdx[a.div], dz = dIdx[z.div] == null ? 99 : dIdx[z.div]; return da - dz || (a.ps < z.ps ? -1 : a.ps > z.ps ? 1 : 0) || (+a.no - +z.no); });
    para_(b, (++k) + ') ' + h[1] + ' - ' + g.length, { align: A.LEFT, bold: true, size: 11, after: 2 });
    var rows = [h[0] === 'AD' ? ADC : COLS], divRows = [], cur = null, n = 0;
    g.forEach(function (c) {
      if ((!scope || scope === 'ALL') && c.div !== cur) { cur = c.div; divRows.push(rows.length); rows.push(['', (cur || 'विभाग नमूद नाही') + (cur ? ' विभाग' : ''), '', '', '', '', '', '', '', '']); }
      rows.push(bRow_(c, ++n));
    });
    var t = b.appendTable(rows);
    t.editAsText().setFontSize(8);
    W.forEach(function (w, i) { t.setColumnWidth(i, w); });
    t.getRow(0).editAsText().setBold(true);
    for (var i = 0; i < 10; i++) t.getRow(0).getCell(i).setBackgroundColor('#e8eef9');
    divRows.forEach(function (ri) { t.getRow(ri).editAsText().setBold(true); for (var i = 0; i < 10; i++) t.getRow(ri).getCell(i).setBackgroundColor('#f6f3e6'); });
    para_(b, '', { after: 6 });
  });
  b.editAsText().setFontFamily(font);
  doc.saveAndClose();
  var out = exportDoc_(doc.getId(), folder, base);
  sh_('BPART').appendRow(toRow_('BPART', { month: month, scope: scope || 'ALL', cases: cs.length, docxUrl: out.docxUrl, pdfUrl: out.pdfUrl, created: now_() }));
  out.cases = cs.length;
  return out;
}

/* ---------- सेटिंग्ज ---------- */
function saveSettings_(o) {
  Object.keys(o || {}).forEach(function (k) { if (k in DEF || k === 'paras') setSetting_(k, o[k]); });
  return settings_();
}
function savePs_(list) {
  var s = sh_('PS'); if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, 2).clearContent();
  var v = (list || []).filter(function (r) { return r.ps; }).map(function (r) { return [r.ps, r.div || '']; });
  if (v.length) s.getRange(2, 1, v.length, 2).setValues(v);
  return rows_('PS');
}
/** पहिल्यांदा: Apps Script संपादकात हे एकदा चालवा (परवानग्या देण्यासाठी). फोल्डर व Sheet तयार होतात. */
function setup() {
  Object.keys(SHEETS).forEach(sh_);
  Logger.log('फोल्डर: ' + root_().getUrl());
  Logger.log('नोंदवही: ' + ss_().getUrl());
  UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true }); // Word export साठी परवानगी
}


/* DCR FORMAT पत्रातील तीन लोगो (डावा पोलीस लोगो, मध्य राजमुद्रा, उजवा). सेटिंग्जमध्ये दुसरा लोगो अपलोड केल्यास तो वापरला जातो. */
var LOGOS = {
  L: { mime: 'image/jpeg', b64: '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' },
  C: { mime: 'image/png', b64: 'iVBORw0KGgoAAAANSUhEUgAAADUAAABICAMAAABhhE1dAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAK1UExURf////7+/v39/fv7+/Hx8dnZ2c3Nzb+/v6mpqZ6enqysrLCwsLm5uba2trq6utvb2/z8/PT09Hp6ekdHR1dXV2xsbHt7e5eXl6CgoIeHh5KSko+Pj5ubm2dnZ1hYWNfX1+zs7Gpqaqurq5GRkXh4eJqampOTk6GhoZ+fn4KCgrKysvf399zc3N/f3/j4+Hl5eaKiopmZmXd3d+Dg4OLi4ubm5tjY2K2tra6uru/v78XFxYiIiFZWVmVlZfPz8/r6+uTk5IGBgZWVlc7OzoCAgG5ubvX19b29vYuLi3Z2du7u7vDw8F1dXYqKitXV1d3d3cPDw8DAwJycnH9/f4yMjCgoKMrKyoaGhnBwcKSkpI6Ojurq6nNzc+np6fb29m1tbUVFRcLCwsbGxr6+vi4uLtTU1MTExIWFhaqqqrGxsby8vJSUlGRkZKenp8nJya+vr9LS0qamppCQkGlpaSsrK4SEhNDQ0HJycllZWZaWlsvLy2FhYbi4uIODg8HBwVNTU6WlpUlJSUBAQPLy8k1NTbe3t8zMzLOzs3R0dF9fX4mJiaOjo/n5+dra2ufn53FxcZiYmH5+fjo6OmNjY2hoaFJSUnV1dXx8fFBQUMjIyFVVVeXl5U5OTm9vb0hISH19fUtLS0RERI2NjT09PV5eXjY2Nru7u6ioqOPj4y0tLSYmJlxcXDIyMmtra7S0tEJCQiMjIxMTEyAgICEhIR4eHrW1tT8/Pzc3N9PT01RUVDU1NSwsLFpaWmZmZmBgYJ2dnT4+PsfHxzs7O+vr6zMzM1FRUUZGRjk5OVtbWzg4OE9PT+Hh4UFBQUpKSkxMTCkpKSUlJc/PzyQkJNHR0RwcHO3t7Tw8PBcXFw8PDx0dHWJiYhISEtbW1ujo6BEREd7e3jQ0NDExMUNDQxgYGCoqKhUVFQAAACHrdzkAAADndFJOU///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////ALLDDRIAAAAJcEhZcwAAIdUAACHVAQSctJ0AAAW6SURBVFhHxZexeSM5DIWnJSJFB0zYBiphEciYMGYDyFiHarkfo9m1bGl2L7qD5c/yiCDIh4cH6Ph/rbY+rRg252zXw7+aDA8pRbUUV/+XbjWsiJWpVrTMKNfjW2vnvhUXKbPV1qwUk3p+dmdTPYwlVXV2jVJ5Ugj23OqzbxXX1kK0yyxTVFfHKVSFI9bJLee7YzXptR2mEqPP3rtufq10Y5dZOGZ3u9b+tjyL2mzWuwVYVI0lcUY2V4JnPO/fozXlQt1nFzGLbsOEoEGssTOillrZS9nu8kibgYeGNelTwkixcE2iJyhajAxYcC82vjywabO0WtV6lNyYSFuUWEVlNlGYUkgGq/wrey266BCu0KUVGaWeIHI8YakM3g/cWuNp6ZcToPPfsConiQpvZcvsEey9OCrH5qnOCT17+c2w7j5WxoKvfDxKjwGSxXzBLpKWTlJiRfgz6aeBoY8JeNy6NPduOtbjsdYAUghpGZclw/dr0nro1F34CGTy02bmUUjXXiY+LdEUoP5eAZMDgLR1ThMaOzgaGxXfFA3vQ0Ig6neno5IoiiOr0Z1NNZEHPuDn9OdOBaZcq7+sAQY46RSLAfcODZJeiG4SLXnc2jt/OQdgQO5kEqmAj0Z8eM/RubHs90jUPezl9ixSri+EgfBwDzDgJLvtD1VNaVETPVck11nuwMEmAgdUyVnItfTFYAZn4sbuvMGJwxLRwmIPCDlg/IeyBFvt6QVcgwSgM3A+kXSSQJ7lQ6yjCEcCZgqmsEFJcoIQJxvEpDj3S5X8NhubDZ06DGA+SYmbDIjCHyfqjxSfxhFJLWAkEmQn80ypCA84xBgR18LvBhcHa1MuCOmUs7rCD0iRr/EmNk+z4fkKzyVig/uTBdTbR9i4VVMuLqXa3mDt1LfKysIL54Dj061Ooy5Sq7Y1ECPEXtr7LjYGxLgLdVTEgaJJiHWvtYTKaGxka/utEwpvOBCLwz4Gaoe24HUc43EDRRqx7OgTtGo8yoFwz8MWG8XjE99/GXgfvUPT06vP1g9ZOEi63ppxEkRYavX0stnb3lz1dL21NvpBZwAESS/Kt68BDpqut9aIgthrPZQzNY9pD7yqfGbTZb0cLauFNpm7IwW2BlcqHzrel02QnqOQ6un8ixLTRaxWnv/BDMzBi04+k3ZzWqNmcqs/xCJfvQ2v9ABJENCsVhAgvG7RqCnnxShKRDe9KmIxhTZIo7zjRs2GfCoHPZzuXMEQTiFsozWk7XO4bOiMAWeHRP28p2yhvTQ/rSTcPrG+IhaIoVNbFOPYS2jFHjSy4QJJsvNeS1+soRiOqhGRBfzJXkzfRRJoJxht81r6ZTMePpHAbP1l0IZTbk5BzO6ZcuKP93IGNOqWy1tvSGjGSlgwkMlhEQVmZLpWX9aca9Emke2ueQeUlD6GltKqew6X+ThHqldr22mgPGd6SARRxE2jhYioJ62f5m9MYz90qiw6+KxQiHb09ILvGQ2FMwqGEwM/5Ho1A2j6Z26FVqeUOj1hMOEI51LoS1k3Rp/n8ssopYgGWZlAtdHqaF8YkggeSWaWVOaeH3DEymnLGYGeQxP3yZ5C8hDH3oAETF6mqKdxMVAsOfRwVqa0hD1f2ipTUGc+iP3WmetY8Ifu05hzzh6EW3rjBO9pnIYSvzFxxqBz5XgBe2glebactRvdn7MXl/3uRLTaZGx+aHNg43wVgH+Q6Wwt5POunnU9YGIeL2e0ZxviRR3sx6em/DSGWHpXBoB5eT2MvMPIGB/K5JdVjpnWe35zqC8/DEO30vHLyPT17svQjveLMRAmCZOoOQ1Svoz3HPp6kJ+c30e+W2N+Tsmg4fMmh0XITkLYgRo95/RAgX4Y3xbYDF4wHNpgPOEfFhEoaNFQOPlyrX2xLrRKagzAEU0kl6kwN2MAawzdOQV+SjP1ysgOjtnh+EpSz68lfOUhvTU/olj+ezuOfwBN+phHYt1vmgAAAABJRU5ErkJggg==' },
  R: { mime: 'image/png', b64: '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' }
};


/* ===== सामायिक कोड (ॲपमध्येही हाच): Kruti Dev → Unicode, DCR वाचन, पत्र ===== */
// Kruti Dev (010/016/050/055 — एकच मांडणी) → Unicode
var K1=["ñ","Q+Z","sas","aa",")Z","ZZ","å","ƒ","„","…","†","‡","ˆ","‰","Š","‹",
"¶+","d+","[+k","[+","x+","T+","t+","M+","<+","Q+",";+","j+","u+",
"Ùk","Ù","ä","–","—","é","™","=kk","f=k",
"à","á","â","ã","ºz","º","í","{k","{","=","«",
"Nî","Vî","Bî","Mî","<î","|","K","}",
"J","Vª","Mª","<ªª","Nª","Ø","Ý","nzZ","æ","ç","Á","xz","#",":",
"v‚","vks","vkS","vk","v","b±","Ã","bZ","b","m","Å",",s",",","_",
"ô","d","Dk","D","[k","[","x","Xk","X","Ä","?k","?","³",
"pkS","p","Pk","P","N","t","Tk","T",">","÷","¥",
"ê","ë","V","B","ì","ï","M+","<+","M","<",".k",".",
"r","Rk","R","Fk","F",")","n","/k","èk","/","Ë","è","u","Uk","U",
"i","Ik","I","Q","¶","c","Ck","C","Hk","H","e","Ek","E",
";","¸","j","y","Yk","Y","G","o","Ok","O","Ú","ª",
"'k","'","\"k","\"","l","Lk","L","g",
"È","z",
"Ì","Í","Î","Ï","Ñ","Ò","Ó","Ô","Ö","Ù","Ük","Ü",
"‚","ks","kS","k","h","q","w","`","s","S",
"a","¡","%","W","•","·","∙","~j","~","\\","+"," ः",
"^","*","Þ","ß","(","¼","½","¿","À","¾","A","-","&","Œ","]","~ ","@"];
var K2=["॰","QZ+","sa","a","र्द्ध","Z","०","१","२","३","४","५","६","७","८","९",
"फ़्","क़","ख़","ख़्","ग़","ज़्","ज़","ड़","ढ़","फ़","य़","ऱ","ऩ",
"त्त","त्त्","क्त","दृ","कृ","न्न","न्न्","=k","f=",
"ह्न","ह्य","हृ","ह्म","ह्र","ह्","द्द","क्ष","क्ष्","त्र","त्र्",
"छ्य","ट्य","ठ्य","ड्य","ढ्य","द्य","ज्ञ","द्व",
"श्र","ट्र","ड्र","ढ्र","छ्र","क्र","फ्र","र्द्र","द्र","प्र","प्र","ग्र","रु","रू",
"ऑ","ओ","औ","आ","अ","ईं","ई","ई","इ","उ","ऊ","ऐ","ए","ऋ",
"क्क","क","क","क्","ख","ख्","ग","ग","ग्","घ","घ","घ्","ङ",
"चै","च","च","च्","छ","ज","ज","ज्","झ","झ्","ञ",
"ट्ट","ट्ठ","ट","ठ","ड्ड","ड्ढ","ड़","ढ़","ड","ढ","ण","ण्",
"त","त","त्","थ","थ्","द्ध","द","ध","ध","ध्","ध्","ध्","न","न","न्",
"प","प","प्","फ","फ्","ब","ब","ब्","भ","भ्","म","म","म्",
"य","य्","र","ल","ल","ल्","ळ","व","व","व्","ऱ्","्र",
"श","श्","ष","ष्","स","स","स्","ह",
"ीं","्र",
"द्द","ट्ट","ट्ठ","ड्ड","कृ","भ","्य","ड्ढ","झ्","त्त्","श","श्",
"ॉ","ो","ौ","ा","ी","ु","ू","ृ","े","ै",
"ं","ँ","ः","ॅ","ऽ","ऽ","ऽ","्र","्","?","़",":",
"‘","’","“","”",";","(",")","{","}","=","।",".","-","॰",",","् ","/"];
function kruti2uni(s){
  if(!s) return s;
  var t=String(s).replace(/[‘’]/g,"'").replace(/[“”]/g,'"');
  // 'f' (ि) स्थान बदल: f+अक्षर → अक्षर+ि
  for(var i=0;i<K1.length;i++){ t=t.split(K1[i]).join(K2[i]); }
  // ि
  t=t.replace(/f((?:[क-ह]़?्)*[क-ह]़?)/g,function(m,c){return c+'ि';});
  t=t.replace(/f/g,'ि');
  // र्‌ (Z): अक्षर+मात्रा+Z → र्+अक्षर+मात्रा
  t=t.replace(/((?:[क-ह]़?्)*[क-हअ-औ]़?[ा-ौंँः]*)Z/g,function(m,c){return 'र्'+c;});
  t=t.replace(/Z/g,'र्');
  return t;
}

/* Unicode → Kruti Dev (उलट रूपांतर): पत्र Kruti Dev फॉन्टमध्ये छापण्यासाठी. kruti2uni चा उलटा; जोड्या round-trip तपासून निवडतो.
   फक्त देवनागरी अक्षरे व दंड (।) बदलतो; इंग्रजी अक्षरे, अंक व विरामचिन्हे जशीच्या तशी राहतात. */
var UNI2K_MAP=null, UNI2K_RE=null;
function uni2kruti(s){
  if(!s) return s;
  if(!UNI2K_MAP){
    UNI2K_MAP={};
    for(var i=0;i<K1.length;i++){
      var u=K2[i];
      if(!u||UNI2K_MAP[u]||!(/[\u0900-\u097F]/.test(u)||u==='।')) continue;
      if(kruti2uni(K1[i])===u) UNI2K_MAP[u]=K1[i];
    }
    UNI2K_MAP['ि']='f';
    UNI2K_RE=new RegExp(Object.keys(UNI2K_MAP).sort(function(a,b){return b.length-a.length;}).join('|'),'g');
  }
  var t=String(s);
  // reph: र् + अक्षर+मात्रा → अक्षर+मात्रा+Z
  t=t.replace(/र्((?:[क-ह]़?्)*[क-हअ-औ]़?[ा-ौंँः]*)/g,function(m,c){return c+'Z';});
  // ि: अक्षर+ि → f+अक्षर
  t=t.replace(/((?:[क-ह]़?्)*[क-ह]़?)ि/g,function(m,c){return 'f'+c;});
  // एक पास: जुळणारे सर्वात लांब तुकडे एकदाच बदलले जातात (साखळी-बदल नाही)
  return t.replace(UNI2K_RE,function(k){return UNI2K_MAP[k];});
}
if(typeof module!=='undefined') module.exports={kruti2uni:kruti2uni,uni2kruti:uni2kruti};

/* DCR (Word) वाचन: Kruti Dev → Unicode, गुन्ह्यांच्या ओळी, B भाग शीर्षक, त्रुटी तपासणी.
 * ब्राउझरमध्ये (JSZip + DOMParser) व node चाचणीत (jszip + @xmldom) दोन्हीकडे चालते. */
var DCR = (function () {
  var W = 'w:';

  /* ---------- docx → blocks ---------- */
  function kids(el, name) {
    var o = [];
    for (var n = el.firstChild; n; n = n.nextSibling) if (n.nodeType === 1 && (!name || n.nodeName === name)) o.push(n);
    return o;
  }
  function kid(el, name) { var k = kids(el, name); return k.length ? k[0] : null; }
  function fontOf(rPr) {
    if (!rPr) return null;
    var f = kid(rPr, W + 'rFonts');
    if (!f) return null;
    return f.getAttribute(W + 'ascii') || f.getAttribute(W + 'hAnsi') || (f.getAttribute(W + 'asciiTheme') ? 'theme' : null);
  }
  function readStyles(doc) {
    var st = { map: {}, defPara: null, defFont: null };
    if (!doc) return st;
    var root = doc.documentElement;
    var dd = kid(root, W + 'docDefaults');
    if (dd) { var rd = kid(dd, W + 'rPrDefault'); if (rd) st.defFont = fontOf(kid(rd, W + 'rPr')); }
    kids(root, W + 'style').forEach(function (s) {
      var id = s.getAttribute(W + 'styleId'), b = kid(s, W + 'basedOn');
      st.map[id] = { based: b ? b.getAttribute(W + 'val') : null, font: fontOf(kid(s, W + 'rPr')) };
      if (s.getAttribute(W + 'type') === 'paragraph' && s.getAttribute(W + 'default') === '1') st.defPara = id;
    });
    return st;
  }
  function styleFont(st, id) {
    for (var i = 0; id && i < 10; i++) { var s = st.map[id]; if (!s) break; if (s.font) return s.font; id = s.based; }
    return null;
  }
  /* Unicode मध्ये Kruti पद्धतीने टाइप केल्याने आलेले चुकीचे क्रम */
  function fixUni(t) {
    return t.replace(/ाॅ/g, 'ॉ').replace(/ेा|ाे/g, 'ो').replace(/ाै|ैा/g, 'ौ').replace(/्ा/g, '')
      .replace(/ं([ािीुूेैोौ])/g, '$1ं').replace(/([ेुूृी])\1/g, '$1');
  }
  function isKruti(f) { return /kruti/i.test(f || ''); }

  function paraText(p, st) {
    var pPr = kid(p, W + 'pPr'), ps = pPr && kid(pPr, W + 'pStyle');
    var pFont = styleFont(st, ps ? ps.getAttribute(W + 'val') : st.defPara) || st.defFont;
    var runs = [];
    (function walk(el) {
      kids(el).forEach(function (n) {
        var nm = n.nodeName;
        if (nm === W + 'r') runs.push(n);
        else if (nm === W + 'hyperlink' || nm === W + 'smartTag' || nm === W + 'ins' || nm === W + 'sdt' || nm === W + 'sdtContent' || nm === W + 'fldSimple') walk(n);
      });
    })(p);
    var out = '', buf = '', bufK = null;
    function flush() { if (buf) out += bufK ? kruti2uni(buf) : buf; buf = ''; }
    runs.forEach(function (r) {
      var rPr = kid(r, W + 'rPr'), rs = rPr && kid(rPr, W + 'rStyle');
      var f = fontOf(rPr) || (rs && styleFont(st, rs.getAttribute(W + 'val'))) || pFont;
      var k = isKruti(f), t = '';
      kids(r).forEach(function (n) {
        var nm = n.nodeName;
        if (nm === W + 't') t += n.textContent;
        else if (nm === W + 'tab') t += ' ';
        else if (nm === W + 'br' || nm === W + 'cr') t += '\n';
        else if (nm === W + 'noBreakHyphen') t += '-';
      });
      if (!t) return;
      if (k !== bufK) { flush(); bufK = k; }
      buf += t;
    });
    flush();
    return fixUni(out.replace(/\u00a0/g, ' '));
  }
  function cellLines(tc, st) {
    var o = [];
    kids(tc).forEach(function (n) {
      if (n.nodeName === W + 'p') { var t = paraText(n, st).replace(/\s+/g, ' ').trim(); if (t) o.push(t); }
      else if (n.nodeName === W + 'tbl') kids(n, W + 'tr').forEach(function (tr) { kids(tr, W + 'tc').forEach(function (c) { o = o.concat(cellLines(c, st)); }); });
      else if (n.nodeName === W + 'sdt') { var c = kid(n, W + 'sdtContent'); if (c) o = o.concat(cellLines(c, st)); }
    });
    return o;
  }
  function tableRows(tbl, st) {
    var rows = [];
    kids(tbl, W + 'tr').forEach(function (tr) {
      var row = [];
      var tcs = [];
      kids(tr).forEach(function (n) { if (n.nodeName === W + 'tc') tcs.push(n); else if (n.nodeName === W + 'sdt') { var c = kid(n, W + 'sdtContent'); if (c) tcs = tcs.concat(kids(c, W + 'tc')); } });
      tcs.forEach(function (tc) {
        var pr = kid(tc, W + 'tcPr'), gs = pr && kid(pr, W + 'gridSpan'), vm = pr && kid(pr, W + 'vMerge');
        var span = gs ? +gs.getAttribute(W + 'val') || 1 : 1;
        var cont = vm && vm.getAttribute(W + 'val') !== 'restart';
        row.push(cont ? [] : cellLines(tc, st));
        for (var i = 1; i < span; i++) row.push(null);
      });
      rows.push(row);
    });
    return rows;
  }
  function blocksFromXml(docXml, stylesXml, DP) {
    var dp = new DP();
    var doc = dp.parseFromString(docXml, 'text/xml');
    var st = readStyles(stylesXml ? dp.parseFromString(stylesXml, 'text/xml') : null);
    var body = doc.getElementsByTagName(W + 'body')[0], out = [];
    (function walk(el) {
      kids(el).forEach(function (n) {
        if (n.nodeName === W + 'p') out.push({ t: 'p', text: paraText(n, st).replace(/\s+/g, ' ').trim() });
        else if (n.nodeName === W + 'tbl') out.push({ t: 'tbl', rows: tableRows(n, st) });
        else if (n.nodeName === W + 'sdt') { var c = kid(n, W + 'sdtContent'); if (c) walk(c); }
      });
    })(body);
    return out;
  }
  function readDocx(buf, JSZip, DP) {
    return JSZip.loadAsync(buf).then(function (z) {
      var d = z.file('word/document.xml'), s = z.file('word/styles.xml');
      if (!d) throw new Error('ही Word (.docx) फाइल नाही');
      return Promise.all([d.async('string'), s ? s.async('string') : '']).then(function (a) { return blocksFromXml(a[0], a[1], DP); });
    });
  }

  /* ---------- मदतनीस ---------- */
  var DEVD = '०१२३४५६७८९';
  function asc(s) { return String(s || '').replace(/[०-९]/g, function (c) { return DEVD.indexOf(c); }); }
  function clean(s) { return String(s || '').replace(/\s+/g, ' ').trim(); }
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function dateOf(s) {
    var m = asc(s).match(/(\d{1,2})\s*[\/.\-]+\s*(\d{1,2})\s*[\/.\-]+\s*(\d{2,4})/);
    if (!m) return '';
    var d = +m[1], mo = +m[2], y = +m[3];
    if (y < 100) y += 2000;
    if (!d || d > 31 || !mo || mo > 12 || y < 2000 || y > 2100) return '';
    return y + '-' + pad(mo) + '-' + pad(d);
  }
  function dmy(iso) { return iso ? iso.slice(8, 10) + '/' + iso.slice(5, 7) + '/' + iso.slice(0, 4) : ''; }
  function normDiv(s) {
    s = clean(s).replace(/उपविभाग|विभाग|उप\s*वि\.?/g, '').replace(/[.:\-–]/g, ' ').trim();
    return s.replace(/नागपुर/g, 'नागपूर');
  }
  function normPs(s) {
    s = clean(s).replace(/^(पो\.?\s*स्टे\.?|पोस्टे|पोलीस\s*(स्टेशन|ठाणे)|पो\.?\s*ठाणे)\s*/, '').replace(/[,.:\-–\/]+$/, '').trim();
    return s;
  }

  /* ---------- विभाग (DCR मधील शीर्षके) ---------- */
  var SEC = [
    ['MVT', /वाहन\s*चोरी|मोटार\s*सायकल\s*\/?\s*वाहन/], ['HBT', /फोडी/], ['SLL', /स्थानिक\s*कायदे|\(SLL\)|SLL/],
    ['AD', /अकस्मात\s*मृत्यू|^मर्ग/], ['COMM', /जातीय/], ['GAMB', /^जुगार/], ['PROH', /^दारूबंदी|^दारुबंदी/],
    ['IPC', /न्याय\s*संहिता|भादंवि|भाग\s*5|भाग\s*1\s*ते\s*5/]
  ];
  var SEC_MR = { IPC: 'भाग 1 ते 5', MVT: 'वाहन चोरी', HBT: 'घरफोडी', SLL: 'भाग 6 (SLL)', AD: 'मर्ग', COMM: 'जातीय', GAMB: 'जुगार', PROH: 'दारूबंदी' };
  function secOf(t) {
    var s = t.replace(/्ा/g, 'ा').replace(/^ि/, '');
    for (var i = 0; i < SEC.length; i++) if (SEC[i][1].test(s)) return SEC[i][0];
    return null;
  }
  function headCount(t) {
    if (/निरंक/.test(t)) return 0;
    var m = asc(t).match(/[-–]\s*(\d{1,3})\s*$/);
    return m ? +m[1] : null;
  }

  var COLS = [
    ['sr', /^अ\.?\s*क्र|^क्र\.?$/], ['psCr', /मर्ग\s*क्र|पो(लीस)?\.?\s*स्टे|पो\.\s*स्टेशन|पोलीस\s*स्टेशन/],
    ['priorAction', /केलेली\s*प्रति|प्रति\.?\s*कार्यवाही/], ['prior', /पु[रू]्?वी\s*(चे)?\s*नोंद|पूर्वीचे\s*नोंद/],
    ['comp', /फिर्यादी/], ['victim', /पिडीत|पीडित/], ['deceased', /मृतक/],
    ['arrestReason', /नसल्याचे\s*कारण/], ['arrest', /अटक\s*आरोपी/], ['acc', /आरोपी/],
    ['place', /स्थळ/], ['occ', /घटना\s*ता/], ['reg', /दाखल\s*ता|रिपोर्ट\s*ता/],
    ['prop', /माल/], ['cause', /मरणाचे\s*कारण/], ['regOff', /दाखल\s*अधि/], ['io', /तपास\s*अधि/], ['incharge', /प्रभारी/],
    ['visit', /भेट/], ['facts', /हकी[कग]त/], ['nCases', /एकूण\s*केस/], ['nAcc', /एकूण\s*आरोपी/]
  ];
  function colMap(row) {
    var m = {}, hit = 0, ne = row.filter(function (c) { return c && c.length; });
    if (!ne.length || !/^अ\.?\s*क्र|^क्र/.test(clean(ne[0].join(' '))) || ne.some(function (c) { return c.join(' ').length > 90; })) return null;
    row.forEach(function (c, i) {
      if (!c) return;
      var t = clean(c.join(' '));
      if (!t) return;
      for (var j = 0; j < COLS.length; j++) if (COLS[j][1].test(t) && !(COLS[j][0] in m)) { m[COLS[j][0]] = i; hit++; return; }
    });
    return hit >= 4 && ('psCr' in m) ? m : null;
  }

  /* ---------- पो.स्टे. / अप.क्र. / कलम ---------- */
  function parsePsCr(lines) {
    var full = lines.join(' / '), a = asc(full);
    var first = lines[0] || '';
    var ps = first.split(/\s*(?:अप\.?\s*क्र|अप\.?\s*क|मर्ग|एन\.?\s*सी|अदखलपात्र|\d|\/)/)[0];
    ps = normPs(ps);
    var kind = /मर्ग/.test(full) ? 'AD' : (/एन\.?\s*सी\.?\s*क्र|अदखलपात्र|एनसी\s*क्र/.test(full) ? 'NC' : 'FIR');
    var rest = asc(full.slice(full.indexOf(first) + (ps ? first.indexOf(ps) + ps.length : 0)));
    var m = rest.match(/(\d{1,5})\s*\/\s*(20\d{2}|\d{2})(?!\d)/) || rest.match(/(?:क्र|क्रं|नं)\.?\s*[\-:]?\s*(\d{1,5})\s+(20\d{2}|\d{2})(?![\d\/])/);
    var no = m ? +m[1] : '', yr = m ? +m[2] : '';
    if (yr && yr < 100) yr += 2000;
    var km = a.match(/क\s*्?\s*[लज]\s*म\s*[:\-]?\s*/), kalam = '';
    if (km) {
      kalam = a.slice(km.index + km[0].length);
      kalam = kalam.split(/\s*\/\s*\(|\s*\/\s*(?:एस\s*आय\s*डी|एसआयडी|sid|SID|ई\s*साक्ष|वयस्क|अल्पवयीन)/)[0];
      kalam = clean(kalam.replace(/\s*\/\s*/g, ' ')).replace(/[\/,\s]+$/, '');
    }
    var lab = '', lm, re = /\(([^()0-9]{3,40})\)\s*(?=\/|$)/g;
    while ((lm = re.exec(full))) lab = clean(lm[1]);
    var sm = a.match(/(?:SID|sid|एस\s*\.?\s*आय\s*\.?\s*डी|एसआयडी|ई\s*-?\s*साक्ष)[^\d]{0,15}(\d{10,})/);
    return { ps: ps, kind: kind, no: no, yr: yr, kalam: kalam, label: lab, sid: sm ? sm[1] : '' };
  }

  /* ---------- B भाग शीर्षके ---------- */
  var HEADS = [
    ['MUR', 'खुन कलम 103(1) बीएनएस'], ['CH', 'सदोष मनुष्यवध कलम 105 बीएनएस'], ['ATM', 'खुनाचा प्रयत्न कलम 109 बीएनएस'],
    ['SUI', 'आत्महत्येस प्रवृत्त करणे कलम 108 बीएनएस'], ['NEG', 'निष्काळजीपणाने मृत्यू कलम 106(1) बीएनएस'], ['RAPE', 'बलात्कार / पोक्सो'],
    ['MOL', 'विनयभंग'], ['KID', 'पळवून नेणे'], ['ROB', 'जबरी चोरी / दरोडा'], ['HURT', 'दुखापत कलम 118(1) बीएनएस'],
    ['MVT', 'मोटार सायकल / वाहन चोरी'], ['THEFT', 'इतर चोरी'], ['HBTN', 'घरफोडी रात्री'], ['HBTD', 'घरफोडी दिवसा'],
    ['CBT', 'विश्वासघात कलम 316 बीएनएस'], ['CHEAT', 'फसवणूक कलम 318 बीएनएस'], ['ACC', '336 भादवि (281, 125/324/185 बीएनएस)'],
    ['ACC1', '337 भादवि (125 अ बीएनएस)'], ['ACC2', '338 भादवि (125 ब बीएनएस)'], ['DO', 'कलम 285 बीएनएस'],
    ['OBS', 'अश्लील शिवीगाळ कलम 296 बीएनएस'], ['OTH5', 'इतर (भाग 1 ते 5)'], ['ANI', 'प्राणीवाका'], ['ARMS', 'आर्म ॲक्ट'],
    ['ECA', 'जीवनावश्यक वस्तू अधिनियम 1955'], ['NDPS', 'एनडीपीएस ॲक्ट'], ['MPA', 'मपोका'], ['LAB', 'वेठबिगारी / बाल मजुरी'],
    ['IT', 'माहिती तंत्रज्ञान अधिनियम 2000'], ['MV', 'मोटार वाहन कायदा 1988'], ['OTH6', 'इतर विशेष व स्थानिक कायदे'],
    ['GAMB', 'जुगार कायदा'], ['PROH', 'महाराष्ट्र दारूबंदी कायदा'], ['AD', 'मर्ग']
  ];
  var HEAD_MR = {}; HEADS.forEach(function (h) { HEAD_MR[h[0]] = h[1]; });
  function has(K, n) { return new RegExp('(^|[^0-9(./])' + n + '(?![0-9])').test(K); }
  function classify(c) {
    if (c.sec === 'AD' || c.kind === 'AD') return 'AD';
    if (c.sec === 'GAMB') return 'GAMB';
    if (c.sec === 'PROH') return 'PROH';
    var K = asc(c.kalam || '').replace(/बी|b|B/g, 'ब').replace(/\(\s*ए\s*\)|\(\s*a\s*\)/gi, '(अ)');
    var T = (c.label || '') + ' ' + (c.facts || '').slice(0, 300);
    var any = function () { for (var i = 0; i < arguments.length; i++) if (has(K, arguments[i])) return true; return false; };
    if (/पोक्सो|POCSO|पोस्को/i.test(K) && !any(74, 75, 76, 77, 78, 79, 354) || any(64, 70, 376)) return 'RAPE';
    if (any(103, 302)) return 'MUR';
    if (any(105) || (any(304) && !/304\s*\(\s*(अ|ब)/.test(K))) return 'CH';
    if (any(109, 307)) return 'ATM';
    if (any(108, 306)) return 'SUI';
    if (any(106) || /304\s*\(\s*अ/.test(K)) return 'NEG';
    if (/पोक्सो|POCSO|पोस्को/i.test(K) || any(74, 75, 76, 77, 78, 79, 354)) return 'MOL';
    if (any(137, 140, 87, 363, 366)) return 'KID';
    if (any(309, 310, 311, 392, 394, 395)) return 'ROB';
    if (c.sec === 'HBT' || any(331, 454, 457, 380)) return /331\s*\(\s*[2468]\s*\)/.test(K) || any(457) ? 'HBTN' : 'HBTD';
    if (c.sec === 'MVT') return 'MVT';
    if (any(303, 379)) return /मोटार\s*सायकल|मो\.?\s*सा\.?|दुचाकी|वाहन\s*चोरी|मोपेड|स्कुटी|अॅक्टीवा/.test(T) ? 'MVT' : 'THEFT';
    if (any(305, 306, 381)) return 'THEFT';
    if (any(316, 406, 409)) return 'CBT';
    if (any(318, 420)) return 'CHEAT';
    if (any(117, 118, 323, 324, 325, 326) && !(any(324) && /बीएनएस|भान्यासं|भा\.?\s*न्या/.test(K) && !any(117, 118))) return 'HURT';
    if (/125\s*\(\s*ब/.test(K) || any(338)) return 'ACC2';
    if (/125\s*\(\s*अ/.test(K) || any(337)) return 'ACC1';
    if (any(281, 279, 336) || (any(125) && /मोवाका|मो\.?\s*वा/.test(K))) return 'ACC';
    if (any(285, 283)) return 'DO';
    if (any(296, 294)) return 'OBS';
    if (/प्राणी|पशु|प्राण्यां|जनावर|गोवंश/.test(K + ' ' + T)) return 'ANI';
    if (/आर्म|शस्त्र|हत्यार|भा\.?\s*ह\.?\s*का|arms/i.test(K)) return 'ARMS';
    if (/जी?ि?वनावश्यक|ई\.?\s*सी\.?\s*अॅ?क्ट|E\.?C\.?\s*Act/i.test(K)) return 'ECA';
    if (/एन[एन.\s]*ड[ीि][.\s]*प[ीि][.\s]*एस|NDPS|अंमली|अमली/i.test(K + ' ' + (c.label || ''))) return 'NDPS';
    if (/मपोका|म\.?\s*पो\.?\s*का|पोलीस\s*कायदा/.test(K)) return 'MPA';
    if (/बाल\s*मजु|वेठबिगार|बालकामगार/.test(K)) return 'LAB';
    if (/आय\.?\s*टी|IT\s*Act|माहिती\s*तंत्रज्ञान|66\s*\(\s*[सडcdCD]/i.test(K)) return 'IT';
    if (/मोवाका|मो\.?\s*वा\.?\s*का|M\.?\s*V\.?\s*Act/i.test(K)) return 'MV';
    if (/12\s*\(\s*अ|जुगार/.test(K)) return 'GAMB';
    if (/मदाका|म\.?\s*दा\.?\s*का|दारू/.test(K)) return 'PROH';
    return c.sec === 'SLL' ? 'OTH6' : 'OTH5';
  }
  /* 7 वर्षांवरील शिक्षा असलेली शीर्षके (ई-साक्ष / SID आवश्यक) */
  var GRAVE = { MUR: 1, CH: 1, ATM: 1, RAPE: 1, ROB: 1, KID: 1 };

  /* ---------- पूर्ण DCR ---------- */
  function extract(blocks, fileName) {
    var R = { reportDate: '', issueDate: '', jaKr: '', cases: [], issues: [], counts: {}, heads: {}, totals: {} };
    var sec = null, maps = {}, div = '', seen = {};
    function issue(level, msg, key) { R.issues.push({ level: level, msg: msg, key: key || '' }); }
    blocks.forEach(function (b, bi) {
      if (b.t === 'p') {
        var t = b.text;
        if (!t) return;
        if (!R.issueDate && bi < 12) { var m = t.match(/^दि\.?\s*([०-९\d\/.\-]+)/); if (m) R.issueDate = dateOf(m[1]); }
        var s = secOf(t);
        if (s) { sec = s; div = ''; var hc = headCount(t); if (hc !== null) R.heads[s] = (R.heads[s] || 0) + hc; }
        else if (sec && t.replace(/[-\s0]/g, '').length > 2) sec = 'OTHER';
        return;
      }
      if (!sec) {
        b.rows.forEach(function (row) {
          row.forEach(function (c) {
            if (!c) return;
            var t = c.join(' ');
            var dm = t.match(/दिनांक\s*[–\-:]*\s*([०-९\d\/.\-]{8,10})/); if (dm && !R.reportDate) R.reportDate = dateOf(dm[1]);
            var jm = asc(t).match(/जा\.?\s*क्र\.?\s*[:\-]*\s*([\d]+\s*\/\s*\d{4})/); if (jm && !R.jaKr) R.jaKr = jm[1].replace(/\s/g, '');
          });
        });
        return;
      }
      if (sec === 'OTHER') return;
      var map = maps[sec] || null;
      b.rows.forEach(function (row) {
        var cm = colMap(row);
        if (cm) { map = maps[sec] = cm; return; }
        var ne = row.filter(function (c) { return c && c.length; });
        if (!ne.length) return;
        var joined = clean(ne.map(function (c) { return c.join(' '); }).join(' | '));
        if (ne.length >= 5 && ne.every(function (c) { return /^\d{1,2}$/.test(asc(c.join(''))); })) return;
        if (joined.length < 120 && /एकु[णण]|एकूण/.test(joined) && !/कलम|\d+\s*\/\s*\d{2}/.test(joined)) {
          var tm = asc(joined).match(/एकु[णण]\s*[-–:]?\s*(\d{1,3})/); if (tm) R.totals[sec] = +tm[1];
          return;
        }
        if (ne.length <= 2 && /विभाग/.test(joined) && joined.length < 40) { div = normDiv(joined.split('|')[0]); return; }
        if (!map) return;
        var get = function (f) { var i = map[f]; return i === undefined || !row[i] ? [] : row[i]; };
        var pc = get('psCr');
        if (!pc.length) return;
        if (/एकु[णण]|एकूण/.test(pc.join(' ')) && pc.join(' ').length < 40) { var tm2 = asc(joined).match(/एकु[णण]\s*[-–:]?\s*(\d{1,3})/); if (tm2) R.totals[sec] = +tm2[1]; return; }
        var p = parsePsCr(pc);
        if (!p.ps && !p.no) return;
        var c = {
          sec: sec, kind: p.kind, ps: p.ps, div: div, no: p.no, yr: p.yr, kalam: p.kalam, label: p.label, sid: p.sid,
          sr: clean(get('sr').join(' ')), psCr: pc.join(' / ')
        };
        ['comp', 'acc', 'victim', 'deceased', 'place', 'occ', 'reg', 'prop', 'io', 'incharge', 'regOff', 'visit', 'facts', 'prior', 'priorAction', 'arrest', 'arrestReason', 'cause', 'nCases', 'nAcc']
          .forEach(function (f) { c[f] = get(f).join(' / '); });
        c.regDate = dateOf(c.reg);
        c.occDate = dateOf(c.occ);
        if (sec === 'COMM' && c.kind === 'FIR' && /एन\.?\s*सी/.test(c.psCr)) c.kind = 'NC';
        c.key = c.kind + '|' + c.ps + '|' + c.no + '|' + c.yr;
        c.head = c.kind === 'NC' ? '' : classify(c);
        R.cases.push(c);
        R.counts[sec] = (R.counts[sec] || 0) + 1;
      });
    });
    /* अहवालाची तारीख: DCR मधील, फाइलच्या नावातील आणि बहुतेक गुन्ह्यांची दाखल तारीख यांची तुलना */
    var fd = fileName ? dateOf(String(fileName).replace(/_/g, ' ')) : '', freq = {}, mode = '', mx = 0;
    R.cases.forEach(function (c) { if (c.regDate) { freq[c.regDate] = (freq[c.regDate] || 0) + 1; if (freq[c.regDate] > mx) { mx = freq[c.regDate]; mode = c.regDate; } } });
    R.docDate = R.reportDate;
    if (!R.reportDate) R.reportDate = fd || mode;
    else if (mode && mode !== R.reportDate && (fd === mode || mx >= R.cases.length / 2)) {
      R.issues.push({ level: 'warn', msg: 'DCR मध्ये "दिनांक – ' + dmy(R.reportDate) + '" आहे, पण बहुतेक गुन्हे ' + dmy(mode) + ' रोजी दाखल आहेत. अहवालाची तारीख ' + dmy(mode) + ' धरली.', key: '' });
      R.reportDate = mode;
    }
    check(R);
    return R;
  }

  function check(R) {
    var I = R.issues, seen = {};
    function add(level, msg, key) { I.push({ level: level, msg: msg, key: key || '' }); }
    if (!R.reportDate) add('bad', 'DCR ची तारीख (दिनांक – ...) सापडली नाही. फाइलचे नाव CRO_DCR_दिवस-महिना-वर्ष ठेवा.');
    if (!R.cases.length) add('bad', 'या फाइलमध्ये एकही गुन्हा सापडला नाही. ही दैनिक गुन्हे अहवालाची Word फाइल आहे का?');
    var cnt = R.counts, H = R.heads;
    if ('IPC' in H) {
      var n5 = (cnt.IPC || 0) + (cnt.MVT || 0) + (cnt.HBT || 0);
      if (H.IPC !== n5) add('bad', 'भाग 1 ते 5 शीर्षकात ' + H.IPC + ' लिहिले आहे, पण तक्त्यांत ' + n5 + ' गुन्हे आहेत (वाहन चोरी व घरफोडीसह).');
    }
    ['MVT', 'HBT', 'SLL', 'AD', 'GAMB', 'PROH'].forEach(function (s) {
      if (s in H && H[s] !== (cnt[s] || 0)) add('bad', SEC_MR[s] + ' शीर्षकात ' + H[s] + ' लिहिले आहे, पण तक्त्यात ' + (cnt[s] || 0) + ' ओळी आहेत.');
      if (s in R.totals && R.totals[s] !== (cnt[s] || 0)) add('bad', SEC_MR[s] + ' तक्त्याच्या एकूण ओळीत ' + R.totals[s] + ' आहे, पण ' + (cnt[s] || 0) + ' ओळी आहेत.');
    });
    R.cases.forEach(function (c) {
      var nm = c.ps + ' ' + (c.kind === 'AD' ? 'मर्ग ' : '') + (c.no || '?') + '/' + (c.yr ? String(c.yr).slice(2) : '?');
      if (seen[c.key]) add('bad', nm + ' हा गुन्हा याच DCR मध्ये दोनदा आला आहे.', c.key);
      seen[c.key] = 1;
      if (!c.ps) add('bad', 'एका ओळीत पो.स्टे.चे नाव वाचता आले नाही: ' + c.psCr.slice(0, 60), c.key);
      if (!c.no) add('bad', nm + ': अप.क्र. / वर्ष वाचता आले नाही.', c.key);
      if (!c.kalam && c.kind !== 'AD') add('warn', nm + ': कलम नमूद नाही.', c.key);
      if (!c.div) add('warn', nm + ': उपविभाग (विभाग) ओळ नाही.', c.key);
      if (!c.io && c.kind !== 'NC') add('warn', nm + ': तपास अधिकारी नमूद नाही.', c.key);
      if (!c.place) add('warn', nm + ': घटनास्थळ नमूद नाही.', c.key);
      if (!c.regDate) add('warn', nm + ': दाखल तारीख वाचता आली नाही.', c.key);
      else if (R.reportDate) {
        if (c.regDate > R.reportDate) add('bad', nm + ': दाखल तारीख ' + dmy(c.regDate) + ' ही DCR तारखेनंतरची आहे.', c.key);
        else if (c.regDate < R.reportDate) add('info', nm + ': दाखल ' + dmy(c.regDate) + ', पण DCR ' + dmy(R.reportDate) + ' चा आहे (उशिरा नोंद?).', c.key);
      }
      if (c.occDate && c.regDate && c.occDate > c.regDate) add('warn', nm + ': घटना तारीख दाखल तारखेनंतरची दिसते.', c.key);
      if (GRAVE[c.head] && !c.sid) add('warn', nm + ' (' + HEAD_MR[c.head] + '): ई-साक्ष / SID क्रमांक नमूद नाही.', c.key);
      if (c.head === 'OTH5' || c.head === 'OTH6') add('info', nm + ' (कलम ' + (c.kalam || '-') + '): B भागाचे शीर्षक ठरले नाही, "इतर" मध्ये टाकले.', c.key);
    });
  }

  /* ---------- DCR चेक पत्र: वर्णन व सूचना ---------- */
  function summary(c, full) {
    var f = clean(c.facts || '').replace(/^(थोडक्यात\s*)?(हकी[कग]त|हकिकत)\s*(अशा|अश्या)\s*(प्रकारे|प्रकारची)\s*आहे\s*की[,\s]*/, '').replace(/^यातील\s*/, 'यातील ');
    f = f.split(' / ')[0];
    f = f.replace(/[,.\s]*(अशा|अश्या)?\s*(फिर्यादी|फि\.|फिर्यादिच्या|फिर्यादीच्या|फिर्यादीचे)[^.]{0,60}(रिपोर्ट|तक्रार)[^.]*(नोंद|दाखल)[^.]*\.?\s*$/, '');
    f = f.replace(/[,.\s]*सदर(चा)?\s*गुन्हा\s*(नोंद|दाखल)[^.]*\.?\s*$/, '');
    if (!full && f.length > 500) { var cut = f.slice(0, 500), i = cut.lastIndexOf('.'); if (i < 250) i = cut.lastIndexOf(' '); f = cut.slice(0, i > 250 ? i : 500); }
    return scrubNames(f, c).replace(/\s+/g, ' ').replace(/[.\s]+$/, '');
  }
  /* पत्रात फिर्यादी / आरोपी / इतरांची नावे जाऊ नयेत: "नामे <नाव>" व DCR मधील नावे काढतो (वय, पत्ता, गाव तसेच राहतात) */
  function scrubNames(t, c) {
    t = String(t || '').replace(/मो\.?\s*(नं\.?|क्र\.?)?\s*[\d ]{10,}/g, '').replace(/नामे\s+(?:फिर्यादी\s+|आरोपी\s+)?(?:श्री\.?\s*|श्रीमती\s*|कु\.?\s*)?[^\d,.()]+?(?=\s+(?:वय|रा\.?\s|राहणार|मो\.?\s|मो\b|ता\.?\s|जि\.?\s|यांची|यांचा|यांचे|हा\s|ही\s|दि\.?\s|\d))/g, '');
    t = t.replace(/\d{10,}/g, '').replace(/(मुलगी|मुलगा|मुलीचे|मुलाचे|मुलीची|मुलाची|मुलीला|मुलाला|पत्नी|पती|भाऊ|बहीण|आई|वडील|मुलीच्या|मुलाच्या)\s+(?:कु\.?\s*|श्री\.?\s*|श्रीमती\s*)?[\u0900-\u097F]+(?:\s+[\u0900-\u097F]+){0,3}?(?=\s+(?:हिचे|हिचा|हिची|हिला|हिने|वय|रा\b|नामे|यांची|यांचा|यांचे|याचा|याचे|चे|च्या|हा|ही))/g, '$1');
    [c.comp, c.acc, c.victim, c.deceased].forEach(function (x) {
      var n = String(x || '').split(/\s+वय|\s+रा[.\s]|\s+मो|\s+ता[.\s]|\s+जि[.\s]|\d|,|\(/)[0].replace(/^(श्री|श्रीमती|कु)\.?\s*/, '').trim();
      if (n.length >= 4 && !/अज्ञात|अनोळखी|^-$/.test(n)) t = t.split(n).join('');
    });
    return t.replace(/\s{2,}/g, ' ').replace(/\s+([,.])/g, '$1');
  }
  return {
    readDocx: readDocx, blocksFromXml: blocksFromXml, extract: extract, classify: classify, parsePsCr: parsePsCr,
    HEADS: HEADS, HEAD_MR: HEAD_MR, SEC_MR: SEC_MR, dateOf: dateOf, dmy: dmy, asc: asc, summary: summary, scrubNames: scrubNames, normDiv: normDiv
  };
})();

/* DCR चेक (तपास सूचना) पत्र: जुन्या 75 पत्रांवरून घेतलेली सूचना-वाक्ये, शीर्षकानुसार निवड, शेऱ्यावरून सुचवण.
 * { } मधील जागा ॲप आपोआप भरते. सेटिंग्जमधून ही वाक्ये बदलता / वाढवता येतात. */
var LETTER = (function () {
  var ADDL = 'अपर पोलीस अधीक्षक, नागपूर जिल्हा (ग्रामीण) तथा विशेष कार्यकारी दंडाधिकारी, नागपूर ग्रामीण यांच्याकडे सादर करावे. अपर पोलीस अधीक्षक, नागपूर जिल्हा (ग्रामीण) यांनी त्याच्या वर्तणुकीस प्रतिबंध व्हावा अशा प्रकारे कार्यवाही करावी.';
  var PARAS = [
    { k: 'shera', t: 'मा. पोलीस अधीक्षक यांचा शेरा', x: 'सदर गुन्ह्याच्या तपासात {शेरा}.' },
    { k: 'naka', t: 'वाहन चोरी: अचानक नाकाबंदी', x: 'मासिक गुन्हे बैठकीत सर्व ठाणे प्रभारी अधिकारी यांना वाहन चोरीच्या गुन्ह्यांवर प्रतिबंध व्हावा याकरिता पोलीस ठाणे हद्दीतील चोरीची संवेदनशील ठिकाणे चिन्हित करून पोलीस ठाणे हद्दीत अचानकपणे नाकाबंदी आयोजित करून वाहन मालकीसंबंधी कागदपत्रांची तपासणी करण्याचे, परंतु अचानक नाकाबंदी आयोजित करण्याच्या दोन दिवसांपूर्वी स्थानिक स्तरावर योग्य प्रसिद्धी करावी, असे निर्देश देण्यात आलेले आहेत. तरी पोलीस ठाणे हद्दीतून मोटार सायकल चोरीचे गुन्हे घडत आहेत, यावरून आपण निर्देशांची अवहेलना करून कर्तव्यात दुर्लक्ष करीत असल्याचे स्पष्टपणे दिसून येत आहे. आपण शेवटची अचानक नाकाबंदी केव्हा राबवली व नाकाबंदी दरम्यान केलेल्या कार्यवाहीची माहिती, तसेच नाकाबंदीकरिता जाताना व परत येतानाचे स्टेशन डायरी सान्हे सादर करावे.' },
    { k: 'known', t: 'ओळखीचे / नातेवाईक: प्रतिबंधक कार्यवाही', x: 'सदर गुन्ह्यातील फिर्यादी / आरोपी हे {नाते} असून एकाच {ठिकाण} राहतात. तरी आरोपीविरुद्ध त्वरित आवश्यक ती प्रतिबंधात्मक कार्यवाही करावी.' },
    { k: 'unknown', t: 'अनोळखी आरोपी: शोध व प्रतिबंधक', x: 'सदर गुन्ह्यातील फिर्यादी / आरोपी हे अनोळखी असून आरोपींचा तात्काळ शोध घेऊन त्यांच्यावर त्वरित आवश्यक ती प्रतिबंधात्मक कार्यवाही करावी.' },
    { k: 'prior', t: 'आरोपीवर पूर्वीचे गुन्हे', x: 'सदर गुन्ह्याचे गांभीर्य पाहता गुन्ह्यातील आरोपीवर {पूर्वगुन्हे} अन्वये गुन्हे नोंद असल्याचे दैनिक गुन्हे अहवालावरून दिसून येत आहे.' },
    { k: 'prev', t: 'प्रतिबंधक कार्यवाही: अपर पो. अधीक्षक', x: 'तरी सदर गुन्ह्यातील आरोपीविरुद्ध त्वरित आवश्यक ती प्रतिबंधात्मक कार्यवाही करून पुढील कार्यवाहीकरिता ' + ADDL, addl: 1 },
    { k: 'arrest', t: 'आरोपीस तात्काळ अटक', x: 'सदर गुन्ह्याचे गांभीर्य पाहता गुन्ह्यातील आरोपीस तात्काळ अटक करावी.' },
    { k: 'bail', t: 'जामीन रद्द करणे', x: 'आरोपीस पूर्वीच्या गुन्ह्यात मिळालेला जामीन रद्द करण्याबाबत मा. न्यायालयात कार्यवाही करावी.' },
    { k: 'sand', t: 'रेती / गौण खनिज चोरी', x: 'सदरहू प्रकरणात वाहन मालकास आरोपी करावे. गुन्ह्यातील आरोपींचा पूर्व गुन्हे अभिलेख तपासून त्यांच्याविरुद्ध त्वरित आवश्यक ती प्रतिबंधात्मक कार्यवाही करून ' + ADDL + ' तसेच आरोपीने विना रॉयल्टी रेती चोरी करून शासनाच्या महसूल विभागाची फसवणूक केली असून त्याबाबत दंड वसूल करणेबाबत तहसील कार्यालयाशी पत्रव्यवहार करावा.', addl: 1 },
    { k: 'traffic', t: 'अपघात: वाहतूक शाखा भेट', x: 'सदरहू प्रकरणी प्रभारी अधिकारी, वाहतूक शाखा यांनी घटनास्थळी प्रत्यक्ष भेट देऊन कार्यवाहीचा अहवाल सादर करावा.' },
    { k: 'sdpo', t: 'गंभीर गुन्हा: SDPO भेट व मार्गदर्शन', x: 'सदर गुन्ह्याचे गांभीर्य पाहता उपविभागीय पोलीस अधिकारी, {विभाग} विभाग यांनी गुन्ह्यात भेट द्यावी व गुन्हा उघडकीस आणण्याबाबत तपासी अधिकारी यांना योग्य मार्गदर्शन करावे.' },
    { k: 'ad', t: 'मर्ग: SDPO भेट (घातपात?)', x: 'सदरहू मर्गचे गांभीर्य पाहता उपविभागीय पोलीस अधिकारी, {विभाग} विभाग यांनी सदर मर्गच्या घटनास्थळी प्रत्यक्ष भेट देऊन सदर प्रकरणात घातपातासारखा प्रकार आहे काय याची शहानिशा करावी व चौकशी अधिकारी यांना मर्गच्या चौकशीत योग्य मार्गदर्शन करावे.' },
    { k: 'minor', t: 'अल्पवयीन अपहरण: स्थायी आदेश 21/2012', x: 'सदर गुन्ह्याच्या तपासात हरविलेल्या बालकांच्या तपासासंदर्भात स्थायी आदेश क्र. 21/2012 दि. 05/10/2012 अन्वये तपासासंबंधी कार्यप्रणाली (SOP) निर्गमित केली असून त्या मार्गदर्शक तत्त्वानुसार तपास करून अल्पवयीन पीडितेचा शोध घेण्याचे सर्वतोपरी प्रयत्न करावे.' },
    { k: 'missing', t: 'सज्ञान व्यक्ती अपहरण: शोध', x: 'सदर गुन्ह्यातील पीडित व्यक्तीचा नातेवाईक, मित्र, मोबाईल लोकेशन (CDR) व सीसीटीव्ही फुटेजच्या आधारे तात्काळ शोध घ्यावा.' },
    { k: 'team', t: 'SDPO पथक, 7 दिवसांत प्रगती अहवाल', x: 'सदरहू गुन्ह्याचे गांभीर्य पाहता उपविभागीय पोलीस अधिकारी, {विभाग} विभाग यांनी सदर गुन्ह्याच्या तपासासंबंधाने एक पथक तयार करून गुन्ह्याचा तपास करावा व गुन्हा उघडकीस आणावा, तसेच सदर गुन्ह्यासंबंधाने प्रगती अहवाल 7 दिवसांत या कार्यालयास सादर करावा.' },
    { k: 'transfer', t: 'तपास दुसऱ्या अधिकाऱ्याकडे वर्ग', x: 'सदर गुन्ह्याचा तपास {अधिकारी} यांच्याकडे वर्ग करण्यात यावा.' },
    { k: 'cattle', t: 'अवैध जनावर वाहतूक SOP', x: 'सदर प्रकरणात वाहन मालकास आरोपी करावे. अशा प्रकारच्या गुन्ह्यांच्या तपासासंदर्भात या कार्यालयाने निर्गमित केलेल्या कार्यप्रणालीनुसार (SOP) गुन्ह्याचा तपास करावा.' },
    { k: 'cctv', t: 'सीसीटीव्ही / तांत्रिक तपास', x: 'घटनास्थळ परिसरातील सीसीटीव्ही फुटेज, मोबाईल टॉवर डम्प व तांत्रिक विश्लेषणाच्या आधारे आरोपीचा तात्काळ शोध घ्यावा.' },
    { k: 'property', t: 'मुद्देमाल हस्तगत', x: 'गुन्ह्यातील चोरीस गेलेला मुद्देमाल हस्तगत करण्याचे सर्वतोपरी प्रयत्न करावे.' },
    { k: 'reward', t: 'चांगली कामगिरी: बक्षीस अहवाल', x: 'सदर गुन्ह्यात कार्यवाहीकरिता हजर असलेले अधिकारी व अंमलदार यांच्या सदर कामगिरीबद्दल बक्षीस अहवाल उपविभागीय पोलीस अधिकारी, {विभाग} विभाग यांच्या मार्फतीने सादर करावा.' }
  ];
  /* B भाग शीर्षक → आपोआप निवडली जाणारी वाक्ये */
  var BY_HEAD = {
    MUR: ['sdpo', 'arrest'], CH: ['sdpo'], ATM: ['sdpo', 'arrest', 'prev'], SUI: ['sdpo'], NEG: ['traffic'], RAPE: ['arrest', 'prev'], MOL: ['prev'],
    KID: ['minor'], ROB: ['sdpo', 'cctv'], HURT: ['known'], MVT: ['naka'], THEFT: ['cctv', 'property'], HBTN: ['sdpo', 'cctv'], HBTD: ['sdpo', 'cctv'],
    CBT: ['property'], CHEAT: ['property'], ACC: ['traffic'], ACC1: ['traffic'], ACC2: ['traffic'], OBS: ['known'], ANI: ['cattle'], NDPS: ['reward'], AD: ['ad']
  };
  /* शेऱ्यातील शब्द → सुचवायची वाक्ये */
  var BY_WORD = [
    [/भेट|visit/i, 'sdpo'], [/प्रतिबंध|MPDA|तडीपार|चॅप्टर|prevent/i, 'prev'], [/अटक|arrest/i, 'arrest'], [/नाकाबंदी/, 'naka'],
    [/वाहतूक|ट्राफिक|traffic/i, 'traffic'], [/पथक|टीम|team/i, 'team'], [/वर्ग|transfer/i, 'transfer'], [/बक्षीस|reward/i, 'reward'],
    [/जामीन|bail/i, 'bail'], [/सीसीटीव्ही|CCTV|तांत्रिक|CDR/i, 'cctv'], [/मुद्देमाल|माल\s*हस्त/, 'property'], [/शोध/, 'unknown'], [/मर्ग|घातपात/, 'ad']
  ];
  function para(k, list) { list = list || PARAS; for (var i = 0; i < list.length; i++) if (list[i].k === k) return list[i]; return null; }
  /* शीर्षक व DCR मधील माहितीवरून आपोआप निवडलेली वाक्ये (शेरा असल्यास तो सर्वात आधी) */
  function suggest(c, shera, list) {
    var ks = (BY_HEAD[c.head] || []).slice();
    if (c.head === 'KID' && /वयस्क|सज्ञान/.test(c.psCr + ' ' + c.label) && !/अल्पवयीन/.test(c.psCr)) ks = ['missing'];
    if (c.head === 'THEFT' && /रेती|गौण\s*खनिज|खाण|रॉयल्टी|राॅयल्टी|48\s*\(\s*7/.test(c.kalam + ' ' + c.facts)) ks = ['sand'];
    if (c.head === 'HURT' && /अज्ञात|अनोळखी/.test(c.acc)) ks = ['unknown'];
    if (hasPrior(c)) { ks.push('prior'); if (ks.indexOf('prev') < 0) ks.push('prev'); }
    if (shera && shera.trim()) ks.unshift('shera');
    var seen = {};
    return ks.filter(function (k) { if (seen[k] || !para(k, list)) return false; seen[k] = 1; return true; });
  }
  /* शेऱ्यातील शब्दांवरून "हेही जोडता येईल" अशी वाक्ये (आपोआप टिक होत नाहीत, कारण शेरा स्वतः पत्रात जातो) */
  function hints(shera, list) {
    var o = [];
    (shera || '').split(/[,.;\n]/).forEach(function (w) { BY_WORD.forEach(function (b) { if (b[0].test(w) && o.indexOf(b[1]) < 0 && para(b[1], list)) o.push(b[1]); }); });
    return o;
  }
  function hasPrior(c) {
    var p = String(c.prior || '').replace(/\s*\/\s*/g, ' ');
    return /\d+\s*\/\s*\d{2}/.test(p) && !/अप्राप्त|निरंक|नाही|मिळून\s*आले\s*नाही/.test(p.slice(0, 40));
  }
  function priorText(c) { return String(c.prior || '').replace(/\s*\/\s*(?=अप)/g, ', ').replace(/\s*\/\s*/g, ' ').replace(/\s+/g, ' ').trim(); }

  /* पत्राचा मसुदा (सर्व भाग वेगळे, म्हणजे स्क्रीनवर बदलता येतात) */
  function draft(c, o) {
    o = o || {};
    var S = o.settings || {}, list = o.paras || PARAS;
    var ps = c.ps, div = c.div || '', crTxt = c.kind === 'AD' ? 'मर्ग क्र. ' + c.no + '/' + c.yr : 'अप.क्र. ' + c.no + '/' + c.yr;
    var kal = c.kalam ? ' कलम ' + c.kalam : '';
    var vars = { 'विभाग': div, 'नाते': o.relation || 'एकमेकांच्या ओळखीचे', 'ठिकाण': o.place || 'गावात', 'पूर्वगुन्हे': priorText(c), 'अधिकारी': o.officer || '______', 'शेरा': (o.shera || '').trim().replace(/[.।\s]+$/, '') };
    function fill(t) { return t.replace(/\{([^}]+)\}/g, function (m, k) { return k in vars ? vars[k] : m; }); }
    var keys = o.keys || suggest(c, o.shera, list);
    var dd = DCR.dmy(o.refDate || c.issueDate || c.dcrIssue || ''), rd = DCR.dmy(c.regDate || '');
    var sum = o.summary != null ? o.summary : DCR.summary(c, true);
    var intro = 'दिनांक ' + (dd || '__/__/____') + ' च्या दैनिक गुन्हे अहवालाचे अवलोकन केले असता दिनांक ' + (rd || '__/__/____') + ' रोजी पोलीस ठाणे ' + ps + ' ' + crTxt + kal +
      (c.kind === 'AD' ? ' अन्वये मर्ग दाखल असून' : ' अन्वये गुन्हा नोंद असून') + ' दैनिक गुन्हे अहवालात नमूद केल्याप्रमाणे ' + (sum ? sum.replace(/[.।\s]+$/, '') : '__________') +
      (c.kind === 'AD' ? ', अशा रिपोर्टवरून सदर मर्ग नोंद करण्यात आला आहे.' : ', अशा फिर्यादीच्या तक्रारीवरून सदर गुन्हा नोंद करण्यात आला आहे.');
    var body = keys.map(function (k) { var p = para(k, list); return p ? fill(p.x) : ''; }).filter(Boolean);
    var addl = keys.some(function (k) { var p = para(k, list); return p && p.addl; });
    var days = S.days || 3;
    var prati = ['ठाणेप्रभारी अधिकारी, पोलीस ठाणे ' + ps];
    if (addl) prati.unshift('अपर पोलीस अधीक्षक, नागपूर जिल्हा (ग्रामीण)');
    var R = o.rules || [];
    var ap = function (t) { return applyRules(t, R); };
    return {
      jaKr: o.jaKr || '', date: o.date || '',
      ref: ap('दिनांक ' + (dd || '__/__/____') + ' चा दैनिक गुन्हे अहवाल.'),
      subject: ap('पोलीस ठाणे ' + ps + ' ' + crTxt + kal + (c.kind === 'AD' ? ' च्या मर्गच्या चौकशीकरिता सूचना.' : ' या गुन्ह्याच्या तपासाकरिता सूचना.')),
      intro: ap(intro), body: body.map(ap), close: ap('केलेल्या कार्यवाहीचा अहवाल ' + days + ' दिवसांत या कार्यालयात सादर करावा.'),
      prati: prati, prat: div ? ['उपविभागीय पोलीस अधिकारी, ' + div + ' विभाग'] : [], keys: keys
    };
  }
  /* शिकलेले नियम: { f: जुना मजकूर, t: नवा मजकूर }. SP/ठाणे पत्रात केलेले बदल पुढील पत्रांत आपोआप लागू होतात. */
  function applyRules(text, rules) {
    var t = String(text || '');
    (rules || []).forEach(function (r) { if (r && r.f && r.f.length >= 3 && t.indexOf(r.f) >= 0) t = t.split(r.f).join(r.t || ''); });
    return t.replace(/\s{2,}/g, ' ').replace(/\s+([,.।])/g, '$1');
  }
  /* मूळ व बदललेल्या मजकुरातील फरक (पहिल्या व शेवटच्या समान शब्दांनंतरचा भाग). अंक असलेले फरक (तारीख, क्रमांक) नियम होत नाहीत. */
  function diffRule(a, b) {
    a = String(a || '').trim(); b = String(b || '').trim();
    if (a === b) return null;
    var x = a.split(/\s+/), y = b.split(/\s+/), i = 0, j = 0;
    while (i < x.length && i < y.length && x[i] === y[i]) i++;
    while (j < x.length - i && j < y.length - i && x[x.length - 1 - j] === y[y.length - 1 - j]) j++;
    var f = x.slice(i, x.length - j).join(' '), t = y.slice(i, y.length - j).join(' ');
    if (!f) { /* फक्त भर पडला: आधीच्या दोन शब्दांना धरून घालतो */
      var k = Math.max(0, i - 2), pre = x.slice(k, i).join(' ');
      if (!pre) return null;
      f = pre; t = pre + ' ' + t;
    }
    if (/\d/.test(f) || /\d/.test(t) || f.length < 3) return null;
    return { f: f, t: t };
  }
  return { PARAS: PARAS, BY_HEAD: BY_HEAD, suggest: suggest, hints: hints, draft: draft, para: para, hasPrior: hasPrior, applyRules: applyRules, diffRule: diffRule };
})();
