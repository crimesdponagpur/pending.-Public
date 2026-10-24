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
