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
  try { return id ? DriveApp.getFileById(id).getBlob() : null; } catch (e) { return null; }
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
    var p = cells[k].getChild(0).asParagraph(), w = k === 'C' ? 48 : 70;
    p.setAlignment(A.CENTER);
    var im = p.appendInlineImage(bl); im.setWidth(w); im.setHeight(w * im.getHeight() / im.getWidth());
  });
  var hasC = !!logoOf_(S, 'C');
  off.forEach(function (l, i) {
    var p = (i === 0 && !hasC) ? cells.C.getChild(0).asParagraph() : cells.C.appendParagraph('');
    p.setAlignment(A.CENTER).setSpacingAfter(0);
    rich_(p, l, { kr: kr, font: F, bold: i === 0, size: i === 0 ? sz(14, 20) : (i === off.length - 1 ? sz(9, 13) : sz(11, 16)) });
  });
  ht.setColumnWidth(0, 90); ht.setColumnWidth(2, 90); ht.setColumnWidth(1, 612 - 115 - 90 - 60);
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
