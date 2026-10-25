/**
 * जिल्हा DCR चेक व B भाग ॲप (नागपूर ग्रामीण) • स्वतंत्र Google Apps Script Web app.
 * रोजचा DCR (Word) → गुन्हे नोंदवही + त्रुटी, गुन्ह्यावर SP शेरा → DCR चेक (तपास सूचना) पत्र Word + PDF,
 * महिन्याचे सर्व DCR → हेडवाइज B भाग Word + PDF. सर्व फाइल्स "DCR_चेक_ॲप" या Drive फोल्डरमध्ये आपोआप जतन.
 * MCR / प्रलंबित ॲपशी याचा काही संबंध नाही: स्वतःचा फोल्डर, स्वतःची Sheet.
 */
var VERSION = '2.0.0';
var ROOT_NAME = 'DCR_चेक_ॲप';
var SS_NAME = 'DCR_नोंदवही';
var CASE_COLS = ['key', 'kind', 'sec', 'ps', 'div', 'no', 'yr', 'kalam', 'label', 'sid', 'head', 'headSet', 'comp', 'acc', 'victim', 'deceased', 'place',
  'occ', 'reg', 'regDate', 'occDate', 'prop', 'io', 'incharge', 'regOff', 'visit', 'facts', 'prior', 'priorAction', 'arrest', 'arrestReason', 'cause',
  'nCases', 'nAcc', 'psCr', 'dcrDate', 'dcrIssue', 'dcrList', 'shera', 'sheraAt', 'letters', 'updated', 'bRow'];
var SHEETS = {
  CASES: CASE_COLS,
  DCRS: ['reportDate', 'issueDate', 'jaKr', 'fileName', 'fileUrl', 'pdfUrl', 'cases', 'bad', 'warn', 'info', 'uploaded', 'bAdded'],
  ISSUES: ['reportDate', 'level', 'msg', 'key'],
  LETTERS: ['id', 'jaKr', 'date', 'key', 'ps', 'div', 'crime', 'head', 'shera', 'keys', 'docxUrl', 'pdfUrl', 'created', 'status', 'replyDate', 'note', 'krutiUrl'],
  BPART: ['month', 'scope', 'cases', 'docxUrl', 'pdfUrl', 'created'],
  PS: ['ps', 'div'],
  SETTINGS: ['key', 'value']
};
var DEF = {
  spName: '', spDesig: 'भा.पो.से. पोलीस अधीक्षक,\nनागपूर जिल्हा (ग्रामीण), नागपूर.',
  office: 'कार्यालय पोलीस अधीक्षक, नागपूर जिल्हा (ग्रामीण)\nसिव्हील लाईन, नागपूर - 440001\nई-मेल: spreader-ngp.r@mahapolice.gov.in • दूरध्वनी क्र. 0712-2566650 • फॅक्स क्र. 0712-2560779',
  jaKrPrefix: 'नाजिग्रा/डी-32/वाशा/दैगुअ तपास सूचना', jaKrFile: '26-1962', autoNo: 'होय', letterSeq: '0', rules: '[]', logoL: '', logoC: '', logoR: '', days: '3', font: 'Mukta',
  divOrder: 'काटोल,सावनेर,नागपूर,उमरेड,रामटेक,कन्हान', distName: 'नागपूर ग्रामीण जिल्हा',
  quick: 'आरोपीस तात्काळ अटक करावी\nउपविभागीय पोलीस अधिकारी यांनी घटनास्थळी भेट द्यावी\nआरोपीविरुद्ध प्रतिबंधक कार्यवाही करावी\nसीसीटीव्ही फुटेज तपासावे\nमुद्देमाल हस्तगत करावा\nगुन्हा लवकर उघडकीस आणावा\nपीडितेचा तात्काळ शोध घ्यावा\nतपास दुसऱ्या अधिकाऱ्याकडे वर्ग करावा',
  paras: '', quickHead: ''
};
/* शीर्षकनिहाय नेहमीचे शेरे (सेटिंग्जमध्ये बदलता येतात) */
DEF.quickHead = (function () {
  var G = [
    ['MUR,CH,ATM', 'आरोपीस तात्काळ अटक करावी\nउपविभागीय पोलीस अधिकारी यांनी घटनास्थळी भेट द्यावी\nशवविच्छेदन / वैद्यकीय अहवाल तात्काळ प्राप्त करावा\nगुन्ह्यात वापरलेले हत्यार जप्त करावे\nप्रत्यक्षदर्शी साक्षीदारांचे जबाब नोंदवावे\nदोषारोपपत्र मुदतीत पाठवावे'],
    ['RAPE,MOL', 'पीडितेची वैद्यकीय तपासणी तात्काळ करावी\nपीडितेचा जबाब मा. न्यायदंडाधिकारी यांचेसमोर नोंदवावा\nआरोपीस तात्काळ अटक करावी\nपोक्सो कलमे लागू होत असल्यास वाढवावी\nदोषारोपपत्र 60 दिवसांत पाठवावे'],
    ['KID', 'पीडितेचा तात्काळ शोध घ्यावा\nमोबाईल CDR / लोकेशन तपासावे\nनातेवाईक व मित्रांकडे चौकशी करावी\nबाल कल्याण समितीस कळवावे'],
    ['ROB,THEFT,HBTN,HBTD', 'सीसीटीव्ही फुटेज तपासावे\nमुद्देमाल हस्तगत करावा\nअभिलेखावरील आरोपींची तपासणी करावी\nश्वान पथक / ठसे तज्ञांची मदत घ्यावी\nगुन्हा लवकर उघडकीस आणावा'],
    ['MVT', 'वाहनाचा क्रमांक सर्व पोलीस ठाण्यांना कळवावा\nसीसीटीव्ही फुटेज तपासावे\nवाहन हस्तगत करावे\nअभिलेखावरील वाहनचोरांची तपासणी करावी'],
    ['CHEAT,CBT,IT', 'संबंधित बँक खाते तात्काळ गोठवावे\nसायबर सेलची मदत घ्यावी\nआरोपीचे मोबाईल व बँक तपशील मिळवावे\nफसवणुकीची रक्कम परत मिळवावी'],
    ['ACC,ACC1,ACC2,NEG,DO', 'वाहन जप्त करावे\nचालकाचा वाहन परवाना तपासावा\nRTO तपासणी अहवाल मिळवावा\nघटनास्थळ पंचनामा व नकाशा तयार करावा'],
    ['SUI', 'आत्महत्येस प्रवृत्त करणाऱ्यांचा शोध घ्यावा\nचिठ्ठी / मोबाईल तपासावा\nनातेवाईकांचे जबाब नोंदवावे'],
    ['AD', 'शवविच्छेदन अहवाल प्राप्त करावा\nमृत्यूचे कारण स्पष्ट करावे\nनातेवाईकांचे जबाब नोंदवावे\nघातपात नसल्याची खात्री करावी'],
    ['HURT,OBS', 'वैद्यकीय अहवाल प्राप्त करावा\nआरोपीविरुद्ध प्रतिबंधक कार्यवाही करावी'],
    ['NDPS,ARMS', 'मूळ स्रोताचा शोध घ्यावा\nमुद्देमाल रासायनिक विश्लेषणास पाठवावा\nआरोपीचा पूर्वइतिहास तपासावा'],
    ['PROH,GAMB', 'आरोपीविरुद्ध प्रतिबंधक कार्यवाही करावी\nसराईत आरोपीविरुद्ध हद्दपारी / MPDA प्रस्ताव पाठवावा\nमूळ स्रोताचा शोध घ्यावा']
  ], o = {};
  G.forEach(function (g) { g[0].split(',').forEach(function (h) { o[h] = g[1]; }); });
  return JSON.stringify(o);
})();

/* ---------- Web app ---------- */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('App').setTitle('DCR चेक • B भाग')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
/** ॲपमधील सर्व विनंत्या इथून जातात: उत्तर नेहमी { ok, d } किंवा { ok:false, e } */
function api(fn, args) {
  var F = { init: init_, saveDcr: saveDcr_, savePdf: savePdf_, search: search_, saveShera: saveShera_, setHead: setHead_, makeLetter: makeLetter_, makeLetters: makeLetters_, saveLogo: saveLogo_, dayBPart: dayBPart_, addDay: addDay_, bDays: bDays_, bView: bView_, saveBRow: saveBRow_,
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
    if (s.getMaxColumns() < SHEETS[name].length) s.insertColumnsAfter(s.getMaxColumns(), SHEETS[name].length - s.getMaxColumns());
    s.getRange(1, 1, 1, SHEETS[name].length).setValues([SHEETS[name]]).setFontWeight('bold');
    s.setFrozenRows(1);
    if (name === 'SETTINGS') s.getRange(2, 1, Object.keys(DEF).length, 2).setValues(Object.keys(DEF).map(function (k) { return [k, DEF[k]]; }));
    var d = ss.getSheetByName('Sheet1') || ss.getSheetByName('पत्रक1');
    if (d && ss.getSheets().length > 1 && d.getLastRow() === 0) ss.deleteSheet(d);
  }
  migrate_(s, SHEETS[name]);
  return s;
}
/* जुन्या Sheet मध्ये नवीन स्तंभ नसल्यास शेवटी आपोआप जोडतो */
function migrate_(s, cols) {
  if (s.getMaxColumns() < cols.length) s.insertColumnsAfter(s.getMaxColumns(), cols.length - s.getMaxColumns());
  var n = s.getLastColumn();
  if (n >= cols.length) return;
  var h = n ? s.getRange(1, 1, 1, n).getDisplayValues()[0] : [];
  for (var i = 0; i < h.length; i++) if (h[i] !== cols[i]) return;
  s.getRange(1, h.length + 1, 1, cols.length - h.length).setValues([cols.slice(h.length)]).setFontWeight('bold');
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
  var L = logos_(S).L;
  return { version: VERSION, settings: S, dcrs: dcrs, ps: rows_('PS'), rootUrl: root_().getUrl(), ssUrl: ss_().getUrl(), today: today_(), brand: L ? 'data:' + L.mime + ';base64,' + L.b64 : '' };
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
        ['shera', 'sheraAt', 'letters', 'bRow'].forEach(function (k) { c[k] = old[k]; });
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
      pdfUrl: dR ? dR.pdfUrl : '', cases: (p.cases || []).length, bad: cnt('bad'), warn: cnt('warn'), info: cnt('info'), uploaded: ts, bAdded: dR ? dR.bAdded : '' };
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

function jaKrText_(S, date, seq) {
  return S.jaKrPrefix + '/' + S.jaKrFile + '/ ' + (seq || '______') + ' /' + date.slice(0, 4);
}
/* तीन लोगो {mime,b64}: सेटिंग्जमध्ये अपलोड केलेला असल्यास तो, नाहीतर DCR FORMAT मधील तयार लोगो */
function logos_(S) {
  var o = {};
  ['L', 'C', 'R'].forEach(function (k) {
    var id = S['logo' + k];
    try { if (id) { var bl = DriveApp.getFileById(id).getBlob(); o[k] = { mime: bl.getContentType(), b64: Utilities.base64Encode(bl.getBytes()) }; return; } } catch (e) { /* तयार लोगो वापरा */ }
    o[k] = LOGOS[k];
  });
  return o;
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
    if (rules.some(function (x) { return x.f === r.f; })) return; // सेटिंग्जमध्ये तुम्ही बदललेला / बंद केलेला नियम तसाच राहतो
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
  var logos = logos_(S), letters = built.map(function (x) { return x.L; });
  var mk = function (font, kr, name) { return docxBlob_(DOCXB.build(letters, S, { kr: kr, font: font, logos: logos }), name); };
  var fU = saveBlob_(folder, mk('Mangal', false, base + '.docx'));
  var fK = saveBlob_(folder, mk('Kruti Dev 010', true, base + '_कृतिदेव.docx'));
  var pdfUrl = '', pdfNote = '';
  try { pdfUrl = pdfFromDocx_(mk('Mukta', false, base + '_tmp.docx'), folder, base).getUrl(); } /* PDF साठी Google मध्ये असलेला Mukta फॉन्ट */
  catch (e) { pdfNote = 'PDF बनली नाही: appsscript.json मध्ये Drive सेवा (dependencies) जोडा. ' + e; }
  var outU = { docxUrl: fU.getUrl(), pdfUrl: pdfUrl }, outK = { docxUrl: fK.getUrl() };
  var sL = sh_('LETTERS'), sC = sh_('CASES'), now = now_(), ids = [];
  built.forEach(function (x) {
    var id = Utilities.getUuid().slice(0, 8), r = findCase_(x.key);
    ids.push(id);
    sL.appendRow(toRow_('LETTERS', { id: id, jaKr: x.L.jaKr, date: date, key: x.key, ps: x.c.ps, div: x.c.div, crime: (x.c.kind === 'AD' ? 'मर्ग ' : '') + x.c.no + '/' + x.c.yr,
      head: x.c.head, shera: x.L.shera || '', keys: (x.L.keys || []).join(','), docxUrl: outU.docxUrl, krutiUrl: outK.docxUrl, pdfUrl: outU.pdfUrl, created: now, status: '', replyDate: '', note: '' }));
    sC.getRange(r._row, CASE_COLS.indexOf('shera') + 1, 1, 2).setValues([[x.L.shera || r.shera || '', now]]);
    sC.getRange(r._row, CASE_COLS.indexOf('letters') + 1).setValue((+r.letters || 0) + 1);
  });
  return { docxUrl: outU.docxUrl, krutiUrl: outK.docxUrl, pdfUrl: outU.pdfUrl, pdfNote: pdfNote, count: built.length, jaKrs: built.map(function (x) { return x.L.jaKr; }), ids: ids };
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
function bOv_(c) { try { return JSON.parse(c.bRow || '{}') || {}; } catch (e) { return {}; } }
function bRow_(c, n) {
  var ov = bOv_(c);
  if (ov.cells) return [String(n)].concat(ov.cells.slice(1));
  return bRowAuto_(c, n);
}
function bRowAuto_(c, n) {
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
/* DOCXB च्या फाईल्स → .docx blob; आणि .docx → PDF (Drive सेवेद्वारे Google Doc करून) */
function docxBlob_(files, name) {
  return Utilities.zip(files.map(function (f) {
    return f.text != null ? Utilities.newBlob(f.text, 'text/xml', f.name) : Utilities.newBlob(Utilities.base64Decode(f.b64), 'image/png', f.name);
  })).setContentType('application/vnd.openxmlformats-officedocument.wordprocessingml.document').setName(name);
}
function pdfFromDocx_(blob, folder, base) {
  var tmp = Drive.Files.create({ name: base + '_tmp', mimeType: 'application/vnd.google-apps.document', parents: [folder.getId()] }, blob);
  var pf = saveBlob_(folder, DriveApp.getFileById(tmp.id).getAs('application/pdf').setName(base + '.pdf'));
  DriveApp.getFileById(tmp.id).setTrashed(true);
  return pf;
}
/* हेडवाइज B भाग (महिना): विभागाचे नाव नाही, थेट पो.स्टे.; गुन्हे DCR च्या दिनांकानुसार क्रमाने (रोजचे गुन्हे शेवटी जोडले जातात).
   एकच फाईल "B_भाग_जिल्हा_MM-YYYY" दर वेळी अद्ययावत होते. */
/* हेडवाइज तक्ते (विभागाचे नाव नाही, थेट पो.स्टे.; DCR दिनांकानुसार क्रम) → Word + PDF */
function bpartFile_(cs, titles, folder, base) {
  var HD = DCR.HEADS, groups = {};
  cs = cs.filter(function (c) { return !bOv_(c).skip; });
  cs.forEach(function (c) { var h = c.head || 'OTH5'; (groups[h] = groups[h] || []).push(c); });
  var sum = [['अ.क्र.', 'शीर्षक', 'गुन्हे']], k = 0, counts = [];
  HD.forEach(function (h) { var g = groups[h[0]]; if (g) { sum.push([String(++k), h[1], String(g.length)]); counts.push([h[0], h[1], g.length]); } });
  sum.push(['', 'एकूण', String(cs.length)]);
  var COLS = ['अ.क्र.', 'पोस्टे', 'अप.क्र. व कलम', 'घटनास्थळ', 'घ. ता. वेळ', 'दा. ता. वेळ', 'फिर्यादी व आरोपीचे नाव व पत्ता', 'अटक / प्रतिबंधक', 'तपास अधिकारी', 'घटनेचे थोडक्यात विवरण'];
  var ADC = ['अ.क्र.', 'पोस्टे', 'मर्ग क्र. व कलम', 'घटनास्थळ', 'घ. ता. वेळ', 'दा. ता. वेळ', 'फिर्यादीचे नाव', 'मृतकाचे नाव', 'दाखल / तपास', 'मरणाचे कारण'];
  var W = [440, 1000, 1600, 1360, 1120, 1120, 2400, 1200, 1400, 4238], sections = [];
  k = 0;
  HD.forEach(function (h) {
    var g = groups[h[0]]; if (!g) return;
    g.sort(function (a, z) { var x = String(a.dcrDate || '').localeCompare(String(z.dcrDate || '')); return x || (a.ps < z.ps ? -1 : a.ps > z.ps ? 1 : 0) || (+a.no - +z.no); });
    sections.push({ title: (++k) + ') ' + h[1] + ' - ' + g.length, widths: W, rows: [h[0] === 'AD' ? ADC : COLS].concat(g.map(function (c, i) { return bRow_(c, i + 1); })) });
  });
  var m = { titles: titles, summary: { rows: sum, widths: [800, 6000, 1500] }, sections: sections };
  var fD = saveBlob_(folder, docxBlob_(DOCXB.bpart(m, { font: 'Mangal' }), base + '.docx'));
  var pdfUrl = '';
  try { pdfUrl = pdfFromDocx_(docxBlob_(DOCXB.bpart(m, { font: 'Mukta' }), base + '_tmp.docx'), folder, base).getUrl(); } catch (e) { Logger.log('B भाग PDF: ' + e); }
  return { docxUrl: fD.getUrl(), pdfUrl: pdfUrl, cases: cs.length, counts: counts };
}
function dd_(iso) { return iso.slice(8, 10) + '-' + iso.slice(5, 7) + '-' + iso.slice(0, 4); }
/* एका दिवसाचा (DCR) हेडवाइज B भाग: तपासण्यासाठी */
function dayBPart_(date) {
  var S = settings_(), cs = rows_('CASES').filter(function (r) { return r.kind !== 'NC' && r.dcrDate === date; }).map(caseObj_);
  if (!cs.length) throw new Error(dmy_(date) + ' च्या DCR मधील गुन्हे सापडले नाहीत.');
  var out = bpartFile_(cs, [S.distName, 'दैनिक बी पार्ट (हेडवाइज)', 'DCR दि. ' + dmy_(date)], folder_(['B_भाग', date.slice(0, 7), 'दैनिक']), 'B_भाग_दैनिक_' + dd_(date));
  out.issues = rows_('ISSUES').filter(function (r) { return r.reportDate === date && r.level !== 'info'; }).map(function (r) { return { level: r.level, msg: r.msg }; });
  var d = rows_('DCRS').filter(function (r) { return r.reportDate === date; })[0];
  out.bAdded = d ? d.bAdded : '';
  out.others = cs.filter(function (c) { return c.head === 'OTH5' || c.head === 'OTH6'; }).map(function (c) { return { key: c.key, ps: c.ps, no: c.no, yr: c.yr, kalam: c.kalam, label: c.label, head: c.head }; });
  return out;
}
/* तपासलेला दिवस मासिक B भागात जोडा (on=true) / काढा; नंतर महिन्याची एकच फाईल पुन्हा बनते */
function addDay_(date, on) {
  var d = rows_('DCRS').filter(function (r) { return r.reportDate === date; })[0];
  if (!d) throw new Error(dmy_(date) + ' चा DCR नोंदवहीत नाही.');
  sh_('DCRS').getRange(d._row, SHEETS.DCRS.indexOf('bAdded') + 1).setValue(on ? now_() : '');
  var month = date.slice(0, 7), any = rows_('DCRS').some(function (r) { return String(r.reportDate).slice(0, 7) === month && r.bAdded; });
  return any ? buildBPart_(month, 'ALL') : { cases: 0 };
}
/* महिन्यातील DCR दिवस व त्यांची स्थिती, आणि सध्याची मासिक फाईल */
function bDays_(month) {
  var days = rows_('DCRS').filter(function (r) { return String(r.reportDate).slice(0, 7) === month && +r.cases > 0; })
    .sort(function (a, b) { return a.reportDate < b.reportDate ? -1 : 1; })
    .map(function (r) { return { date: r.reportDate, cases: +r.cases, bad: +r.bad || 0, warn: +r.warn || 0, bAdded: r.bAdded }; });
  var bp = rows_('BPART').filter(function (r) { return r.month === month && r.scope === 'ALL'; })[0];
  return { days: days, monthly: bp ? { docxUrl: bp.docxUrl, pdfUrl: bp.pdfUrl, cases: bp.cases, created: bp.created } : null };
}
/* मासिक हेडवाइज B भाग: फक्त "मासिक B भागात जोडा" केलेल्या दिवसांचे गुन्हे, एकाच फाईलमध्ये ("B_भाग_जिल्हा_MM-YYYY") */
function buildBPart_(month, scope) {
  var S = settings_(), all = !scope || scope === 'ALL';
  var added = {}; rows_('DCRS').forEach(function (r) { if (String(r.reportDate).slice(0, 7) === month && r.bAdded) added[r.reportDate] = 1; });
  var dates = Object.keys(added).sort();
  if (!dates.length) throw new Error(month.slice(5) + '/' + month.slice(0, 4) + ' मध्ये अजून कोणताही दिवस मासिक B भागात जोडलेला नाही. आधी दैनिक B भाग तपासून "मासिक B भागात जोडा" करा.');
  var cs = monthCases_(month, scope).filter(function (c) { return added[c.dcrDate]; });
  if (!cs.length) throw new Error('जोडलेल्या दिवसांचे गुन्हे सापडले नाहीत.');
  var mm = month.slice(5) + '/' + month.slice(0, 4);
  var span = 'DCR दि. ' + dmy_(dates[0]).slice(0, 5) + (dates.length > 1 ? ' ते ' + dmy_(dates[dates.length - 1]).slice(0, 5) : '') + ', ' + dates.length + ' दिवस';
  var base = 'B_भाग_' + (all ? 'जिल्हा' : scope) + '_' + month.slice(5) + '-' + month.slice(0, 4);
  var out = bpartFile_(cs, [all ? S.distName : scope + ' विभाग', 'बी पार्ट (हेडवाइज)', 'माहे ' + mm + ' (' + span + ')'], folder_(['B_भाग', month]), base);
  out.days = dates.length;
  var bs = sh_('BPART'), old = rows_('BPART').filter(function (r) { return r.month === month && r.scope === (scope || 'ALL'); })[0];
  var row = toRow_('BPART', { month: month, scope: scope || 'ALL', cases: out.cases, docxUrl: out.docxUrl, pdfUrl: out.pdfUrl, created: now_() });
  if (old) bs.getRange(old._row, 1, 1, row.length).setValues([row]); else bs.appendRow(row);
  return out;
}

/* B भाग ॲपमध्ये तपासणे/संपादन: q = {date} (दैनिक) किंवा {month} (मासिक: जोडलेले दिवस) */
function bView_(q) {
  var cs;
  if (q.date) cs = rows_('CASES').filter(function (r) { return r.kind !== 'NC' && r.dcrDate === q.date; }).map(caseObj_);
  else {
    var added = {}; rows_('DCRS').forEach(function (r) { if (String(r.reportDate).slice(0, 7) === q.month && r.bAdded) added[r.reportDate] = 1; });
    cs = monthCases_(q.month, 'ALL').filter(function (c) { return added[c.dcrDate]; });
  }
  var HD = DCR.HEADS, groups = {};
  cs.forEach(function (c) { var h = c.head || 'OTH5'; (groups[h] = groups[h] || []).push(c); });
  var secs = [];
  HD.forEach(function (h) {
    var g = groups[h[0]]; if (!g) return;
    g.sort(function (a, z) { var x = String(a.dcrDate || '').localeCompare(String(z.dcrDate || '')); return x || (a.ps < z.ps ? -1 : a.ps > z.ps ? 1 : 0) || (+a.no - +z.no); });
    secs.push({ head: h[0], title: h[1], rows: g.map(function (c, i) {
      var ov = bOv_(c);
      return { key: c.key, dcrDate: c.dcrDate, cells: bRow_(c, i + 1), auto: bRowAuto_(c, i + 1), edited: !!ov.cells, skip: !!ov.skip };
    }) });
  });
  return { secs: secs, total: cs.length };
}
/* एका गुन्ह्याची B भाग ओळ जतन: o = {cells, skip, head}; cells:null → मूळ (DCR मधील) मजकूर */
function saveBRow_(key, o) {
  var r = findCase_(key), s = sh_('CASES'), ov = bOv_(r);
  if ('cells' in o) { if (o.cells) ov.cells = o.cells; else delete ov.cells; }
  if ('skip' in o) { if (o.skip) ov.skip = 1; else delete ov.skip; }
  s.getRange(r._row, CASE_COLS.indexOf('bRow') + 1).setValue(Object.keys(ov).length ? JSON.stringify(ov) : '');
  if (o.head && o.head !== r.head) setHead_(key, o.head);
  return { ok: 1 };
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
