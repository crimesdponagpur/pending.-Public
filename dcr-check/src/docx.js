/* DCR चेक पत्र थेट Word (.docx) म्हणून बनवतो: DCR FORMAT नमुन्याप्रमाणे Letter पान, तीन लोगो असलेले शीर्ष (दुहेरी रेषेसह),
 * जा.क्र. / संदर्भ / विषय / मजकूर / सही / प्रति. Unicode (Mangal) व Kruti Dev 010 दोन्ही. Google Docs वर अवलंबून नाही. */
var DOCXB = (function () {
  var NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"';
  /* नमुन्यातील शीर्ष: ओळ, ठळक?, आकार (अर्ध-पॉइंट: Kruti, Unicode) */
  var OFFICE = [
    ['कार्यालय पोलीस अधीक्षक', 1, 32, 24],
    ['नागपूर जिल्हा (ग्रामीण) सिव्हील लाईन नागपूर- 440001.', 1, 32, 24],
    ['E Mail Id - spreader-ngp.r@mahapolice.gov.in', 1, 28, 22],
    ['दूरध्वनी क्रमांक 0712-2566650, फॅक्स क्र. 0712-2560779', 0, 28, 22]
  ];
  var SZ = { k: 32, u: 26 };               /* मजकूर: Kruti 16pt, Unicode 13pt */
  var IMG = { L: [1126439, 962167], C: [310960, 417575], R: [1168306, 1003240] }; /* EMU (नमुन्याप्रमाणे) */
  function esc(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
  /* मजकूर देवनागरी / इंग्रजी तुकड्यांत; Kruti मध्ये फक्त देवनागरी तुकडे बदलतात, इंग्रजी Arial मध्ये */
  function segs(text) {
    var out = [], re = /[A-Za-z][A-Za-z0-9@._\-\/]*/g, last = 0, m;
    text = String(text || '');
    while ((m = re.exec(text))) { if (m.index > last) out.push([text.slice(last, m.index), 0]); out.push([m[0], 1]); last = re.lastIndex; }
    if (last < text.length) out.push([text.slice(last), 0]);
    return out;
  }
  function runs(text, o) {
    return segs(text).map(function (s) {
      var f = o.kr ? (s[1] ? 'Arial' : 'Kruti Dev 010') : o.font, t = o.kr && !s[1] ? uni2kruti(s[0]) : s[0], sz = o.sz;
      return '<w:r><w:rPr><w:rFonts w:ascii="' + f + '" w:hAnsi="' + f + '" w:eastAsia="' + f + '" w:cs="' + f + '"/>' + (o.b ? '<w:b/><w:bCs/>' : '') +
        '<w:sz w:val="' + sz + '"/><w:szCs w:val="' + sz + '"/><w:lang w:val="en-IN" w:bidi="hi-IN"/></w:rPr><w:t xml:space="preserve">' + esc(t) + '</w:t></w:r>';
    }).join('');
  }
  function para(inner, ppr) { return '<w:p><w:pPr>' + (ppr || '') + '</w:pPr>' + inner + '</w:p>'; }
  var did = 0;
  function image(slot, rId) {
    var cx = IMG[slot][0], cy = IMG[slot][1]; did++;
    return '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="' + cx + '" cy="' + cy + '"/><wp:docPr id="' + did + '" name="Picture ' + did + '"/>' +
      '<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr>' +
      '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="' + did + '" name="img' + did + '"/><pic:cNvPicPr/></pic:nvPicPr>' +
      '<pic:blipFill><a:blip r:embed="' + rId + '"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
      '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
  }
  var CEN = '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/>';
  function cell(w, inner) {
    return '<w:tc><w:tcPr><w:tcW w:w="' + w + '" w:type="dxa"/><w:tcBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="thinThickSmallGap" w:sz="24" w:space="0" w:color="auto"/><w:right w:val="nil"/></w:tcBorders></w:tcPr>' + inner + '</w:tc>';
  }
  function header(o) {
    var mid = (o.logos.C ? para(image('C', 'rIdLogoC'), CEN) : '') + OFFICE.map(function (l) {
      return para(runs(l[0], { kr: o.kr, font: o.font, sz: o.kr ? l[2] : l[3], b: l[1] }), CEN);
    }).join('');
    return '<w:tbl><w:tblPr><w:tblW w:w="9889" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="2042"/><w:gridCol w:w="5505"/><w:gridCol w:w="2342"/></w:tblGrid><w:tr>' +
      cell(2042, para(o.logos.L ? image('L', 'rIdLogoL') : '', CEN)) + cell(5505, mid) + cell(2342, para(o.logos.R ? image('R', 'rIdLogoR') : '', CEN)) + '</w:tr></w:tbl>';
  }
  function letter(L, S, o, first) {
    var sz = o.kr ? SZ.k : SZ.u, R = function (t, b) { return runs(t, { kr: o.kr, font: o.font, sz: sz, b: b }); };
    var body = (L.body || []).filter(function (x) { return x && x.trim(); });
    var IND = '<w:ind w:firstLine="720"/>', SP = function (a) { return '<w:spacing w:after="' + a + '" w:line="300" w:lineRule="auto"/>'; };
    var sig = '(' + (S.spName || '                    ') + ') ' + String(S.spDesig || '').split('\n').join(' ');
    var pr = (L.prati || []).filter(String), pt = (L.prat || []).filter(String);
    var x = para('', (first ? '' : '<w:pageBreakBefore/>') + '<w:spacing w:after="0" w:line="120" w:lineRule="exact"/>');
    x += header(o);
    x += para(R('जा.क्र.:- ' + L.jaKr, 1), '<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr><w:spacing w:before="120" w:after="160"/>');
    x += para(R('संदर्भ:- ', 1) + R(L.ref || '', 0), '<w:spacing w:after="160"/>');
    x += para(R('विषय:- ' + (L.subject || ''), 1), '<w:spacing w:after="280"/>');
    x += para(R(L.intro || '', 0), IND + SP(120) + '<w:jc w:val="both"/>');
    body.forEach(function (t) { x += para(R(t.trim(), 0), IND + SP(120) + '<w:jc w:val="both"/>'); });
    if (L.close) x += para(R(L.close, 1), IND + SP(480));
    x += para(R(sig, 1), '<w:spacing w:after="480"/><w:ind w:left="6480"/><w:jc w:val="center"/>');
    if (pr.length) {
      x += para(R('प्रति,', 1), '<w:keepNext/><w:spacing w:after="120"/>');
      pr.forEach(function (t, i) { x += para(R((pr.length > 1 ? (i + 1) + ') ' : '') + t, 0), '<w:spacing w:after="' + (i === pr.length - 1 ? 240 : 60) + '"/>'); });
    }
    if (pt.length) x += para(R('प्रत कार्यवाहीस, ' + pt.join(' / '), 1), '<w:spacing w:after="0"/>');
    return x;
  }
  /* letters: [{ jaKr, ref, subject, intro, body[], close, prati[], prat[] }]; S: सेटिंग्ज; o: { kr, font, logos: { L: {mime,b64}, C, R } }
     उत्तर: [{ name, text } | { name, b64 }] — .docx ची आतील फाईल्स (zip करायच्या) */
  function build(letters, S, o) {
    did = 0;
    var doc = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS + '><w:body>' +
      letters.map(function (L, i) { return letter(L, S, o, i === 0); }).join('') +
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1152" w:right="1152" w:bottom="1152" w:left="1152" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>';
    var files = [], rels = '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>', exts = {};
    ['L', 'C', 'R'].forEach(function (k) {
      var g = o.logos[k]; if (!g) return;
      var ext = /png/.test(g.mime) ? 'png' : 'jpeg'; exts[ext] = 1;
      files.push({ name: 'word/media/logo' + k + '.' + ext, b64: g.b64 });
      rels += '<Relationship Id="rIdLogo' + k + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/logo' + k + '.' + ext + '"/>';
    });
    var H = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
    files.unshift(
      { name: '[Content_Types].xml', text: H + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
        Object.keys(exts).map(function (e) { return '<Default Extension="' + e + '" ContentType="image/' + e + '"/>'; }).join('') +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>' },
      { name: '_rels/.rels', text: H + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>' },
      { name: 'word/document.xml', text: doc },
      { name: 'word/_rels/document.xml.rels', text: H + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + rels + '</Relationships>' },
      { name: 'word/styles.xml', text: H + '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="' + o.font + '" w:hAnsi="' + o.font + '" w:cs="' + o.font + '"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-IN" w:bidi="hi-IN"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>' });
    return files;
  }
  return { build: build };
})();
if (typeof module !== 'undefined') module.exports = DOCXB;
