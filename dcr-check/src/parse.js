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
  function summary(c) {
    var f = clean(c.facts || '').replace(/^(थोडक्यात\s*)?(हकी[कग]त|हकिकत)\s*(अशा|अश्या)\s*(प्रकारे|प्रकारची)\s*आहे\s*की[,\s]*/, '').replace(/^यातील\s*/, 'यातील ');
    f = f.split(' / ')[0];
    f = f.replace(/[,.\s]*(अशा|अश्या)?\s*(फिर्यादी|फि\.|फिर्यादिच्या|फिर्यादीच्या|फिर्यादीचे)[^.]{0,60}(रिपोर्ट|तक्रार)[^.]*(नोंद|दाखल)[^.]*\.?\s*$/, '');
    f = f.replace(/[,.\s]*सदर(चा)?\s*गुन्हा\s*(नोंद|दाखल)[^.]*\.?\s*$/, '');
    if (f.length > 500) { var cut = f.slice(0, 500), i = cut.lastIndexOf('.'); if (i < 250) i = cut.lastIndexOf(' '); f = cut.slice(0, i > 250 ? i : 500); }
    return f.replace(/[.\s]+$/, '');
  }
  return {
    readDocx: readDocx, blocksFromXml: blocksFromXml, extract: extract, classify: classify, parsePsCr: parsePsCr,
    HEADS: HEADS, HEAD_MR: HEAD_MR, SEC_MR: SEC_MR, dateOf: dateOf, dmy: dmy, asc: asc, summary: summary, normDiv: normDiv
  };
})();
if (typeof module !== 'undefined') module.exports = DCR;
