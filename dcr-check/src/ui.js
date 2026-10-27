/* DCR चेक ॲप: स्क्रीन */
(function () {
  var S = { settings: {}, dcrs: [], ps: [], today: '', cases: [], tab: 'dcr', picked: {} };
  var $ = function (s, r) { return (r || document).querySelector(s); };
  var esc = function (s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  var dmy = DCR.dmy;
  var SEC_ORDER = ['IPC', 'MVT', 'HBT', 'SLL', 'AD', 'COMM', 'GAMB', 'PROH'];
  var RELS = ['एकमेकांच्या ओळखीचे', 'पती-पत्नी', 'नातेवाईक', 'शेजारी', 'वडील व मुलगा', 'सख्खे भाऊ'];

  function call(fn) {
    var args = [].slice.call(arguments, 1);
    return new Promise(function (res, rej) {
      if (!(window.google && google.script && google.script.run)) return window.MOCK ? MOCK(fn, args).then(res, rej) : rej(new Error('हे पान Apps Script मधून उघडा.'));
      google.script.run.withSuccessHandler(function (s) { var j = JSON.parse(s); j.ok ? res(j.d) : rej(new Error(j.e)); })
        .withFailureHandler(function (e) { rej(new Error(e && e.message || 'सर्व्हरशी संपर्क झाला नाही')); }).api(fn, args);
    });
  }
  function toast(m, bad) {
    document.querySelectorAll('.toast').forEach(function (x) { x.remove(); });
    var t = document.createElement('div'); t.className = 'toast' + (bad ? ' bad' : ''); t.textContent = m; document.body.appendChild(t);
    setTimeout(function () { t.remove(); }, bad ? 6000 : 3000);
  }
  function busy(btn, on, txt) { if (!btn) return; btn.disabled = on; if (on) { btn._t = btn.innerHTML; btn.innerHTML = '<span class="spin"></span> ' + (txt || 'थांबा...'); } else if (btn._t) btn.innerHTML = btn._t; }
  function b64(buf) { var u = new Uint8Array(buf), s = '', CH = 0x8000; for (var i = 0; i < u.length; i += CH) s += String.fromCharCode.apply(null, u.subarray(i, i + CH)); return btoa(s); }
  function headSel(v, cls) { return '<select class="' + (cls || '') + '">' + DCR.HEADS.map(function (h) { return '<option value="' + h[0] + '"' + (h[0] === v ? ' selected' : '') + '>' + esc(h[1]) + '</option>'; }).join('') + '</select>'; }
  function crTxt(c) { return (c.kind === 'AD' ? 'मर्ग ' : c.kind === 'NC' ? 'NC ' : '') + c.no + '/' + String(c.yr).slice(2); }
  function divs() { return String(S.settings.divOrder || '').split(',').filter(String); }
  function paras() { try { var p = JSON.parse(S.settings.paras || ''); if (p && p.length) return p; } catch (e) {} return LETTER.PARAS; }
  function curMonth() { return (S.today || new Date().toISOString()).slice(0, 7); }

  /* ---------- सुरुवात ---------- */
  function boot() {
    document.querySelectorAll('nav button').forEach(function (b) { b.onclick = function () { show(b.dataset.t); }; });
    call('init').then(function (d) {
      S.settings = d.settings; S.dcrs = d.dcrs; S.ps = d.ps; S.today = d.today;
      $('#lnkRoot').href = d.rootUrl; $('#lnkSs').href = d.ssUrl; $('#ver').textContent = 'आवृत्ती ' + d.version; if (d.brand) $('#brand').src = d.brand;
      imeBar(); show('dcr');
    }).catch(function (e) { $('#main').innerHTML = '<div class="card"><h2>ॲप सुरू झाले नाही</h2><p>' + esc(e.message) + '</p></div>'; });
  }
  function show(t) {
    S.tab = t;
    document.querySelectorAll('nav button').forEach(function (b) { b.classList.toggle('on', b.dataset.t === t); });
    var P = { dcr: ['DCR अपलोड', 'रोजचा DCR (Word) तपासा व नोंदवहीत जतन करा'], cases: ['गुन्हे व शेरा', 'SP शेरा, DCR चेक पत्रे व एकत्र पत्रे'], letters: ['DCR चेक नोंदवही', 'दिलेली पत्रे, अनुपालन अहवाल व मुदत'],
      bpart: ['हेडवाइज B भाग', 'दैनिक तपासणी, संपादन व मासिक B भाग'], set: ['सेटिंग्ज', 'पत्राचे तपशील, शेरे, लोगो व पो.स्टे. यादी'] }[t];
    $('#ptitle').textContent = P[0]; $('#psub').textContent = P[1]; window.scrollTo(0, 0);
    ({ dcr: vDcr, cases: vCases, letters: vLetters, bpart: vBpart, set: vSet })[t]();
  }

  /* ---------- 1. DCR अपलोड ---------- */
  /* महिन्याचा आढावा: वरच्या चार चौकटी */
  function stats() {
    var m = curMonth(), ds = S.dcrs.filter(function (r) { return String(r.reportDate).slice(0, 7) === m && +r.cases > 0; });
    var n = 0, bad = 0, add = 0; ds.forEach(function (r) { n += +r.cases || 0; if (+r.bad) bad++; if (r.bAdded) add++; });
    var last = ds.map(function (r) { return r.reportDate; }).sort().pop();
    var T = [['या महिन्याचे DCR', ds.length, last ? 'शेवटचा: ' + dmy(last) : 'अजून नाही'], ['गुन्हे (महिना)', n, 'NC सोडून सर्व नोंदी'],
      ['त्रुटी असलेले DCR', bad, bad ? 'तपासून PS ला कळवा' : 'सर्व ठीक', bad ? 'bad' : 'ok'], ['मासिक B भागात', add + '/' + ds.length, 'दिवस जोडले', add < ds.length ? 'warn' : 'ok']];
    return '<div class="kpis">' + T.map(function (t) { return '<div class="kpi ' + (t[3] || '') + '"><span>' + t[0] + '</span><b>' + t[1] + '</b><small>' + esc(t[2]) + '</small></div>'; }).join('') + '</div>';
  }
  function vDcr() {
    $('#main').innerHTML = stats() +
      '<div class="card"><h2>रोजचा DCR टाका</h2><p class="muted">Word (.docx) फाइल द्या: ॲप Kruti Dev मजकूर Unicode मध्ये बदलून गुन्हे नोंदवहीत घेईल, त्रुटी सांगेल आणि मूळ फाइल Drive मध्ये ठेवेल. PDF फक्त Drive मध्ये जपली जाते (वाचली जात नाही). एकाच वेळी अनेक तारखांच्या फाइल्स देता येतात.</p>' +
      '<label class="drop" id="drop">येथे फाइल्स ओढा किंवा <b>निवडण्यासाठी दाबा</b><input type="file" id="fi" accept=".docx,.pdf" multiple hidden></label><div id="res"></div></div>' +
      '<div class="card"><h2>आलेले DCR</h2><div id="dl"></div></div>';
    var dr = $('#drop'), fi = $('#fi');
    fi.onchange = function () { take(fi.files); fi.value = ''; };
    dr.ondragover = function (e) { e.preventDefault(); dr.classList.add('over'); };
    dr.ondragleave = function () { dr.classList.remove('over'); };
    dr.ondrop = function (e) { e.preventDefault(); dr.classList.remove('over'); take(e.dataTransfer.files); };
    dcrList();
  }
  function dcrList() {
    var el = $('#dl'); if (!el) return;
    if (!S.dcrs.length) { el.innerHTML = '<p class="muted">अजून एकही DCR नाही.</p>'; return; }
    el.innerHTML = '<div class="scroll"><table class="t"><tr><th>अहवाल दि.</th><th>गुन्हे</th><th>त्रुटी</th><th>फाइल</th><th></th></tr>' + S.dcrs.map(function (r) {
      return '<tr><td>' + dmy(r.reportDate) + '</td><td>' + esc(r.cases) + '</td><td>' + (+r.bad ? '<span class="chip bad">' + r.bad + ' त्रुटी</span>' : '') + (+r.warn ? '<span class="chip warn">' + r.warn + ' कमी</span>' : '') +
        '</td><td class="links">' + (r.fileUrl ? '<a target="_blank" href="' + esc(r.fileUrl) + '">Word</a>' : '') + (r.pdfUrl ? '<a target="_blank" href="' + esc(r.pdfUrl) + '">PDF</a>' : '') +
        '</td><td><button class="btn sm ghost" data-d="' + r.reportDate + '">शेरा →</button></td></tr>';
    }).join('') + '</table></div>';
    el.querySelectorAll('button[data-d]').forEach(function (b) { b.onclick = function () { S.pickDate = b.dataset.d; show('cases'); }; });
  }
  function take(files) {
    files = [].slice.call(files || []);
    var chain = Promise.resolve();
    files.forEach(function (f) { chain = chain.then(function () { return one(f); }); });
  }
  function one(f) {
    var box = document.createElement('div'); box.className = 'case'; box.innerHTML = '<b>' + esc(f.name) + '</b> <span class="spin"></span> वाचत आहे...';
    $('#res').prepend(box);
    return f.arrayBuffer().then(function (buf) {
      if (/\.pdf$/i.test(f.name)) {
        var d = DCR.dateOf(f.name.replace(/_/g, ' '));
        return call('savePdf', { fileName: f.name, b64: b64(buf), reportDate: d }).then(function (r) {
          box.innerHTML = '<b>' + esc(f.name) + '</b> <span class="chip ok">Drive मध्ये जतन</span> <a target="_blank" href="' + esc(r.url) + '">उघडा</a><div class="muted">PDF मधून गुन्हे वाचता येत नाहीत; त्याच तारखेची Word फाइलही द्या.</div>';
          refreshDcrs();
        });
      }
      if (!/\.docx$/i.test(f.name)) throw new Error('फक्त .docx किंवा .pdf फाइल चालते. जुनी .doc फाइल Word मध्ये उघडून .docx म्हणून जतन करा.');
      return DCR.readDocx(buf, JSZip, DOMParser).then(function (blocks) {
        var R = DCR.extract(blocks, f.name);
        if (!R.reportDate) {
          var d = prompt('या DCR ची तारीख (दिवस/महिना/वर्ष) लिहा:', '');
          R.reportDate = DCR.dateOf(d || '');
          if (!R.reportDate) throw new Error('तारीख नसल्याने जतन केले नाही.');
        }
        box.innerHTML = '<b>' + esc(f.name) + '</b> <span class="spin"></span> ' + R.cases.length + ' गुन्हे सापडले, जतन करत आहे...';
        var saveName = 'CRO_DCR_' + dmy(R.reportDate).replace(/\//g, '-') + '.docx';
        return call('saveDcr', { fileName: f.name, saveName: saveName, b64: b64(buf), reportDate: R.reportDate, issueDate: R.issueDate, jaKr: R.jaKr, cases: R.cases, issues: R.issues })
          .then(function (r) { box.outerHTML = dcrCard(R, r); bindCard(R.reportDate); refreshDcrs(); });
      });
    }).catch(function (e) { box.innerHTML = '<b>' + esc(f.name) + '</b> <span class="chip bad">अडचण</span> ' + esc(e.message); });
  }
  function dcrCard(R, r) {
    var lv = { bad: [], warn: [], info: [] };
    R.issues.forEach(function (i) { (lv[i.level] || lv.info).push(i.msg); });
    var cnt = SEC_ORDER.filter(function (s) { return R.counts[s]; }).map(function (s) { return '<span class="chip">' + DCR.SEC_MR[s] + ' ' + R.counts[s] + '</span>'; }).join('');
    var li = function (a, c) { return a.map(function (m) { return '<li class="' + c + '">' + esc(m) + '</li>'; }).join(''); };
    return '<div class="case" data-card="' + R.reportDate + '"><div class="row"><b class="grow">DCR दि. ' + dmy(R.reportDate) + (R.jaKr ? ' • जा.क्र. ' + esc(R.jaKr) : '') + '</b>' +
      '<span class="chip ok">' + r.added + ' नवे, ' + r.updated + ' आधीचे</span><a target="_blank" href="' + esc(r.fileUrl) + '">Drive</a></div><div>' + cnt + '</div>' +
      (lv.bad.length || lv.warn.length ? '<ul class="iss">' + li(lv.bad, 'bad') + li(lv.warn, 'warn') + '</ul>' : '<p class="muted">शीर्षकातील संख्या व ओळी जुळल्या. मोठी त्रुटी नाही.</p>') +
      (lv.info.length ? '<details><summary>माहितीसाठी (' + lv.info.length + ')</summary><ul class="iss">' + li(lv.info, 'info') + '</ul></details>' : '') +
      '<div class="row links" style="margin-top:8px"><button class="btn sm" data-go="' + R.reportDate + '">या DCR वर शेरा / पत्र →</button>' +
      '<button class="btn sm ghost" data-bp="' + R.reportDate + '">दैनिक B भाग तपासा →</button></div></div>';
  }
  function bindCard(d) { var bp = document.querySelector('[data-bp="' + d + '"]'); if (bp) bp.onclick = function () { S.pickDay = d; show('bpart'); }; var b = document.querySelector('[data-go="' + d + '"]'); if (b) b.onclick = function () { S.pickDate = d; show('cases'); }; }
  function refreshDcrs() { call('init').then(function (d) { S.dcrs = d.dcrs; S.ps = d.ps; dcrList(); var k = $('.kpis'); if (k) k.outerHTML = stats(); }); }

  /* ---------- 2. गुन्हे व शेरा ---------- */
  function vCases() {
    var dates = S.dcrs.filter(function (r) { return +r.cases > 0; }).map(function (r) { return r.reportDate; });
    var d0 = S.pickDate || dates[0] || ''; S.pickDate = null;
    $('#main').innerHTML = '<div class="card"><div class="row">' +
      '<select id="qd"><option value="">सर्व तारखा</option>' + dates.map(function (d) { return '<option value="' + d + '"' + (d === d0 ? ' selected' : '') + '>DCR ' + dmy(d) + '</option>'; }).join('') + '</select>' +
      '<input type="text" id="qt" class="grow" placeholder="शोधा: पो.स्टे., अप.क्र. (1048/26), नाव, कलम...">' +
      '<select id="qh"><option value="">सर्व शीर्षके</option>' + DCR.HEADS.map(function (h) { return '<option value="' + h[0] + '">' + esc(h[1]) + '</option>'; }).join('') + '</select>' +
      '<label class="muted"><input type="checkbox" id="qs"> फक्त शेरा असलेले</label><button class="btn" id="qb">शोधा</button><button class="btn ghost" id="qp" title="SP साहेबांसमोर ठेवण्यासाठी">शेरा पत्रक छापा</button></div></div><div id="bulk"></div><div id="cl"></div>';
    $('#qb').onclick = find; $('#qt').onkeydown = function (e) { if (e.key === 'Enter') find(); };
    $('#qd').onchange = find; $('#qh').onchange = find; $('#qs').onchange = find;
    $('#qp').onclick = printSheet;
    find();
  }
  function find() {
    var q = { date: $('#qd').value, text: $('#qt').value.trim(), head: $('#qh').value, shera: $('#qs').checked ? 'yes' : '' };
    if (q.text) q.date = '';
    $('#cl').innerHTML = '<p class="muted"><span class="spin"></span> शोधत आहे...</p>';
    call('search', q).then(function (r) {
      S.cases = r.cases;
      if (!r.cases.length) { $('#cl').innerHTML = '<div class="card muted">काही सापडले नाही.</div>'; return; }
      var list = r.cases.slice();
      if (q.date && !q.text) list.sort(function (a, b) { return (+a.dcrSeq && +b.dcrSeq) ? a.dcrSeq - b.dcrSeq : SEC_ORDER.indexOf(a.sec) - SEC_ORDER.indexOf(b.sec) || divs().indexOf(a.div) - divs().indexOf(b.div); });
      $('#cl').innerHTML = '<p class="muted">' + r.total + ' गुन्हे' + (r.more ? ' (पहिले 400 दाखवले)' : '') + '</p>' + list.map(caseHtml).join('');
      list.forEach(bindCase); updateBulk();
    }).catch(function (e) { $('#cl').innerHTML = '<div class="card"><ul class="iss"><li class="bad">' + esc(e.message) + '</li></ul></div>'; });
  }
  /* शीर्षकनिहाय सूचना (सेटिंग्ज): { HEAD: { p: [जुन्या पत्रांतील वाक्ये], q: ['स्वतःची सूचना', ...] } }.
     सेटिंग्जमध्ये जतन केले नसल्यास जुन्या DCR चेक पत्रांवरून (BY_HEAD) व आधीच्या शीर्षकनिहाय शेऱ्यांवरून भरलेले. */
  function hsAll() {
    var o = null; try { o = JSON.parse(S.settings.headSuchna || 'null'); } catch (e) {}
    if (o) return o;
    var qh = {}; try { qh = JSON.parse(S.settings.quickHead || '{}') || {}; } catch (e) {}
    o = {};
    DCR.HEADS.forEach(function (h) { o[h[0]] = { p: (LETTER.BY_HEAD[h[0]] || []).slice(), q: String(qh[h[0]] || '').split('\n').map(function (x) { return x.trim(); }).filter(String) }; });
    return o;
  }
  function hsOf(head) { var o = hsAll()[head] || {}; return { p: o.p || [], q: o.q || [] }; }
  /* गुन्ह्याखालील यादी: शेरा + त्या शीर्षकाच्या सूचना (जुन्या पत्रांतील व स्वतःच्या) */
  function caseItems(c) { var h = hsOf(c.head); return ['shera'].concat(h.p).concat(h.q.map(function (t) { return 'u:' + t; })); }
  /* फक्त ✓ केलेल्या सूचना पत्रात जातात (ॲप स्वतःहून काही टिकत नाही) */
  function caseKeys(c) {
    var s = String(c.sKeys || '');
    if (!s) return [];
    if (s.charAt(0) === '[') { try { return JSON.parse(s) || []; } catch (e) { return []; } }
    return s.split(',').filter(function (k) { return k && k !== '-'; });
  }
  function quickChips(head) {
    return String(S.settings.quick || '').split('\n').map(function (x) { return x.trim(); }).filter(String).map(function (q) { return '<span class="chip q">' + esc(q) + '</span>'; }).join('') +
      '<span class="chip qadd" title="लिहिलेला शेरा या शीर्षकाच्या सूचनांत कायमचा जोडा">+ हा शेरा "' + esc(DCR.HEAD_MR[head] || 'इतर') + '" च्या सूचनांत जोडा</span>';
  }
  function caseHtml(c) {
    var id = 'c' + btoa(unescape(encodeURIComponent(c.key))).replace(/[^a-z0-9]/gi, '');
    c._id = id;
    return '<div class="case' + (c.shera ? ' has' : '') + '" id="' + id + '"><div class="row"><label class="pickl" title="एकत्र पत्रांसाठी निवडा"><input type="checkbox" class="pk1"' + (S.picked[c.key] ? ' checked' : '') + '></label><span class="t grow">' + (c.dcrSeq ? '<span class="seq">' + c.dcrSeq + '</span>' : '') + esc(c.ps) + ' ' + crTxt(c) + (c.kalam ? ' • ' + esc(c.kalam) : '') + '</span>' +
      '<span class="chip">' + esc(DCR.HEAD_MR[c.head] || DCR.SEC_MR[c.sec] || '') + '</span>' + (+c.letters ? '<span class="chip ok">पत्र ' + c.letters + '</span>' : '') + '</div>' +
      '<div class="muted">' + (c.div ? esc(c.div) + ' विभाग • ' : '') + 'दाखल ' + dmy(c.regDate) + (c.label ? ' • ' + esc(c.label) : '') + (c.io ? ' • तपास: ' + esc(c.io.slice(0, 60)) : '') + '</div>' +
      '<div class="facts">' + esc(c.facts || ('फिर्यादी: ' + c.comp + ' • आरोपी: ' + c.acc)) + '</div>' +
      '<textarea placeholder="SP साहेबांचा शेरा येथे लिहा...">' + esc(c.shera) + '</textarea><div class="qk">' + quickChips(c.head) + '</div>' +
      '<div class="sk"><div class="skh"><b>पत्रातील सूचना</b> <span class="muted">(फक्त ✓ केलेल्या पत्रात जातील)</span> <span class="chip skn"></span></div><div class="skl"></div>' +
      '<details class="sko"><summary>इतर सूचना (जुन्या पत्रांतील सर्व)</summary><div class="skl2"></div></details></div>' +
      '<div class="row" style="margin-top:8px"><button class="btn sm ghost" data-a="save">शेरा जतन</button><button class="btn sm gold" data-a="letter">DCR चेक पत्र बनवा</button>' +
      '<span class="grow"></span><span class="muted">शीर्षक:</span>' + headSel(c.head, 'hs') + '</div></div>';
  }
  function bindCase(c) {
    var el = document.getElementById(c._id); if (!el) return;
    var ta = $('textarea', el);
    function item(k, ks, H) {
      var p = LETTER.para(k, paras()); if (!p) return '';
      var on = ks.indexOf(k) >= 0;
      return '<label class="pk' + (on ? ' on' : '') + '"><input type="checkbox" value="' + esc(k) + '"' + (on ? ' checked' : '') + '> <span>' + esc(p.t) +
        (/^u:/.test(k) ? '' : '') + (H.indexOf(k) >= 0 ? ' <span class="chip info">शेऱ्यानुसार सुचवलेले</span>' : '') + '</span></label>';
    }
    function skDraw() {
      var P = paras(), ks = caseKeys(c), H = LETTER.hints(ta.value, P), main = caseItems(c);
      H.concat(ks).forEach(function (k) { if (main.indexOf(k) < 0) main.push(k); });
      $('.skn', el).textContent = '✓ ' + ks.length;
      $('.skl', el).innerHTML = main.map(function (k) { return item(k, ks, H); }).join('');
      $('.skl2', el).innerHTML = P.filter(function (p) { return main.indexOf(p.k) < 0; }).map(function (p) { return item(p.k, ks, H); }).join('');
      el.querySelectorAll('.sk input').forEach(function (cb) {
        cb.onchange = function () {
          var sel = []; el.querySelectorAll('.sk input:checked').forEach(function (x) { if (sel.indexOf(x.value) < 0) sel.push(x.value); });
          if (cb.value === 'shera') c._shTouched = true;
          c.sKeys = JSON.stringify(sel); skDraw(); skSave(sel);
        };
      });
    }
    function skSave(sel) { call('saveShera', c.key, null, sel).catch(function (e) { toast(e.message, 1); }); }
    /* शेरा लिहिला की "SP शेरा" ✓ होतो (तुम्ही लिहिलेलाच मजकूर; नको असल्यास ✓ काढा) */
    function shTick() {
      var ks = caseKeys(c);
      if (ta.value.trim() && !c._shTouched && ks.indexOf('shera') < 0) { ks.unshift('shera'); c.sKeys = JSON.stringify(ks); skSave(ks); }
      skDraw();
    }
    skDraw();
    ta.addEventListener('input', function () { c.shera = ta.value.trim(); shTick(); });
    $('.facts', el).onclick = function () { this.classList.toggle('open'); };
    $('.pk1', el).onchange = function () { if (this.checked) S.picked[c.key] = c; else delete S.picked[c.key]; updateBulk(); };
    ta.onfocus = function () { el.classList.add('qo'); };
    function bindQ() {
      el.querySelectorAll('.chip.q').forEach(function (q) { q.onclick = function () { ta.value = (ta.value.trim() ? ta.value.trim().replace(/[.।]?$/, '. ') : '') + q.textContent; c.shera = ta.value.trim(); shTick(); }; });
      $('.chip.qadd', el).onclick = function () {
        var t = ta.value.trim().replace(/[.।]$/, ''); if (!t) return toast('आधी शेरा लिहा', 1);
        var o = hsAll(), h = o[c.head] = o[c.head] || { p: [], q: [] };
        h.q = h.q || []; if (h.q.indexOf(t) < 0) h.q.push(t);
        call('saveSettings', { headSuchna: JSON.stringify(o) }).then(function (d) { S.settings = d; skDraw(); toast('"' + (DCR.HEAD_MR[c.head] || '') + '" च्या सूचनांत जोडला. पुढे अशा सर्व गुन्ह्यांखाली दिसेल.'); }).catch(function (e) { toast(e.message, 1); });
      };
    }
    bindQ();
    $('[data-a=save]', el).onclick = function () {
      var b = this; busy(b, true);
      call('saveShera', c.key, ta.value.trim(), caseKeys(c)).then(function () { c.shera = ta.value.trim(); el.classList.toggle('has', !!c.shera); toast('शेरा व सूचना जतन झाल्या'); })
        .catch(function (e) { toast(e.message, 1); }).then(function () { busy(b, false); });
    };
    $('[data-a=letter]', el).onclick = function () { c.shera = ta.value.trim(); letterEditor(c, el); };
    $('.hs', el).onchange = function () {
      var v = this.value; call('setHead', c.key, v).then(function () { c.head = v; $('.qk', el).innerHTML = quickChips(v); bindQ(); skDraw(); toast('शीर्षक बदलले'); }).catch(function (e) { toast(e.message, 1); });
    };
  }
  function printSheet() {
    var cs = S.cases.filter(function (c) { return c.kind !== 'NC'; });
    if (!cs.length) return toast('आधी गुन्हे शोधा', 1);
    var d = $('#qd').value;
    $('#print').innerHTML = '<h2>' + esc(S.settings.distName || '') + ' • DCR ' + (d ? dmy(d) : '') + ' • शेरा पत्रक</h2><table><tr><th>क्र.</th><th>पो.स्टे. / अप.क्र. / कलम</th><th>थोडक्यात</th><th style="width:34%">मा. पोलीस अधीक्षक यांचा शेरा</th></tr>' +
      cs.map(function (c, i) { return '<tr><td>' + (i + 1) + '</td><td><b>' + esc(c.ps) + ' ' + crTxt(c) + '</b><br>' + esc(c.kalam) + '<br><i>' + esc(DCR.HEAD_MR[c.head] || '') + '</i></td><td>' + esc(DCR.summary(c).slice(0, 260)) + '</td><td>' + esc(c.shera) + '</td></tr>'; }).join('') + '</table>';
    window.print();
  }

  /* ---------- पत्र संपादक ---------- */
  function letterEditor(c, el) {
    var P = paras(), opt = { shera: c.shera, paras: P, settings: { days: S.settings.days }, relation: RELS[0], place: 'गावात', refDate: c.dcrIssue || '', rules: rules() };
    opt.keys = caseKeys(c);
    caseItems(c).concat(opt.keys).forEach(function (k) { if (/^u:/.test(k) && !P.some(function (p) { return p.k === k; })) P.push(LETTER.para(k)); });
    var D = LETTER.draft(c, opt), on = {};
    delete opt.keys;
    D.keys.forEach(function (k) { on[k] = 1; });
    var texts = {};
    function fillOne(k) { var o = {}; for (var x in opt) o[x] = opt[x]; o.keys = [k]; return LETTER.draft(c, o).body[0] || ''; }
    P.forEach(function (p) { texts[p.k] = fillOne(p.k); });
    var m = document.createElement('div'); m.className = 'modal';
    var H = LETTER.hints(c.shera, P).filter(function (k) { return D.keys.indexOf(k) < 0; });
    var order = D.keys.concat(H).concat(P.map(function (p) { return p.k; }).filter(function (k) { return D.keys.indexOf(k) < 0 && H.indexOf(k) < 0; }));
    m.innerHTML = '<div class="sheet"><div class="row"><h2 class="grow" style="margin:0">DCR चेक पत्र: ' + esc(c.ps) + ' ' + crTxt(c) + '</h2><button class="btn ghost sm" data-x>बंद</button></div>' +
      '<div class="row"><div class="grow"><label>जा.क्र. (रिकामे ठेवल्यास आपोआप)</label><input type="text" id="lj" placeholder="' + esc(S.settings.jaKrPrefix) + '/...">' + '</div><div><label>पत्राचा दिनांक</label><input type="text" id="ldt" value="' + dmy(S.today) + '" style="width:130px"></div></div>' +
      '<label>संदर्भ</label><input type="text" id="lr" value="' + esc(D.ref) + '"><label>विषय</label><input type="text" id="ls" value="' + esc(D.subject) + '">' +
      '<label>पहिला परिच्छेद (DCR मधील घटना)</label><textarea id="li" style="min-height:120px">' + esc(D.intro) + '</textarea>' +
      '<label>SP शेरा</label><textarea id="lsh" style="min-height:44px">' + esc(c.shera) + '</textarea>' +
      '<div class="row"><span class="muted">नाते:</span><select id="lrel">' + RELS.map(function (r) { return '<option>' + r + '</option>'; }).join('') + '</select><select id="lpl"><option>गावात</option><option>घरात</option><option>परिसरात</option></select></div>' +
      '<label>सूचना (जुन्या पत्रांप्रमाणे; ✓ केलेले पत्रात जातील, मजकूर बदलता येतो)</label><div id="lp">' +
      order.map(function (k) { var p = LETTER.para(k, P); return '<div class="pk' + (on[k] ? ' on' : '') + '" data-k="' + k + '"><label><input type="checkbox"' + (on[k] ? ' checked' : '') + '> ' + esc(p.t) + (H.indexOf(k) >= 0 ? ' <span class="chip info">शेऱ्यानुसार सुचवलेले</span>' : '') + '</label>' + (on[k] ? '<textarea>' + esc(texts[k]) + '</textarea>' : '') + '</div>'; }).join('') + '</div>' +
      '<label>शेवटचे वाक्य</label><input type="text" id="lc" value="' + esc(D.close) + '">' +
      '<div class="row"><div class="grow"><label>प्रति (एका ओळीत एक)</label><textarea id="lpr">' + esc(D.prati.join('\n')) + '</textarea></div><div class="grow"><label>प्रत माहितीस</label><textarea id="lpt">' + esc(D.prat.join('\n')) + '</textarea></div></div>' +
      '<div class="row" style="margin-top:12px"><button class="btn gold" id="lgo">पत्र बनवा (Word: Unicode व कृतिदेव, PDF) व Drive मध्ये जतन</button><span id="lout" class="links"></span></div></div>';
    document.body.appendChild(m);
    $('[data-x]', m).onclick = function () { m.remove(); };
    function regen(k) { var box = $('[data-k="' + k + '"]', m); if (box && box.classList.contains('on')) $('textarea', box).value = texts[k] = fillOne(k); }
    $('#lsh', m).oninput = function () { opt.shera = this.value; regen('shera'); };
    $('#lrel', m).onchange = function () { opt.relation = this.value; regen('known'); };
    $('#lpl', m).onchange = function () { opt.place = this.value; regen('known'); };
    m.querySelectorAll('.pk').forEach(function (box) {
      var k = box.dataset.k, cb = $('input', box);
      cb.onchange = function () {
        box.classList.toggle('on', cb.checked);
        var ta = $('textarea', box);
        if (cb.checked && !ta) { ta = document.createElement('textarea'); ta.value = texts[k] = fillOne(k); box.appendChild(ta); }
        if (!cb.checked && ta) ta.remove();
      };
    });
    $('#lgo', m).onclick = function () {
      var b = this, keys = [], body = [];
      m.querySelectorAll('.pk.on').forEach(function (box) { keys.push(box.dataset.k); body.push($('textarea', box).value.trim()); });
      var dt = $('#ldt', m).value.trim(), iso = DCR.dateOf(dt);
      var L = { jaKr: $('#lj', m).value.trim(), date: iso || S.today, dateText: dt, ref: $('#lr', m).value.trim(), subject: $('#ls', m).value.trim(), intro: $('#li', m).value.trim(),
        body: body, close: $('#lc', m).value.trim(), prati: $('#lpr', m).value.split('\n').map(function (x) { return x.trim(); }).filter(String),
        prat: $('#lpt', m).value.split('\n').map(function (x) { return x.trim(); }).filter(String), keys: keys, shera: $('#lsh', m).value.trim() };
      busy(b, true, 'पत्र बनत आहे...');
      call('makeLetters', [{ key: c.key, L: L, orig: D }]).then(function (r) {
        $('#lout', m).innerHTML = '<span class="chip ok">जा.क्र. ' + esc(r.jaKrs[0]) + '</span>' + links(r);
        c.letters = (+c.letters || 0) + 1; c.shera = L.shera;
        if (el) { var t = $('textarea', el); if (t) t.value = L.shera; el.classList.toggle('has', !!L.shera); }
        toast('पत्र Drive मध्ये जतन झाले');
      }).catch(function (e) { toast(e.message, 1); }).then(function () { busy(b, false); });
    };
  }
  function links(r) {
    return '<a target="_blank" href="' + esc(r.docxUrl) + '">Word (Unicode)</a><a target="_blank" href="' + esc(r.krutiUrl) + '">Word (कृतिदेव)</a>' + (r.pdfUrl ? '<a target="_blank" href="' + esc(r.pdfUrl) + '">PDF</a>' : '<span class="chip bad" title="' + esc(r.pdfNote || '') + '">PDF बनली नाही (appsscript.json तपासा)</span>');
  }
  function rules() { try { return JSON.parse(S.settings.rules || '[]') || []; } catch (e) { return []; } }
  /* निवडलेल्या गुन्ह्यांसाठी एकच शेरा व एकत्र सूचना; सर्व पत्रे एका Word फाईलमध्ये */
  function updateBulk() {
    var n = Object.keys(S.picked).length, el = $('#bulk');
    if (!el) return;
    el.innerHTML = n ? '<div class="card row"><b>' + n + ' गुन्हे निवडले</b><span class="grow"></span><button class="btn gold sm" id="bgo">एकत्र पत्रे बनवा</button><button class="btn ghost sm" id="bcl">निवड रद्द</button></div>' : '';
    if (!n) return;
    $('#bgo').onclick = function () { bulkEditor(Object.keys(S.picked).map(function (k) { return S.picked[k]; })); };
    $('#bcl').onclick = function () { S.picked = {}; document.querySelectorAll('.pk1').forEach(function (x) { x.checked = false; }); updateBulk(); };
  }
  function bulkEditor(list) {
    list = list.slice().sort(function (a, b) { return String(a.dcrDate).localeCompare(String(b.dcrDate)) || (+a.dcrSeq || 0) - (+b.dcrSeq || 0); }); // DCR मधील क्रमानेच पत्रे
    var P = paras(), shera = '', rel = RELS[0], place = 'गावात';
    var on = {}, sugg = {};
    var own = list.filter(function (c) { return caseKeys(c).length; }).length;
    var m = document.createElement('div'); m.className = 'modal';
    m.innerHTML = '<div class="sheet"><div class="row"><h2 class="grow" style="margin:0">एकत्र पत्रे: ' + list.length + ' गुन्हे</h2><button class="btn ghost sm" data-x>बंद</button></div>' +
      '<div class="scroll" style="margin:8px 0"><table class="t"><tr><th>गुन्हा</th><th>शेरा</th><th>सूचना</th></tr>' + list.map(function (c) {
        return '<tr><td><b>' + esc(c.ps) + ' ' + crTxt(c) + '</b></td><td>' + (c.shera ? esc(c.shera.slice(0, 90)) : '<span class="muted">खालील सामायिक शेरा</span>') + '</td><td>' +
          (caseKeys(c).length ? '<span class="chip ok">✓ ' + caseKeys(c).length + '</span> ' + esc(caseKeys(c).map(function (k) { var p = LETTER.para(k, P); return p ? p.t : ''; }).join(', ').slice(0, 120)) : '<span class="chip warn">एकही ✓ नाही</span>') + '</td></tr>';
      }).join('') + '</table></div>' +
      '<label>सामायिक SP शेरा (रिकामा ठेवल्यास प्रत्येक गुन्ह्याचा स्वतःचा शेरा वापरला जातो)</label><textarea id="bsh" style="min-height:60px"></textarea>' +
      '<div class="row"><span class="muted">नाते:</span><select id="brel">' + RELS.map(function (r) { return '<option>' + r + '</option>'; }).join('') + '</select><select id="bpl"><option>गावात</option><option>घरात</option><option>परिसरात</option></select></div>' +
      '<label>सर्व पत्रांत जोडायच्या सामायिक सूचना (ऐच्छिक; प्रत्येक गुन्ह्याखाली ✓ केलेल्या सूचना त्या त्या पत्रात आधीच जातात)</label><div id="bp">' + P.map(function (p) {
        return '<div class="pk' + (on[p.k] ? ' on' : '') + '" data-k="' + p.k + '"><label><input type="checkbox"' + (on[p.k] ? ' checked' : '') + '> ' + esc(p.t) + '</label></div>';
      }).join('') + '</div>' +
      '<div class="row" style="margin-top:12px"><button class="btn gold" id="bmk">सर्व पत्रे बनवा (एका Word फाईलमध्ये)</button><span id="bout" class="links"></span></div></div>';
    document.body.appendChild(m);
    $('[data-x]', m).onclick = function () { m.remove(); };
    $('#brel', m).onchange = function () { rel = this.value; };
    $('#bpl', m).onchange = function () { place = this.value; };
    m.querySelectorAll('.pk input').forEach(function (cb) { cb.onchange = function () { cb.parentNode.parentNode.classList.toggle('on', cb.checked); }; });
    $('#bmk', m).onclick = function () {
      var b = this, keys = [];
      m.querySelectorAll('.pk.on').forEach(function (x) { keys.push(x.dataset.k); });
      shera = $('#bsh', m).value.trim();
      var items = list.map(function (c) {
        var sh = shera || c.shera || '', ks = caseKeys(c).slice();
        if (shera && ks.indexOf('shera') < 0) ks.unshift('shera');
        keys.forEach(function (k) { if (ks.indexOf(k) < 0) ks.push(k); });
        var o = { shera: sh, paras: P, settings: { days: S.settings.days }, relation: rel, place: place, refDate: c.dcrIssue || '', keys: ks, rules: rules() };
        var D = LETTER.draft(c, o);
        return { key: c.key, orig: D, L: { jaKr: '', date: S.today, dateText: dmy(S.today), ref: D.ref, subject: D.subject, intro: D.intro, body: D.body, close: D.close,
          prati: D.prati, prat: D.prat, keys: ks, shera: sh } };
      });
      busy(b, true, 'पत्रे बनत आहेत...');
      call('makeLetters', items).then(function (r) {
        $('#bout', m).innerHTML = '<span class="chip ok">' + r.count + ' पत्रे</span>' + links(r);
        list.forEach(function (c) { if (shera) c.shera = shera; c.letters = (+c.letters || 0) + 1; });
        toast(r.count + ' पत्रे Drive मध्ये जतन झाली');
      }).catch(function (e) { toast(e.message, 1); }).then(function () { busy(b, false); });
    };
  }

  /* ---------- 3. DCR चेक नोंदवही ---------- */
  function vLetters() {
    $('#main').innerHTML = '<div class="card"><div class="row"><h2 class="grow" style="margin:0">DCR चेक पत्रांची नोंदवही</h2><input type="month" id="lm" value="' + curMonth() + '"></div><div id="ll" class="scroll"></div></div>';
    $('#lm').onchange = loadLetters; loadLetters();
  }
  function loadLetters() {
    $('#ll').innerHTML = '<p class="muted"><span class="spin"></span></p>';
    call('letters', { month: $('#lm').value }).then(function (rs) {
      if (!rs.length) { $('#ll').innerHTML = '<p class="muted">या महिन्यात पत्र नाही.</p>'; return; }
      var days = +S.settings.days || 3, td = S.today, late = 0;
      $('#ll').innerHTML = '<table class="t"><tr><th>दि.</th><th>जा.क्र.</th><th>पो.स्टे. / गुन्हा</th><th>शेरा</th><th>फाइल</th><th>अनुपालन अहवाल</th></tr>' + rs.map(function (r) {
        var due = addDays(r.date, days), isLate = !r.status && due < td; if (isLate) late++;
        return '<tr class="' + (isLate ? 'late' : '') + '" data-id="' + r.id + '"><td>' + dmy(r.date) + '</td><td>' + esc(r.jaKr) + '</td><td><b>' + esc(r.ps) + '</b> ' + esc(r.crime) + '<div class="muted">' + esc(DCR.HEAD_MR[r.head] || '') + '</div></td><td>' + esc(r.shera) +
          '</td><td class="links"><a target="_blank" href="' + esc(r.docxUrl) + '">Word</a><a target="_blank" href="' + esc(r.pdfUrl) + '">PDF</a></td><td><select class="st"><option value="">बाकी' + (isLate ? ' (मुदत संपली)' : '') + '</option>' +
          ['अहवाल आला', 'स्मरणपत्र दिले', 'बंद'].map(function (o) { return '<option' + (r.status === o ? ' selected' : '') + '>' + o + '</option>'; }).join('') + '</select> <input type="text" class="rd" placeholder="दि." value="' + esc(r.replyDate) + '" style="width:100px"></td></tr>';
      }).join('') + '</table>' + (late ? '<p><span class="chip bad">' + late + ' पत्रांचा अहवाल मुदतीत आला नाही</span></p>' : '');
      $('#ll').querySelectorAll('tr[data-id]').forEach(function (tr) {
        var save = function () { call('setLetter', tr.dataset.id, { status: $('.st', tr).value, replyDate: $('.rd', tr).value }).then(function () { toast('नोंद जतन'); tr.classList.remove('late'); }).catch(function (e) { toast(e.message, 1); }); };
        $('.st', tr).onchange = save; $('.rd', tr).onchange = save;
      });
    }).catch(function (e) { $('#ll').innerHTML = '<ul class="iss"><li class="bad">' + esc(e.message) + '</li></ul>'; });
  }
  function addDays(iso, n) { var d = new Date(iso + 'T00:00:00'); d.setDate(d.getDate() + n); return d.toISOString().slice(0, 10); }

  /* ---------- 4. B भाग ---------- */
  function vBpart() {
    var m = (S.pickDay || '').slice(0, 7) || curMonth();
    if (!S.pickDay && +String(S.today).slice(8, 10) <= 7) { var d = new Date(m + '-01T00:00:00'); d.setMonth(d.getMonth() - 1); m = d.toISOString().slice(0, 7); }
    $('#main').innerHTML = '<div class="card"><h2>हेडवाइज B भाग</h2><p class="muted">रोज: DCR अपलोड → "दैनिक B भाग" पाहा व तपासा → "मासिक B भागात जोडा". जोडलेल्या सर्व दिवसांचे गुन्हे शीर्षकानुसार महिन्याच्या एकाच फाईलमध्ये जमा होतात.</p>' +
      '<div class="row"><input type="month" id="bm" value="' + m + '"><button class="btn ghost" id="bc">महिन्याची तपासणी</button></div></div>' +
      '<div id="bmo"></div><div id="bdl"></div><div id="bdv"></div><div id="br"></div><div id="bal"></div>';
    $('#bm').onchange = function () { $('#bdv').innerHTML = ''; $('#br').innerHTML = ''; days(); };
    $('#bc').onclick = check;
    days(S.pickDay); S.pickDay = null;
  }
  function monthName(m) { var M = ['जानेवारी', 'फेब्रुवारी', 'मार्च', 'एप्रिल', 'मे', 'जून', 'जुलै', 'ऑगस्ट', 'सप्टेंबर', 'ऑक्टोबर', 'नोव्हेंबर', 'डिसेंबर']; return M[+m.slice(5, 7) - 1] + ' ' + m.slice(0, 4); }
  function bLinks(r) { return '<a target="_blank" href="' + esc(r.docxUrl) + '">Word</a>' + (r.pdfUrl ? '<a target="_blank" href="' + esc(r.pdfUrl) + '">PDF</a>' : ''); }
  function days(open) {
    var mo = $('#bm').value;
    $('#bdl').innerHTML = '<p class="muted"><span class="spin"></span> दिवस शोधत आहे...</p>';
    call('bDays', mo).then(function (r) {
      var nAdd = r.days.filter(function (x) { return x.bAdded; }).length;
      $('#bmo').innerHTML = '<div class="hero"><div class="hero-l"><span class="eyebrow">मासिक हेडवाइज B भाग</span><h2>' + monthName(mo) + '</h2>' +
        (r.monthly ? '<p>अद्ययावत ' + esc(r.monthly.created) + '</p><div class="links">' + bLinks(r.monthly) + '<button class="btn sm" id="bme">ॲपमध्ये पाहा / संपादन</button></div>' : '<p>अजून बनलेला नाही. दैनिक B भाग तपासून "मासिक मध्ये जोडा" करा.</p>') +
        '</div><div class="hero-n"><div><b>' + (r.monthly ? r.monthly.cases : 0) + '</b><span>गुन्हे</span></div><div><b>' + nAdd + '/' + r.days.length + '</b><span>दिवस जोडले</span></div></div></div>';
      if (r.monthly) $('#bme').onclick = function () { bEditor({ month: mo }); };
      $('#bal').innerHTML = r.all.length ? '<div class="card"><h2>बनलेले मासिक B भाग</h2><div class="tiles">' + r.all.map(function (x) {
        return '<div class="tile"><span class="eyebrow">B भाग</span><b>' + monthName(x.month) + '</b><small>' + x.cases + ' गुन्हे • ' + esc(x.created) + '</small><div class="links">' + bLinks(x) + '<button class="btn sm ghost" data-mo="' + x.month + '">पाहा</button></div></div>';
      }).join('') + '</div></div>' : '';
      $('#bal').querySelectorAll('[data-mo]').forEach(function (b) { b.onclick = function () { bEditor({ month: b.dataset.mo }); }; });
      if (!r.days.length) { $('#bdl').innerHTML = '<div class="card muted">या महिन्याचे DCR अपलोड झालेले नाहीत.</div>'; return; }
      var added = r.days.filter(function (x) { return x.bAdded; }).length;
      $('#bdl').innerHTML = '<div class="card"><h3>' + mo.slice(5) + '/' + mo.slice(0, 4) + ' चे DCR दिवस (' + added + '/' + r.days.length + ' मासिक B भागात)</h3><div class="scroll"><table class="t"><tr><th>DCR दि.</th><th>गुन्हे</th><th>त्रुटी</th><th>दैनिक B भाग</th><th>स्थिती</th><th></th></tr>' +
        r.days.map(function (x) {
          return '<tr data-d="' + x.date + '"><td>' + dmy(x.date) + '</td><td>' + x.cases + '</td><td>' + (x.bad ? '<span class="chip bad">' + x.bad + '</span>' : '') + (x.warn ? '<span class="chip warn">' + x.warn + '</span>' : '') + '</td>' +
            '<td class="links">' + (x.file ? bLinks(x.file) + '<button class="btn sm ghost" data-a="see">पाहा</button>' : '<span class="muted">बनलेला नाही</span>') + '</td>' +
            '<td>' + (x.bAdded ? '<span class="chip ok">जोडला</span>' : '<span class="chip warn">बाकी</span>') + '</td>' +
            '<td><button class="btn sm ghost" data-a="view">दैनिक B भाग</button> ' + (x.bAdded ? '<button class="btn sm ghost" data-a="rm">काढा</button>' : '<button class="btn sm gold" data-a="add">मासिक मध्ये जोडा</button>') + '</td></tr>';
        }).join('') + '</table></div></div>';
      $('#bdl').querySelectorAll('tr[data-d]').forEach(function (tr) {
        var d = tr.dataset.d;
        $('[data-a=view]', tr).onclick = function () { dayView(d); };
        var see = $('[data-a=see]', tr); if (see) see.onclick = function () { bEditor({ date: d }); };
        var ad = $('[data-a=add]', tr), rm = $('[data-a=rm]', tr);
        if (ad) ad.onclick = function () { addDay(d, true, ad); };
        if (rm) rm.onclick = function () { addDay(d, false, rm); };
      });
      if (open) dayView(open);
    }).catch(function (e) { $('#bdl').innerHTML = '<div class="card"><ul class="iss"><li class="bad">' + esc(e.message) + '</li></ul></div>'; });
  }
  function addDay(d, on, b) {
    busy(b, true, on ? 'जोडत आहे...' : 'काढत आहे...');
    call('addDay', d, on).then(function (r) {
      toast(on ? dmy(d) + ' मासिक B भागात जोडला (' + r.cases + ' गुन्हे)' : dmy(d) + ' मासिक B भागातून काढला');
      S.dcrs.forEach(function (x) { if (x.reportDate === d) x.bAdded = on ? 'होय' : ''; });
      days(); if ($('#bdv [data-day]') && $('#bdv [data-day]').dataset.day === d) dayView(d);
    }).catch(function (e) { toast(e.message, 1); busy(b, false); });
  }
  function dayView(d) {
    $('#bdv').innerHTML = '<div class="card"><p class="muted"><span class="spin"></span> ' + dmy(d) + ' चा दैनिक B भाग बनत आहे...</p></div>';
    call('dayBPart', d).then(function (r) {
      if (r.docxUrl) days();
      var li = r.issues.map(function (i) { return '<li class="' + i.level + '">' + esc(i.msg) + '</li>'; }).join('');
      $('#bdv').innerHTML = '<div class="card" data-day="' + d + '"><div class="row"><h3 class="grow" style="margin:0">दैनिक B भाग: DCR दि. ' + dmy(d) + ' (' + r.cases + ' गुन्हे)</h3><span class="links">' + bLinks(r) + '</span><button class="btn sm gold" id="bde">ॲपमध्ये तपासा / संपादन</button></div>' +
        '<div class="scroll"><table class="t" style="margin-top:8px"><tr><th>शीर्षक</th><th>गुन्हे</th></tr>' + r.counts.map(function (c) { return '<tr><td>' + esc(c[1]) + '</td><td>' + c[2] + '</td></tr>'; }).join('') + '</table></div>' +
        (li ? '<h3>DCR मधील त्रुटी (जोडण्यापूर्वी तपासा)</h3><ul class="iss">' + li + '</ul>' : '<p class="muted">मोठ्या त्रुटी नाहीत.</p>') +
        (r.others.length ? '<h3>"इतर" मधील गुन्हे: योग्य शीर्षक निवडा</h3>' + r.others.map(function (o) { return '<div class="case row" data-k="' + esc(o.key) + '"><span class="grow"><b>' + esc(o.ps) + ' ' + o.no + '/' + String(o.yr).slice(2) + '</b> ' + esc(o.kalam) + ' (' + esc(o.label) + ')</span>' + headSel(o.head, 'hs') + '</div>'; }).join('') : '') +
        '<div class="row" style="margin-top:10px">' + (r.bAdded ? '<span class="chip ok">हा दिवस मासिक B भागात आहे</span><button class="btn ghost" id="bdr">मासिक मधून काढा</button>' : '<button class="btn gold" id="bda">तपासले, मासिक B भागात जोडा</button>') + '</div></div>';
      $('#bdv').querySelectorAll('[data-k]').forEach(function (row) {
        $('.hs', row).onchange = function () { var v = this.value; call('setHead', row.dataset.k, v).then(function () { toast('शीर्षक बदलले. दैनिक B भाग पुन्हा पाहा.'); }).catch(function (e) { toast(e.message, 1); }); };
      });
      $('#bde').onclick = function () { bEditor({ date: d }); };
      var a = $('#bda'), rr = $('#bdr');
      if (a) a.onclick = function () { addDay(d, true, a); };
      if (rr) rr.onclick = function () { addDay(d, false, rr); };
    }).catch(function (e) { $('#bdv').innerHTML = '<div class="card"><ul class="iss"><li class="bad">' + esc(e.message) + '</li></ul></div>'; });
  }
  /* B भाग ॲपमध्येच तपासणे व संपादन (दैनिक: {date}, मासिक: {month}). बदल गुन्ह्यावर जतन होतात, त्यामुळे दैनिक व मासिक दोन्हीत येतात. */
  var BCOLS = ['अ.क्र.', 'पोस्टे', 'अप.क्र. व कलम', 'घटनास्थळ', 'घ. ता. वेळ', 'दा. ता. वेळ', 'फिर्यादी व आरोपीचे नाव व पत्ता', 'अटक / प्रतिबंधक', 'तपास अधिकारी', 'घटनेचे थोडक्यात विवरण'];
  var BADC = ['अ.क्र.', 'पोस्टे', 'मर्ग क्र. व कलम', 'घटनास्थळ', 'घ. ता. वेळ', 'दा. ता. वेळ', 'फिर्यादीचे नाव', 'मृतकाचे नाव', 'दाखल / तपास', 'मरणाचे कारण'];
  function bEditor(q) {
    var name = q.date ? 'दैनिक B भाग: DCR दि. ' + dmy(q.date) : 'मासिक B भाग: ' + q.month.slice(5) + '/' + q.month.slice(0, 4);
    $('#bdv').innerHTML = '<div class="card"><p class="muted"><span class="spin"></span> ' + name + ' उघडत आहे...</p></div>';
    $('#bdv').scrollIntoView({ behavior: 'smooth' });
    call('bView', q).then(function (r) {
      S.bv = r;
      $('#bdv').innerHTML = '<div class="card bed"><div class="row"><h3 class="grow" style="margin:0">' + name + ' (' + r.total + ' गुन्हे) • तपासणी व संपादन</h3>' +
        '<input type="search" id="bef" placeholder="पो.स्टे. / अप.क्र. / शब्द शोधा" style="width:220px"><button class="btn gold" id="bgo">Word / PDF पुन्हा बनवा</button></div>' +
        '<p class="muted">ओळीवर "बदला" दाबून मजकूर, शीर्षक दुरुस्त करा किंवा गुन्हा B भागातून वगळा. बदल त्या गुन्ह्यावर जतन होतात; दैनिक व मासिक दोन्ही B भागात येतात. शेवटी "Word / PDF पुन्हा बनवा" दाबा.</p><div id="bgl" class="links"></div>' +
        r.secs.map(function (sc, si) {
          var cols = sc.head === 'AD' ? BADC : BCOLS;
          return '<h3>' + (si + 1) + ') ' + esc(sc.title) + ' - ' + sc.rows.filter(function (x) { return !x.skip; }).length + '</h3><div class="scroll"><table class="t bt"><tr><th></th>' + cols.map(function (c) { return '<th>' + esc(c) + '</th>'; }).join('') + '</tr>' +
            sc.rows.map(function (x, ri) {
              return '<tr data-s="' + si + '" data-r="' + ri + '" class="' + (x.skip ? 'skip' : '') + '"><td class="act"><button class="btn sm ghost" data-a="ed">बदला</button>' + (x.edited ? '<span class="chip warn">बदललेले</span>' : '') + (x.skip ? '<span class="chip bad">वगळले</span>' : '') + '</td>' +
                x.cells.map(function (c) { return '<td>' + esc(c).replace(/\n/g, '<br>') + '</td>'; }).join('') + '</tr>';
            }).join('') + '</table></div>';
        }).join('') + '</div>';
      $('#bef').oninput = function () {
        var t = this.value.trim();
        $('#bdv').querySelectorAll('tr[data-r]').forEach(function (tr) { tr.style.display = !t || tr.textContent.indexOf(t) >= 0 ? '' : 'none'; });
      };
      $('#bdv').querySelectorAll('tr[data-r]').forEach(function (tr) {
        var sc = r.secs[+tr.dataset.s], x = sc.rows[+tr.dataset.r];
        $('[data-a=ed]', tr).onclick = function () { bRowEditor(q, sc, x); };
      });
      $('#bgo').onclick = function () {
        var b = this; busy(b, true, 'बनत आहे...');
        (q.date ? call('dayBPart', q.date) : call('buildBPart', q.month, 'ALL')).then(function (o) {
          $('#bgl').innerHTML = '<span class="chip ok">नवीन फाईल तयार: ' + o.cases + ' गुन्हे</span>' + bLinks(o); toast('Word / PDF अद्ययावत झाले');
          if (q.month) days();
        }).catch(function (e) { toast(e.message, 1); }).then(function () { busy(b, false); });
      };
    }).catch(function (e) { $('#bdv').innerHTML = '<div class="card"><ul class="iss"><li class="bad">' + esc(e.message) + '</li></ul></div>'; });
  }
  function bRowEditor(q, sc, x) {
    var cols = sc.head === 'AD' ? BADC : BCOLS, m = document.createElement('div');
    m.className = 'modal';
    m.innerHTML = '<div class="sheet"><div class="row"><h2 class="grow" style="margin:0">B भाग ओळ: ' + esc(x.cells[1]) + ' ' + esc(String(x.cells[2]).split('\n')[0]) + '</h2><button class="btn ghost sm" data-x>बंद</button></div>' +
      '<div class="row"><span class="muted">शीर्षक:</span>' + headSel(sc.head, 'bh') + '<label class="row" style="margin:0"><input type="checkbox" id="bsk"' + (x.skip ? ' checked' : '') + '> हा गुन्हा B भागातून वगळा</label></div>' +
      cols.slice(1).map(function (c, i) { return '<label>' + esc(c) + '</label><textarea data-c="' + (i + 1) + '" style="min-height:' + (i === 8 ? 150 : 44) + 'px">' + esc(x.cells[i + 1]) + '</textarea>'; }).join('') +
      '<div class="row" style="margin-top:12px"><button class="btn gold" id="bsv">जतन करा</button>' + (x.edited ? '<button class="btn ghost" id="brs">DCR मधील मूळ मजकूर परत आणा</button>' : '') + '</div></div>';
    document.body.appendChild(m);
    $('[data-x]', m).onclick = function () { m.remove(); };
    function send(o, b) {
      busy(b, true);
      call('saveBRow', x.key, o).then(function () { m.remove(); toast('जतन झाले. शेवटी "Word / PDF पुन्हा बनवा" दाबा.'); bEditor(q); }).catch(function (e) { toast(e.message, 1); busy(b, false); });
    }
    $('#bsv', m).onclick = function () {
      var cells = [x.cells[0]], ch = false;
      m.querySelectorAll('textarea[data-c]').forEach(function (t) { cells[+t.dataset.c] = t.value.trim(); });
      for (var i = 1; i < cells.length; i++) if (cells[i] !== String(x.auto[i]).trim()) ch = true;
      var o = { skip: $('#bsk', m).checked, head: $('.bh', m).value };
      o.cells = ch ? cells : null;
      send(o, this);
    };
    var rs = $('#brs', m); if (rs) rs.onclick = function () { send({ cells: null }, this); };
  }
  function check() {
    $('#br').innerHTML = '<div class="card muted"><span class="spin"></span> तपासत आहे...</div>';
    call('monthCheck', $('#bm').value, 'ALL').then(function (r) {
      var ds = [];
      var rows = DCR.HEADS.filter(function (h) { return r.counts[h[0]]; }).map(function (h) {
        var c = r.counts[h[0]], t = 0; Object.keys(c).forEach(function (k) { t += c[k]; });
        return '<tr><td>' + esc(h[1]) + '</td>' + ds.map(function (d) { return '<td>' + (c[d] || '') + '</td>'; }).join('') + '<td><b>' + t + '</b></td></tr>';
      }).join('');
      $('#br').innerHTML = '<div class="card"><h3>कमी / त्रुटी</h3>' + (r.issues.length ? '<ul class="iss">' + r.issues.map(function (i) { return '<li class="' + i.level + '">' + esc(i.msg) + '</li>'; }).join('') + '</ul>' : '<p class="muted">काही कमी आढळली नाही.</p>') + '</div>' +
        '<div class="card"><h3>शीर्षकवार संख्या (एकूण ' + r.total + ')</h3><div class="scroll"><table class="t"><tr><th>शीर्षक</th>' + ds.map(function (d) { return '<th>' + esc(d) + '</th>'; }).join('') + '<th>एकूण</th></tr>' + rows + '</table></div></div>' +
        (r.others.length ? '<div class="card"><h3>"इतर" मधील गुन्हे: योग्य शीर्षक निवडा</h3>' + r.others.map(function (o) {
          return '<div class="row case" data-k="' + esc(o.key) + '"><span class="grow"><b>' + esc(o.ps) + ' ' + o.no + '/' + String(o.yr).slice(2) + '</b> ' + esc(o.kalam) + (o.label ? ' (' + esc(o.label) + ')' : '') + '</span>' + headSel(o.head, 'hs') + '</div>';
        }).join('') + '</div>' : '');
      $('#br').querySelectorAll('[data-k]').forEach(function (row) {
        $('.hs', row).onchange = function () { var v = this.value; call('setHead', row.dataset.k, v).then(function () { toast('शीर्षक बदलले'); }).catch(function (e) { toast(e.message, 1); }); };
      });
    }).catch(function (e) { $('#br').innerHTML = '<div class="card"><ul class="iss"><li class="bad">' + esc(e.message) + '</li></ul></div>'; });
  }

  /* ---------- 5. सेटिंग्ज ---------- */
  function vSet() {
    var s = S.settings, f = function (k, l, ta) { return '<label>' + l + '</label>' + (ta ? '<textarea data-s="' + k + '">' + esc(s[k]) + '</textarea>' : '<input type="text" data-s="' + k + '" value="' + esc(s[k]) + '">'); };
    $('#main').innerHTML = '<div class="card sheet" style="max-width:none;box-shadow:none"><h2>पत्राचे तपशील</h2>' +
      f('spName', 'पोलीस अधीक्षकांचे नाव (सही खाली)') + f('spDesig', 'पदनाम (ओळीनुसार)', 1) +
      f('jaKrPrefix', 'जा.क्र. सुरुवात') + f('jaKrFile', 'जा.क्र. फाईल क्रमांक (उदा. 26-1962)') + '<div class="row"><div><label>जा.क्र. आपोआप क्रमांक</label><select data-s="autoNo"><option' + (s.autoNo === 'होय' ? ' selected' : '') + '>होय</option><option' + (s.autoNo !== 'होय' ? ' selected' : '') + '>नाही</option></select></div>' +
      '<div><label>शेवटचा जा.क्र. क्रमांक (पुढचा आपोआप +1)</label><input type="text" data-s="letterSeq" value="' + esc(s.letterSeq) + '" style="width:90px"></div><div><label>अहवालाची मुदत (दिवस)</label><input type="text" data-s="days" value="' + esc(s.days) + '" style="width:70px"></div>' +
      '<div><label>फॉन्ट (Word/PDF)</label><input type="text" data-s="font" value="' + esc(s.font) + '" style="width:160px"></div></div>' +
      f('quick', 'नेहमीचे शेरे, सर्व गुन्ह्यांसाठी (एका ओळीत एक)', 1) + f('distName', 'B भागावरील घटकाचे नाव') + f('divOrder', 'उपविभागांचा क्रम (स्वल्पविरामाने)') +
      '<div class="row" style="margin-top:12px"><button class="btn" id="ss">जतन करा</button></div></div>' +
      '<div class="card"><h2>पो.स्टे. → उपविभाग</h2><p class="muted">DCR मधील "... विभाग" ओळींवरून ॲप हे आपोआप शिकते. इथे दुरुस्ती करता येते.</p><div id="pst"></div><button class="btn ghost" id="psv">पो.स्टे. यादी जतन</button></div>';
    $('#ss').onclick = function () {
      var o = {}, b = this; document.querySelectorAll('[data-s]').forEach(function (e) { o[e.dataset.s] = e.value; });
      busy(b, true); call('saveSettings', o).then(function (d) { S.settings = d; toast('सेटिंग्ज जतन'); }).catch(function (e) { toast(e.message, 1); }).then(function () { busy(b, false); });
    };
    $('#main').insertAdjacentHTML('beforeend', quickCard() + logoCard() + rulesCard());
    var hs = hsAll(), qs = $('#qhs'), qt = $('#qht');
    function qShow() {
      var h = hs[qs.value] = hs[qs.value] || { p: [], q: [] };
      $('#qhp').innerHTML = paras().filter(function (p) { return p.k !== 'shera'; }).map(function (p) {
        var on = (h.p || []).indexOf(p.k) >= 0;
        return '<label class="pk' + (on ? ' on' : '') + '"><input type="checkbox" value="' + p.k + '"' + (on ? ' checked' : '') + '> ' + esc(p.t) + '</label>';
      }).join('');
      $('#qhp').querySelectorAll('input').forEach(function (cb) {
        cb.onchange = function () { h.p = []; $('#qhp').querySelectorAll('input:checked').forEach(function (x) { h.p.push(x.value); }); cb.parentNode.classList.toggle('on', cb.checked); };
      });
      qt.value = (h.q || []).join('\n');
    }
    qs.onchange = qShow; qShow();
    qt.oninput = function () { hs[qs.value].q = qt.value.split('\n').map(function (x) { return x.trim(); }).filter(String); };
    $('#qhv').onclick = function () {
      var b = this; busy(b, true);
      call('saveSettings', { headSuchna: JSON.stringify(hs) }).then(function (d) { S.settings = d; toast('शीर्षकनिहाय सूचना जतन. आता त्या प्रकारच्या प्रत्येक गुन्ह्याखाली याच सूचना दिसतील.'); }).catch(function (e) { toast(e.message, 1); }).then(function () { busy(b, false); });
    };
    $('#main').querySelectorAll('[data-lsave]').forEach(function (bt) {
      bt.onclick = function () {
        var k = bt.dataset.lsave, fi = $('[data-logo="' + k + '"]'), f = fi.files && fi.files[0];
        if (!f) return toast('आधी चित्र निवडा', 1);
        var rd = new FileReader();
        rd.onload = function () {
          busy(bt, true);
          call('saveLogo', k, String(rd.result).split(',')[1], f.type).then(function () { toast('लोगो जतन झाला'); vSet(); })
            .catch(function (e) { toast(e.message, 1); }).then(function () { busy(bt, false); });
        };
        rd.readAsDataURL(f);
      };
    });
    var rbox = $('#rul');
    function rRow(r) {
      return '<tr><td><input type="checkbox" class="ron"' + (r.off ? '' : ' checked') + ' title="चालू / बंद"></td><td><textarea class="rf" rows="2">' + esc(r.f) + '</textarea></td><td><textarea class="rt" rows="2">' + esc(r.t) + '</textarea></td><td><button class="btn ghost sm" data-rdel>काढा</button></td></tr>';
    }
    function rBind() { rbox.querySelectorAll('[data-rdel]').forEach(function (bt) { bt.onclick = function () { bt.closest('tr').remove(); }; }); }
    rbox.innerHTML = '<div class="scroll"><table class="t"><tr><th>चालू</th><th>पत्रात हा मजकूर असल्यास</th><th>त्याऐवजी हा मजकूर लिहा</th><th></th></tr>' + rules().map(rRow).join('') + '</table></div>';
    rBind();
    $('#radd').onclick = function () { $('table', rbox).insertAdjacentHTML('beforeend', rRow({ f: '', t: '' })); rBind(); var t = rbox.querySelectorAll('.rf'); t[t.length - 1].focus(); };
    $('#rsv').onclick = function () {
      var b = this, rs = [], short = 0;
      rbox.querySelectorAll('tr').forEach(function (tr) {
        var f = $('.rf', tr); if (!f) return;
        var x = { f: f.value.trim(), t: $('.rt', tr).value.trim() };
        if (!x.f) return;
        if (x.f.length < 2) { short++; return; }
        if (!$('.ron', tr).checked) x.off = 1;
        rs.push(x);
      });
      busy(b, true);
      call('saveSettings', { rules: JSON.stringify(rs) }).then(function (d) { S.settings = d; toast(rs.length + ' नियम जतन झाले' + (short ? ' (' + short + ' खूप लहान म्हणून सोडले)' : '') + '. पुढची पत्रे यानुसार बनतील.'); })
        .catch(function (e) { toast(e.message, 1); }).then(function () { busy(b, false); });
    };
    var dv = divs();
    $('#pst').innerHTML = '<div class="scroll"><table class="t"><tr><th>पो.स्टे.</th><th>उपविभाग</th></tr>' + S.ps.map(function (p) {
      return '<tr><td><input type="text" class="pn" value="' + esc(p.ps) + '"></td><td><select class="pd">' + dv.concat(dv.indexOf(p.div) < 0 && p.div ? [p.div] : []).map(function (d) { return '<option' + (d === p.div ? ' selected' : '') + '>' + esc(d) + '</option>'; }).join('') + '</select></td></tr>';
    }).join('') + '</table></div>';
    $('#psv').onclick = function () {
      var list = []; document.querySelectorAll('#pst tr').forEach(function (tr) { var n = $('.pn', tr); if (n && n.value.trim()) list.push({ ps: n.value.trim(), div: $('.pd', tr).value }); });
      call('savePs', list).then(function (d) { S.ps = d; toast('यादी जतन'); }).catch(function (e) { toast(e.message, 1); });
    };
  }

  function quickCard() {
    return '<div class="card"><h2>शीर्षकनिहाय सूचना</h2><p class="muted">गुन्ह्याच्या प्रकारानुसार (खुन, चोरी, अपघात, मर्ग इ.) त्या गुन्ह्याखाली कोणत्या सूचना ✓ करण्यासाठी दिसाव्यात ते इथे ठरवा. सुरुवातीला जुन्या DCR चेक पत्रांवरून भरलेले आहे. पत्रात फक्त गुन्ह्याखाली ✓ केलेल्याच सूचना जातात.</p>' +
      '<div class="row"><label style="margin:0">गुन्ह्याचा प्रकार</label><select id="qhs">' + DCR.HEADS.map(function (h) { return '<option value="' + h[0] + '">' + esc(h[1]) + '</option>'; }).join('') + '</select><span class="grow"></span><button class="btn" id="qhv">शीर्षकनिहाय सूचना जतन</button></div>' +
      '<h3>जुन्या पत्रांतील सूचना (✓ केलेल्या या प्रकारच्या गुन्ह्याखाली दिसतील)</h3><div id="qhp" class="qgrid"></div>' +
      '<h3>तुमच्या स्वतःच्या सूचना (एका ओळीत एक; पत्रात याच शब्दांत जातील)</h3><textarea id="qht" style="min-height:130px"></textarea></div>';
  }
  function logoCard() {
    var s = S.settings, nm = { L: 'डावा (पोलीस लोगो)', C: 'मध्य (राजमुद्रा)', R: 'उजवा' };
    return '<div class="card"><h2>पत्राचे लोगो</h2><p class="muted">DCR FORMAT मधील तीन लोगो आधीपासून पत्रांत येतात. वेगळा लोगो हवा असेल तरच लहान PNG/JPG निवडून "जतन" दाबा.</p>' +
      ['L', 'C', 'R'].map(function (k) {
        return '<div class="row" style="margin:6px 0"><b style="width:170px">' + nm[k] + '</b><input type="file" accept="image/*" data-logo="' + k + '"><button class="btn ghost sm" data-lsave="' + k + '">जतन</button><span class="muted">' + (s['logo' + k] ? 'तुमचा लोगो जतन आहे' : 'तयार लोगो वापरात') + '</span></div>';
      }).join('') + '</div>';
  }
  function rulesCard() {
    return '<div class="card"><h2>पत्रांचे नियम (शिकलेले व तुमचे)</h2><p class="muted">पत्र संपादित करून जतन केलेले बदल इथे नियम म्हणून येतात. कोणताही नियम बदला, बंद करा, काढा किंवा "+ नवीन नियम" ने स्वतः लिहा. पत्रातील (संदर्भ, विषय, मजकूर, सूचना, शेवटचे वाक्य) डाव्या रकान्यातील मजकूर आपोआप उजव्या रकान्यातील मजकुराने बदलतो. उजवा रकाना रिकामा ठेवल्यास तो मजकूर काढला जातो.</p>' +
      '<div id="rul"></div><div class="row" style="margin-top:10px"><button class="btn ghost" id="radd">+ नवीन नियम</button><button class="btn" id="rsv">नियम जतन करा</button></div></div>';
  }

  /* ---------- टायपिंग: युनिकोड / कृतिदेव / रोमन (ॲपमधील सर्व मजकूर चौकटींत) ---------- */
  var IME = { mode: 'uni', st: null };
  try { IME.mode = localStorage.getItem('dcrIme') || 'uni'; } catch (e) {}
  var IME_T = { uni: 'युनिकोड (मराठी कीबोर्ड)', kruti: 'कृतिदेव कीबोर्डप्रमाणे टाइप करा, युनिकोडमध्ये दिसेल', roman: 'Roman: gunha → गुन्हा, aaropi → आरोपी, ThaaNe → ठाणे. aa=ा ii=ी uu=ू T=ट D=ड N=ण L=ळ sh=श Sh=ष dny=ज्ञ M=ं | = ।' };
  function imeBar() {
    var top = $('.bar .sp'); if (!top || $('#ime')) return;
    top.insertAdjacentHTML('afterend', '<span id="ime" class="ime" title="' + esc(IME_T[IME.mode]) + '">टायपिंग: ' +
      [['uni', 'युनिकोड'], ['kruti', 'कृतिदेव'], ['roman', 'Roman']].map(function (x) { return '<button data-ime="' + x[0] + '"' + (IME.mode === x[0] ? ' class="on"' : '') + '>' + x[1] + '</button>'; }).join('') + '</span>');
    $('#ime').querySelectorAll('button').forEach(function (b) {
      b.onclick = function () {
        IME.mode = b.dataset.ime; IME.st = null;
        try { localStorage.setItem('dcrIme', IME.mode); } catch (e) {}
        $('#ime').querySelectorAll('button').forEach(function (x) { x.classList.toggle('on', x === b); });
        $('#ime').title = IME_T[IME.mode]; toast(IME_T[IME.mode]);
      };
    });
  }
  function imeOk(el) { return el && !el.hasAttribute('data-noime') && (el.tagName === 'TEXTAREA' || (el.tagName === 'INPUT' && /^(text|search)$/.test(el.type))); }
  function imeConv(buf) { return IME.mode === 'kruti' ? kruti2uni(buf) : roman2uni(buf); }
  function imeBuf(ch) { return IME.mode === 'kruti' ? !/\s/.test(ch) : /[A-Za-z~]/.test(ch); }
  /* चालू शब्दाची मूळ अक्षरे (buf) लक्षात ठेवून तो शब्द प्रत्येक अक्षराला पुन्हा रूपांतरित होतो; जागा/Enter/कर्सर हलवल्यास शब्द पूर्ण */
  document.addEventListener('beforeinput', function (e) {
    var el = e.target;
    if (IME.mode === 'uni' || !imeOk(el)) return;
    var st = IME.st, a = el.selectionStart, z = el.selectionEnd;
    var live = st && st.el === el && a === z && a === st.at + st.out.length && el.value.substr(st.at, st.out.length) === st.out;
    function put(out) { el.setRangeText(out, st.at, st.at + st.out.length, 'end'); st.out = out; }
    if (e.inputType === 'insertText' && e.data) {
      e.preventDefault();
      if (a !== z) { el.setRangeText('', a, z, 'end'); live = false; }
      if (!live) st = IME.st = { el: el, at: el.selectionStart, buf: '', out: '' };
      e.data.split('').forEach(function (ch) {
        if (imeBuf(ch)) { st.buf += ch; put(imeConv(st.buf)); }
        else {
          el.setRangeText(IME.mode === 'roman' && ch === '|' ? '।' : ch, el.selectionStart, el.selectionEnd, 'end');
          st = IME.st = { el: el, at: el.selectionStart, buf: '', out: '' };
        }
      });
      el.dispatchEvent(new Event('input', { bubbles: true }));
    } else if (e.inputType === 'deleteContentBackward' && live && st.buf) {
      e.preventDefault(); st.buf = st.buf.slice(0, -1); put(st.buf ? imeConv(st.buf) : '');
      el.dispatchEvent(new Event('input', { bubbles: true }));
    } else IME.st = null;
  }, true);
  document.addEventListener('mousedown', function () { IME.st = null; }, true);

  window.addEventListener('load', boot);
})();
