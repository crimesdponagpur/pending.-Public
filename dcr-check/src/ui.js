/* DCR चेक ॲप: स्क्रीन */
(function () {
  var S = { settings: {}, dcrs: [], ps: [], today: '', cases: [], tab: 'dcr' };
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
      $('#lnkRoot').href = d.rootUrl; $('#lnkSs').href = d.ssUrl; $('#ver').textContent = 'v' + d.version;
      show('dcr');
    }).catch(function (e) { $('#main').innerHTML = '<div class="card"><h2>ॲप सुरू झाले नाही</h2><p>' + esc(e.message) + '</p></div>'; });
  }
  function show(t) {
    S.tab = t;
    document.querySelectorAll('nav button').forEach(function (b) { b.classList.toggle('on', b.dataset.t === t); });
    ({ dcr: vDcr, cases: vCases, letters: vLetters, bpart: vBpart, set: vSet })[t]();
  }

  /* ---------- 1. DCR अपलोड ---------- */
  function vDcr() {
    $('#main').innerHTML =
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
      '<div class="row" style="margin-top:8px"><button class="btn sm" data-go="' + R.reportDate + '">या DCR वर शेरा / पत्र →</button></div></div>';
  }
  function bindCard(d) { var b = document.querySelector('[data-go="' + d + '"]'); if (b) b.onclick = function () { S.pickDate = d; show('cases'); }; }
  function refreshDcrs() { call('init').then(function (d) { S.dcrs = d.dcrs; S.ps = d.ps; dcrList(); }); }

  /* ---------- 2. गुन्हे व शेरा ---------- */
  function vCases() {
    var dates = S.dcrs.filter(function (r) { return +r.cases > 0; }).map(function (r) { return r.reportDate; });
    var d0 = S.pickDate || dates[0] || ''; S.pickDate = null;
    $('#main').innerHTML = '<div class="card"><div class="row">' +
      '<select id="qd"><option value="">सर्व तारखा</option>' + dates.map(function (d) { return '<option value="' + d + '"' + (d === d0 ? ' selected' : '') + '>DCR ' + dmy(d) + '</option>'; }).join('') + '</select>' +
      '<input type="text" id="qt" class="grow" placeholder="शोधा: पो.स्टे., अप.क्र. (1048/26), नाव, कलम...">' +
      '<select id="qh"><option value="">सर्व शीर्षके</option>' + DCR.HEADS.map(function (h) { return '<option value="' + h[0] + '">' + esc(h[1]) + '</option>'; }).join('') + '</select>' +
      '<label class="muted"><input type="checkbox" id="qs"> फक्त शेरा असलेले</label><button class="btn" id="qb">शोधा</button><button class="btn ghost" id="qp" title="SP साहेबांसमोर ठेवण्यासाठी">शेरा पत्रक छापा</button></div></div><div id="cl"></div>';
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
      if (q.date && !q.text) list.sort(function (a, b) { return SEC_ORDER.indexOf(a.sec) - SEC_ORDER.indexOf(b.sec) || divs().indexOf(a.div) - divs().indexOf(b.div); });
      $('#cl').innerHTML = '<p class="muted">' + r.total + ' गुन्हे' + (r.more ? ' (पहिले 400 दाखवले)' : '') + '</p>' + list.map(caseHtml).join('');
      list.forEach(bindCase);
    }).catch(function (e) { $('#cl').innerHTML = '<div class="card"><ul class="iss"><li class="bad">' + esc(e.message) + '</li></ul></div>'; });
  }
  function quickChips() { return String(S.settings.quick || '').split('\n').filter(String).map(function (q) { return '<span class="chip q">' + esc(q) + '</span>'; }).join(''); }
  function caseHtml(c) {
    var id = 'c' + btoa(unescape(encodeURIComponent(c.key))).replace(/[^a-z0-9]/gi, '');
    c._id = id;
    return '<div class="case' + (c.shera ? ' has' : '') + '" id="' + id + '"><div class="row"><span class="t grow">' + esc(c.ps) + ' ' + crTxt(c) + (c.kalam ? ' • ' + esc(c.kalam) : '') + '</span>' +
      '<span class="chip">' + esc(DCR.HEAD_MR[c.head] || DCR.SEC_MR[c.sec] || '') + '</span>' + (+c.letters ? '<span class="chip ok">पत्र ' + c.letters + '</span>' : '') + '</div>' +
      '<div class="muted">' + (c.div ? esc(c.div) + ' विभाग • ' : '') + 'दाखल ' + dmy(c.regDate) + (c.label ? ' • ' + esc(c.label) : '') + (c.io ? ' • तपास: ' + esc(c.io.slice(0, 60)) : '') + '</div>' +
      '<div class="facts">' + esc(c.facts || ('फिर्यादी: ' + c.comp + ' • आरोपी: ' + c.acc)) + '</div>' +
      '<textarea placeholder="SP साहेबांचा शेरा येथे लिहा...">' + esc(c.shera) + '</textarea><div class="qk">' + quickChips() + '</div>' +
      '<div class="row" style="margin-top:6px"><button class="btn sm ghost" data-a="save">शेरा जतन</button><button class="btn sm gold" data-a="letter">DCR चेक पत्र बनवा</button>' +
      '<span class="grow"></span><span class="muted">शीर्षक:</span>' + headSel(c.head, 'hs') + '</div></div>';
  }
  function bindCase(c) {
    var el = document.getElementById(c._id); if (!el) return;
    var ta = $('textarea', el);
    $('.facts', el).onclick = function () { this.classList.toggle('open'); };
    ta.onfocus = function () { el.classList.add('qo'); };
    el.querySelectorAll('.chip.q').forEach(function (q) { q.onclick = function () { ta.value = (ta.value.trim() ? ta.value.trim().replace(/[.।]?$/, '. ') : '') + q.textContent; }; });
    $('[data-a=save]', el).onclick = function () {
      var b = this; busy(b, true);
      call('saveShera', c.key, ta.value.trim()).then(function () { c.shera = ta.value.trim(); el.classList.toggle('has', !!c.shera); toast('शेरा जतन झाला'); })
        .catch(function (e) { toast(e.message, 1); }).then(function () { busy(b, false); });
    };
    $('[data-a=letter]', el).onclick = function () { c.shera = ta.value.trim(); letterEditor(c, el); };
    $('.hs', el).onchange = function () {
      var v = this.value; call('setHead', c.key, v).then(function () { c.head = v; toast('शीर्षक बदलले'); }).catch(function (e) { toast(e.message, 1); });
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
    var P = paras(), opt = { shera: c.shera, paras: P, settings: { days: S.settings.days }, relation: RELS[0], place: 'गावात', refDate: c.dcrIssue || '' };
    var D = LETTER.draft(c, opt), on = {};
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
      '<div class="row" style="margin-top:12px"><button class="btn gold" id="lgo">Word + PDF बनवा व Drive मध्ये जतन करा</button><span id="lout" class="links"></span></div></div>';
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
      call('makeLetter', c.key, L).then(function (r) {
        $('#lout', m).innerHTML = '<span class="chip ok">जा.क्र. ' + esc(r.jaKr) + '</span><a target="_blank" href="' + esc(r.docxUrl) + '">Word</a><a target="_blank" href="' + esc(r.pdfUrl) + '">PDF</a>';
        c.letters = (+c.letters || 0) + 1; c.shera = L.shera;
        if (el) { var t = $('textarea', el); if (t) t.value = L.shera; el.classList.toggle('has', !!L.shera); }
        toast('पत्र Drive मध्ये जतन झाले');
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
    var m = curMonth(); if (+String(S.today).slice(8, 10) <= 7) { var d = new Date(m + '-01T00:00:00'); d.setMonth(d.getMonth() - 1); m = d.toISOString().slice(0, 7); }
    $('#main').innerHTML = '<div class="card"><h2>हेडवाइज B भाग (मासिक)</h2><div class="row"><input type="month" id="bm" value="' + m + '"><select id="bs"><option value="ALL">पूर्ण जिल्हा</option>' +
      divs().map(function (d) { return '<option value="' + esc(d) + '">' + esc(d) + ' विभाग</option>'; }).join('') + '</select><button class="btn ghost" id="bc">तपासणी</button><button class="btn gold" id="bb">B भाग बनवा (Word + PDF)</button></div><div id="bo" class="links" style="margin-top:8px"></div></div><div id="br"></div>';
    $('#bc').onclick = check; $('#bm').onchange = check; $('#bs').onchange = check;
    $('#bb').onclick = function () {
      var b = this; busy(b, true, 'B भाग बनत आहे (1-2 मिनिटे)...');
      call('buildBPart', $('#bm').value, $('#bs').value).then(function (r) {
        $('#bo').innerHTML = '<span class="chip ok">' + r.cases + ' गुन्हे</span><a target="_blank" href="' + esc(r.docxUrl) + '">Word</a><a target="_blank" href="' + esc(r.pdfUrl) + '">PDF</a>';
        toast('B भाग Drive मध्ये जतन झाला');
      }).catch(function (e) { toast(e.message, 1); }).then(function () { busy(b, false); });
    };
    check();
  }
  function check() {
    $('#br').innerHTML = '<div class="card muted"><span class="spin"></span> तपासत आहे...</div>';
    call('monthCheck', $('#bm').value, $('#bs').value).then(function (r) {
      var ds = r.divs.filter(function (d) { return $('#bs').value === 'ALL' || d === $('#bs').value; });
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
      f('spName', 'पोलीस अधीक्षकांचे नाव (सही खाली)') + f('spDesig', 'पदनाम (ओळीनुसार)', 1) + f('office', 'कार्यालयाचे शीर्ष (ओळीनुसार)', 1) +
      f('jaKrPrefix', 'जा.क्र. सुरुवात') + '<div class="row"><div><label>जा.क्र. आपोआप क्रमांक</label><select data-s="autoNo"><option' + (s.autoNo === 'होय' ? ' selected' : '') + '>होय</option><option' + (s.autoNo !== 'होय' ? ' selected' : '') + '>नाही</option></select></div>' +
      '<div><label>शेवटचा क्रमांक</label><input type="text" data-s="letterSeq" value="' + esc(s.letterSeq) + '" style="width:90px"></div><div><label>अहवालाची मुदत (दिवस)</label><input type="text" data-s="days" value="' + esc(s.days) + '" style="width:70px"></div>' +
      '<div><label>फॉन्ट (Word/PDF)</label><input type="text" data-s="font" value="' + esc(s.font) + '" style="width:160px"></div></div>' +
      f('quick', 'नेहमीचे शेरे (एका ओळीत एक)', 1) + f('distName', 'B भागावरील घटकाचे नाव') + f('divOrder', 'उपविभागांचा क्रम (स्वल्पविरामाने)') +
      '<div class="row" style="margin-top:12px"><button class="btn" id="ss">जतन करा</button></div></div>' +
      '<div class="card"><h2>पो.स्टे. → उपविभाग</h2><p class="muted">DCR मधील "... विभाग" ओळींवरून ॲप हे आपोआप शिकते. इथे दुरुस्ती करता येते.</p><div id="pst"></div><button class="btn ghost" id="psv">पो.स्टे. यादी जतन</button></div>';
    $('#ss').onclick = function () {
      var o = {}, b = this; document.querySelectorAll('[data-s]').forEach(function (e) { o[e.dataset.s] = e.value; });
      busy(b, true); call('saveSettings', o).then(function (d) { S.settings = d; toast('सेटिंग्ज जतन'); }).catch(function (e) { toast(e.message, 1); }).then(function () { busy(b, false); });
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

  window.addEventListener('load', boot);
})();
