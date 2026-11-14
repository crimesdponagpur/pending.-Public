/* MCR सुरक्षित तिजोरी (v10.4)
 * लॉगिन token व गुन्ह्यांची साठवलेली माहिती फोन / संगणकात उघड्या स्वरूपात न ठेवता AES-GCM 256 ने encrypt करून IndexedDB मध्ये.
 * किल्ली (CryptoKey) "non-extractable": ती ब्राउझरच्या बाहेर काढता / कॉपी करता येत नाही; storage च्या फाइली कॉपी केल्या तरी माहिती वाचता येत नाही.
 * IndexedDB / WebCrypto नसल्यास (खाजगी विंडो, जुना ब्राउझर): काहीही साठवले जात नाही, फक्त या पानाच्या मेमरीत (सुरक्षित पर्याय).
 * मुख्य ॲप व Noting ॲप दोन्ही हीच फाइल वापरतात.
 */
(function () {
  var DBN = 'pl_vault', ST = 'kv', KEYN = '__k', MEM = {}, KEY = null, DB = null, OK = false, Q = {}, T = {};
  var SENS = ['pl_tok', 'pl_data', 'pl_ntg1', 'pl_rec1', 'pl_undo', 'pl_q', 'pl_pv'];
  /* pl_q = नेट नसताना केलेल्या, अजून सर्व्हरवर न गेलेल्या नोंदी; pl_pv = ऑफलाइन lock उघडण्यासाठी पासवर्ड पडताळणी (hash). हे लॉगआउटला पुसले जात नाहीत (नोंदी हरवू नयेत). */
  var KEEP = ['pl_q', 'pl_pv'];
  function req(r) { return new Promise(function (res, rej) { r.onsuccess = function () { res(r.result); }; r.onerror = function () { rej(r.error); }; }); }
  function store(mode) { return DB.transaction(ST, mode).objectStore(ST); }
  function open() {
    return new Promise(function (res, rej) {
      var r = indexedDB.open(DBN, 1);
      r.onupgradeneeded = function () { r.result.createObjectStore(ST); };
      r.onsuccess = function () { res(r.result); }; r.onerror = function () { rej(r.error); }; r.onblocked = function () { rej(new Error('blocked')); };
    });
  }
  function enc(s) {
    var iv = crypto.getRandomValues(new Uint8Array(12));
    return crypto.subtle.encrypt({ name: 'AES-GCM', iv: iv }, KEY, new TextEncoder().encode(s)).then(function (ct) { return { iv: iv, ct: ct }; });
  }
  function dec(o) { return crypto.subtle.decrypt({ name: 'AES-GCM', iv: o.iv }, KEY, o.ct).then(function (b) { return new TextDecoder().decode(b); }); }
  function persist(k) {
    if (!OK) return;
    var v = MEM[k];
    Q[k] = (Q[k] || Promise.resolve()).then(function () {
      if (v == null) return req(store('readwrite').delete(k));
      return enc(v).then(function (o) { return req(store('readwrite').put(o, k)); });
    }).catch(function () { });
  }
  function rmPlain(k) { try { localStorage.removeItem(k); sessionStorage.removeItem(k); } catch (e) { } }
  var init = (function () {
    if (!window.indexedDB || !window.crypto || !crypto.subtle || !window.isSecureContext) return Promise.resolve();
    var t = new Promise(function (res) { setTimeout(res, 4000, 'timeout'); });
    return Promise.race([t, open().then(function (db) {
      DB = db;
      return req(store('readonly').get(KEYN)).then(function (k) {
        if (k) return k;
        return crypto.subtle.generateKey({ name: 'AES-GCM', length: 256 }, false, ['encrypt', 'decrypt']).then(function (nk) { return req(store('readwrite').put(nk, KEYN)).then(function () { return nk; }); });
      }).then(function (k) {
        KEY = k;
        return Promise.all(SENS.map(function (n) {
          return req(store('readonly').get(n)).then(function (o) { if (o && o.ct) return dec(o).then(function (s) { MEM[n] = s; }, function () { }); });
        }));
      }).then(function () {
        OK = true;
        /* जुनी उघडी (plaintext) माहिती तिजोरीत हलवा व मूळ जागेतून काढा. sessionStorage मधील token (लक्षात न ठेवलेला) तिथेच राहतो. */
        SENS.forEach(function (n) {
          var v = null; try { v = localStorage.getItem(n); } catch (e) { }
          if (v != null) { if (MEM[n] == null) { MEM[n] = v; persist(n); } try { localStorage.removeItem(n); } catch (e) { } }
        });
      });
    })]).then(function (r) { if (r === 'timeout') { OK = false; KEY = null; } }, function () { OK = false; KEY = null; });
  })();
  window.VAULT = {
    ready: init,
    ok: function () { return OK; },
    sens: function (k) { return SENS.indexOf(k) >= 0; },
    get: function (k) { return MEM[k] == null ? null : MEM[k]; },
    /* मोठी माहिती (गुन्ह्यांची यादी) थोडी थांबून एकदाच encrypt; लहान (token) लगेच */
    /* temp = true: फक्त या पानाच्या मेमरीत (लक्षात न ठेवलेला लॉगिन), साठवलेली प्रत पुसली जाते */
    set: function (k, v, temp) {
      MEM[k] = v == null ? null : String(v); rmPlain(k);
      clearTimeout(T[k]);
      if (temp && MEM[k] != null) { var keep = MEM[k]; MEM[k] = null; persist(k); MEM[k] = keep; return; }
      if (MEM[k] && MEM[k].length > 50000) T[k] = setTimeout(function () { persist(k); }, 400); else persist(k);
    },
    del: function (k) { VAULT.set(k, null); },
    /* लॉगआउट: सर्व संवेदनशील माहिती पुसा (किल्ली तशीच; पुढच्या लॉगिनला वापरली जाते) */
    wipe: function () { SENS.forEach(function (n) { if (KEEP.indexOf(n) >= 0) return; clearTimeout(T[n]); MEM[n] = null; rmPlain(n); persist(n); }); }
  };
})();
