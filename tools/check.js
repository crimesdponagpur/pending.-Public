#!/usr/bin/env node
/* MCR सुरक्षा तपासणी (GitHub वर प्रत्येक बदलावर आपोआप; हाताने: node tools/check.js, दुरुस्ती: node tools/check.js --fix)
 * 1. CSP: पानातील प्रत्येक inline <script> चा sha256 hash CSP मध्ये आहे का (नसेल तर ब्राउझर तो code चालवणार नाही)
 * 2. lib/: ॲपसोबतच्या लायब्ररी फाइली lib/README.md मधील sha256 शी जुळतात का (कोणी बदल केला तर पकडले जाते)
 * 3. inline JavaScript मध्ये syntax चूक नाही ना
 * 4. पानात बाहेरच्या CDN वरून script / inline event handler (onclick= इ.) नाही ना
 */
'use strict';
var fs = require('fs'), path = require('path'), crypto = require('crypto'), vm = require('vm');
var ROOT = path.resolve(__dirname, '..'), FIX = process.argv.indexOf('--fix') >= 0;
var PAGES = process.argv.slice(2).filter(function (a) { return a[0] !== '-'; });
if (!PAGES.length) PAGES = ['index.html', 'noting/index.html'];
var bad = 0;
function fail(m) { bad++; console.error('✗ ' + m); }
function ok(m) { console.log('✓ ' + m); }

PAGES.forEach(function (rel) {
  var f = path.resolve(ROOT, rel), html = fs.readFileSync(f, 'utf8');
  var re = /<script>([\s\S]*?)<\/script>/g, m, hashes = [], n = 0;
  while ((m = re.exec(html))) {
    n++;
    hashes.push("'sha256-" + crypto.createHash('sha256').update(m[1], 'utf8').digest('base64') + "'");
    try { new vm.Script(m[1], { filename: rel + ' <script> #' + n }); } catch (e) { fail(rel + ': inline script #' + n + ' मध्ये syntax चूक: ' + e.message); }
  }
  var cspRe = /(<meta http-equiv="Content-Security-Policy" content=")([^"]*)(">)/;
  var cm = cspRe.exec(html);
  if (!cm) { fail(rel + ': CSP meta नाही'); return; }
  var pol = cm[2], sm = /script-src ([^;]*)/.exec(pol);
  if (!sm) { fail(rel + ': CSP मध्ये script-src नाही'); return; }
  var keep = sm[1].split(/\s+/).filter(function (t) { return t && !/^'sha256-/.test(t); });
  if (keep.some(function (t) { return t === "'unsafe-inline'"; })) fail(rel + ": script-src मध्ये 'unsafe-inline' चालणार नाही");
  var want = keep.concat(hashes).join(' ');
  if (sm[1].trim() === want) ok(rel + ': CSP मध्ये ' + n + ' inline script चे hash बरोबर');
  else if (FIX) { html = html.replace(cspRe, function (_, a, p, z) { return a + p.replace(/script-src [^;]*/, 'script-src ' + want) + z; }); fs.writeFileSync(f, html); ok(rel + ': CSP hash नवीन केले (' + n + ')'); }
  else fail(rel + ': CSP hash जुने आहेत. चालवा: node tools/check.js --fix');
  var ext = html.match(/<script[^>]+src="https?:\/\/[^"]+"/g);
  if (ext) fail(rel + ': बाहेरून script: ' + ext.join(', '));
  var cdn = html.match(/['"]https:\/\/(cdn\.jsdelivr\.net|cdnjs\.cloudflare\.com|unpkg\.com)[^'"]*['"]/g);
  if (cdn) fail(rel + ': CDN वरून लायब्ररी: ' + cdn.join(', '));
  var inl = html.match(/<[a-z][^>]*\son[a-z]+=["'][^>]*>/gi);
  if (inl) fail(rel + ': inline event handler (CSP मुळे चालणार नाही): ' + inl.map(function (x) { return x.slice(0, 80); }).join(' | '));
});

var readme = fs.readFileSync(path.join(ROOT, 'lib/README.md'), 'utf8'), rows = 0;
readme.split('\n').forEach(function (l) {
  var c = l.split('|').map(function (x) { return x.trim(); });
  if (c.length < 6 || !/\.js$/.test(c[1])) return;
  rows++;
  var p = path.join(ROOT, 'lib', c[1]);
  if (!fs.existsSync(p)) return fail('lib/' + c[1] + ' सापडली नाही');
  var h = crypto.createHash('sha256').update(fs.readFileSync(p)).digest('hex');
  if (h !== c[4]) fail('lib/' + c[1] + ' बदललेली आहे! (sha256 ' + h + ')'); else ok('lib/' + c[1]);
});
if (!rows) fail('lib/README.md मध्ये फाइलींची यादी नाही');

if (bad) { console.error('\n' + bad + ' अडचणी'); process.exit(1); }
console.log('\nसर्व तपासण्या बरोबर');
