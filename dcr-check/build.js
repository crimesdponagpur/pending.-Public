// node build.js → gas/ (Apps Script मध्ये चिकटवायच्या फाइल्स: Code.gs, App.html, appsscript.json)
var fs = require('fs'), H = __dirname + '/src/', OUT = __dirname + '/gas/';
function rd(f) { return fs.readFileSync(H + f, 'utf8'); }
function lib(f) { return rd(f).replace(/^if \(typeof module[^\n]*\n?/m, ''); }
var docx = lib('docx.js');
var shared = lib('kruti.js') + '\n' + lib('parse.js') + '\n' + lib('letters.js');
var html = rd('shell.html');
html = html.replace('/*CSS*/', function () { return rd('ui.css'); }).replace('/*JS*/', function () { return shared + '\n' + rd('ui.js'); });
fs.mkdirSync(OUT, { recursive: true });
fs.writeFileSync(OUT + 'App.html', html);
fs.writeFileSync(OUT + 'Code.gs', rd('Code.gs') + '\n\n' + rd('logos.js') + '\n\n' + docx + '\n\n/* ===== सामायिक कोड (ॲपमध्येही हाच): Kruti Dev → Unicode, DCR वाचन, पत्र ===== */\n' + shared);
fs.writeFileSync(OUT + 'appsscript.json', JSON.stringify({ timeZone: 'Asia/Kolkata', exceptionLogging: 'STACKDRIVER', runtimeVersion: 'V8',
  dependencies: { enabledAdvancedServices: [{ userSymbol: 'Drive', serviceId: 'drive', version: 'v3' }] },
  webapp: { executeAs: 'USER_DEPLOYING', access: 'MYSELF' } }, null, 2) + '\n');
console.log('gas/App.html', html.length, 'bytes; gas/Code.gs ok');
