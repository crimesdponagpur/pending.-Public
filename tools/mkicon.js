/* MCR आयकॉन: चालू महिन्याइतक्या सोनेरी कडी. वापर: node mkicon.js [महिना 1-12] [बाहेरील फोल्डर...] */
const sharp = require('sharp'), fs = require('fs'), path = require('path');
const mn = +process.argv[2] || (new Date().getUTCMonth() + 1 + (new Date().getUTCHours() >= 18 ? 0 : 0));
const dirs = process.argv.slice(3).length ? process.argv.slice(3) : ['.', 'test'];
function logo(mn) {
  var mo = '';
  for (var i = 0; i < 12; i++) {
    var a0 = (i * 30 - 90 + 5) * Math.PI / 180, a1 = (i * 30 - 90 + 25) * Math.PI / 180, r = 30;
    mo += '<path d="M' + (60 + r * Math.cos(a0)).toFixed(2) + ' ' + (57 + r * Math.sin(a0)).toFixed(2) + 'A' + r + ' ' + r + ' 0 0 1 ' + (60 + r * Math.cos(a1)).toFixed(2) + ' ' + (57 + r * Math.sin(a1)).toFixed(2) + '" fill="none" stroke="' + (i < mn ? 'url(#g)' : '#7db4ff') + '" stroke-width="6.5" stroke-linecap="round"' + (i < mn ? '' : ' opacity=".45"') + '/>';
  }
  var sh = 'M60 6L103 21V58C103 87 85 105 60 115C35 105 17 87 17 58V21Z';
  return '<defs><linearGradient id="b" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#2a7de1"/><stop offset=".55" stop-color="#0b46a8"/><stop offset="1" stop-color="#071f55"/></linearGradient>' +
    '<linearGradient id="g" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#ffd966"/><stop offset=".5" stop-color="#f2a516"/><stop offset="1" stop-color="#c46a0a"/></linearGradient></defs>' +
    '<path d="' + sh + '" fill="url(#b)"/><path d="' + sh + '" fill="none" stroke="url(#g)" stroke-width="5" stroke-linejoin="round"/>' +
    '<circle cx="60" cy="57" r="38" fill="#fff" opacity=".07"/>' + mo + '<circle cx="60" cy="57" r="19" fill="#fff"/>' +
    '<path d="M50.5 57.5l7 7 13-14.5" fill="none" stroke="#0b46a8" stroke-width="6" stroke-linecap="round" stroke-linejoin="round"/>';
}
function icon(mask, size) {
  const sh = mask ? 66 : 86, o = (100 - sh) / 2;
  return '<svg xmlns="http://www.w3.org/2000/svg" width="' + size + '" height="' + size + '" viewBox="0 0 100 100"><defs><radialGradient id="bg" cx=".3" cy=".2" r=".9"><stop offset="0" stop-color="#1b4fb5"/><stop offset=".7" stop-color="#0a1a3f"/></radialGradient></defs>' +
    '<rect width="100" height="100" fill="url(#bg)"/><svg x="' + o + '" y="' + o + '" width="' + sh + '" height="' + sh + '" viewBox="0 0 120 120">' + logo(mn) + '</svg></svg>';
}
(async () => {
  for (const d of dirs) for (const [n, m, s] of [['icon-192', 0, 192], ['icon-512', 0, 512], ['icon-m-192', 1, 192], ['icon-m-512', 1, 512]]) {
    if (!fs.existsSync(d)) continue;
    await sharp(Buffer.from(icon(m, s))).png().toFile(path.join(d, n + '.png'));
  }
  console.log('महिना', mn);
})();
