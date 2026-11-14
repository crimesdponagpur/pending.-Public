var C='pl-v7-v1080';
var RT='pl-rt-1';
var CDN=/^https:\/\/(fonts\.googleapis\.com|fonts\.gstatic\.com)\//;
self.addEventListener('install',function(e){e.waitUntil(caches.open(C).then(function(c){return Promise.all(['./','./index.html','./config.js','./vault.js','./manifest.webmanifest','./icon-192.png','./icon-512.png'].map(function(u){return fetch(new Request(u,{cache:'reload'})).then(function(r){return r.ok?c.put(u,r):0})}))}));self.skipWaiting()});
self.addEventListener('activate',function(e){e.waitUntil(caches.keys().then(function(k){return Promise.all(k.filter(function(x){return x!==C&&x.indexOf('pl-v7-')===0}).map(function(x){return caches.delete(x)}))}));self.clients.claim()});
self.addEventListener('fetch',function(e){if(e.request.method!=='GET')return;var u=new URL(e.request.url);
/* फॉन्ट व Excel/PDF लायब्ररी: फोनमध्ये जतन, पुढच्या वेळी लगेच (इंटरनेट नसतानाही) */
if(CDN.test(e.request.url)){e.respondWith(caches.open(RT).then(function(c){return c.match(e.request).then(function(hit){var net=fetch(e.request).then(function(r){if(r.ok||r.type==='opaque')c.put(e.request,r.clone());return r});if(hit){if(u.hostname!=='fonts.gstatic.com')e.waitUntil(net.catch(function(){}));return hit}return net})}));return}
if(u.origin!==location.origin)return;e.respondWith(caches.open(C).then(function(c){return c.match(e.request,{ignoreSearch:true}).then(function(hit){var net=fetch(e.request,{cache:'no-cache'}).then(function(r){if(r.ok)c.put(e.request,r.clone());return r});if(hit){e.waitUntil(net.catch(function(){}));return hit}return net.catch(function(){return c.match('./index.html')})})}))});
