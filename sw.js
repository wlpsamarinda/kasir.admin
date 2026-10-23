const CACHE = 'wonderland-v2';
const SHELL = ['./', './index.html', './config.js', './gas-bridge.js', './manifest.webmanifest', './icons/icon-192.png', './icons/icon-512.png'];
self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (/google\.com|googleusercontent\.com/.test(url.hostname) && !/fonts\./.test(url.hostname)) return; // API: jangan di-cache
  if (/^(localhost|127\.|\[::1\]|192\.168\.|10\.)/.test(url.hostname)) return; // API Hermes lokal: jangan di-cache
  const sameOrigin = url.origin === location.origin;
  if (sameOrigin && (req.mode === 'navigate' || /\.(html|js)$/.test(url.pathname) || url.pathname.endsWith('/'))) {
    // network-first untuk file aplikasi supaya update langsung terbaca
    e.respondWith(fetch(req).then(r => { const cp = r.clone(); caches.open(CACHE).then(c => c.put(req, cp)); return r; })
      .catch(() => caches.match(req).then(m => m || caches.match('./index.html'))));
    return;
  }
  // cache-first untuk CDN & aset statis
  e.respondWith(caches.match(req).then(m => m || fetch(req).then(r => {
    if (r && (r.ok || r.type === 'opaque')) { const cp = r.clone(); caches.open(CACHE).then(c => c.put(req, cp)); }
    return r;
  })));
});
