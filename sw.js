const V = 'tpa-nav-v4';
const LIBS = ['https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.js', 'https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.css', 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js', 'https://cdn.jsdelivr.net/npm/vazirmatn@33.0.3/Vazirmatn-font-face.css'];
self.addEventListener('install', e => e.waitUntil(caches.open(V).then(c => Promise.allSettled(['./', 'index.html', ...LIBS].map(u => c.add(u)))).then(() => self.skipWaiting())));
self.addEventListener('activate', e => e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== V).map(k => caches.delete(k)))).then(() => self.clients.claim())));
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET') return;
  if (u.origin === location.origin) { // app files: network first, cache when offline
    e.respondWith(fetch(e.request).then(r => { if (r.ok) { const c = r.clone(); caches.open(V).then(x => x.put(e.request, c)); } return r; }).catch(() => caches.match(e.request).then(h => h || caches.match('index.html'))));
  } else if (u.hostname === 'cdnjs.cloudflare.com' || u.hostname === 'cdn.jsdelivr.net') { // libraries and font: cache first
    e.respondWith(caches.match(e.request).then(hit => hit || fetch(e.request).then(r => { if (r.ok) { const c = r.clone(); caches.open(V).then(x => x.put(e.request, c)); } return r; })));
  }
  // Google Drive/OAuth and map tiles: always live
});
