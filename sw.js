// Service Worker — Residencia Villa Teresa
// vt-v4: estrategia NETWORK-FIRST para el index.html (evita quedar pegado a una version vieja).
// Los iconos/manifest siguen cache-first porque casi no cambian.
const CACHE = 'vt-v4';
const BASE = '/medicamentos';
const ASSETS = [
  BASE + '/manifest.json',
  BASE + '/icon-192.png',
  BASE + '/icon-512.png',
  BASE + '/icon.svg'
];

self.addEventListener('install', (e) => {
  self.skipWaiting();
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(ASSETS)).catch(() => {}));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);

  // index.html y navegacion: SIEMPRE intentar la red primero, cache solo si no hay internet.
  const esApp = req.mode === 'navigate'
    || url.pathname === BASE + '/'
    || url.pathname.endsWith('/index.html');

  if (esApp) {
    e.respondWith(
      fetch(req)
        .then((res) => {
          const copy = res.clone();
          caches.open(CACHE).then((c) => c.put(BASE + '/index.html', copy));
          return res;
        })
        .catch(() => caches.match(BASE + '/index.html'))
    );
    return;
  }

  // Resto (iconos, manifest): cache-first.
  e.respondWith(caches.match(req).then((r) => r || fetch(req)));
});
