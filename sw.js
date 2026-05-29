const CACHE = 'vt-v2';
const ASSETS = [
  '/medicamentos/',
  '/medicamentos/index.html',
  '/medicamentos/manifest.json',
  '/medicamentos/icon-192.png',
  '/medicamentos/icon-512.png',
  '/medicamentos/icon.svg'
];
self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS)));
  self.skipWaiting();
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys =>
    Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
  ));
  self.clients.claim();
});
self.addEventListener('fetch', e => {
  e.respondWith(fetch(e.request).catch(() => caches.match(e.request)));
});
