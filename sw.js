const CACHE = 'nz2026-v2';
const ASSETS = [
  './',
  './index.html',
  './manifest.json',
  './icon.svg',
  './images/christchurch.jpg',
  './images/tekapo.jpg',
  './images/queenstown.jpg',
  './images/milford.jpg',
  './images/wanaka.jpg'
];

self.addEventListener('install', (e) => {
  e.waitUntil(
    caches.open(CACHE).then((c) => c.addAll(ASSETS)).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (e) => {
  if (e.request.method !== 'GET') return;
  const url = new URL(e.request.url);
  // Bypass cache for live APIs (weather, currency)
  if (url.hostname.includes('open-meteo.com') || url.hostname.includes('exchangerate.host')) return;

  e.respondWith(
    caches.match(e.request).then((cached) => {
      if (cached) {
        // background revalidate
        fetch(e.request).then((fresh) => {
          if (fresh && fresh.ok) caches.open(CACHE).then((c) => c.put(e.request, fresh));
        }).catch(() => {});
        return cached;
      }
      return fetch(e.request).then((fresh) => {
        if (fresh && fresh.ok && url.origin === location.origin) {
          const copy = fresh.clone();
          caches.open(CACHE).then((c) => c.put(e.request, copy));
        }
        return fresh;
      }).catch(() => cached);
    })
  );
});
