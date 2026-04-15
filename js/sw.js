const CACHE_NAME = 'monitor-pracy-v1';
const ASSETS = [
  './index.html',
  './css/style.css',
  './js/app.js',
  './js/calculations.js',
  './js/store.js',
  './manifest.json'
];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => {
      return cache.addAll(ASSETS);
    })
  );
});

self.addEventListener('fetch', event => {
  event.respondWith(
    caches.match(event.request).then(cachedResponse => {
      return cachedResponse || fetch(event.request);
    })
  );
});
