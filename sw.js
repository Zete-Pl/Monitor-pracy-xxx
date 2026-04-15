const CACHE_NAME = 'monitor-pracy-v2';
const ASSETS = [
  './index.html',
  './css/style.css',
  './js/app.js',
  './js/calculations.js',
  './js/store.js',
  './manifest.json'
];

self.addEventListener('install', event => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => {
      return cache.addAll(ASSETS);
    })
  );
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys => {
      return Promise.all(
        keys.filter(key => key !== CACHE_NAME).map(key => caches.delete(key))
      );
    }).then(() => clients.claim())
  );
});

self.addEventListener('fetch', event => {
  event.respondWith(
    fetch(event.request).then(response => {
      // Jeśli mamy poprawną odpowiedź z serwera, klonujemy ją do cache
      if (response && response.status === 200 && response.type === 'basic') {
        const responseToCache = response.clone();
        caches.open(CACHE_NAME).then(cache => {
          cache.put(event.request, responseToCache);
        });
      }
      return response;
    }).catch(() => {
      // Jeśli sieć padnie (offline), używamy cache jako awaryjnego rozwiązania
      return caches.match(event.request);
    })
  );
});
