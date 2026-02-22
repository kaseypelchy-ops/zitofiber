// Zito Fiber Sales — Service Worker
const CACHE = 'zito-fiber-v1';

// Core assets to cache on install
const PRECACHE = [
  './index.html',
  './manifest.json',
  'https://fonts.googleapis.com/css2?family=Space+Mono:wght@400;700&family=Syne:wght@400;600;700;800&display=swap',
  'https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.min.css',
  'https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/PapaParse/5.4.1/papaparse.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js'
];

// Install: cache everything upfront
self.addEventListener('install', function(e) {
  e.waitUntil(
    caches.open(CACHE).then(function(cache) {
      return cache.addAll(PRECACHE);
    }).then(function() {
      return self.skipWaiting();
    })
  );
});

// Activate: delete old caches
self.addEventListener('activate', function(e) {
  e.waitUntil(
    caches.keys().then(function(keys) {
      return Promise.all(
        keys.filter(function(k) { return k !== CACHE; })
            .map(function(k) { return caches.delete(k); })
      );
    }).then(function() {
      return self.clients.claim();
    })
  );
});

// Fetch: cache-first for static assets, network-first for API calls
self.addEventListener('fetch', function(e) {
  var url = e.request.url;

  // Always go network-first for Google Sheets API calls
  if (url.includes('script.google.com') || url.includes('nominatim.openstreetmap.org')) {
    e.respondWith(
      fetch(e.request).catch(function() {
        return new Response(JSON.stringify({ error: 'Offline — no network available' }), {
          headers: { 'Content-Type': 'application/json' }
        });
      })
    );
    return;
  }

  // Cache-first for everything else (app shell, libraries, fonts)
  e.respondWith(
    caches.match(e.request).then(function(cached) {
      if (cached) return cached;
      return fetch(e.request).then(function(response) {
        // Only cache valid responses for same-origin or CDN assets
        if (!response || response.status !== 200 || response.type === 'error') {
          return response;
        }
        var toCache = response.clone();
        caches.open(CACHE).then(function(cache) {
          cache.put(e.request, toCache);
        });
        return response;
      });
    })
  );
});
