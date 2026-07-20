var CACHE_NAME = 'gumboot-trial-v4';
var ASSETS = ['./', './index.html', './manifest.json', './app-icon.png', './app-icon-192.png'];

self.addEventListener('install', function(event) {
  event.waitUntil(
    caches.open(CACHE_NAME).then(function(cache) {
      // cache: 'reload' bypasses the browser's HTTP cache so installs always
      // pick up the real latest files, not a stale cached copy of them.
      return Promise.all(ASSETS.map(function(url) {
        return fetch(url, { cache: 'reload' }).then(function(res) { return cache.put(url, res); });
      }));
    })
  );
  self.skipWaiting();
});

self.addEventListener('activate', function(event) {
  event.waitUntil(
    caches.keys().then(function(keys) {
      return Promise.all(keys.filter(function(k) { return k !== CACHE_NAME; }).map(function(k) { return caches.delete(k); }));
    })
  );
  self.clients.claim();
});

self.addEventListener('fetch', function(event) {
  if (event.request.method !== 'GET') return;
  var url = new URL(event.request.url);
  if (url.origin !== self.location.origin) return; // let cross-origin API calls pass through untouched
  // Network-first: always serve the latest version when online (this app is
  // still actively changing), falling back to the cached copy only when
  // offline. Cache-first would otherwise keep serving stale content until
  // CACHE_NAME is manually bumped on every single update.
  event.respondWith(
    fetch(event.request).then(function(res) {
      var resClone = res.clone();
      caches.open(CACHE_NAME).then(function(cache) { cache.put(event.request, resClone); });
      return res;
    }).catch(function() {
      return caches.match(event.request);
    })
  );
});
