// SERVICE WORKER FORCÉ - REMU-CI
const CACHE_NAME = 'remuci-force-v3';
const urlsToCache = [
  '/',
  '/manifest.json?v=3',
  '/images/android-icon-192x192.png?v=3',
  '/images/android-icon-512x512.png?v=3',
  '/images/apple-icon-180x180.png?v=3',
  '/force-reinstall'
];

self.addEventListener('install', function(event) {
  console.log('🔄 SW FORCE INSTALL v3');
  self.skipWaiting();
  
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(function(cache) {
        console.log('🗂️ Cache forcé ouvert');
        return cache.addAll(urlsToCache);
      })
      .then(() => {
        console.log('✅ Pré-cache complet');
        return self.skipWaiting();
      })
  );
});

self.addEventListener('activate', function(event) {
  console.log('🔥 SW FORCE ACTIVATION v3');
  
  event.waitUntil(
    caches.keys().then(function(cacheNames) {
      return Promise.all(
        cacheNames.map(function(cacheName) {
          // SUPPRIMER TOUS LES ANCIENS CACHES
          if (cacheName !== CACHE_NAME) {
            console.log('🗑️ Suppression cache:', cacheName);
            return caches.delete(cacheName);
          }
        })
      );
    }).then(() => {
      console.log('✅ Anciens caches supprimés');
      return self.clients.claim();
    })
  );
});

self.addEventListener('fetch', function(event) {
  // FORCER le rechargement des icônes et du manifest
  if (event.request.url.includes('manifest') || event.request.url.includes('icon') || event.request.url.includes('images/')) {
    event.respondWith(
      fetch(event.request).catch(function() {
        return caches.match(event.request);
      })
    );
    return;
  }

  event.respondWith(
    caches.match(event.request).then(function(response) {
      return response || fetch(event.request);
    })
  );
});