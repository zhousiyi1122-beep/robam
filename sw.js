// ROBAM Service Worker
const CACHE = 'robam-v1';
const ASSETS = [
  './ROBAM.html',
  './manifest.json'
];

// Install - cache core files
self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE).then(cache => cache.addAll(ASSETS))
  );
  self.skipWaiting();
});

// Activate - clean old caches
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// Fetch - serve from cache, fallback to network
self.addEventListener('fetch', e => {
  // Don't cache Google Apps Script calls
  if(e.request.url.includes('script.google.com')){
    e.respondWith(fetch(e.request).catch(()=>new Response(JSON.stringify({ok:false,msg:'Offline'}),{headers:{'Content-Type':'application/json'}})));
    return;
  }
  e.respondWith(
    caches.match(e.request).then(cached => {
      if(cached) return cached;
      return fetch(e.request).then(res => {
        // Cache new resources
        const clone = res.clone();
        caches.open(CACHE).then(cache => cache.put(e.request, clone));
        return res;
      }).catch(() => caches.match('./ROBAM.html'));
    })
  );
});
