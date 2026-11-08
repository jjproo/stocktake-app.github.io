const CACHE_NAME = 'inventory-mobile-offline-v3';
const CACHE_PREFIX = 'inventory-mobile-offline-';
const scope = new URL(self.registration.scope);
const index = new URL('index.html', scope);
const assets = [
  'index.html',
  'manifest.json',
  'css/style.css?v=abf24629ac',
  'js/app.js?v=acd55e1b4b',
  'sample-data.json',
  'icons/icon-192.png',
  'icons/icon-512.png'
];

self.addEventListener('install', event => {
  event.waitUntil((async () => {
    const cache = await caches.open(CACHE_NAME);
    await cache.addAll(assets.map(path => new Request(new URL(path, scope), { cache: 'reload' })));
    await self.skipWaiting();
  })());
});

self.addEventListener('activate', event => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.filter(key => key.startsWith(CACHE_PREFIX) && key !== CACHE_NAME)
      .map(key => caches.delete(key)));
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', event => {
  const request = event.request;
  const url = new URL(request.url);
  if (request.method !== 'GET' || url.origin !== scope.origin || !url.pathname.startsWith(scope.pathname)) return;

  const isAppPage = request.mode === 'navigate' && (url.pathname === scope.pathname || url.pathname === index.pathname);
  const isAppAsset = assets.some(path => new URL(path, scope).pathname === url.pathname);
  if (!isAppPage && !isAppAsset) return;

  event.respondWith((async () => {
    const cache = await caches.open(CACHE_NAME);
    if (isAppPage) {
      try {
        const response = await fetch(request, { cache: 'no-cache' });
        if (response.ok) await cache.put(index, response.clone()).catch(() => {});
        return response;
      } catch (error) {
        const offline = await cache.match(index);
        if (offline) return offline;
        throw error;
      }
    }

    const cached = await cache.match(request);
    if (cached) return cached;
    const response = await fetch(request);
    if (response.ok) await cache.put(request, response.clone()).catch(() => {});
    return response;
  })());
});
