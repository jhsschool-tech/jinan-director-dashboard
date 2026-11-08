// Jinan director dashboard — offline cache for public app shell only.
// Never cache Google API responses or private data.
const CACHE_NAME = 'jinan-pwa-logo-v3-20261108';
const SHELL = [
  './', './index.html', './styles.css', './pwa-brand.css', './app.js',
  './manifest.json', './logo.png', './icon-192.png', './icon-512.png',
  './icon-maskable-512.png', './apple-touch-icon.png', './favicon-32.png'
];
const SHELL_PATHS = new Set(SHELL.map(path => new URL(path, self.registration.scope).pathname));
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => cache.addAll(SHELL))
      .then(() => self.skipWaiting())
  );
});
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(
        keys.filter(key => (key.startsWith('jinan-pwa-') || key.startsWith('jinan-director-')) && key !== CACHE_NAME)
            .map(key => caches.delete(key))
      ))
      .then(() => self.clients.claim())
  );
});
self.addEventListener('fetch', event => {
  if (event.request.method !== 'GET') return;
  const url = new URL(event.request.url);
  if (url.origin !== self.location.origin || !SHELL_PATHS.has(url.pathname)) return;
  event.respondWith((async () => {
    const cache = await caches.open(CACHE_NAME);
    try {
      const live = await fetch(event.request);
      if (live.ok) await cache.put(event.request, live.clone());
      return live;
    } catch (error) {
      return (await cache.match(event.request)) ||
        (event.request.mode === 'navigate' ? await cache.match('./index.html') : undefined) ||
        Response.error();
    }
  })());
});
