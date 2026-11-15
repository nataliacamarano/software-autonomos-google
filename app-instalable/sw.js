// Service worker mínimo: guarda la "puerta de entrada" para que abra rápido.
// El programa en sí (Google) siempre se pide en línea; aquí no se guarda nada suyo.
const CACHE = 'alvaro-gestion-v2';
const ARCHIVOS = ['./', 'index.html', 'manifest.webmanifest', 'icon-192.png', 'icon-512.png', 'apple-touch-icon.png'];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(ARCHIVOS)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys()
      .then((ks) => Promise.all(ks.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (e) => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.origin !== self.location.origin) return;   // lo de Google no se toca
  e.respondWith(
    // "no-cache": siempre pregunta a la web si hay una versión nueva (GitHub Pages guarda 10 min por su cuenta).
    fetch(e.request, { cache: 'no-cache' })
      .then((r) => {
        const copia = r.clone();
        caches.open(CACHE).then((c) => c.put(e.request, copia));
        return r;
      })
      .catch(() => caches.match(e.request).then((r) => r || caches.match('index.html')))
  );
});
