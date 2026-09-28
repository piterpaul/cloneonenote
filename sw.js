/**
 * sw.js — Service Worker para la Aplicación Web Progresiva (PWA) en Android, Chromebook y Mac
 * Estrategia Network-First con limpieza automática de cachés antiguas.
 */

const CACHE_NAME = 'onenote-pwa-cache-v5';
const ASSETS_TO_CACHE = [
  './',
  './index.html',
  './styles.css?v=5',
  './app.js?v=5',
  './drive.js?v=5',
  './manifest.json',
  './icon-192.png',
  './icon-512.png',
  './favicon.ico'
];

self.addEventListener('install', (event) => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(ASSETS_TO_CACHE);
    })
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys.map((key) => {
          if (key !== CACHE_NAME) {
            return caches.delete(key);
          }
        })
      )
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET') return;

  const url = new URL(event.request.url);
  if (
    url.hostname.includes('googleapis.com') ||
    url.hostname.includes('accounts.google.com') ||
    url.pathname.startsWith('/api/')
  ) {
    return;
  }

  event.respondWith(
    fetch(event.request, { cache: 'no-store' })
      .then((networkResponse) => {
        if (networkResponse && networkResponse.status === 200) {
          const cloned = networkResponse.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, cloned));
        }
        return networkResponse;
      })
      .catch(() =>
        caches
          .match(event.request, { ignoreSearch: true })
          .then((cached) => cached || caches.match('./index.html'))
      )
  );
});
