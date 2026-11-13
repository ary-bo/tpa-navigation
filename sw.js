'use strict';

/* TPA Navigation — Service Worker
   Cache version: v9
   This version includes the app icon referenced by index.html and manifest. */

const CACHE_PREFIX = 'tpa-navigation-';
const SHELL_CACHE = `${CACHE_PREFIX}shell-v10`;
const RUNTIME_CACHE = `${CACHE_PREFIX}runtime-v10`;

const APP_SHELL = [
  './',
  './index.html',
  './style.css',
  './app.js',
  './manifest.webmanifest',
  './icon.svg'
];

self.addEventListener('install', event => {
  event.waitUntil((async () => {
    const cache = await caches.open(SHELL_CACHE);
    await cache.addAll(APP_SHELL);
    await self.skipWaiting();
  })());
});

self.addEventListener('activate', event => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(
      keys
        .filter(key => key.startsWith(CACHE_PREFIX))
        .filter(key => key !== SHELL_CACHE && key !== RUNTIME_CACHE)
        .map(key => caches.delete(key))
    );
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', event => {
  const request = event.request;
  if (request.method !== 'GET') return;

  const url = new URL(request.url);

  if (url.origin === self.location.origin) {
    if (request.mode === 'navigate') {
      event.respondWith(
        fetch(request)
          .then(response => {
            if (response && response.ok) {
              const copy = response.clone();
              caches.open(SHELL_CACHE).then(cache => cache.put('./index.html', copy));
            }
            return response;
          })
          .catch(async () => {
            const cached = await caches.match('./index.html');
            return cached || Response.error();
          })
      );
      return;
    }

    event.respondWith(
      caches.match(request).then(cached => {
        if (cached) return cached;
        return fetch(request).then(response => {
          if (response && response.ok) {
            const copy = response.clone();
            caches.open(SHELL_CACHE).then(cache => cache.put(request, copy));
          }
          return response;
        });
      })
    );
    return;
  }

  if ([
    'cdn.jsdelivr.net',
    'unpkg.com',
    'accounts.google.com'
  ].includes(url.hostname)) {
    event.respondWith((async () => {
      const cache = await caches.open(RUNTIME_CACHE);
      try {
        const response = await fetch(request);
        if (response && response.ok) {
          await cache.put(request, response.clone());
        }
        return response;
      } catch (error) {
        const cached = await cache.match(request);
        if (cached) return cached;
        throw error;
      }
    })());
  }
});
