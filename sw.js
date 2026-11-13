/* TPA Navigation - resilient PWA cache */
const CACHE_NAME = "tpa-navigation-v5";
const APP_ASSETS = ["./", "./index.html", "./manifest.webmanifest"];
const LIBRARY_HOSTS = new Set(["unpkg.com", "cdn.jsdelivr.net"]);

self.addEventListener("install", event => {
  event.waitUntil((async () => {
    const cache = await caches.open(CACHE_NAME);
    // A missing optional shell URL must not prevent the Service Worker installing.
    await Promise.allSettled(APP_ASSETS.map(async asset => {
      try {
        const response = await fetch(asset, { cache: "reload" });
        if (response.ok) await cache.put(asset, response);
      } catch (error) {
        // The browser may still have an older cached shell to use.
      }
    }));
    await self.skipWaiting();
  })());
});

self.addEventListener("activate", event => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys
      .filter(key => key.startsWith("tpa-navigation-") && key !== CACHE_NAME)
      .map(key => caches.delete(key)));
    await self.clients.claim();
  })());
});

self.addEventListener("fetch", event => {
  const request = event.request;
  if (request.method !== "GET") return;

  const url = new URL(request.url);

  // App files: network first, cached copy when offline.
  if (url.origin === self.location.origin) {
    event.respondWith((async () => {
      try {
        const response = await fetch(request);
        if (response && response.ok) {
          const cache = await caches.open(CACHE_NAME);
          await cache.put(request, response.clone());
        }
        return response;
      } catch (error) {
        const cached = await caches.match(request);
        if (cached) return cached;
        if (request.mode === "navigate") {
          return (await caches.match("./index.html")) || Response.error();
        }
        return Response.error();
      }
    })());
    return;
  }

  // CDN libraries: cache-first for faster repeat loads, then network.
  if (LIBRARY_HOSTS.has(url.hostname)) {
    event.respondWith((async () => {
      const cached = await caches.match(request);
      if (cached) {
        // Refresh the cache in the background when possible.
        fetch(request).then(async response => {
          if (response && response.ok) {
            const cache = await caches.open(CACHE_NAME);
            await cache.put(request, response);
          }
        }).catch(() => {});
        return cached;
      }
      try {
        const response = await fetch(request);
        if (response && response.ok) {
          const cache = await caches.open(CACHE_NAME);
          await cache.put(request, response.clone());
        }
        return response;
      } catch (error) {
        return Response.error();
      }
    })());
  }

  // Google OAuth/Drive and map tiles remain online services and are not cached.
});
