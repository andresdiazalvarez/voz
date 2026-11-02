const CACHE_NAME = "edificios-v24-datos-por-registro";
const ASSETS = [
  "./",
  "./index.html",
  "./styles.css",
  "./app.js",
  "./seed-data.js",
  "./manifest.webmanifest",
  "./icon.svg",
  "./exceljs.min.js"
];

self.addEventListener("install", (event) => {
  event.waitUntil(caches.open(CACHE_NAME).then((cache) => cache.addAll(ASSETS.map((url) => new Request(url, { cache: "reload" })))));
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key)))
    )
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  if (event.request.method !== "GET") return;
  const url = new URL(event.request.url);
  if (url.origin !== self.location.origin) return;
  event.respondWith((async () => {
    const cache = await caches.open(CACHE_NAME);
    try {
      const response = await fetch(event.request, { cache: "no-cache" });
      if (response.ok) {
        await cache.put(event.request, response.clone());
        return response;
      }
      return (await cache.match(event.request)) || response;
    } catch (error) {
      const cached = await cache.match(event.request);
      if (cached) return cached;
      throw error;
    }
  })());
});
