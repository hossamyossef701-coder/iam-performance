// Network-first: the dashboard must never show yesterday's numbers from a cache. The cache
// exists only so the app still opens when the phone is offline.
const CACHE = "iam-perf-v3";
const SHELL = ["./index.html", "./report.html", "./manifest.webmanifest", "./report.webmanifest",
  "./icon-192.png", "./icon-512.png", "./icon-maskable-512.png"];
const CACHE_BUST = "logo";   // bumping this retires the cache holding the old icons

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()).catch(() => self.skipWaiting()));
});
self.addEventListener("activate", (e) => {
  e.waitUntil(caches.keys().then((ks) => Promise.all(ks.filter((k) => k !== CACHE).map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener("fetch", (e) => {
  const req = e.request;
  if (req.method !== "GET") return;
  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return;                 // CDN scripts and the GitHub API go straight to the network
  if (url.pathname.endsWith("live.enc.json")) return;              // always fresh, never cached
  if (req.mode === "navigate") {                                   // the page itself: network first, cache only as a backstop
    e.respondWith(fetch(req).then((res) => {
      if (res && res.ok) { const copy = res.clone(); caches.open(CACHE).then((c) => c.put(req, copy)).catch(() => {}); }
      return res;
    }).catch(() => caches.match(req, { ignoreSearch: true }).then((hit) => hit || Response.error())));
    return;
  }
  e.respondWith(
    fetch(req).then((res) => {
      if (res && res.ok) { const copy = res.clone(); caches.open(CACHE).then((c) => c.put(req, copy)).catch(() => {}); }
      return res;
    }).catch(() => caches.match(req, { ignoreSearch: true }).then((hit) => hit || Response.error()))
  );
});
