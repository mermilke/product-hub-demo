// Service worker — makes the Hub an installable app (Android WebAPK). Offline it can serve the cached
// shell page, but NOT a working app: the Microsoft sign-in library is cross-origin and never cached, so an
// offline launch shows the sign-in card with a "couldn't load the sign-in library" note. It is
// deliberately MINIMAL and SAFE:
//   • NETWORK-FIRST for same-origin GETs, so a fresh Vercel deploy ALWAYS wins — the cache is only a
//     fallback when the device is offline. There is no stale-code risk.
//   • It NEVER touches auth or data: non-GET requests, every cross-origin request (MSAL login.microsoft*,
//     Microsoft Graph, SharePoint), and /api/* are all ignored (it doesn't call respondWith), so the
//     browser handles them exactly as if no service worker existed.
const CACHE = "ph-shell-v1";
const SHELL = ["/", "/app.css", "/app.js", "/pf-shared.js", "/demo/demo-signin.js", "/pre-init.js", "/site.webmanifest", "/icons/logo-192.png", "/fonts/inter-400.woff2", "/fonts/inter-600.woff2", "/fonts/inter-700.woff2"];

self.addEventListener("install", function (e) {
  e.waitUntil(caches.open(CACHE).then(function (c) { return c.addAll(SHELL); }).catch(function () {}).then(function () { return self.skipWaiting(); }));
});

self.addEventListener("activate", function (e) {
  e.waitUntil(
    caches.keys().then(function (keys) { return Promise.all(keys.filter(function (k) { return k !== CACHE; }).map(function (k) { return caches.delete(k); })); })
      .then(function () { return self.clients.claim(); })
  );
});

self.addEventListener("fetch", function (e) {
  var req = e.request;
  if (req.method !== "GET") return;                         // never intercept writes / auth POSTs
  var url;
  try { url = new URL(req.url); } catch (err) { return; }
  if (url.origin !== self.location.origin) return;          // ignore Graph / MSAL / SharePoint (cross-origin)
  if (url.pathname.indexOf("/api/") === 0) return;          // never cache dynamic API responses
  // Network-first: fresh content always wins; the cache is only used when the network is unavailable.
  // Query-string URLs are never stored: the version nudge probes "/?cb=<n>" and each one would be a new,
  // never-read cache entry.
  var cacheable = !url.search;
  e.respondWith(
    fetch(req).then(function (res) {
      if (cacheable && res && res.ok && res.type === "basic") { var copy = res.clone(); caches.open(CACHE).then(function (c) { c.put(req, copy); }); }
      return res;
    }).catch(function () {
      return caches.match(req).then(function (c) {
        if (c) return c;
        // Only a NAVIGATION gets the cached shell page. Handing index.html to a <script>/<link>/font
        // request just produces a MIME-type error in the console; a clean network error is what the
        // browser would have reported without a service worker.
        return req.mode === "navigate" ? caches.match("/") : Response.error();
      });
    })
  );
});
