/* Service worker: caches the app shell and data so the app works offline.
 * Bump VERSION whenever you change app files so phones pick up the update. */
var VERSION = "h2r-v2";
var SHELL = [
  "./",
  "index.html",
  "styles.css",
  "app.js",
  "manifest.webmanifest",
  "icons/icon.svg",
  "icons/icon-180.png",
  "icons/icon-192.png",
  "icons/icon-512.png",
  "data/h2r_data.json"
];

self.addEventListener("install", function (event) {
  event.waitUntil(caches.open(VERSION).then(function (c) { return c.addAll(SHELL); }));
});

self.addEventListener("activate", function (event) {
  event.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(keys.filter(function (k) { return k !== VERSION; }).map(function (k) { return caches.delete(k); }));
    }).then(function () { return self.clients.claim(); })
  );
});

self.addEventListener("fetch", function (event) {
  var req = event.request;
  if (req.method !== "GET" || new URL(req.url).origin !== self.location.origin) return;

  // Data: network first so new survey rounds show up, cache as fallback.
  if (req.url.indexOf("/data/") !== -1) {
    event.respondWith(
      fetch(req).then(function (res) {
        var copy = res.clone();
        caches.open(VERSION).then(function (c) { c.put(req, copy); });
        return res;
      }).catch(function () { return caches.match(req, { ignoreSearch: true }); })
    );
    return;
  }

  // App shell: cache first, then network.
  event.respondWith(
    caches.match(req, { ignoreSearch: true }).then(function (hit) {
      return hit || fetch(req).then(function (res) {
        if (res.ok) {
          var copy = res.clone();
          caches.open(VERSION).then(function (c) { c.put(req, copy); });
        }
        return res;
      }).catch(function () {
        if (req.mode === "navigate") return caches.match("index.html");
      });
    })
  );
});
