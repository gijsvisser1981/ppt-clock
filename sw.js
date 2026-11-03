// Network-first met cache-fallback: updates komen gewoon door,
// maar zonder internet (bv. op locatie) laadt de klok uit de cache.
var CACHE = "avpclock-v1";
var FILES = [
  "clock.html",
  "clock.css",
  "clock.js",
  "https://appsforoffice.microsoft.com/lib/1/hosted/office.js"
];

self.addEventListener("install", function (e) {
  e.waitUntil(
    caches.open(CACHE).then(function (c) {
      return Promise.all(FILES.map(function (f) {
        return c.add(new Request(f, { mode: f.indexOf("http") === 0 ? "no-cors" : "same-origin" }))
          .catch(function () { /* één mislukt bestand mag de install niet blokkeren */ });
      }));
    }).then(function () { return self.skipWaiting(); })
  );
});

self.addEventListener("activate", function (e) {
  e.waitUntil(self.clients.claim());
});

self.addEventListener("fetch", function (e) {
  if (e.request.method !== "GET") return;
  e.respondWith(
    fetch(e.request).then(function (res) {
      var copy = res.clone();
      caches.open(CACHE).then(function (c) { c.put(e.request, copy); });
      return res;
    }).catch(function () {
      return caches.match(e.request, { ignoreSearch: true });
    })
  );
});
