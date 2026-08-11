// TaskDeck service worker
// アプリ本体（HTML/manifest/アイコン）をキャッシュし、オフラインでも起動できるようにする。
// データ自体は localStorage に保存されるため、このSWはキャッシュ更新のみを担当する。

var CACHE_NAME = "taskdeck-cache-v2";
var CORE_ASSETS = [
  "./",
  "./index.html",
  "./manifest.json"
];

self.addEventListener("install", function(event){
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME).then(function(cache){
      return cache.addAll(CORE_ASSETS);
    })
  );
});

self.addEventListener("activate", function(event){
  event.waitUntil(
    caches.keys().then(function(keys){
      return Promise.all(
        keys.filter(function(k){ return k !== CACHE_NAME; })
            .map(function(k){ return caches.delete(k); })
      );
    }).then(function(){ return self.clients.claim(); })
  );
});

// Network-first for the HTML shell (so updates show up quickly),
// cache-first for everything else (icons, manifest).
self.addEventListener("fetch", function(event){
  if(event.request.method !== "GET") return;

  var url = new URL(event.request.url);
  if(url.origin !== self.location.origin) return;

  if(event.request.mode === "navigate" || url.pathname.endsWith("/index.html")){
    event.respondWith(
      fetch(event.request)
        .then(function(res){
          var copy = res.clone();
          caches.open(CACHE_NAME).then(function(cache){ cache.put(event.request, copy); });
          return res;
        })
        .catch(function(){ return caches.match(event.request).then(function(r){ return r || caches.match("./index.html"); }); })
    );
    return;
  }

  event.respondWith(
    caches.match(event.request).then(function(cached){
      return cached || fetch(event.request).then(function(res){
        var copy = res.clone();
        caches.open(CACHE_NAME).then(function(cache){ cache.put(event.request, copy); });
        return res;
      });
    })
  );
});
