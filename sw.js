// Rede primeiro para a página (dados do dia), cache como reserva para abrir offline.
const VERSAO = "meu-rpps-v2";
self.addEventListener("install", (e) => { self.skipWaiting(); });
self.addEventListener("activate", (e) => {
  e.waitUntil(caches.keys().then((ks) => Promise.all(ks.filter((k) => k !== VERSAO).map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener("fetch", (e) => {
  if (e.request.method !== "GET" || new URL(e.request.url).origin !== location.origin) return;
  e.respondWith(fetch(e.request).then((r) => { const cp = r.clone(); caches.open(VERSAO).then((c) => c.put(e.request, cp)); return r; })
    .catch(() => caches.match(e.request).then((m) => m || caches.match("./"))));
});
