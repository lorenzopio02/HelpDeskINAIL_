// Salva l'app sul telefono per farla funzionare anche senza campo.
const CACHE = "helpdesk-v27";
const FILES = ["./", "index.html", "manifest.webmanifest", "logo_header.png", "icon-192.png", "icon-512.png", "apple-touch-icon.png", "export.js", "verbali.js", "stanze.js", "bolle.js", "toscana.js", "barcode.js", "qr.js", "scan.js", "corridoio.js", "stampante-c250i.png", "stampante-c3350i.png", "disp-lap13.png", "disp-lap15.png", "disp-s24.png", "disp-s24plus.png", "disp-s25.png", "disp-teams.png"];
self.addEventListener("install", e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES)).then(() => self.skipWaiting())); });
self.addEventListener("activate", e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
// Prima prova la rete (così gli aggiornamenti arrivano), se non c'è campo usa la copia salvata.
self.addEventListener("fetch", e => {
  if (e.request.method !== "GET") return;
  e.respondWith(fetch(e.request).then(r => {
    const copy = r.clone(); caches.open(CACHE).then(c => c.put(e.request, copy)); return r;
  }).catch(() => caches.match(e.request).then(m => m || caches.match("./"))));
});
