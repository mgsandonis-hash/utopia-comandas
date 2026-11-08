/* Utopía Comandas · funcionamiento sin conexión
   Guarda la app en el dispositivo para que abra aunque no haya cobertura.
   Los datos van aparte: la app los guarda en el móvil y los envía al volver la conexión. */
const VERSION = "utopia-v53";
const SHELL = [
  "./", "./index.html", "./store.js", "./config.js", "./informe.js", "./manifest.webmanifest",
  "./icons/icon-192.png", "./icons/icon-512.png", "./icons/apple-touch-icon.png",
  "https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2.45.4/dist/umd/supabase.min.js"
];

self.addEventListener("install", e => {
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== VERSION).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});

self.addEventListener("fetch", e => {
  const req = e.request;
  if (req.method !== "GET") return;
  const url = new URL(req.url);
  // La base de datos siempre va a internet
  if (url.hostname.endsWith("supabase.co") || url.hostname.endsWith("supabase.in")) return;
  // La página y los archivos de la app: primero internet (para tener la última versión), si no, lo guardado
  if (req.mode === "navigate" || (url.origin === location.origin && /\.(html|js|webmanifest)$|\/$/.test(url.pathname))) {
    e.respondWith(
      fetch(req, {cache: "no-cache"}).then(res => { const copy = res.clone(); caches.open(VERSION).then(c => c.put(req, copy)); return res; })
        .catch(() => caches.match(req).then(r => r || caches.match("./index.html")))
    );
    return;
  }
  // Iconos, fuentes y librerías: lo guardado al momento y se actualiza por detrás
  e.respondWith(
    caches.match(req).then(cached => {
      const net = fetch(req).then(res => { if (res && (res.ok || res.type === "opaque")) { const copy = res.clone(); caches.open(VERSION).then(c => c.put(req, copy)); } return res; }).catch(() => cached);
      return cached || net;
    })
  );
});
