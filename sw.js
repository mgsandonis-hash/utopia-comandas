/* Utopía Coffee Club se ha mudado a https://utopiacoffeeclub.github.io/ : este service worker se borra solo */
self.addEventListener("install", () => self.skipWaiting());
self.addEventListener("activate", e => {
  e.waitUntil(caches.keys().then(k => Promise.all(k.map(x => caches.delete(x))))
    .then(() => self.registration.unregister())
    .then(() => self.clients.matchAll({type:"window"}))
    .then(cs => cs.forEach(c => c.navigate(c.url))));
});
