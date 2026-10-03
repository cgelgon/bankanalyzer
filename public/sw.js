/* BankAnalyzer — service worker de l'appli (PWA)
   Principe : TOUJOURS le réseau d'abord, pour que chaque mise à jour du site
   arrive immédiatement dans l'appli. Le cache ne sert qu'à afficher une page
   "Pas de réseau" quand le téléphone est hors connexion.
   Les appels au serveur d'analyse (Railway) et les sites externes ne sont jamais touchés. */
const VERSION = 'bankanalyzer-v1';
const HORS_LIGNE = '/offline.html';

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(VERSION).then((c) => c.addAll([HORS_LIGNE])).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys()
      .then((noms) => Promise.all(noms.filter((n) => n !== VERSION).map((n) => caches.delete(n))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;
  if (new URL(req.url).origin !== self.location.origin) return;
  if (req.mode === 'navigate') {
    e.respondWith(fetch(req).catch(() => caches.match(HORS_LIGNE)));
  }
});
