// Chaque appli ne nettoie que SES anciens caches : les 3 applis partagent la même adresse GitHub.
const PREFIXE = 'jauge-devlin-v';
const CACHE = PREFIXE + '87';
const ASSETS = ['./', './index.html', './manifest.json', './icon-192.png', './icon-512.png', './icon-maskable-512.png', './moteur.js', './sauvegarde.js', './apps-script/Code.js'];
// (9.27) Installation : on va chercher les fichiers frais sur le serveur (pas une copie du navigateur vieille de quelques minutes)
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS.map(u => new Request(u, { cache: 'reload' }))))); self.skipWaiting(); });
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(k => Promise.all(k.filter(x => x !== CACHE && x.startsWith(PREFIXE)).map(x => caches.delete(x)))));
  self.clients.claim();
});
// Ouvre instantanément depuis le cache, puis met à jour en arrière-plan.
// Hors ligne, la page d'accueil ne sert de repli qu'à l'ouverture d'une page (pas à une icône ou au manifeste).
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.origin !== location.origin) return;
  if (u.searchParams.has('maj')) return; // (9.47) contrôle « nouvelle version ? » : toujours sur le serveur, jamais depuis le cache
  e.respondWith(caches.open(CACHE).then(async c => {
    const hit = await c.match(e.request, { ignoreSearch: true });
    const net = fetch(e.request).then(r => { if (r.ok) c.put(e.request, r.clone()); return r; }).catch(() => hit || (e.request.mode === 'navigate' ? caches.match('./index.html') : Response.error()));
    return hit || net;
  }));
});
