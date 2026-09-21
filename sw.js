// Service worker: o app abre sem internet e recebe atualizações quando houver rede.
const VERSAO = '1.32.0';
const CACHE = 'pesquisa-maxipark-' + VERSAO;
const FILES = ['./', './index.html', './manifest.webmanifest', './icon-192.png', './icon-512.png', './equipe/', './equipe/index.html', './equipe/manifest.webmanifest', './cliente/', './cliente/index.html', './cliente/manifest.webmanifest', './cliente/qr.html', './conselho/', './conselho/index.html', './conselho/manifest.webmanifest', './painel/', './painel/index.html', './painel/manifest.webmanifest'];   // inclui a pesquisa com clientes (pasta cliente/)

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET' || !req.url.startsWith(self.location.origin)) return;
  const isPage = req.mode === 'navigate' || req.url.endsWith('/') || req.url.endsWith('index.html') || req.url.endsWith('sw.js');
  if (isPage) {
    // rede primeiro (pega a versão nova), cache se estiver offline
    e.respondWith(fetch(req).then(res => { if (res && res.ok) { const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)); } return res; })
      .catch(() => caches.match(req, { ignoreSearch: true }).then(hit => hit || caches.match('./index.html'))));
  } else {
    // ícones e manifest: cache primeiro
    e.respondWith(caches.match(req, { ignoreSearch: true }).then(hit => hit || fetch(req).then(res => { if (res && res.ok) { const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)); } return res; })));
  }
});
