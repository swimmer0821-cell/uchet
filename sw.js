/* Учёт денег — работа как приложение. Свои файлы берёт из сети, а без сети — из памяти телефона.
   Запросы к Google (Apps Script) не трогает. Папку sklad/ не трогает — у склада свой sw.js. */
const CACHE = 'uchet-v1';
const FILES = ['./', './index.html', './manifest.webmanifest', './icons/icon-192.png', './icons/icon-512.png', './icons/apple-touch-icon.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k.startsWith('uchet-') && k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET' || url.origin !== location.origin) return;
  const base = new URL('./', self.registration.scope).pathname;
  if (url.pathname.startsWith(base + 'sklad/')) return;
  // Сначала сеть (чтобы обновления приходили сразу), без сети — сохранённая копия
  e.respondWith(fetch(e.request).then(res => {
    if (res.ok) { const copy = res.clone(); caches.open(CACHE).then(c => c.put(e.request, copy)); }
    return res;
  }).catch(() => caches.match(e.request).then(r => r || caches.match('./index.html'))));
});
