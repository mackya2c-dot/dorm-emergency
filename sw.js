/* 寮 緊急時対応ガイド Service Worker
 * アプリ本体（HTML・CSS・JS・アイコン）を変更したら、CACHE の番号を必ず上げること。
 * データ（data/）は GAS からの公開で更新され、ここを変える必要はありません。 */
const CACHE = 'emg-shell-v1';
const SHELL = [
  './', './index.html', './style.css', './app.js', './manifest.webmanifest',
  './icons/icon-192.png', './icons/icon-512.png', './icons/apple-touch-icon.png'
];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)));
});

self.addEventListener('activate', (e) => {
  e.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.filter((k) => k.startsWith('emg-') && k !== CACHE && k !== 'emg-data').map((k) => caches.delete(k)));
    await self.clients.claim();
  })());
});

self.addEventListener('message', (e) => {
  if (e.data === 'SKIP_WAITING') self.skipWaiting();
});

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return;

  // データ：通信を優先し、失敗したら保存済みのものを返す
  if (url.pathname.includes('/data/')) {
    e.respondWith((async () => {
      try {
        const res = await fetch(req);
        if (res.ok) {
          const c = await caches.open('emg-data');
          await c.put(url.origin + url.pathname, res.clone());
        }
        return res;
      } catch (err) {
        const hit = await caches.match(url.origin + url.pathname, { cacheName: 'emg-data' });
        return hit || Response.error();
      }
    })());
    return;
  }

  // ページ遷移：保存済みの index.html を優先
  if (req.mode === 'navigate') {
    e.respondWith((async () => {
      const hit = await caches.match('./index.html', { cacheName: CACHE });
      if (hit) return hit;
      try { return await fetch(req); } catch (err) { return Response.error(); }
    })());
    return;
  }

  // アプリ本体：保存済みを優先
  e.respondWith((async () => {
    const hit = await caches.match(req, { ignoreSearch: true, cacheName: CACHE });
    if (hit) return hit;
    try {
      const res = await fetch(req);
      if (res.ok) (await caches.open(CACHE)).put(req, res.clone());
      return res;
    } catch (err) {
      return Response.error();
    }
  })());
});
