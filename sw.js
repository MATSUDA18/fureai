// ふれあいホームみずほ アプリ用 Service Worker
// 役割：常に最新のページを取得（ネットワーク優先）＝更新を自動反映。
// オフライン時のみ、直近に取得したページを表示（フォールバック）。
const CACHE = 'mizuho-cache-v1';

self.addEventListener('install', (e) => {
  self.skipWaiting();
});

self.addEventListener('activate', (e) => {
  e.waitUntil(self.clients.claim());
});

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const accept = req.headers.get('accept') || '';
  const isHTML = req.mode === 'navigate' || accept.includes('text/html');
  if (isHTML) {
    e.respondWith(
      fetch(req)
        .then((res) => {
          const copy = res.clone();
          caches.open(CACHE).then((c) => c.put('./', copy)).catch(() => {});
          return res;
        })
        .catch(() => caches.match('./').then((r) => r || caches.match(req)))
    );
  }
});
