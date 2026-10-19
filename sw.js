const CACHE = 'ledger-v16';
const ASSETS = [
  './',
  './index.html',
  './manifest.webmanifest',
  './icons/icon-192.png',
  './icons/icon-512.png'
];

self.addEventListener('install', e => {
  // cache:'reload' 繞過 HTTP 快取，避免抓到 CDN 上的舊檔
  e.waitUntil(
    caches.open(CACHE).then(c =>
      Promise.all(ASSETS.map(u =>
        fetch(new Request(u, { cache: 'reload' }))
          .then(r => { if (r && r.ok) return c.put(u, r) })
          .catch(() => {})
      ))
    ).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
      // 新版本接手後把已開著的頁面重載一次，避免還停在舊畫面
      .then(() => self.clients.matchAll({ type: 'window' }))
      .then(list => list.forEach(c => { try { c.navigate(c.url) } catch (_) {} }))
  );
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  // 只處理自家資源，Apps Script 的請求一律走網路
  if (new URL(req.url).origin !== location.origin) return;

  const isDoc = req.mode === 'navigate' ||
                req.destination === 'document' ||
                /(\/|index\.html)$/.test(new URL(req.url).pathname);

  // 頁面本體：先網路、拿不到才用快取（改版一定拿得到最新）
  if (isDoc) {
    e.respondWith(
      fetch(req, { cache: 'reload' })
        .then(res => {
          if (res && res.ok) {
            const copy = res.clone();
            caches.open(CACHE).then(c => c.put(req, copy)).catch(() => {});
          }
          return res;
        })
        .catch(() => caches.match(req).then(hit => hit || caches.match('./index.html')))
    );
    return;
  }

  // 圖示、manifest 等：快取優先，背景更新
  e.respondWith(
    caches.match(req).then(hit => {
      if (hit) {
        fetch(req).then(res => {
          if (res && res.ok) caches.open(CACHE).then(c => c.put(req, res)).catch(() => {});
        }).catch(() => {});
        return hit;
      }
      return fetch(req).then(res => {
        if (res && res.ok) {
          const copy = res.clone();
          caches.open(CACHE).then(c => c.put(req, copy)).catch(() => {});
        }
        return res;
      }).catch(() => caches.match('./index.html'));
    })
  );
});
