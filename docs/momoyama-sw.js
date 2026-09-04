// 桃山出願材料チェックリスト — 只管 momoyama 相关请求，其他页面一律不拦截
const CACHE = 'momo-v1';
const ASSETS = [
  './momoyama_docs.html',
  './momoyama.webmanifest',
  './momoyama-icon-192.png',
  './momoyama-icon-512.png',
  './momoyama-icon-maskable.png',
  './momoyama-apple-touch.png'
];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

// 网络优先：在线永远拿最新版；断网时用缓存兜底
self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET') return;
  if (url.origin !== location.origin) return;
  if (!url.pathname.includes('momoyama')) return;
  e.respondWith(
    fetch(e.request)
      .then(res => {
        const copy = res.clone();
        caches.open(CACHE).then(c => c.put(e.request, copy)).catch(() => {});
        return res;
      })
      .catch(() => caches.match(e.request).then(r => r || caches.match('./momoyama_docs.html')))
  );
});
