// 士林電機 FA 業務戰情與報價系統 - Service Worker
const CACHE_NAME = 'shihlin-fa-pwa-v1.0.30';
const ASSETS_TO_CACHE = [
  './',
  './index.html',
  './css/style.css',
  './js/app.js',
  './manifest.json',
  './icon-192x192-v2.png',
  './icon-512x512-v2.png',
  './apple-touch-icon-v2.png',
  './favicon-v2.png'
];

// 安裝 Service Worker 並快取核心靜態資源
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(ASSETS_TO_CACHE).catch((err) => {
        console.warn('[SW] 快取部分資源失敗:', err);
      });
    })
  );
  self.skipWaiting();
});

// 啟動 Service Worker 並自動清理舊版本快取
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((cacheNames) => {
      return Promise.all(
        cacheNames.map((name) => {
          if (name !== CACHE_NAME) {
            console.log('[SW] 清理舊快取:', name);
            return caches.delete(name);
          }
        })
      );
    })
  );
  self.clients.claim();
});

// 網路請求攔截策略：
// 1. Google 試算表 (GAS) / Google 登入 / 本地 API / POST 請求：直接放行 (Network Only)，確保資料即時
// 2. HTML 主頁面與核心 JS/CSS：聯網優先 (Network First)，確保隨時載入最新功能
// 3. 圖示等靜態素材：快取優先 (Stale-While-Revalidate)
self.addEventListener('fetch', (event) => {
  const url = event.request.url;

  // 1. 外部 API 與 POST 請求直接放行
  if (
    url.includes('script.google.com') ||
    url.includes('accounts.google.com') ||
    url.includes('googleapis.com') ||
    url.includes('/api/') ||
    event.request.method === 'POST'
  ) {
    return;
  }

  const isNavigate = event.request.mode === 'navigate';
  const isHtml = event.request.headers.get('accept')?.includes('text/html');
  const isIndex = url.endsWith('/') || url.includes('index.html');
  const isCode = url.includes('.js') || url.includes('.css');

  // 2. HTML 與程式碼採用 Network First (聯網優先，斷網才用快取)
  if (isNavigate || isHtml || isIndex || isCode) {
    event.respondWith(
      fetch(event.request)
        .then((networkResponse) => {
          if (networkResponse && networkResponse.status === 200) {
            const responseToCache = networkResponse.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(event.request, responseToCache));
          }
          return networkResponse;
        })
        .catch(() => caches.match(event.request))
    );
    return;
  }

  // 3. 其他靜態資源採用 Stale-While-Revalidate
  event.respondWith(
    caches.match(event.request).then((cachedResponse) => {
      const fetchPromise = fetch(event.request)
        .then((networkResponse) => {
          if (networkResponse && networkResponse.status === 200 && networkResponse.type === 'basic') {
            const responseToCache = networkResponse.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(event.request, responseToCache));
          }
          return networkResponse;
        })
        .catch((err) => {
          console.warn('[SW] 背景擷取資源失敗:', err);
        });

      return cachedResponse || fetchPromise;
    })
  );
});
