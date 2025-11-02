// Service Worker 版本號
const CACHE_VERSION = 'v1';
const CACHE_NAME = `妙華蓮華經-${CACHE_VERSION}`;

// 需要預先快取的資源
const STATIC_CACHE_URLS = [
  '/',
  '/index.html'
];

// 外部資源（CDN 等），需要動態快取
const EXTERNAL_RESOURCES = [
  'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css',
  'https://img.freepik.com/free-photo/old-paper-texture_1194-5415.jpg'
];

// 安裝 Service Worker
self.addEventListener('install', function(event) {
  console.log('[Service Worker] 安裝中...');
  event.waitUntil(
    caches.open(CACHE_NAME).then(function(cache) {
      console.log('[Service Worker] 快取靜態資源');
      // 只快取基本的 HTML 文件
      return cache.addAll(STATIC_CACHE_URLS).catch(function(error) {
        console.error('[Service Worker] 部分資源快取失敗:', error);
      });
    })
  );
  // 強制激活新的 Service Worker
  self.skipWaiting();
});

// 激活 Service Worker
self.addEventListener('activate', function(event) {
  console.log('[Service Worker] 激活中...');
  event.waitUntil(
    caches.keys().then(function(cacheNames) {
      return Promise.all(
        cacheNames.map(function(cacheName) {
          // 刪除舊版本的快取
          if (cacheName !== CACHE_NAME) {
            console.log('[Service Worker] 刪除舊快取:', cacheName);
            return caches.delete(cacheName);
          }
        })
      );
    }).then(function() {
      // 立即控制所有頁面
      return self.clients.claim();
    })
  );
});

// 攔截網路請求
self.addEventListener('fetch', function(event) {
  const url = new URL(event.request.url);
  
  // 只處理 GET 請求
  if (event.request.method !== 'GET') {
    return;
  }

  // 對於 HTML 頁面，使用快取優先策略
  if (url.pathname === '/' || url.pathname === '/index.html' || url.pathname.endsWith('.html')) {
    event.respondWith(
      caches.match(event.request).then(function(cachedResponse) {
        // 如果快取中有，先返回快取版本，同時在背景更新
        if (cachedResponse) {
          // 在背景嘗試更新快取
          fetch(event.request).then(function(networkResponse) {
            if (networkResponse.ok) {
              caches.open(CACHE_NAME).then(function(cache) {
                cache.put(event.request, networkResponse.clone());
              });
            }
          }).catch(function() {
            // 網路請求失敗時忽略錯誤
          });
          return cachedResponse;
        }
        // 如果快取中沒有，從網路獲取
        return fetch(event.request).then(function(networkResponse) {
          if (networkResponse.ok) {
            const responseClone = networkResponse.clone();
            caches.open(CACHE_NAME).then(function(cache) {
              cache.put(event.request, responseClone);
            });
          }
          return networkResponse;
        }).catch(function(error) {
          console.error('[Service Worker] 網路請求失敗:', error);
          // 如果網路失敗且快取也沒有，返回基本的離線頁面（可選）
          return new Response('離線狀態，無法載入頁面', {
            status: 503,
            statusText: 'Service Unavailable',
            headers: new Headers({
              'Content-Type': 'text/html; charset=utf-8'
            })
          });
        });
      })
    );
    return;
  }

  // 對於 CSS、JS、圖片等資源，使用網路優先策略
  if (url.pathname.match(/\.(css|js|png|jpg|jpeg|gif|svg|ico|woff|woff2|ttf|eot)$/i)) {
    event.respondWith(
      fetch(event.request).then(function(networkResponse) {
        // 如果網路請求成功，快取響應
        if (networkResponse.ok) {
          const responseClone = networkResponse.clone();
          caches.open(CACHE_NAME).then(function(cache) {
            cache.put(event.request, responseClone);
          });
        }
        return networkResponse;
      }).catch(function() {
        // 如果網路請求失敗，嘗試從快取中獲取
        return caches.match(event.request).then(function(cachedResponse) {
          if (cachedResponse) {
            return cachedResponse;
          }
          // 如果快取也沒有，返回 404
          return new Response('資源未找到', {
            status: 404,
            statusText: 'Not Found'
          });
        });
      })
    );
    return;
  }

  // 對於外部資源（CDN），使用網路優先策略
  if (EXTERNAL_RESOURCES.some(function(resource) {
    return event.request.url.startsWith(resource.split('?')[0]);
  })) {
    event.respondWith(
      fetch(event.request).then(function(networkResponse) {
        if (networkResponse.ok) {
          const responseClone = networkResponse.clone();
          caches.open(CACHE_NAME).then(function(cache) {
            cache.put(event.request, responseClone);
          });
        }
        return networkResponse;
      }).catch(function() {
        return caches.match(event.request).catch(function() {
          return new Response('資源載入失敗', {
            status: 503,
            statusText: 'Service Unavailable'
          });
        });
      })
    );
    return;
  }

  // 對於其他請求（如 API 請求），使用網路優先，失敗時不返回快取
  event.respondWith(
    fetch(event.request).catch(function() {
      // 對於 API 等動態內容，不從快取返回
      return new Response('網路連接失敗', {
        status: 503,
        statusText: 'Service Unavailable'
      });
    })
  );
});

