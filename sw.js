/* Service Worker for 우성아이디 통합 현장 대시보드 (PWA) */
const CACHE_NAME = 'ws-dashboard-pwa-v12';

const PRECACHE_URLS = [
  './',
  './index.html',
  './client.html',
  './manifest.json',
  './data/dashboard_data.js',
  './icon-192.png',
  './icon-512.png',
  './apple-touch-icon.png',
  './partner_icon_640x640.png'
];

// 1. 설치(Install): 핵심 정적 자산 사전 캐싱 및 즉시 활성화
self.addEventListener('install', function(event) {
  event.waitUntil(
    caches.open(CACHE_NAME).then(function(cache) {
      return cache.addAll(PRECACHE_URLS);
    }).then(function() {
      return self.skipWaiting();
    }).catch(function(err) {
      console.warn('[SW] Pre-cache failed during install:', err);
    })
  );
});

// 2. 활성화(Activate): 이전 버전 캐시 정리 및 즉시 클라이언트 제어 (skipWaiting + claim)
self.addEventListener('activate', function(event) {
  event.waitUntil(
    caches.keys().then(function(cacheNames) {
      return Promise.all(
        cacheNames.map(function(name) {
          if (name !== CACHE_NAME) {
            return caches.delete(name);
          }
        })
      );
    }).then(function() {
      return self.clients.claim();
    })
  );
});

// 3. 네트워크 요청 처리(Fetch): 4초 타임아웃 Network-First 전략 (지하/음영지역 Lie-Fi 즉시 캐시 폴백)
self.addEventListener('fetch', function(event) {
  var request = event.request;

  // GET 요청만 캐싱 및 처리 (POST/PUT/DELETE 등은 무조건 브라우저 기본 네트워크 통과)
  if (request.method !== 'GET') {
    return;
  }

  if (request.cache === 'only-if-cached' && request.mode !== 'same-origin') {
    return;
  }

  var url;
  try {
    url = new URL(request.url);
  } catch (e) {
    return;
  }

  // http/https 스킴 외(chrome-extension 등)는 처리 제외
  if (!url.protocol.startsWith('http')) {
    return;
  }

  // Google Apps Script, Supabase 및 외부 API, 타 도메인은 서비스워커에서 가로채지 않고 100% 네트워크 직접 통과
  if (url.origin !== self.location.origin ||
      url.hostname.includes('script.google.com') ||
      url.hostname.includes('googleusercontent.com') ||
      url.hostname.includes('googleapis.com') ||
      url.hostname.includes('supabase.co')) {
    return;
  }

  // 동일 출처(Same-Origin) 정적 자산 및 HTML 요청: 4000ms AbortController 타임아웃 Network-First
  event.respondWith((function() {
    var controller = typeof AbortController === 'function' ? new AbortController() : null;
    var timeoutId = controller
      ? setTimeout(function() { controller.abort(); }, 4000)
      : null;

    var fetchOptions = {};
    if (controller) {
      fetchOptions.signal = controller.signal;
    }
    if (request.mode === 'navigate') {
      fetchOptions.cache = 'no-cache';
    }

    return fetch(request, fetchOptions)
      .then(function(networkResponse) {
        if (timeoutId) clearTimeout(timeoutId);
        // 유효한 정상 응답(200)이면 캐시를 최신 상태로 갱신 (배포 즉시 반영)
        if (networkResponse && networkResponse.status === 200 && networkResponse.type === 'basic') {
          var responseToCache = networkResponse.clone();
          caches.open(CACHE_NAME).then(function(cache) {
            cache.put(request, responseToCache).catch(function() {});
          });
        }
        return networkResponse;
      })
      .catch(function() {
        if (timeoutId) clearTimeout(timeoutId);
        // 오프라인이거나 4초 네트워크 지연(Lie-Fi) 시 캐시에서 즉시 반환
        return caches.match(request, { ignoreSearch: true }).then(function(cachedResponse) {
          if (cachedResponse) {
            return cachedResponse;
          }
          // 캐시 미스 시 네비게이션 요청 폴백 (고객용 client.html은 절대 내부 index.html로 노출하지 않음)
          if (request.mode === 'navigate') {
            if (url.pathname.endsWith('/client.html')) {
              return caches.match('./client.html');
            }
            return caches.match('./index.html').then(function(idxRes) {
              return idxRes || caches.match('./');
            });
          }
          return new Response('Network error (offline)', {
            status: 503,
            statusText: 'Service Unavailable',
            headers: new Headers({ 'Content-Type': 'text/plain; charset=utf-8' })
          });
        });
      });
  })());
});
