/* Service Worker for 우성아이디 통합 현장 대시보드 (PWA) */
const CACHE_NAME = 'ws-dashboard-pwa-v1';

const PRECACHE_URLS = [
  './',
  './v2.html',
  './index.html',
  './manifest.json',
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

// 3. 네트워크 요청 처리(Fetch): Network-First 전략 (오프라인 시 캐시 폴백)
self.addEventListener('fetch', function(event) {
  var request = event.request;

  // GET 요청만 캐싱 및 처리 (POST/PUT/DELETE 등은 무조건 브라우저 기본 네트워크 통과)
  if (request.method !== 'GET') {
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

  // Google Apps Script 및 외부 API, 타 도메인은 서비스워커에서 가로채지 않고 100% 네트워크 직접 통과
  if (url.origin !== self.location.origin ||
      url.hostname.includes('script.google.com') ||
      url.hostname.includes('googleusercontent.com') ||
      url.hostname.includes('googleapis.com')) {
    return;
  }

  // 동일 출처(Same-Origin) 정적 자산 및 HTML 요청: Network-First 전략
  event.respondWith(
    fetch(request)
      .then(function(networkResponse) {
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
        // 오프라인이거나 네트워크 오류 시 캐시에서 탐색
        return caches.match(request, { ignoreSearch: true }).then(function(cachedResponse) {
          if (cachedResponse) {
            return cachedResponse;
          }
          // 캐시 미스 시 네비게이션 요청인 경우 v2.html 또는 index.html 오프라인 폴백
          if (request.mode === 'navigate') {
            return caches.match('./v2.html').then(function(v2Res) {
              return v2Res || caches.match('./index.html');
            });
          }
          return new Response('Network error (offline)', {
            status: 503,
            statusText: 'Service Unavailable',
            headers: new Headers({ 'Content-Type': 'text/plain; charset=utf-8' })
          });
        });
      })
  );
});
