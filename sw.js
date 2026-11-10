/* Service Worker for Woosung Techboard v207 */
self.addEventListener('install', (event) => {
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(self.clients.claim());
});

self.addEventListener('message', (event) => {
  if (!event.data) return;
  if (event.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  } else if (event.data.type === 'SHOW_NOTIFICATION') {
    const { title, body, icon, badge, data } = event.data;
    event.waitUntil(
      self.registration.showNotification(title || '우성 일정 알림', {
        body: body || '',
        icon: icon || 'partner_icon_640x640.png',
        badge: badge || 'partner_icon_640x640.png',
        data: data || {},
        tag: event.data.tag || (data && data.tag) || ('memo-' + Date.now()),
        renotify: true,
        vibrate: [150, 70, 150]
      })
    );
  }
});

self.addEventListener('notificationclick', (event) => {
  event.notification.close();
  const notifData = event.notification.data || {};
  let targetUrl = './index.html';
  const queryParams = [];
  if (notifData.siteName) queryParams.push('openSite=' + encodeURIComponent(notifData.siteName));
  if (notifData.taskName) queryParams.push('openTask=' + encodeURIComponent(notifData.taskName));
  if (notifData.taskId) queryParams.push('openTaskId=' + encodeURIComponent(notifData.taskId));
  if (notifData.row) queryParams.push('openRow=' + encodeURIComponent(notifData.row));
  if (notifData.scope) queryParams.push('openScope=' + encodeURIComponent(notifData.scope));
  if (queryParams.length > 0) {
    targetUrl += '?' + queryParams.join('&');
  }

  event.waitUntil(
    self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then((clientList) => {
      if (clientList && clientList.length > 0) {
        let targetClient = null;
        for (const client of clientList) {
          try {
            client.postMessage({
              type: 'WS_NOTIFICATION_CLICK',
              payload: notifData
            });
          } catch (ePost) {}
          if (!targetClient || client.focused || client.visibilityState === 'visible') {
            targetClient = client;
          }
        }
        if (targetClient && 'focus' in targetClient) {
          return targetClient.focus();
        }
        return;
      }
      if (self.clients.openWindow) {
        return self.clients.openWindow(targetUrl);
      }
    })
  );
});
