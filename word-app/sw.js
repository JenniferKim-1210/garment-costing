// 단어카드 service worker: 앱 파일을 저장해 두고 오프라인에서도 열리게 합니다.
// 앱을 수정해서 다시 배포할 때는 CACHE 버전 숫자를 올려 주세요.
const CACHE = 'wordcard-v2';
const ASSETS = [
  './',
  './index.html',
  './app.js',
  './words.js',
  './manifest.webmanifest',
  './icons/icon.svg',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/apple-touch-icon.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil(caches.open(CACHE).then((c) => c.addAll(ASSETS)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);

  // 페이지: 온라인이면 최신 버전, 오프라인이면 저장된 버전
  if (req.mode === 'navigate') {
    event.respondWith(
      fetch(req)
        .then((res) => {
          const copy = res.clone();
          caches.open(CACHE).then((c) => c.put('./index.html', copy));
          return res;
        })
        .catch(() => caches.match('./index.html'))
    );
    return;
  }

  // 글꼴: 저장된 것을 먼저 쓰고 뒤에서 갱신
  if (url.hostname === 'fonts.googleapis.com' || url.hostname === 'fonts.gstatic.com') {
    event.respondWith(
      caches.open(CACHE).then((c) =>
        c.match(req).then((hit) => {
          const net = fetch(req).then((res) => { c.put(req, res.clone()); return res; }).catch(() => hit);
          return hit || net;
        })
      )
    );
    return;
  }

  // 그 밖의 앱 파일: 저장본을 바로 쓰고, 뒤에서 최신 파일로 갱신
  if (url.origin === self.location.origin) {
    event.respondWith(
      caches.open(CACHE).then((c) =>
        c.match(req).then((hit) => {
          const net = fetch(req).then((res) => { if (res.ok) c.put(req, res.clone()); return res; }).catch(() => hit);
          return hit || net;
        })
      )
    );
  }
});

// 공부 안 한 날 알림 (안드로이드 Chrome에 설치한 경우, 브라우저가 허락할 때만 동작)
async function readState() {
  const c = await caches.open('wordcard-state');
  const r = await c.match('./__state');
  return r ? r.json() : null;
}
async function maybeRemind() {
  const s = await readState();
  if (!s || !s.remind) return;
  const now = new Date();
  const today = Math.floor((now.getTime() - now.getTimezoneOffset() * 60000) / 86400000);
  if (s.studied === today || now.getHours() < s.hour || s.notified === today) return;
  await self.registration.showNotification('단어카드', {
    body: s.body || '오늘 단어를 아직 안 했어요. 5분이면 충분해요!',
    icon: 'icons/icon-192.png', badge: 'icons/icon-192.png', tag: 'study-reminder'
  });
  const c = await caches.open('wordcard-state');
  await c.put('./__state', new Response(JSON.stringify(Object.assign({}, s, { notified: today })), { headers: { 'Content-Type': 'application/json' } }));
}
self.addEventListener('periodicsync', (event) => {
  if (event.tag === 'study-reminder') event.waitUntil(maybeRemind());
});
self.addEventListener('notificationclick', (event) => {
  event.notification.close();
  event.waitUntil(
    self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then((list) => {
      for (const c of list) { if ('focus' in c) return c.focus(); }
      return self.clients.openWindow('./');
    })
  );
});
