// PWA shell — ban xuat ban 01/11/2026 (kem fix ngay that + nap truoc thang).
// Service worker PWA Hồng Phúc — chỉ lo "vỏ app": khi mất mạng vẫn mở được app
// bằng bản HTML gần nhất. Dữ liệu luôn đi qua API (script.google.com) — tuyệt đối
// không chặn hay lưu đệm các request đó.
const CACHE = 'hp-shell-v1';
const SHELL = ['.', 'index.html', 'manifest.webmanifest', 'icon-192.png', 'icon-512.png', 'icon-maskable-512.png', 'apple-touch-icon.png'];
self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET' || url.origin !== self.location.origin) return;
  if (e.request.mode === 'navigate' || url.pathname.endsWith('/index.html') || url.pathname.endsWith('/')) {
    // HTML: MẠNG TRƯỚC — có mạng luôn nhận bản mới nhất; mất mạng mới dùng bản đã lưu.
    e.respondWith(fetch(e.request).then(r => {
      const copy = r.clone();
      caches.open(CACHE).then(c => c.put('index.html', copy));
      return r;
    }).catch(() => caches.match('index.html')));
    return;
  }
  // Tài nguyên tĩnh (icon/manifest): ưu tiên bản đã lưu.
  e.respondWith(caches.match(e.request).then(hit => hit || fetch(e.request)));
});
