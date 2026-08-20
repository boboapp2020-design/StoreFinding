/* Service Worker — ผู้ช่วยคลังพัสดุ (offline support) */
const CACHE = 'warehouse-app-v25';
const ASSETS = [
  './',
  './index.html',
  './manifest.webmanifest',
  './xlsx.full.min.js',
  './na.png',
  './icon-192.png',
  './icon-512.png',
  './icon-512-maskable.png',
  './apple-touch-icon.png'
];

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE).then(c => c.addAll(ASSETS)).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET') return;
  var req = e.request;
  // ข้ามคำขอข้ามโดเมน (เช่น Google Apps Script) — ปล่อยให้เบราว์เซอร์จัดการเอง ไม่ให้ SW แคช/ทำพัง
  if (new URL(req.url).origin !== self.location.origin) return;
  var isHTML = req.mode === 'navigate' || (req.headers.get('accept') || '').indexOf('text/html') !== -1;

  if (isHTML) {
    // network-first: เปิดแอปได้ตัวล่าสุดเสมอเมื่อออนไลน์ (ออฟไลน์ค่อยใช้ cache)
    e.respondWith(
      fetch(req).then(res => {
        var copy = res.clone();
        caches.open(CACHE).then(c => c.put(req, copy)).catch(() => {});
        return res;
      }).catch(() => caches.match(req).then(r => r || caches.match('./index.html')))
    );
    return;
  }
  // ไฟล์อื่น (ไอคอน/สคริปต์) → cache-first เพื่อความเร็ว
  e.respondWith(
    caches.match(req).then(hit => hit || fetch(req).then(res => {
      var copy = res.clone();
      caches.open(CACHE).then(c => c.put(req, copy)).catch(() => {});
      return res;
    }))
  );
});
