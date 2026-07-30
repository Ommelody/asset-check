/* Service worker — ระบบตรวจนับครุภัณฑ์ QR Code
   กลยุทธ์: network-first สำหรับหน้าแอป (ได้เวอร์ชันใหม่เสมอเมื่อออนไลน์),
   cache-first สำหรับไอคอน/รูปภาพ, และไม่แคชคำขอไป Google Apps Script เลย */
const CACHE = 'krupan-v3';
const SHELL = [
  './',
  './index.html',
  './manifest.json',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/apple-touch-icon.png',
  './icons/favicon-32.png',
];

self.addEventListener('install', (e) => {
  e.waitUntil(
    caches.open(CACHE)
      .then((c) => Promise.allSettled(SHELL.map((u) => c.add(new Request(u, { cache: 'reload' })))))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('message', (e) => {
  if (e.data === 'skipWaiting') self.skipWaiting();
});

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);

  // อย่าแตะคำขอไป Google Apps Script / API ภายนอก — ต้องสดเสมอ
  if (url.hostname.includes('script.google') || url.hostname.includes('googleusercontent')) return;

  // หน้าแอป: network-first เผื่อมีเวอร์ชันใหม่ ถ้าออฟไลน์ใช้ของในแคช
  const isDoc = req.mode === 'navigate' || (req.headers.get('accept') || '').includes('text/html');
  if (isDoc) {
    e.respondWith(
      fetch(req)
        .then((res) => { caches.open(CACHE).then((c) => c.put(req, res.clone())); return res; })
        .catch(() => caches.match(req).then((r) => r || caches.match('./index.html')))
    );
    return;
  }

  // ที่เหลือ (ไอคอน รูปครุภัณฑ์ ฟอนต์): cache-first แล้วเติมแคชเบื้องหลัง
  e.respondWith(
    caches.match(req).then((hit) => {
      if (hit) return hit;
      return fetch(req).then((res) => {
        if (res && res.ok && (url.origin === location.origin || url.hostname.includes('gstatic') || url.hostname.includes('jsdelivr'))) {
          caches.open(CACHE).then((c) => c.put(req, res.clone()));
        }
        return res;
      }).catch(() => hit);
    })
  );
});
