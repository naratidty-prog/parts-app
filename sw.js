// เก็บไฟล์แอปไว้ในเครื่อง เพื่อเปิดได้แม้ไม่มีอินเทอร์เน็ต
// เปลี่ยนเลข VERSION ทุกครั้งที่แก้ไฟล์ในโฟลเดอร์ web เพื่อให้เครื่องลูกโหลดของใหม่
const VERSION = 'parts-app-v4';
const FILES = ['./', 'index.html', 'styles.css', 'app.js', 'config.js', 'manifest.webmanifest',
  'vendor/xlsx.full.min.js', 'icons/icon-192.png', 'icons/icon-512.png'];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(VERSION).then((c) => c.addAll(FILES)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== VERSION).map((k) => caches.delete(k))))
    .then(() => self.clients.claim()));
});

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== location.origin) return;
  // ใช้ของใหม่จากเน็ตก่อน ถ้าไม่มีเน็ตใช้ของในเครื่อง
  e.respondWith(
    fetch(req).then((res) => {
      const copy = res.clone();
      caches.open(VERSION).then((c) => c.put(req, copy));
      return res;
    }).catch(() => caches.match(req, { ignoreSearch: true }).then((r) => r || caches.match('index.html')))
  );
});
