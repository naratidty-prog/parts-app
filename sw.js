// เก็บไฟล์แอปไว้ในเครื่อง เพื่อเปิดได้แม้ไม่มีอินเทอร์เน็ต
// เปลี่ยนเลข VERSION ทุกครั้งที่แก้ไฟล์ในโฟลเดอร์ web เพื่อให้เครื่องลูกโหลดของใหม่
const VERSION = 'parts-app-v13';
const FILES = ['./', 'index.html', 'styles.css', 'app.js', 'config.js', 'manifest.webmanifest',
  'vendor/xlsx.full.min.js', 'icons/icon-192.png', 'icons/icon-512.png'];

self.addEventListener('install', (e) => {
  // cache: 'reload' = โหลดจากเซิร์ฟเวอร์จริง ไม่ใช้ไฟล์เก่าที่เบราว์เซอร์จำไว้
  e.waitUntil(caches.open(VERSION).then((c) => c.addAll(FILES.map((f) => new Request(f, { cache: 'reload' })))).then(() => self.skipWaiting()));
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
    // no-cache: ถามเซิร์ฟเวอร์ทุกครั้งว่าไฟล์เปลี่ยนไหม กันได้ไฟล์ app.js เก่าปนกับ index.html ใหม่หลังอัปเดต
    fetch(req, { cache: 'no-cache' }).then((res) => {
      const copy = res.clone();
      caches.open(VERSION).then((c) => c.put(req, copy));
      return res;
    }).catch(() => caches.match(req, { ignoreSearch: true }).then((r) => r || caches.match('index.html')))
  );
});
