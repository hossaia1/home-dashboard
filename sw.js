const CACHE = 'hd-v2';
// adhan.mp3 removed — audio now streamed from cdn.islamic.network (no local file needed)
const SHELL = ['/', '/index.html', '/manifest.json'];

self.addEventListener('install', e =>
  e.waitUntil(
    caches.open(CACHE).then(c =>
      Promise.allSettled(SHELL.map(url =>
        c.add(url).catch(() => console.log('[SW] Skipping:', url))
      ))
    )
  )
);

self.addEventListener('activate', e =>
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    )
  )
);

self.addEventListener('fetch', e => {
  // Only cache-first for same-origin shell; network-first for API calls
  const url = new URL(e.request.url);
  const isAPI = url.hostname.includes('aladhan.com') ||
                url.hostname.includes('open-meteo.com') ||
                url.hostname.includes('corsproxy.io') ||
                url.hostname.includes('calendar.google.com') ||
                url.hostname.includes('cdn.islamic.network') ||  // adhan audio
                url.hostname.includes('cdn.jsdelivr.net');        // Tabler icons / ical.js

  if (isAPI) {
    e.respondWith(fetch(e.request).catch(() => new Response('{}', { headers: { 'Content-Type': 'application/json' } })));
    return;
  }

  e.respondWith(
    caches.match(e.request).then(r => r || fetch(e.request).then(res => {
      if (res.ok && e.request.method === 'GET') {
        const clone = res.clone();
        caches.open(CACHE).then(c => c.put(e.request, clone));
      }
      return res;
    }))
  );
});
