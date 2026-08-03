const CACHE = 'hd-v8';

// Pinned third-party assets. Precached at install rather than left to be picked
// up opportunistically, so the very first offline boot already has them —
// relying on page-load order meant the CSS and ical.js were often missing.
const VENDOR = [
  'https://cdn.jsdelivr.net/npm/@tabler/icons-webfont@3.46.0/dist/tabler-icons.min.css',
  'https://cdn.jsdelivr.net/npm/@tabler/icons-webfont@3.46.0/dist/fonts/tabler-icons.woff2?v3.46.0',
  'https://cdn.jsdelivr.net/npm/ical.js@1.5.0/build/ical.min.js',
];

// Local adhan files are optional — c.add() failures are swallowed below.
const SHELL = ['/', '/index.html', '/manifest.json', ...VENDOR];

self.addEventListener('install', e => {
  // Activate the new worker immediately instead of waiting for every tab to
  // close. A wall-mounted dashboard is never "closed", so without this an
  // update could sit unapplied indefinitely.
  self.skipWaiting();
  e.waitUntil(
    caches.open(CACHE).then(c =>
      Promise.allSettled(SHELL.map(url =>
        c.add(url).catch(() => console.log('[SW] Skipping:', url))
      ))
    )
  );
});

self.addEventListener('activate', e =>
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())   // take over open pages right away
  )
);

self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);

  // Never intercept audio. Media elements rely on Range requests, and a service
  // worker answering those with a full or partial cached body breaks playback
  // in ways that are painful to debug. Let the browser handle it directly.
  if (e.request.headers.has('range') || /\.(mp3|ogg|wav|m4a)$/i.test(url.pathname)) return;

  // Pinned library assets (icon font + CSS, ical.js). Cache-first and kept
  // forever: the versions are fixed, so a stale copy is the correct copy, and
  // this is what lets the dashboard boot with icons after a network-less
  // restart. Previously these were network-first, which meant no network at
  // boot = every icon rendered as a blank box.
  const isVendor = url.hostname === 'cdn.jsdelivr.net';

  const isAPI = !isVendor && (
                url.hostname.includes('aladhan.com')        ||
                url.hostname.includes('open-meteo.com')     ||
                url.hostname.includes('corsproxy.io')       ||
                url.hostname.includes('calendar.google.com'));

  if (isVendor) {
    e.respondWith(
      caches.match(e.request).then(hit => hit || fetch(e.request).then(res => {
        if (res.status === 200) {
          const clone = res.clone();
          caches.open(CACHE).then(c => c.put(e.request, clone));
        }
        return res;
      }))
    );
    return;
  }

  if (isAPI) {
    e.respondWith(fetch(e.request).catch(() => new Response('{}', { headers: { 'Content-Type': 'application/json' } })));
    return;
  }

  // Same-origin shell: network-first so edits show up on reload, falling back
  // to cache when offline. Cache-first was silently serving a stale index.html.
  //
  // cache:'no-cache' forces a revalidation against the server. Without it the
  // browser's own HTTP cache can satisfy this fetch from a heuristically-cached
  // copy — which is why edits sometimes appeared not to apply at all.
  e.respondWith(
    fetch(new Request(e.request.url, {
        method: 'GET', cache: 'no-cache', credentials: 'same-origin'
      }))
      .then(res => {
        if (res.status === 200 && e.request.method === 'GET') {
          const clone = res.clone();
          caches.open(CACHE).then(c => c.put(e.request, clone));
        }
        return res;
      })
      .catch(() => caches.match(e.request))
  );
});
