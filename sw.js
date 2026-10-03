const CACHE = 'hd-v11';  // v11: Quran recitations kept across updates

// Saved recitations live in their own cache, managed by the page. It must
// survive app updates — re-downloading hundreds of MB because the app shell
// changed would be absurd — so activate never deletes it.
const KEEP = ['hd-quran'];

// Pinned third-party assets. Precached at install rather than left to be picked
// up opportunistically, so the very first offline boot already has them —
// relying on page-load order meant the CSS and ical.js were often missing.
const VENDOR = [
  'https://cdn.jsdelivr.net/npm/@tabler/icons-webfont@3.46.0/dist/tabler-icons.min.css',
  'https://cdn.jsdelivr.net/npm/@tabler/icons-webfont@3.46.0/dist/fonts/tabler-icons.woff2?v3.46.0',
  'https://cdn.jsdelivr.net/npm/ical.js@1.5.0/build/ical.min.js',
];

// The adhan recording is precached so the call still plays after a reboot with
// no network — the one thing on this dashboard that absolutely must not depend
// on connectivity. Optional files: c.add() failures are swallowed below.
const AUDIO = ['/adhan.mp3', '/adhan-fajr.mp3'];

const SHELL = ['/', '/index.html', '/manifest.json', ...VENDOR, ...AUDIO];

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
      .then(keys => Promise.all(keys.filter(k => k !== CACHE && !KEEP.includes(k)).map(k => caches.delete(k))))
      .then(() => self.clients.claim())   // take over open pages right away
  )
);

/**
 * Compares the cached recording against the server and replaces it if it changed,
 * then tells any open page so it can reload its in-memory copy.
 *
 * ETag is preferred; GitHub Pages sends one. Content-Length is the fallback,
 * which catches any realistic swap — a different recording is never byte-identical
 * in length. Failure is silent on purpose: this runs on every audio request and
 * being offline is the normal case it must tolerate.
 */
async function revalidateAudio(request, cached) {
  try {
    const fresh = await fetch(request.url, { cache: 'no-cache' });
    if (fresh.status !== 200) return;

    const oldTag = cached.headers.get('etag');
    const newTag = fresh.headers.get('etag');
    const changed = (oldTag && newTag)
      ? oldTag !== newTag
      : cached.headers.get('content-length') !== fresh.headers.get('content-length');
    if (!changed) return;

    const cache = await caches.open(CACHE);
    await cache.put(request, fresh.clone());
    const clients = await self.clients.matchAll({ includeUncontrolled: true });
    clients.forEach(c => c.postMessage({
      type: 'audio-updated', path: new URL(request.url).pathname
    }));
    console.log('[SW] replaced cached audio:', new URL(request.url).pathname);
  } catch (_) { /* offline, or the file is gone — keep what we have */ }
}

self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);

  // Range requests come from <audio> seeking. A service worker answering those
  // with a full or partial cached body breaks playback, so leave them alone.
  // Note this is now keyed on the Range header rather than the file extension:
  // the page fetches the adhan as a plain Blob (no Range), and that request DOES
  // need to be served from cache when offline.
  if (e.request.headers.has('range')) return;

  // Local audio: answered from cache immediately — that is what makes an offline
  // adhan possible — then quietly revalidated in the background. Without the
  // revalidation, replacing adhan.mp3 had no effect until someone remembered to
  // bump CACHE, which is a silent failure of exactly the kind this app must not
  // have. Swapping the file now just works on the next reload.
  if (url.origin === self.location.origin && /\.(mp3|ogg|wav|m4a)$/i.test(url.pathname)) {
    e.respondWith(
      caches.match(e.request).then(hit => {
        if (hit) {
          e.waitUntil(revalidateAudio(e.request, hit));
          return hit;
        }
        return fetch(e.request).then(res => {
          if (res.status === 200) {
            const clone = res.clone();
            caches.open(CACHE).then(c => c.put(e.request, clone));
          }
          return res;
        });
      })
    );
    return;
  }

  // Pinned library assets (icon font + CSS, ical.js). Cache-first and kept
  // forever: the versions are fixed, so a stale copy is the correct copy, and
  // this is what lets the dashboard boot with icons after a network-less
  // restart. Previously these were network-first, which meant no network at
  // boot = every icon rendered as a blank box.
  const isVendor = url.hostname === 'cdn.jsdelivr.net';

  // Live data: always network, never stored. script.google.com matters here —
  // without it the Apps Script bridge fell through to the caching branch below,
  // and because those calls carry a cache-busting timestamp every single one
  // became a new cache entry. On a dashboard that runs for months that is an
  // unbounded leak, and it risks serving a stale calendar too.
  const isAPI = !isVendor && (
                url.hostname.includes('aladhan.com')            ||
                url.hostname.includes('open-meteo.com')          ||
                url.hostname.includes('corsproxy.io')            ||
                url.hostname.includes('calendar.google.com')     ||
                url.hostname.includes('script.google.com')       ||
                url.hostname.includes('script.googleusercontent.com'));

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

  // Anything else cross-origin (recitation audio, surah list) is left entirely
  // to the browser. The branch below used to catch these too and copy every
  // 200 response into the app cache — which for recitations would have stored a
  // second copy of up to 1.6 GB of audio alongside the page's own Quran cache.
  if (url.origin !== self.location.origin) return;

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
