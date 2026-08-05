// Bump this on every deploy. The shell is cache-first, so a new name is what
// evicts the old index.html/icons from everyone's installed copy.
const CACHE = 'pnac-syllabus-v3';
const SHELL = [
  './',
  './index.html',
  './manifest.webmanifest',
  './icon-192.png',
  './icon-512.png',
  './icon-maskable-512.png',
  './apple-touch-icon.png'
];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

// Network-first, but a bad status counts as a failure rather than as data:
// caching a captive-portal page or a 500 over trips.json would leave a member
// permanently unable to see the syllabus offline. When we do serve the saved
// copy instead, tag it so the page can say so out loud.
async function freshTrips(req) {
  try {
    const res = await fetch(req);
    if (!res.ok) throw new Error('status ' + res.status);
    const copy = res.clone();
    caches.open(CACHE).then((c) => c.put(req, copy));
    return res;
  } catch (err) {
    const hit = await caches.match(req);
    if (!hit) throw err;
    const headers = new Headers(hit.headers);
    headers.set('X-From-Cache', '1');
    return new Response(await hit.blob(), { status: 200, headers });
  }
}

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);

  // trips.json: network-first so new syllabus data shows, fall back to cache offline
  if (url.pathname.endsWith('trips.json')) {
    e.respondWith(freshTrips(req));
    return;
  }

  // everything else: cache-first
  e.respondWith(
    caches.match(req).then((hit) => hit || fetch(req).then((res) => {
      // Only ever cache a good response. A 404, a 500 or a wifi captive-portal
      // login page must not be written over a working copy of the shell.
      if (res.ok) {
        const copy = res.clone();
        caches.open(CACHE).then((c) => c.put(req, copy));
      }
      return res;
    }))
  );
});
