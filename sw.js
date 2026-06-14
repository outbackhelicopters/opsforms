/* Heli Ops PWA service worker — caches the app shell so it works fully offline. */
const CACHE = 'heliops-v20';
/* Core files that MUST cache (fail SW install if any missing) */
const CORE = ['./', './index.html', './sw.js'];
/* Extra files — cached best-effort so SW install still succeeds even if temporarily 404 */
const EXTRA = ['./flight-ops.html', './job-advice.html', './job-sw.js'];

self.addEventListener('install', (e) => {
  e.waitUntil(
    caches.open(CACHE).then(async (c) => {
      await c.addAll(CORE);
      /* Cache extras gracefully — don't abort install if they're temporarily unavailable */
      for (const url of EXTRA) {
        try {
          const res = await fetch(url);
          if (res.ok) await c.put(url, res);
        } catch (_) { /* ignore */ }
      }
    }).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

/* Cache-first: the app shell never needs the network once installed. */
self.addEventListener('fetch', (e) => {
  if (e.request.method !== 'GET') return;
  e.respondWith(
    caches.match(e.request, { ignoreSearch: true }).then((hit) => {
      if (hit) return hit;
      return fetch(e.request).then((res) => {
        const copy = res.clone();
        caches.open(CACHE).then((c) => c.put(e.request, copy));
        return res;
      }).catch(() => caches.match('./index.html'));
    })
  );
});
