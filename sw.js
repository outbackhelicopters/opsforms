/* ============================================================
   Outback Helicopter Airwork NT — Service Worker
   Strategy:
     • HTML pages  → network-first (always get latest when online,
                      fall back to cache when offline)
     • All assets  → cache-first  (logo, etc. — fast repeat loads)
   Bump CACHE_NAME on every deploy to purge stale asset cache.
   ============================================================ */

const CACHE_NAME = 'heliops-v54';

/* Files to pre-cache on install (assets only — HTML is network-first) */
const PRECACHE = [
  './logo.png'
];

/* URL patterns that should always be served network-first */
const HTML_PATTERNS = [
  /\/flight-ops\.html(\?.*)?$/,
  /\/job-advice\.html(\?.*)?$/,
  /\/index\.html(\?.*)?$/,
  /\/$/
];

/* ── Install ─────────────────────────────────────────────── */
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => cache.addAll(PRECACHE))
      .catch(() => { /* logo might 404 in dev — don't block install */ })
  );
  self.skipWaiting();
});

/* ── Activate: delete old caches ─────────────────────────── */
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(
        keys
          .filter(key => key !== CACHE_NAME)
          .map(key => caches.delete(key))
      )
    )
  );
  self.clients.claim();
});

/* ── Fetch ───────────────────────────────────────────────── */
self.addEventListener('fetch', event => {
  const req = event.request;
  if (req.method !== 'GET') return;

  const url = new URL(req.url);

  /* Ignore cross-origin requests */
  if (url.origin !== self.location.origin) return;

  const isHTML = HTML_PATTERNS.some(pattern => pattern.test(url.pathname));

  if (isHTML) {
    /* Network-first for HTML: pilots always get the latest version online */
    event.respondWith(
      fetch(req)
        .then(response => {
          if (response && response.status === 200) {
            const clone = response.clone();
            caches.open(CACHE_NAME).then(cache => cache.put(req, clone));
          }
          return response;
        })
        .catch(() =>
          /* Offline fallback: serve cached HTML */
          caches.match(req).then(cached => cached || new Response(
            '<!DOCTYPE html><html><body style="font-family:sans-serif;padding:40px;text-align:center">' +
            '<h2>You\'re offline</h2><p>Open the app while connected at least once to enable offline use.</p>' +
            '</body></html>',
            { headers: { 'Content-Type': 'text/html' } }
          ))
        )
    );
  } else {
    /* Cache-first for all other assets (logo, fonts, etc.) */
    event.respondWith(
      caches.match(req).then(cached => {
        if (cached) return cached;
        return fetch(req).then(response => {
          if (response && response.status === 200 && response.type === 'basic') {
            const clone = response.clone();
            caches.open(CACHE_NAME).then(cache => cache.put(req, clone));
          }
          return response;
        }).catch(() => cached);
      })
    );
  }
});
