const CACHE_VERSION = 'v3';
const CACHE_NAME = `document-editor-${CACHE_VERSION}`;
const FONT_CACHE_NAME = `document-editor-fonts-${CACHE_VERSION}`;

// ─── Install ────────────────────────────────────────────────────────────────
// 1. Fetch the build-time manifest that lists every file in the app.
// 2. Cache all non-font assets synchronously (blocks install completion so the
//    app is immediately fully offline-capable for the UI + editors).
// 3. Start downloading the 613 font binaries in the background without
//    blocking install — they accumulate across sessions until all are cached.
self.addEventListener('install', (event) => {
  event.waitUntil(
    fetch('./precache-manifest.json')
      .then((r) => {
        if (!r.ok) throw new Error('Could not fetch precache-manifest.json');
        return r.json();
      })
      .then(({ assets, fonts }) => {
        return caches.open(CACHE_NAME).then(async (cache) => {
          // Cache all non-font assets — a failed individual fetch is skipped,
          // not fatal, so a missing image never aborts the whole install.
          await Promise.allSettled(
            (assets || []).map((url) =>
              cache.add(url).catch((err) => {
                console.warn('[SW] precache skip:', url, err.message);
              }),
            ),
          );

          // Background-download font binaries without blocking install.
          // Uses a dedicated cache so font data persists across CACHE_VERSION
          // bumps (fonts don't change between app updates).
          caches.open(FONT_CACHE_NAME).then(async (fontCache) => {
            for (const url of (fonts || [])) {
              try {
                const already = await fontCache.match(url);
                if (!already) await fontCache.add(url);
              } catch {
                // Network unavailable etc. — will retry on next visit.
              }
            }
            console.log('[SW] All font binaries cached.');
          });
        });
      })
      .catch(async (err) => {
        // Manifest fetch failed (e.g. first offline install) — fall back to
        // the hard-coded critical-asset list so the app still works.
        console.warn('[SW] Manifest unavailable, using fallback list:', err.message);
        const FALLBACK = [
          './',
          './index.html',
          './manifest.json',
          './web-apps/apps/api/documents/api.js',
          './web-apps/apps/documenteditor/main/index.html',
          './web-apps/apps/documenteditor/main/app.js',
          './web-apps/apps/spreadsheeteditor/main/index.html',
          './web-apps/apps/spreadsheeteditor/main/app.js',
          './web-apps/apps/presentationeditor/main/index.html',
          './web-apps/apps/presentationeditor/main/app.js',
          './sdkjs/word/sdk-all-min.js',
          './sdkjs/cell/sdk-all-min.js',
          './sdkjs/slide/sdk-all-min.js',
          './sdkjs/common/AllFonts.js',
          './wasm/x2t/x2t.js',
        ];
        const cache = await caches.open(CACHE_NAME);
        await Promise.allSettled(FALLBACK.map((u) => cache.add(u).catch(() => {})));
      }),
  );
  self.skipWaiting();
});

// ─── Activate ───────────────────────────────────────────────────────────────
// Delete outdated main caches.  The font cache has its own versioned name and
// is kept as long as the version matches — fonts don't change between deploys.
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((cacheNames) => {
      return Promise.all(
        cacheNames.map((cacheName) => {
          // Keep current main cache and current font cache; delete everything else.
          if (cacheName !== CACHE_NAME && cacheName !== FONT_CACHE_NAME) {
            return caches.delete(cacheName);
          }
        }),
      );
    }),
  );
  self.clients.claim();
});

// ─── Fetch ───────────────────────────────────────────────────────────────────
// Strategy:
//   • HTML / navigation  → network-first, cache fallback (ignores query params)
//   • Font binaries      → font cache first, then main cache, then network
//                          (adds to font cache if fetched from network)
//   • Everything else    → cache-first (main cache), network fallback + store
self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);

  if (event.request.method !== 'GET') return;
  if (url.origin !== self.location.origin) return;

  const isFontBinary = /\/fonts\/\d+$/.test(url.pathname);
  const isHtml = url.pathname.endsWith('.html') || url.pathname.endsWith('/');
  const isNavigation = event.request.mode === 'navigate';

  if (isFontBinary) {
    // Font binaries: check font cache → main cache → network (and store in font cache)
    event.respondWith(
      caches.open(FONT_CACHE_NAME).then(async (fontCache) => {
        const cached = await fontCache.match(event.request);
        if (cached) return cached;
        // Also check the main cache (older cached entries)
        const mainCached = await caches.match(event.request);
        if (mainCached) return mainCached;
        // Fetch from network and store in font cache
        const networkResponse = await fetch(event.request);
        if (networkResponse && networkResponse.status === 200) {
          fontCache.put(event.request, networkResponse.clone());
        }
        return networkResponse;
      }),
    );
    return;
  }

  if (isNavigation || isHtml) {
    // Network-first for HTML/navigation, cache fallback for offline.
    event.respondWith(
      fetch(event.request)
        .then((networkResponse) => {
          if (networkResponse && networkResponse.status === 200) {
            const cleanUrl = url.origin + url.pathname;
            caches.open(CACHE_NAME).then((cache) => cache.put(cleanUrl, networkResponse.clone()));
            return networkResponse;
          }
          return caches.match(event.request, { ignoreSearch: true })
            .then((cached) => cached || networkResponse);
        })
        .catch(() => caches.match(event.request, { ignoreSearch: true })),
    );
    return;
  }

  // All other static assets: cache-first, network fallback + cache for next time.
  event.respondWith(
    caches.match(event.request).then((cachedResponse) => {
      if (cachedResponse) return cachedResponse;
      return fetch(event.request).then((networkResponse) => {
        if (networkResponse && networkResponse.status === 200) {
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, networkResponse.clone()));
        }
        return networkResponse;
      });
    }),
  );
});

