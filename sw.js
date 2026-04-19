const CACHE_VERSION = 'SW_VERSION_PLACEHOLDER'.includes('PLACEHOLDER') ? 'dev-' + Date.now() : 'SW_VERSION_PLACEHOLDER';
const CACHE_NAME = `document-editor-${CACHE_VERSION}`;

// Critical resources that must be pre-cached for offline support.
// These are the core files needed to launch the editor UI and load documents.
const PRECACHE_ASSETS = [
  './',
  './index.html',
  './manifest.json',

  // Icons & images
  './logo.png',
  './img/elementary.png',
  './img/jhs.png',
  './img/doc-formats/docx.png',
  './img/doc-formats/xlsx.png',
  './img/doc-formats/pptx.png',
  './img/icon/word.ico',
  './img/icon/excel.ico',
  './img/icon/ppt.ico',

  // PDF.js viewer
  './pdfjs/web/viewer.html',
  './pdfjs/web/viewer.css',
  './pdfjs/web/viewer.js',
  './pdfjs/build/pdf.mjs',
  './pdfjs/build/pdf.worker.mjs',

  // OnlyOffice API entry point
  './web-apps/apps/api/documents/api.js',
  './web-apps/vendor/jquery/jquery.min.js',
  './web-apps/vendor/requirejs/require.js',

  // Document editor
  './web-apps/apps/documenteditor/main/index.html',
  './web-apps/apps/documenteditor/main/app.js',
  './web-apps/apps/documenteditor/main/loading.js',
  './web-apps/apps/documenteditor/main/locale/en.json',
  './web-apps/apps/documenteditor/main/locale/zh.json',
  './web-apps/apps/documenteditor/main/resources/css/app.css',
  './web-apps/apps/documenteditor/main/resources/watermark/wm-text.json',

  // Spreadsheet editor
  './web-apps/apps/spreadsheeteditor/main/index.html',
  './web-apps/apps/spreadsheeteditor/main/index_internal.html',
  './web-apps/apps/spreadsheeteditor/main/app.js',
  './web-apps/apps/spreadsheeteditor/main/loading.js',
  './web-apps/apps/spreadsheeteditor/main/locale/en.json',
  './web-apps/apps/spreadsheeteditor/main/locale/zh.json',
  './web-apps/apps/spreadsheeteditor/main/resources/css/app.css',
  './web-apps/apps/spreadsheeteditor/main/resources/formula-lang/en.json',
  './web-apps/apps/spreadsheeteditor/main/resources/formula-lang/en_desc.json',
  './web-apps/apps/spreadsheeteditor/main/resources/formula-lang/zh.json',
  './web-apps/apps/spreadsheeteditor/main/resources/formula-lang/zh_desc.json',

  // Presentation editor
  './web-apps/apps/presentationeditor/main/index.html',
  './web-apps/apps/presentationeditor/main/index.reporter.html',
  './web-apps/apps/presentationeditor/main/app.js',
  './web-apps/apps/presentationeditor/main/app.reporter.js',
  './web-apps/apps/presentationeditor/main/loading.js',
  './web-apps/apps/presentationeditor/main/locale/en.json',
  './web-apps/apps/presentationeditor/main/locale/zh.json',
  './web-apps/apps/presentationeditor/main/resources/css/app.css',

  // Common web-apps resources
  './web-apps/apps/common/main/lib/util/min-log.js',
  './web-apps/apps/common/main/resources/themes/themes.json',

  // SDK core JS (min versions for production)
  './sdkjs/word/sdk-all-min.js',
  './sdkjs/cell/sdk-all-min.js',
  './sdkjs/slide/sdk-all-min.js',
  './sdkjs/word/sdk-all.js',
  './sdkjs/cell/sdk-all.js',
  './sdkjs/slide/sdk-all.js',
  './sdkjs/cell/css/main.css',

  // Common SDK resources
  './sdkjs/common/AllFonts.js',
  './sdkjs/common/Charts/ChartStyles.js',
  './sdkjs/common/libfont/engine/fonts.js',
  './sdkjs/common/libfont/engine/fonts.wasm',
  './sdkjs/common/libfont/engine/fonts_native.js',
  './sdkjs/common/Native/jquery_native.js',
  './sdkjs/common/Native/native.js',
  './sdkjs/common/zlib/engine/zlib.js',
  './sdkjs/common/zlib/engine/zlib.wasm',
  './sdkjs/slide/themes/themes.js',

  // WASM converter (JS loader only; the large .wasm binary is runtime-cached on first use)
  './wasm/x2t/x2t.js',

  // Common web-apps font & loading indicator
  './web-apps/apps/common/main/resources/font/ASC.ttf',
  './web-apps/apps/common/main/resources/img/load-mask/loading.svg',

  // SheetJS
  './libs/sheetjs/xlsx.full.min.js',
];

// Install event: Pre-cache critical assets for offline support
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(PRECACHE_ASSETS);
    }),
  );
  self.skipWaiting();
});

// Activate event: Clean up old caches
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((cacheNames) => {
      return Promise.all(
        cacheNames.map((cacheName) => {
          if (cacheName !== CACHE_NAME) {
            return caches.delete(cacheName);
          }
        }),
      );
    }),
  );
  self.clients.claim();
});

// Fetch event: Cache-first for offline PWA support
self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);

  // Only handle GET requests
  if (event.request.method !== 'GET') return;

  // Only handle same-origin requests
  if (url.origin !== self.location.origin) return;

  // The OnlyOffice editor loads iframe HTML with many query parameters
  // (e.g., ?_dc=0&lang=en&_document=...&_config=...).
  // We need to match cache entries ignoring these query params.
  const isHtml = url.pathname.endsWith('.html') || url.pathname.endsWith('/');
  const isNavigation = event.request.mode === 'navigate';
    // Do NOT intercept the 613 font binaries (/fonts/000 to /fonts/612) into service worker cache
    // since they total ~750MB and can freeze the browser storage or exceed quotas.
    if (url.pathname.match(/\/fonts\/\d{3}$/)) {
      return;
    }
  if (isNavigation || isHtml) {
    // Network-first for HTML/navigation, with cache fallback for offline.
    // Use ignoreSearch so cached HTML matches regardless of query parameters.
    event.respondWith(
      fetch(event.request)
        .then((networkResponse) => {
          if (networkResponse && networkResponse.status === 200) {
            const responseToCache = networkResponse.clone();
            // Cache using a clean URL (no query params) so it always matches
            const cleanUrl = url.origin + url.pathname;
            caches.open(CACHE_NAME).then((cache) => {
              cache.put(cleanUrl, responseToCache);
            });
            return networkResponse;
          }
          return caches.match(event.request, { ignoreSearch: true })
            .then((cached) => cached || networkResponse);
        })
        .catch(() => {
          return caches.match(event.request, { ignoreSearch: true });
        })
    );
    return;
  }

  // For all other same-origin static assets: Cache-first with network fallback.
  // This ensures the app works fully offline after the first visit.
  // Fonts, images, JS, CSS, WASM binaries are all cached on first fetch.
  event.respondWith(
    caches.match(event.request).then((cachedResponse) => {
      if (cachedResponse) {
        return cachedResponse;
      }
      // Not in cache — fetch from network and cache for future offline use
      return fetch(event.request).then((networkResponse) => {
        if (networkResponse && networkResponse.status === 200) {
          const responseToCache = networkResponse.clone();
          caches.open(CACHE_NAME).then((cache) => {
            cache.put(event.request, responseToCache);
          });
        }
        return networkResponse;
      });
    }),
  );
});
