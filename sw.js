// -----------------------------------------------------------------------------
//  Service Worker  —  Full Offline PWA
//
//  Storage strategy:
//    • navigator.storage.persist()  — prevent Cache API eviction
//    • Cache API     — all non-font app assets (JS, CSS, HTML, images, WASM…)
//    • OPFS          — the 613 numbered font binary chunks (downloaded in bg)
//
//  Fetch strategy:
//    • Navigation/HTML ? network-first, cache fallback, then inline offline page
//    • Font binaries   ? OPFS-first ? Cache-API ? network (store in OPFS)
//    • Everything else ? Cache-first ? network (store in cache)
//
//  Progress: posted to all open clients via postMessage so the UI can show
//  a progress bar.  Also logged to the console.
// -----------------------------------------------------------------------------

const CACHE_VERSION = 'v4';
const CACHE_NAME    = `doc-editor-${CACHE_VERSION}`;
const OPFS_FONT_DIR = 'fonts';

// --- Broadcast a message to every open window/tab in SW scope ----------------
async function broadcast(msg) {
  const clients = await self.clients.matchAll({ type: 'all', includeUncontrolled: true });
  for (const c of clients) c.postMessage(msg);
}

// --- OPFS helpers (font binaries only) ---------------------------------------
async function opfsDir(create) {
  const root = await navigator.storage.getDirectory();
  return root.getDirectoryHandle(OPFS_FONT_DIR, { create: create || false });
}

async function opfsHasFont(name) {
  try {
    const dir = await opfsDir(false);
    await dir.getFileHandle(name);
    return true;
  } catch { return false; }
}

async function opfsReadFont(name) {
  try {
    const dir = await opfsDir(false);
    const fh  = await dir.getFileHandle(name);
    return await fh.getFile();
  } catch { return null; }
}

async function opfsWriteFont(name, buffer) {
  try {
    const dir      = await opfsDir(true);
    const fh       = await dir.getFileHandle(name, { create: true });
    const writable = await fh.createWritable();
    await writable.write(buffer);
    await writable.close();
  } catch (e) {
    console.warn('[SW] OPFS write failed:', name, e.message);
  }
}

// --- Fetch a URL without throwing on network failure -------------------------
async function safeFetch(url) {
  try {
    const r = await fetch(url);
    return r.ok ? r : null;
  } catch { return null; }
}

// --- Split an array into chunks of `size` ------------------------------------
function chunk(arr, size) {
  const out = [];
  for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
  return out;
}

// -------------------------------------------------------------------------------
//  INSTALL
// -------------------------------------------------------------------------------
self.addEventListener('install', function(event) {
  console.log('[SW] Installing v4...');

  event.waitUntil((async function() {
    // 1. Request persistent storage so the browser wont evict our caches.
    try {
      const persisted = await navigator.storage.persist();
      console.log('[SW] navigator.storage.persist():', persisted);
    } catch (e) {
      console.warn('[SW] persist() unavailable:', e.message);
    }

    // 2. Load the build-time manifest that lists every file in the app.
    let assets = [];
    let fonts  = [];
    try {
      const r = await fetch('./precache-manifest.json');
      if (!r.ok) throw new Error('HTTP ' + r.status);
      const data = await r.json();
      assets = data.assets || [];
      fonts  = data.fonts  || [];
      console.log('[SW] Manifest: ' + assets.length + ' assets + ' + fonts.length + ' font binaries');
    } catch (e) {
      console.warn('[SW] Manifest unavailable, using minimal fallback:', e.message);
      assets = [
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
    }

    // 3. Cache every non-font asset in parallel batches of 20.
    //    Blocks install so the full app shell is cached before activation.
    const cache  = await caches.open(CACHE_NAME);
    const total  = assets.length;
    let   loaded = 0;
    const BATCH  = 20;

    await broadcast({ type: 'INSTALL_PROGRESS', phase: 'assets', loaded: 0, total: total, pct: 0 });
    console.log('[SW] Caching ' + total + ' assets...');

    for (const batch of chunk(assets, BATCH)) {
      await Promise.allSettled(batch.map(async function(url) {
        try { await cache.add(url); }
        catch (e) { console.warn('[SW] cache skip:', url, e.message); }
      }));
      loaded += batch.length;
      if (loaded % 100 === 0 || loaded === total) {
        const pct = Math.round(loaded / total * 100);
        console.log('[SW] Assets ' + loaded + '/' + total + ' (' + pct + '%)');
        await broadcast({ type: 'INSTALL_PROGRESS', phase: 'assets', loaded: loaded, total: total, pct: pct });
      }
    }
    console.log('[SW] All assets cached. Starting background font download...');

    // 4. Download font binaries to OPFS in the background (non-blocking).
    (async function() {
      const fontTotal   = fonts.length;
      let   fontLoaded  = 0;
      const FBATCH      = 5;

      if (fontTotal === 0) return;
      await broadcast({ type: 'INSTALL_PROGRESS', phase: 'fonts', loaded: 0, total: fontTotal, pct: 0 });

      for (const batch of chunk(fonts, FBATCH)) {
        await Promise.allSettled(batch.map(async function(url) {
          const name = url.split('/').pop();
          try {
            if (await opfsHasFont(name)) return;
            const r = await safeFetch(url);
            if (!r) return;
            const buf = await r.arrayBuffer();
            await opfsWriteFont(name, buf);
          } catch (e) {
            console.warn('[SW] font skip:', name, e.message);
          }
        }));
        fontLoaded += batch.length;
        if (fontLoaded % 100 === 0 || fontLoaded === fontTotal) {
          const pct = Math.round(fontLoaded / fontTotal * 100);
          console.log('[SW] Fonts ' + fontLoaded + '/' + fontTotal + ' (' + pct + '%)');
          await broadcast({ type: 'INSTALL_PROGRESS', phase: 'fonts', loaded: fontLoaded, total: fontTotal, pct: pct });
        }
      }
      console.log('[SW] All font binaries stored in OPFS.');
      await broadcast({ type: 'INSTALL_COMPLETE' });
    })();

  })());

  self.skipWaiting();
});

// -------------------------------------------------------------------------------
//  ACTIVATE
// -------------------------------------------------------------------------------
self.addEventListener('activate', function(event) {
  event.waitUntil(
    caches.keys().then(function(names) {
      return Promise.all(
        names
          .filter(function(n) { return n !== CACHE_NAME; })
          .map(function(n) { return caches.delete(n); })
      );
    })
  );
  self.clients.claim();
});

// -------------------------------------------------------------------------------
//  FETCH
// -------------------------------------------------------------------------------
self.addEventListener('fetch', function(event) {
  const url = new URL(event.request.url);

  if (event.request.method !== 'GET') return;
  if (url.origin !== self.location.origin) return;

  const isFontBinary = /\/fonts\/\d+$/.test(url.pathname);
  const isHtml       = url.pathname.endsWith('.html') || url.pathname.endsWith('/');
  const isNavigation = event.request.mode === 'navigate';

  // -- Font binaries: OPFS ? Cache API ? Network ------------------------------
  if (isFontBinary) {
    event.respondWith((async function() {
      const name = url.pathname.split('/').pop();

      // 1. OPFS (primary persistent store)
      const opfsFile = await opfsReadFont(name);
      if (opfsFile) {
        return new Response(opfsFile, {
          status: 200,
          headers: { 'Content-Type': 'application/octet-stream' },
        });
      }

      // 2. Cache API (fallback for older cached entries)
      const cached = await caches.match(event.request);
      if (cached) return cached;

      // 3. Network — store in OPFS for next time
      const r = await safeFetch(event.request);
      if (r) {
        r.clone().arrayBuffer().then(function(buf) { opfsWriteFont(name, buf); });
        return r;
      }

      return new Response('Font unavailable offline', { status: 503 });
    })());
    return;
  }

  // -- Navigation / HTML: Network-first ? Cache ? offline fallback page --------
  if (isNavigation || isHtml) {
    event.respondWith((async function() {
      // Try network first
      try {
        const r = await fetch(event.request);
        if (r && r.ok) {
          const cleanUrl = url.origin + url.pathname;
          caches.open(CACHE_NAME).then(function(c) { c.put(cleanUrl, r.clone()); });
          return r;
        }
      } catch(e) { /* offline */ }

      // Cache lookup (ignoreSearch so ?_dc=... still matches)
      const cached = await caches.match(event.request, { ignoreSearch: true });
      if (cached) return cached;

      // App root as final fallback (handles all SPA routes)
      const scope = self.registration.scope;
      const root  = await caches.match(new Request(scope))
                 || await caches.match(new Request(scope + 'index.html'));
      if (root) return root;

      // Absolute last resort — inline offline page
      return new Response(
        '<!DOCTYPE html><html><head><meta charset="utf-8"><title>Offline</title>' +
        '<style>body{font-family:sans-serif;display:flex;align-items:center;' +
        'justify-content:center;height:100vh;margin:0;background:#f5f5f5}' +
        '.box{text-align:center;padding:2rem;background:#fff;border-radius:8px;' +
        'box-shadow:0 2px 8px rgba(0,0,0,.15)}h1{color:#555}p{color:#888}</style>' +
        '</head><body><div class="box"><h1>You\'re offline</h1>' +
        '<p>The app is still downloading for offline use.<br>' +
        'Please connect to the internet once to complete setup.</p>' +
        '</div></body></html>',
        { status: 200, headers: { 'Content-Type': 'text/html; charset=utf-8' } }
      );
    })());
    return;
  }

  // -- All other assets: Cache-first ? Network ? Store ------------------------
  event.respondWith((async function() {
    const cached = await caches.match(event.request);
    if (cached) return cached;

    const r = await safeFetch(event.request);
    if (r) {
      caches.open(CACHE_NAME).then(function(c) { c.put(event.request, r.clone()); });
      return r;
    }

    return new Response('Unavailable offline', { status: 503 });
  })());
});
