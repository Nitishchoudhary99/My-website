/* DocBrisk service worker.
   1. Makes the site installable as an app.
   2. "Share to DocBrisk": files shared from WhatsApp, Gallery or Files are
      kept briefly in a local cache and handed to the home page.
   3. Opens the site even when offline, once it has been visited.
   Documents opened inside the tools are never touched or stored here. */
const VERSION = 'docbrisk-v1';

self.addEventListener('install', (event) => {
  self.skipWaiting();
  event.waitUntil(caches.open(VERSION).then((c) => c.add('/')).catch(() => {}));
});

self.addEventListener('activate', (event) => {
  event.waitUntil((async () => {
    for (const key of await caches.keys()) {
      if (key !== VERSION && key !== 'docbrisk-share') await caches.delete(key);
    }
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  const url = new URL(req.url);

  /* Share to DocBrisk */
  if (req.method === 'POST' && url.pathname === '/share-target') {
    event.respondWith((async () => {
      try {
        const form = await req.formData();
        const files = form.getAll('files').filter((f) => f && typeof f !== 'string');
        const cache = await caches.open('docbrisk-share');
        await Promise.all(files.map((f, i) => cache.put(
          new Request('/shared/' + Date.now() + '-' + i),
          new Response(f, { headers: {
            'content-type': f.type || 'application/octet-stream',
            'x-filename': encodeURIComponent(f.name || 'file')
          } })
        )));
      } catch (e) { /* open the app anyway */ }
      return Response.redirect('/?shared=1', 303);
    })());
    return;
  }

  /* Pages: always try the network first (so updates show immediately),
     fall back to the last saved copy when offline. */
  if (req.method === 'GET' && req.mode === 'navigate' && url.origin === self.location.origin) {
    event.respondWith((async () => {
      try {
        const res = await fetch(req);
        if (res && res.ok && (res.headers.get('content-type') || '').includes('text/html')) {
          const copy = res.clone();
          caches.open(VERSION).then((c) => c.put('/', copy)).catch(() => {});
        }
        return res;
      } catch (e) {
        return (await caches.match('/')) || Response.error();
      }
    })());
  }
});
