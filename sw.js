/* ArNote service worker: makes the app installable, loads it fast, and receives things shared from other apps. */
const VERSION = 'arnote-v21';
const SHELL = ['./', 'index.html', 'manifest.webmanifest', 'icons/icon-192.png', 'icons/icon-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== VERSION && k !== 'arnote-share').map(k => caches.delete(k)))).then(() => self.clients.claim()));
});

self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  // Share sheet: stash the shared files/text, then open the app, which uploads them.
  if (e.request.method === 'POST' && url.origin === location.origin && url.pathname.endsWith('/share')) {
    e.respondWith((async () => {
      const form = await e.request.formData();
      const files = form.getAll('files').filter(f => f && f.size);
      const c = await caches.open('arnote-share');
      await Promise.all(files.map((f, i) => c.put('file-' + i, new Response(f))));
      await c.put('shared', new Response(JSON.stringify({
        title: form.get('title') || '', text: form.get('text') || '', url: form.get('url') || '',
        files: files.map(f => ({ name: f.name || 'shared', type: f.type || 'application/octet-stream' }))
      })));
      return Response.redirect('./?share=1', 303);
    })());
    return;
  }
  if (e.request.method !== 'GET' || url.origin !== location.origin) return;   // API calls go straight to Google
  // App files: always ask GitHub for the latest copy (no-cache revalidates), fall back to the cached copy when offline.
  e.respondWith(fetch(e.request, { cache: 'no-cache' }).then(r => {
    if (r.ok) { const copy = r.clone(); caches.open(VERSION).then(c => c.put(e.request, copy)); }
    return r;
  }).catch(() => caches.match(e.request).then(r => r || caches.match('index.html'))));
});
