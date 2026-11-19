// Čuva aplikaciju u telefonu da se otvara i bez mreže.
const V = 'ak-v3';
const FILES = ['./', './index.html', './app.css', './app.js', './config.js', './manifest.webmanifest', './icon.svg',
  'https://cdn.jsdelivr.net/npm/@azure/msal-browser@2.38.3/lib/msal-browser.min.js'];
self.addEventListener('install', e => { e.waitUntil(caches.open(V).then(c => c.addAll(FILES)).then(() => self.skipWaiting())); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(k => Promise.all(k.filter(x => x !== V).map(x => caches.delete(x)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET') return;
  if (u.hostname.endsWith('microsoft.com') || u.hostname.endsWith('microsoftonline.com') || u.hostname.endsWith('sharepoint.com')) return;
  // prvo mreža (da stigne nova verzija), pa keš ako nema mreže
  e.respondWith(fetch(e.request, { cache: 'no-cache' }).then(r => { const cp = r.clone(); caches.open(V).then(c => c.put(e.request, cp)); return r; }).catch(() => caches.match(e.request, { ignoreSearch: true })));
});
self.addEventListener('message', e => { if (e.data === 'skip') self.skipWaiting(); });
