// Digital Pilot OPS · offline helper (service worker) · VERSION 2026-11-04-U
// Keeps a copy of the app on the tablet so it still opens without internet.
// With internet: always loads the newest page from GitHub (and refreshes the saved copy).
// Without internet: opens the saved copy instead of the "not found" error.
// It only handles the app's own files. Google Sheets, Zelyna and the voice service are never cached.
const CACHE = 'birriaz-app-v1';
const WAIT = 5000; // weak Wi-Fi: after 5 seconds use the saved copy

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(['./', './index.html']).catch(() => {})).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k.startsWith('birriaz-app-') && k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});

function fromNet(req, page) {
  return new Promise((ok, no) => {
    const t = setTimeout(() => no(new Error('slow')), WAIT);
    (page ? fetch(req.url, { cache: 'no-store', credentials: 'same-origin' }) : fetch(req)).then(r => { clearTimeout(t); ok(r); }, er => { clearTimeout(t); no(er); });
  });
}

self.addEventListener('fetch', e => {
  const req = e.request, u = new URL(req.url);
  if (req.method !== 'GET' || u.origin !== self.location.origin) return; // Google script, Zelyna, voice: untouched
  if (u.pathname.endsWith('/sw.js')) return;
  const page = req.mode === 'navigate' || u.pathname.endsWith('/') || u.pathname.endsWith('.html');
  e.respondWith(
    fromNet(req, page).then(r => {
      if (r && r.ok) { const copy = r.clone(); caches.open(CACHE).then(c => c.put(page ? './index.html' : req, copy)).catch(() => {}); }
      return r;
    }).catch(() =>
      caches.open(CACHE).then(c => (page ? c.match('./index.html').then(m => m || c.match('./')) : c.match(req, { ignoreSearch: true })))
        .then(m => m || new Response('<h1 style="font-family:sans-serif">No internet</h1><p style="font-family:sans-serif;font-size:20px">The app has not been saved on this tablet yet. Connect to Wi-Fi once and open it, then it works without internet.</p>', { status: 503, headers: { 'Content-Type': 'text/html; charset=utf-8' } }))
    )
  );
});
