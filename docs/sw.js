/*
 * sw — offline cache for the weekly-report web app.
 *
 * Network first, so a deploy is picked up on the next load; the cache only
 * answers when offline. Matching ignores the ?v= cache-buster, so this file
 * never needs a version bump. Reports never pass through here: they are read
 * and written via the File System Access API or origin-private storage.
 *
 * Copyright (c) 2026 Yusuf Bunyamin Kahveci <yusuf.kahveci@tii.ae>
 * SPDX-License-Identifier: Apache-2.0
 */

'use strict';

const CACHE = 'weekly-report';
const SHELL = [
  './', 'index.html', 'app.js', 'zip.js', 'manifest.webmanifest',
  'vendor/marked-14.min.js', 'vendor/highlight-11.min.js',
  'icons/icon.svg', 'icons/icon-192.png', 'icons/icon-512.png',
];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(self.clients.claim());
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== location.origin) return;
  e.respondWith(fetch(req).then(res => {
    if (res.ok) {
      const copy = res.clone();
      caches.open(CACHE).then(c => c.put(req, copy));
    }
    return res;
  }).catch(async () => {
    const hit = await caches.match(req, { ignoreSearch: true });
    if (hit) return hit;
    if (req.mode === 'navigate') return caches.match('index.html');
    return Response.error();
  }));
});
