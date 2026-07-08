/* The Millionaires Engine - network-only service worker.
   Purpose: make the app installable on Android (WebAPK: real icon, standalone,
   in the app drawer) WITHOUT caching anything, so it ALWAYS loads the latest
   HTML from the bucket. No offline support by design - the app needs the network
   to reach the webhook regardless. */
self.addEventListener('install', function () { self.skipWaiting(); });
self.addEventListener('activate', function (e) { e.waitUntil(self.clients.claim()); });
self.addEventListener('fetch', function () { /* pass-through: let the network handle every request */ });
