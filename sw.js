const CACHE="ranglekhaa-v4-production";
const CORE=["./","./index.html","./manifest.json","./icon-192.png","./icon-512.png"];
self.addEventListener("install",e=>e.waitUntil(caches.open(CACHE).then(c=>c.addAll(CORE)).then(()=>self.skipWaiting())));
self.addEventListener("activate",e=>e.waitUntil(caches.keys().then(k=>Promise.all(k.filter(x=>x!==CACHE).map(x=>caches.delete(x)))).then(()=>self.clients.claim())));
self.addEventListener("fetch",e=>{
 const u=new URL(e.request.url);
 if(u.origin===location.origin&&(u.pathname.endsWith("/index.html")||u.pathname.endsWith("/manifest.json"))){
   e.respondWith(fetch(e.request,{cache:"no-store"}).catch(()=>caches.match(e.request))); return;
 }
 e.respondWith(caches.match(e.request).then(r=>r||fetch(e.request)));
});