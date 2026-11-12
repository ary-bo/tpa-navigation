'use strict';

const CACHE='tpa-navigation-shell-v7';
const RUNTIME='tpa-navigation-runtime-v7';
const CORE=[
 './',
 './index.html',
 './style.css',
 './js/app.js',
 './js/enhancements.js',
 './manifest.webmanifest',
 './assets/icon.svg'
];

self.addEventListener('install',event=>{
 event.waitUntil((async()=>{
  const cache=await caches.open(CACHE);
  await cache.addAll(CORE);
  await self.skipWaiting();
 })());
});

self.addEventListener('activate',event=>{
 event.waitUntil((async()=>{
  const keys=await caches.keys();
  await Promise.all(
   keys.filter(k=>k.startsWith('tpa-navigation-')&&k!==CACHE&&k!==RUNTIME)
    .map(k=>caches.delete(k))
  );
  await self.clients.claim();
 })());
});

self.addEventListener('fetch',event=>{
 const req=event.request;
 if(req.method!=='GET')return;
 const url=new URL(req.url);

 if(url.origin===self.location.origin){
  if(req.mode==='navigate'){
   event.respondWith(
    fetch(req).then(res=>{
     const copy=res.clone();
     caches.open(CACHE).then(c=>c.put('./index.html',copy));
     return res;
    }).catch(()=>caches.match('./index.html'))
   );
   return;
  }

  event.respondWith(
   caches.match(req).then(hit=>{
    if(hit)return hit;
    return fetch(req).then(res=>{
     if(res.ok)caches.open(CACHE).then(c=>c.put(req,res.clone()));
     return res;
    });
   })
  );
  return;
 }

 if(['cdn.jsdelivr.net','unpkg.com','accounts.google.com'].includes(url.hostname)){
  event.respondWith(
   caches.open(RUNTIME).then(async c=>{
    try{
     const res=await fetch(req);
     if(res.ok)await c.put(req,res.clone());
     return res;
    }catch(e){
     const hit=await c.match(req);
     if(hit)return hit;
     throw e;
    }
   })
  );
 }
});