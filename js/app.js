'use strict';

/* TPA Navigation PWA — core application */
const GOOGLE_CLIENT_ID='1098731892951-2e7jipvf33evn0a69tt0tje94oor54lm.apps.googleusercontent.com';
const DRIVE_FOLDER_ID='1sgTsZu2zzSBi-pz2M9pVBHtjCtQELYhk';
const PLAN_KEY='tpa_auditor_visit_plan_v2';
const SNAPSHOT_KEY='tpa_navigation_snapshot_v2';
const DB_NAME='tpa-navigation-db';
const DB_VERSION=1;
const DB_STORE='snapshots';
const $=id=>document.getElementById(id);
let jpRows=[],salesRows=[],productRows=[],stores=[];
let map=null,markers=[],driveTokenClient=null,lastAccessToken=null;
let auditPlan=[],pendingExternal=[],activeSourceYear=1405;
const DAY_KEYS=['Saturday','Sunday','Monday','Tuesday','Wednesday','Thursday','Friday'];
const DAY_FA=['شنبه','یکشنبه','دوشنبه','سه‌شنبه','چهارشنبه','پنجشنبه','جمعه'];
const FILTER_IDS=['areaFilter','routeFilter','personFilter','classFilter','dayFilter'];

function clean(v){return String(v??'').trim();}
function norm(v){return clean(v).replace(/\s/g,'').replace(/[۰-۹]/g,d=>'۰۱۲۳۴۵۶۷۸۹'.indexOf(d)).replace(/[٠-٩]/g,d=>'٠١٢٣٤٥٦٧٨٩'.indexOf(d)).replace(/\.0$/,'');}
function variants(v){const s=norm(v);if(!s)return[];return [...new Set([s,...(s.includes('|')?[s.split('|').pop()]:[])])];}
function first(row,keys){for(const k of keys){if(row&&row[k]!==undefined&&clean(row[k])!=='')return clean(row[k]);}return '';}
function num(v){if(typeof v==='number')return Number.isFinite(v)?v:0;const n=Number(clean(v).replace(/,/g,'').replace(/[۰-۹]/g,d=>'۰۱۲۳۴۵۶۷۸۹'.indexOf(d)));return Number.isFinite(n)?n:0;}
function coordOK(lat,lng){return lat>=-90&&lat<=90&&lng>=-180&&lng<=180&&(lat!==0||lng!==0);}
function esc(v){return String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
function statusText(row){return first(row,['Status','Order Status','OrderStatus','وضعیت سفارش']);}
function isExcludedStatus(v){return clean(v).toLowerCase()==='closed/expired';}
function setMessage(message,status){if($('message'))$('message').textContent=message;if(status&&$('status'))$('status').textContent=status;}
function readExcel(file,headerRow=0){return new Promise((resolve,reject)=>{const r=new FileReader();r.onload=e=>{try{const wb=XLSX.read(e.target.result,{type:'array',cellDates:true});const sh=wb.Sheets[wb.SheetNames[0]];resolve(XLSX.utils.sheet_to_json(sh,{range:headerRow,defval:'',raw:false}));}catch(err){reject(err);}};r.onerror=()=>reject(new Error('خواندن فایل ممکن نشد.'));r.readAsArrayBuffer(file);});}
function validateSourceRows(jp,sa,pm=[]){if(!Array.isArray(jp)||!jp.length)throw new Error('فایل Journey Plan داده ندارد.');if(!Array.isArray(sa)||!sa.length)throw new Error('فایل TPA داده ندارد.');const jpKeys=new Set(Object.keys(jp[0]||{})),saKeys=new Set(Object.keys(sa[0]||{}));if(!jpKeys.has('CustomerI'))throw new Error('ستون CustomerI در ردیف عنوان Journey Plan پیدا نشد؛ فایل یا ردیف عنوان نادرست است.');if(!saKeys.has('کد مشتری')&&!saKeys.has('Customer Code'))throw new Error('ستون کد مشتری در فایل TPA پیدا نشد؛ فایل یا ردیف عنوان نادرست است.');if(!saKeys.has('Status')&&!saKeys.has('وضعیت سفارش'))throw new Error('ستون Status در فایل TPA پیدا نشد.');if(pm.length&&!Object.keys(pm[0]||{}).includes('Item Code'))throw new Error('ساختار فایل Product Mapping معتبر نیست.');return true;}
function dayIsSet(v){const s=clean(v).toLowerCase();return !['','0','0.0','false','no','n','-','null'].includes(s);}
function todayKey(){const js=new Date().getDay();return ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'][js];}
function dateSortValue(v){return parseSourceDate(v)?.timestamp||0;}
function fileYear(){if(activeSourceYear)return activeSourceYear;const names=[...document.querySelectorAll('input[type=file]')].map(x=>x.files?.[0]?.name||'').join(' ');const m=names.match(/(?:14|20)\d{2}/);return m?Number(m[0]):1405;}
function jalaliToGregorian(jy,jm,jd){let gy=jy-979,jdn=365*gy+Math.floor(gy/33)*8+Math.floor((gy%33+3)/4)+78+jd+(jm<7?(jm-1)*31:(jm-7)*30+186);let gy2=1600+400*Math.floor(jdn/146097);jdn%=146097;let leap=true;if(jdn>=36525){jdn--;gy2+=100*Math.floor(jdn/36524);jdn%=36524;if(jdn>=365)jdn++;else leap=false;}gy2+=4*Math.floor(jdn/1461);jdn%=1461;if(jdn>=366){leap=false;jdn--;gy2+=Math.floor(jdn/365);jdn%=365;}const md=[31,leap?29:28,31,30,31,30,31,31,30,31,30,31];let gm=0;while(gm<12&&jdn>=md[gm])jdn-=md[gm++];return [gy2,gm+1,jdn+1];}
function parseSourceDate(value){const s=clean(value);if(!s)return null;let m=s.match(/^(\d{6})$/);if(m){const year=fileYear(),yy=Number(s.slice(0,2)),month=Number(s.slice(2,4)),day=Number(s.slice(4,6)),jy=Math.floor(year/100)*100+yy;const [gy,gm,gd]=jalaliToGregorian(jy,month,day);return {display:`${jy}/${String(month).padStart(2,'0')}/${String(day).padStart(2,'0')}`,timestamp:Date.UTC(gy,gm-1,gd)};}m=s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);if(m){const mm=Number(m[1]),dd=Number(m[2]),yy=Number(m[3]);const d=new Date(Date.UTC(yy,mm-1,dd));return {display:d.toLocaleDateString('fa-IR-u-ca-persian'),timestamp:d.getTime()};}const d=new Date(s);if(!Number.isNaN(d.getTime()))return {display:d.toLocaleDateString('fa-IR-u-ca-persian'),timestamp:d.getTime()};return {display:s,timestamp:0};}
function displayDate(v){return parseSourceDate(v)?.display||'—';}
function extractStoreName(address){const m=clean(address).match(/(سوپر\s*مارکت|سوپرمارکت|هایپر|سوپر|مارکت|فروشگاه|مواد\s*غذایی|ارزانسرای)\s*[^،,؛;]*/i);return m?m[0].trim():'';}
function formatMoney(v){return num(v).toLocaleString('fa-IR',{maximumFractionDigits:0})+' ریال';}
function productMapBuild(){const out=new Map();for(const r of productRows){const code=norm(first(r,['Item Code']));if(code)out.set(code,first(r,['Product  Name','Product Name','Item'])||first(r,['Item']));}return out;}

/* Excel loading */
async function loadManualFiles(){const jf=$('jpFile')?.files?.[0],sf=$('salesFile')?.files?.[0],mf=$('mappingFile')?.files?.[0];if(!jf||!sf){setMessage('برای بارگذاری دستی، فایل Journey Plan و گزارش فروش هر دو لازم هستند.');return;}const b=$('loadBtn');b.disabled=true;try{setMessage('در حال خواندن فایل‌ها...','در حال پردازش');const [jp,sa,pm]=await Promise.all([readExcel(jf,1),readExcel(sf,3),mf?readExcel(mf,0):Promise.resolve([])]);validateSourceRows(jp,sa,pm);jpRows=jp;salesRows=sa;productRows=pm;activeSourceYear=Number((sf.name.match(/(14\d{2})/)||[])[1]||1405);buildData();renderAll();setMessage(`بارگذاری موفق بود. ${stores.length.toLocaleString('fa-IR')} مشتری پردازش شد.`,'آماده');}catch(e){console.error(e);setMessage('خطا در بارگذاری: '+e.message,'خطا');}finally{b.disabled=false;}}

/* Data model: JP owns customer identity; TPA only enriches orders. */
function buildData(){
 const pmap=productMapBuild();
 const jpMap=new Map(),lookup=new Map();
 for(const row of jpRows){
  const custI=first(row,['CustomerI']),newCode=first(row,['New Customer Code']),oracle=first(row,['Oracle Code']);
  const id=custI||newCode||oracle;if(!id)continue;
  const lat=num(row['Latitude']),lng=num(row['Longitude']);
  const visitDays=DAY_KEYS.filter(d=>dayIsSet(row[d])&&d!=='Friday');
  const store={id,names:variants(custI).concat(variants(newCode),variants(oracle)),customerI:custI,name:first(row,['Customer Name','CustomerName','Name','نام مشتری','نام فروشگاه'])||custI,address:first(row,['Address','Customer Address','آدرس']),className:first(row,['Customer Class','CustomerClass','کلاس مشتری']),salesman:first(row,['Salesman Name','Salesman','نام ویزیتور']),routeName:first(row,['Route Name']),routeId:first(row,['RouteID','Route Id']),area:first(row,['Area Name','Area','ناحیه']),lat,lng,hasGPS:coordOK(lat,lng),visitDays,orders:[],source:'JP',hasSales:false};
  jpMap.set(norm(id),store);
  for(const v of store.names)if(v)lookup.set(v,store);
 }
 const orderGroups=new Map();
 for(const row of salesRows){
  if(isExcludedStatus(statusText(row)))continue;
  const code=first(row,['کد مشتری','Customer Code','CustomerI','Customer Code.1']);
  const match=variants(code).map(v=>lookup.get(v)).find(Boolean);
  const orderNo=first(row,['Order Number','Order No','OrderNumber','شماره سفارش']);
  const groupKey=(match?match.id:variants(code)[0]||'UNKNOWN')+'|'+(orderNo||'NO_ORDER_'+orderGroups.size);
  if(!orderGroups.has(groupKey))orderGroups.set(groupKey,{store:match,code,orderNo,rows:[],date:first(row,['Date','Order Date']),targetDate:first(row,['Target Delivery Date','TargetDeliveryDate']),invoiceDate:first(row,['Invoice Date','invoicedate']),status:statusText(row)});
  orderGroups.get(groupKey).rows.push(row);
 }
 const tpaOnly=new Map();
 for(const group of orderGroups.values()){
  const lines=group.rows.map(r=>{
   const itemCode=first(r,['Item Code','ItemCode','کد کالا']);
   const itemName=pmap.get(norm(itemCode))||first(r,['Item Name','Item Description','Description','نام کالا'])||itemCode;
   const qty=num(first(r,['Quantity','Qty','Order Quantity','تعداد']));
   const unitPrice=num(first(r,['Unit Price','Price','قیمت واحد']));
   return {itemCode,itemName,qty,unitPrice,amount:qty*unitPrice};
  });
  const order={orderNo:group.orderNo,date:group.date,targetDate:group.targetDate,invoiceDate:group.invoiceDate,status:group.status,lines,total:lines.reduce((s,x)=>s+x.amount,0),quantity:lines.reduce((s,x)=>s+x.qty,0),isOpen:!clean(group.invoiceDate)};
  if(group.store){group.store.orders.push(order);group.store.hasSales=true;}
  else{
   const k=norm(group.code)||'unknown-'+tpaOnly.size;
   if(!tpaOnly.has(k))tpaOnly.set(k,{id:group.code||k,names:variants(group.code),customerI:group.code,name:first(group.rows[0],['نام مشتری','Customer Name','CustomerName'])||group.code,address:first(group.rows[0],['آدرس','Address']),className:'',salesman:'',routeName:'',routeId:'',area:'',lat:0,lng:0,hasGPS:false,visitDays:[],orders:[],source:'TPA',hasSales:true});
   tpaOnly.get(k).orders.push(order);
  }
 }
 stores=[...jpMap.values(),...tpaOnly.values()];
 for(const s of stores){
  s.orders.sort((a,b)=>dateSortValue(b.date)-dateSortValue(a.date));
  s.orderCount=s.orders.length;
  s.salesTotal=s.orders.reduce((n,o)=>n+o.total,0);
  s.lastOrder=s.orders[0]||null;
  s.lastPurchaseItems=new Set(s.orders.slice(0,3).flatMap(o=>o.lines.map(l=>norm(l.itemCode)||l.itemName)));
  const allItems=new Map();
  for(const o of s.orders)for(const l of o.lines){const k=norm(l.itemCode)||l.itemName;if(!allItems.has(k))allItems.set(k,{...l,totalQty:0,lastSeen:o.date});allItems.get(k).totalQty+=l.qty;}
  s.inactiveProducts=[...allItems.values()].filter(x=>!s.lastPurchaseItems.has(norm(x.itemCode)||x.itemName));
 }
 saveSnapshot().catch(console.warn);
}

/* IndexedDB persistence */
function openDB(){return new Promise((resolve,reject)=>{if(!('indexedDB'in window)){reject(new Error('IndexedDB در این مرورگر در دسترس نیست.'));return;}const req=indexedDB.open(DB_NAME,DB_VERSION);req.onupgradeneeded=()=>{const db=req.result;if(!db.objectStoreNames.contains(DB_STORE))db.createObjectStore(DB_STORE,{keyPath:'id'});};req.onsuccess=()=>resolve(req.result);req.onerror=()=>reject(req.error);});}
async function idbPut(id,value){const db=await openDB();return new Promise((resolve,reject)=>{const tx=db.transaction(DB_STORE,'readwrite');tx.objectStore(DB_STORE).put({id,value,savedAt:Date.now()});tx.oncomplete=()=>{db.close();resolve(true);};tx.onerror=()=>{db.close();reject(tx.error);};});}
async function idbGet(id){const db=await openDB();return new Promise((resolve,reject)=>{const tx=db.transaction(DB_STORE,'readonly'),req=tx.objectStore(DB_STORE).get(id);req.onsuccess=()=>{db.close();resolve(req.result?.value||null);};req.onerror=()=>{db.close();reject(req.error);};});}
async function saveSnapshot(){if(!jpRows.length||!salesRows.length)return;const payload={jpRows,salesRows,productRows,activeSourceYear,savedAt:Date.now()};try{await idbPut('main',payload);localStorage.setItem(SNAPSHOT_KEY,JSON.stringify({savedAt:payload.savedAt,version:2}));}catch(e){console.warn('ذخیره آفلاین انجام نشد',e);}}
async function restoreSnapshot(){try{const data=await idbGet('main');if(!data)return false;jpRows=data.jpRows||[];salesRows=data.salesRows||[];productRows=data.productRows||[];activeSourceYear=data.activeSourceYear||1405;validateSourceRows(jpRows,salesRows,productRows);buildData();renderAll();setMessage('اطلاعات ذخیره‌شده آفلاین بازیابی شد.','حالت آفلاین');return true;}catch(e){console.warn(e);return false;}}

/* Filter logic — cascading AND */
function getFilterValues(){return Object.fromEntries(FILTER_IDS.map(id=>[id,$(id)?.value||'all']));}
function matchesFilters(store,filters=getFilterValues()){
 if(filters.area!=='all'&&store.area!==filters.area)return false;
 if(filters.route!=='all'&&store.routeName!==filters.route)return false;
 if(filters.person!=='all'&&store.salesman!==filters.person)return false;
 if(filters.class!=='all'&&store.className!==filters.class)return false;
 if(filters.day!=='all'&&filters.day!=='Today'&&!store.visitDays.includes(filters.day))return false;
 if(filters.day==='Today'&&!store.visitDays.includes(todayKey()))return false;
 if(filters.sales==='with'&&!store.hasSales)return false;
 if(filters.sales==='without'&&store.hasSales)return false;
 if(filters.gps==='yes'&&!store.hasGPS)return false;
 if(filters.gps==='no'&&store.hasGPS)return false;
 return true;
}
function filteredStores(){return stores.filter(s=>matchesFilters(s));}
function updateSelectOptions(id,values,current){const el=$(id);if(!el)return;const label=el.options[0]?.textContent||'همه';const unique=[...new Set(values.filter(Boolean))].sort((a,b)=>String(a).localeCompare(String(b),'fa'));el.innerHTML='';const all=document.createElement('option');all.value='all';all.textContent=label;el.appendChild(all);for(const v of unique){const o=document.createElement('option');o.value=v;o.textContent=v;el.appendChild(o);}el.value=unique.includes(current)?current:'all';}
function refreshCascadingFilters(changedId){
 const current=getFilterValues();
 const fields=[['areaFilter','area'],['routeFilter','routeName'],['personFilter','salesman'],['classFilter','className']];
 for(const [id,key] of fields){const other={...current};delete other[key];const candidates=stores.filter(s=>matchesFilters(s,other));updateSelectOptions(id,candidates.map(s=>s[key]),current[key]);}
}
function applyFilters(){refreshCascadingFilters();renderStoreList();renderMarkers();updateStats();}
function resetFilters(){for(const id of FILTER_IDS){if($(id))$(id).value='all';}if($('salesFilter'))$('salesFilter').value='all';if($('gpsFilter'))$('gpsFilter').value='all';applyFilters();}
function updateStats(){const list=filteredStores();if($('storeCount'))$('storeCount').textContent=list.length.toLocaleString('fa-IR');if($('salesCount'))$('salesCount').textContent=list.filter(s=>s.hasSales).length.toLocaleString('fa-IR');if($('noSalesCount'))$('noSalesCount').textContent=list.filter(s=>!s.hasSales).length.toLocaleString('fa-IR');}

/* Map and store list */
function initMap(){if(!window.L||!$('map'))return;map=L.map('map',{zoomControl:false}).setView([35.7,51.4],10);L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',{maxZoom:19,attribution:'© OpenStreetMap'}).addTo(map);L.control.zoom({position:'bottomright'}).addTo(map);}
function markerColor(s){return s.source==='TPA'?'#2878e5':s.hasSales?'#19a56b':'#e79a2e';}
function renderMarkers(){if(!map)return;markers.forEach(m=>map.removeLayer(m));markers=[];for(const s of filteredStores()){if(!s.hasGPS)continue;const marker=L.circleMarker([s.lat,s.lng],{radius:8,color:'#fff',weight:2,fillColor:markerColor(s),fillOpacity:1}).addTo(map);marker.bindTooltip(esc(s.name));marker.on('click',()=>showStoreDetails(s.id));markers.push(marker);}}
function renderStoreList(){const root=$('storeList');if(!root)return;const q=norm($('storeSearch')?.value||'');const list=filteredStores().filter(s=>!q||[s.name,s.id,s.customerI,s.address].some(v=>norm(v).includes(q)));root.innerHTML=list.map(s=>`<button class="store-row" type="button" data-store-id="${esc(s.id)}"><span class="store-dot" style="background:${markerColor(s)}"></span><span class="store-main"><strong>${esc(s.name||s.id)}</strong><small>${esc(s.address||'آدرس ثبت نشده')}</small></span><span class="store-meta">${s.hasSales?'دارای فروش':'بدون فروش'}${s.hasGPS?'':' · بدون GPS'}</span></button>`).join('')||'<div class="empty-state">موردی پیدا نشد.</div>';root.querySelectorAll('[data-store-id]').forEach(b=>b.addEventListener('click',()=>showStoreDetails(b.dataset.storeId)));}
function findStore(id){return stores.find(s=>String(s.id)===String(id));}
function showStoreDetails(id){
 const s=findStore(id);if(!s)return;
 const root=$('modalContent');if(!root)return;
 const orders=s.orders.slice().sort((a,b)=>dateSortValue(b.date)-dateSortValue(a.date));
 root.innerHTML=`
 <div class="detail-head"><div><h2>${esc(s.name||s.id)}</h2><p>${esc(s.address||'آدرس ثبت نشده')}</p></div><button type="button" class="icon-btn" id="closeDetail">✕</button></div>
 <div class="detail-grid">
  <div><small>کد مشتری</small><strong>${esc(s.id)}</strong></div>
  <div><small>کلاس مشتری</small><strong>${esc(s.className||'—')}</strong></div>
  <div><small>ویزیتور</small><strong>${esc(s.salesman||'—')}</strong></div>
  <div><small>مسیر</small><strong>${esc(s.routeName||s.routeId||'—')}</strong></div>
  <div><small>تعداد سفارش</small><strong>${orders.length.toLocaleString('fa-IR')}</strong></div>
  <div><small>مجموع مبلغ</small><strong>${formatMoney(s.salesTotal)}</strong></div>
 </div>
 <div class="detail-section"><h3>روزهای ویزیت</h3><p>${s.visitDays.map(d=>DAY_FA[DAY_KEYS.indexOf(d)]).join('، ')||'ثبت نشده'}</p></div>
 <div class="detail-section"><h3>سفارش‌ها</h3>
 ${orders.length?orders.map(o=>`<article class="order-card"><header><strong>سفارش ${esc(o.orderNo||'—')}</strong><span>${esc(o.status||'')}</span></header><div class="order-dates"><span>تاریخ سفارش: ${displayDate(o.date)}</span><span>تحویل هدف: ${displayDate(o.targetDate)}</span><span>فاکتور: ${displayDate(o.invoiceDate)}</span></div><div class="order-lines">${o.lines.map(l=>`<div><span>${esc(l.itemName)}</span><span>${num(l.qty).toLocaleString('fa-IR')} × ${formatMoney(l.unitPrice)}</span></div>`).join('')}</div><footer><strong>جمع سفارش: ${formatMoney(o.total)}</strong><span>${o.isOpen?'باز':'فاکتور شده'}</span></footer></article>`).join(''):'<p>سفارشی ثبت نشده است.</p>'}
 </div>
 <div class="detail-section"><h3>کالاهای غیرفعال در سه سفارش اخیر</h3><div class="tag-list">${s.inactiveProducts.map(x=>`<span class="tag">${esc(x.itemName)}</span>`).join('')||'موردی شناسایی نشد.'}</div></div>
 <div class="detail-actions"><button type="button" id="addToAudit" class="primary-btn">افزودن به برنامه Auditor</button>${s.hasGPS?`<button type="button" id="openGoogleMaps" class="secondary-btn">مسیریابی</button>`:''}</div>`;
 const modal=$('detailModal');if(modal)modal.hidden=false;
 $('closeDetail')?.addEventListener('click',closeDetails);
 $('addToAudit')?.addEventListener('click',()=>{addAuditStore(s);closeDetails();});
 $('openGoogleMaps')?.addEventListener('click',()=>window.open(`https://www.google.com/maps/dir/?api=1&destination=${s.lat},${s.lng}`,'_blank','noopener'));
}
function closeDetails(){const modal=$('detailModal');if(modal)modal.hidden=true;}
function renderAll(){refreshCascadingFilters();renderStoreList();renderMarkers();updateStats();renderAuditPlan();}

/* Auditor plan */
function loadAuditPlan(){try{auditPlan=JSON.parse(localStorage.getItem(PLAN_KEY)||'[]');if(!Array.isArray(auditPlan))auditPlan=[];}catch{auditPlan=[];}}
function persistAuditPlan(){localStorage.setItem(PLAN_KEY,JSON.stringify(auditPlan));renderAuditPlan();}
function addAuditStore(s,extra={}){if(!s||!s.hasGPS)return setMessage('فروشگاه بدون مختصات GPS قابل افزودن به برنامه Auditor نیست.');const date=extra.visitDate||$('auditDate')?.value||new Date().toISOString().slice(0,10);if(auditPlan.some(x=>x.customerId===s.id&&x.visitDate===date)){setMessage('این فروشگاه برای این تاریخ قبلاً به برنامه اضافه شده است.');return;}auditPlan.push({visitDate:date,customerId:s.id,storeName:s.name,customerName:s.name,address:s.address,routeName:s.routeName,latitude:s.lat,longitude:s.lng,priority:extra.priority||'Normal',selectionReason:extra.selectionReason||'Manual selection',auditorComment:extra.auditorComment||'',selectionSource:extra.selectionSource||'Internal',status:'Pending'});persistAuditPlan();setMessage('فروشگاه به برنامه Auditor اضافه شد.');}
function renderAuditPlan(){
 const root=$('auditPlanList');if(!root)return;
 const date=$('auditDate')?.value||'';
 const list=auditPlan.filter(x=>!date||x.visitDate===date);
 root.innerHTML=list.map((x,i)=>`<article class="audit-row" data-audit-index="${auditPlan.indexOf(x)}"><div class="audit-row-main"><strong>${esc(x.storeName||x.customerName)}</strong><small>${esc(x.address||'')} · ${esc(x.routeName||'بدون مسیر')}</small><small>${esc(x.visitDate)} · ${esc(x.priority)} · ${esc(x.selectionSource)}</small></div><select class="audit-status" data-status-index="${auditPlan.indexOf(x)}"><option value="Pending" ${x.status==='Pending'?'selected':''}>در انتظار</option><option value="Completed" ${x.status==='Completed'?'selected':''}>انجام شد</option><option value="Not completed" ${x.status==='Not completed'?'selected':''}>انجام نشد</option></select><div class="audit-row-actions"><button type="button" data-up="${auditPlan.indexOf(x)}" aria-label="بالا">↑</button><button type="button" data-down="${auditPlan.indexOf(x)}" aria-label="پایین">↓</button><button type="button" data-remove="${auditPlan.indexOf(x)}" aria-label="حذف">حذف</button></div></article>`).join('')||'<div class="empty-state">برای این تاریخ برنامه‌ای ثبت نشده است.</div>';
 root.querySelectorAll('[data-status-index]').forEach(el=>el.addEventListener('change',()=>{auditPlan[Number(el.dataset.statusIndex)].status=el.value;persistAuditPlan();}));
 root.querySelectorAll('[data-remove]').forEach(el=>el.addEventListener('click',()=>{auditPlan.splice(Number(el.dataset.remove),1);persistAuditPlan();}));
 root.querySelectorAll('[data-up]').forEach(el=>el.addEventListener('click',()=>moveAudit(Number(el.dataset.up),-1)));
 root.querySelectorAll('[data-down]').forEach(el=>el.addEventListener('click',()=>moveAudit(Number(el.dataset.down),1)));
}
function moveAudit(index,delta){const next=index+delta;if(next<0||next>=auditPlan.length)return;[auditPlan[index],auditPlan[next]]=[auditPlan[next],auditPlan[index]];persistAuditPlan();}
function addAuditFromSearch(){const q=norm($('auditSearch')?.value||'');if(!q)return;const s=stores.find(x=>[x.name,x.id,x.customerI].some(v=>norm(v).includes(q)));if(!s){setMessage('فروشگاهی با این مشخصات پیدا نشد.');return;}addAuditStore(s);}
function exportAuditExcel(){if(!window.XLSX)return;const rows=auditPlan.map(x=>({'Visit Date':x.visitDate,'Customer ID':x.customerId,'Store Name':x.storeName,'Customer Name':x.customerName,'Address':x.address,'Route Name':x.routeName,'Latitude':x.latitude,'Longitude':x.longitude,'Priority':x.priority,'Selection Reason':x.selectionReason,'Auditor Comment':x.auditorComment,'Selection Source':x.selectionSource,'Status':x.status}));const ws=XLSX.utils.json_to_sheet(rows);const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Audit Visit Plan');XLSX.writeFile(wb,'Audit_Visit_Plan.xlsx');}

/* External plan import: suggestions are staged, not silently merged. */
async function importAuditExcel(file){if(!file)return;try{const rows=await readExcel(file,0);const required=['Visit Date','Customer ID','Store Name'];const headers=Object.keys(rows[0]||{});const missing=required.filter(x=>!headers.includes(x));if(missing.length)throw new Error('ستون‌های ضروری وجود ندارند: '+missing.join(', '));const seen=new Set();pendingExternal=[];for(const r of rows){const id=clean(r['Customer ID']),date=clean(r['Visit Date']),key=id+'|'+date;const s=stores.find(x=>variants(x.id).includes(norm(id)));const validDate=/^\d{4}-\d{2}-\d{2}$/.test(date)||/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4}$/.test(date);const errors=[];if(!s)errors.push('کد مشتری در داده‌های فعلی پیدا نشد');if(!validDate)errors.push('فرمت تاریخ معتبر نیست');if(seen.has(key)||auditPlan.some(x=>x.customerId===id&&x.visitDate===date))errors.push('ردیف تکراری است');if(!num(r.Latitude)||!num(r.Longitude)||!coordOK(num(r.Latitude),num(r.Longitude)))errors.push('مختصات GPS معتبر نیست');seen.add(key);pendingExternal.push({row:{visitDate:date,customerId:id,storeName:r['Store Name']||s?.name||'',customerName:r['Customer Name']||s?.name||'',address:r.Address||s?.address||'',routeName:r['Route Name']||s?.routeName||'',latitude:num(r.Latitude)||s?.lat,longitude:num(r.Longitude)||s?.lng,priority:r.Priority||'Normal',selectionReason:r['Selection Reason']||'External suggestion',auditorComment:r['Auditor Comment']||'',selectionSource:'External',status:'Pending'},store:s,errors,approved:false});}renderExternalSuggestions();}catch(e){setMessage('خطا در ورود برنامه: '+e.message);}}
function renderExternalSuggestions(){const root=$('externalSuggestions');if(!root)return;root.innerHTML=pendingExternal.map((x,i)=>`<article class="external-row"><label><input type="checkbox" data-external="${i}" ${x.approved?'checked':''} ${x.errors.length?'disabled':''}><strong>${esc(x.row.storeName)}</strong></label><small>${esc(x.row.visitDate)} · ${esc(x.row.customerId)}</small><p>${x.errors.map(esc).join('؛ ')||'آماده تأیید'}</p></article>`).join('')||'<p>پیشنهاد خارجی برای بررسی وجود ندارد.</p>';}
function approveExternalSuggestions(){document.querySelectorAll('[data-external]:checked').forEach(el=>{const x=pendingExternal[Number(el.dataset.external)];if(x&&!x.errors.length)x.approved=true;});let count=0;for(const x of pendingExternal){if(x.approved&&!auditPlan.some(a=>a.customerId===x.row.customerId&&a.visitDate===x.row.visitDate)){auditPlan.push(x.row);count++;}}persistAuditPlan();setMessage(`${count} پیشنهاد خارجی تأیید و وارد برنامه شد.`);pendingExternal=[];renderExternalSuggestions();}
/* Google Drive integration */
function initGoogleDrive(){if(!window.google?.accounts?.oauth2)return;driveTokenClient=google.accounts.oauth2.initTokenClient({client_id:GOOGLE_CLIENT_ID,scope:'https://www.googleapis.com/auth/drive.readonly',callback:async response=>{if(response.error){setMessage('اتصال Google Drive ناموفق بود: '+response.error);return;}lastAccessToken=response.access_token;await syncFromDrive(lastAccessToken);}});}
function connectGoogleDrive(){if(!driveTokenClient){setMessage('سرویس ورود Google بارگذاری نشده است. اتصال اینترنت و بارگذاری کتابخانه Google را بررسی کنید.');return;}driveTokenClient.requestAccessToken({prompt:'consent'});}
async function driveList(token){const q=encodeURIComponent(`'${DRIVE_FOLDER_ID}' in parents and trashed=false`);const res=await fetch(`https://www.googleapis.com/drive/v3/files?q=${q}&fields=files(id,name,mimeType,modifiedTime,size)&pageSize=100&orderBy=modifiedTime desc`,{headers:{Authorization:'Bearer '+token}});if(!res.ok)throw new Error('دریافت فهرست Google Drive ناموفق بود: '+res.status);return (await res.json()).files||[];}
async function driveDownload(file,token){const res=await fetch(`https://www.googleapis.com/drive/v3/files/${encodeURIComponent(file.id)}?alt=media`,{headers:{Authorization:'Bearer '+token}});if(!res.ok)throw new Error('دانلود فایل '+file.name+' ناموفق بود.');return await res.arrayBuffer();}
function newestFile(files,prefix){return files.filter(f=>f.name.toLowerCase().startsWith(prefix.toLowerCase())&&f.name.toLowerCase().endsWith('.xlsx')).sort((a,b)=>new Date(b.modifiedTime)-new Date(a.modifiedTime))[0]||null;}
async function parseExcelBuffer(buffer,headerRow){const wb=XLSX.read(buffer,{type:'array',cellDates:true});const sh=wb.Sheets[wb.SheetNames[0]];return XLSX.utils.sheet_to_json(sh,{range:headerRow,defval:'',raw:false});}
async function syncFromDrive(token){
 try{
  setMessage('در حال بررسی فایل‌های Google Drive...','همگام‌سازی');
  const files=await driveList(token);
  const jf=newestFile(files,'JP_'),sf=newestFile(files,'TPA_Alborz'),mf=newestFile(files,'Product Mapping');
  if(!jf||!sf)throw new Error('فایل معتبر JP_ و TPA_Alborz در پوشه تعیین‌شده پیدا نشد.');
  const [jb,sb,mb]=await Promise.all([driveDownload(jf,token),driveDownload(sf,token),mf?driveDownload(mf,token):Promise.resolve(null)]);
  const [jp,sa,pm]=await Promise.all([parseExcelBuffer(jb,1),parseExcelBuffer(sb,3),mb?parseExcelBuffer(mb,0):Promise.resolve([])]);
  validateSourceRows(jp,sa,pm);
  // Replace in-memory and offline data only after all source files pass validation.
  const previous={jpRows,salesRows,productRows,activeSourceYear};
  try{jpRows=jp;salesRows=sa;productRows=pm;activeSourceYear=Number((sf.name.match(/(14\d{2})/)||[])[1]||1405);buildData();await saveSnapshot();}
  catch(err){jpRows=previous.jpRows;salesRows=previous.salesRows;productRows=previous.productRows;activeSourceYear=previous.activeSourceYear;buildData();throw err;}
  renderAll();setMessage(`همگام‌سازی موفق: ${stores.length.toLocaleString('fa-IR')} مشتری از Drive بارگذاری شد.`,'همگام');
 }catch(e){console.error(e);setMessage('همگام‌سازی انجام نشد؛ اطلاعات قبلی حفظ شد. '+e.message,'خطا');}
}

/* Navigation and page switching */
function showPage(pageId){
 document.querySelectorAll('.page').forEach(p=>p.classList.toggle('active',p.id===pageId));
 document.querySelectorAll('.bottom-nav [data-page]').forEach(b=>b.classList.toggle('active',b.dataset.page===pageId));
 if(pageId==='mapPage'&&map)setTimeout(()=>map.invalidateSize(),150);
}
function bindNavigation(){document.querySelectorAll('.bottom-nav [data-page]').forEach(b=>b.addEventListener('click',()=>showPage(b.dataset.page)));}

/* Theme and language */
function setTheme(theme){document.documentElement.dataset.theme=theme;localStorage.setItem('tpa_theme',theme);const b=$('themeToggle');if(b)b.setAttribute('aria-pressed',String(theme==='dark'));}
function toggleTheme(){setTheme(document.documentElement.dataset.theme==='dark'?'light':'dark');}
function setLanguage(lang){document.documentElement.lang=lang;document.documentElement.dir=lang==='fa'?'rtl':'ltr';localStorage.setItem('tpa_language',lang);document.querySelectorAll('[data-fa][data-en]').forEach(el=>{el.textContent=lang==='fa'?el.dataset.fa:el.dataset.en;});}
function toggleLanguage(){setLanguage(document.documentElement.lang==='fa'?'en':'fa');}

/* Search and map controls */
function bindSearch(){ $('storeSearch')?.addEventListener('input',renderStoreList);$('auditSearchBtn')?.addEventListener('click',addAuditFromSearch);$('storeSearchClear')?.addEventListener('click',()=>{if($('storeSearch'))$('storeSearch').value='';renderStoreList();});}
function bindFilters(){for(const id of [...FILTER_IDS,'salesFilter','gpsFilter'])$(id)?.addEventListener('change',()=>{applyFilters();});$('resetFilters')?.addEventListener('click',resetFilters);}
function bindMapActions(){$('locateMe')?.addEventListener('click',()=>{if(!navigator.geolocation)return setMessage('موقعیت‌یابی در این مرورگر پشتیبانی نمی‌شود.');navigator.geolocation.getCurrentPosition(p=>map?.setView([p.coords.latitude,p.coords.longitude],15),e=>setMessage('دریافت موقعیت ناموفق بود: '+e.message),{enableHighAccuracy:true,timeout:12000});});}

/* Excel plan events */
function bindAuditActions(){
 $('exportAudit')?.addEventListener('click',exportAuditExcel);
 $('addAuditSearch')?.addEventListener('click',addAuditFromSearch);
 $('importAuditFile')?.addEventListener('change',e=>importAuditExcel(e.target.files?.[0]));
 $('approveExternal')?.addEventListener('click',approveExternalSuggestions);
 $('auditDate')?.addEventListener('change',renderAuditPlan);
 $('auditDateToday')?.addEventListener('click',()=>{if($('auditDate'))$('auditDate').value=new Date().toISOString().slice(0,10);renderAuditPlan();});
 $('clearAudit')?.addEventListener('click',()=>{if(confirm('تمام برنامه Auditor پاک شود؟')){auditPlan=[];persistAuditPlan();}});
}
/* UI initialization */
function bindFileActions(){$('loadBtn')?.addEventListener('click',loadManualFiles);$('driveConnect')?.addEventListener('click',connectGoogleDrive);}
function bindSettings(){ $('themeToggle')?.addEventListener('click',toggleTheme);$('languageToggle')?.addEventListener('click',toggleLanguage);const theme=localStorage.getItem('tpa_theme')||'light';const lang=localStorage.getItem('tpa_language')||'fa';setTheme(theme);setLanguage(lang);}
function bindModal(){ $('detailModal')?.addEventListener('click',e=>{if(e.target.id==='detailModal')closeDetails();});document.addEventListener('keydown',e=>{if(e.key==='Escape')closeDetails();});}
function bindInstallPrompt(){let deferred=null;window.addEventListener('beforeinstallprompt',e=>{e.preventDefault();deferred=e;const b=$('installApp');if(b)b.hidden=false;});$('installApp')?.addEventListener('click',async()=>{if(!deferred)return;deferred.prompt();await deferred.userChoice;deferred=null;$('installApp').hidden=true;});}
function registerServiceWorker(){if('serviceWorker'in navigator&&location.protocol!=='file:')navigator.serviceWorker.register('./sw.js').catch(e=>console.warn('Service worker registration failed',e));}
function bindAll(){bindNavigation();bindSearch();bindFilters();bindMapActions();bindFileActions();bindAuditActions();bindSettings();bindModal();bindInstallPrompt();}
async function initApp(){
 loadAuditPlan();
 initMap();
 bindAll();
 initGoogleDrive();
 registerServiceWorker();
 const restored=await restoreSnapshot();
 if(!restored)setMessage('برای شروع، فایل‌های Journey Plan و TPA را بارگذاری کنید.','آماده بارگذاری');
 if(navigator.onLine===false&&$('status'))$('status').textContent='آفلاین';
 window.addEventListener('online',()=>{if($('status'))$('status').textContent='آنلاین';});
 window.addEventListener('offline',()=>{if($('status'))$('status').textContent='آفلاین';});
}
document.addEventListener('DOMContentLoaded',initApp);