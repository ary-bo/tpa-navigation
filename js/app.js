'use strict';

/* TPA Navigation PWA — core application */
const GOOGLE_CLIENT_ID='1098731892951-2e7jipvf33evn0a69tt0tje94oor54lm.apps.googleusercontent.com';
const DRIVE_FOLDER_ID='1sgTsZu2zzSBi-pz2M9pVBHtjCtQELYhk';
const PLAN_KEY='tpa_auditor_visit_plan_v2';
const SNAPSHOT_KEY='tpa_navigation_snapshot_v2';
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
function dayIsSet(v){const s=clean(v).toLowerCase();return !['','0','0.0','false','no','n','-','null'].includes(s);}
function todayKey(){const js=new Date().getDay();return ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'][js];}
function dateSortValue(v){return parseSourceDate(v)?.timestamp||0;}
function fileYear(){if(activeSourceYear)return activeSourceYear;const names=[...document.querySelectorAll('input[type=file]')].map(x=>x.files?.[0]?.name||'').join(' ');const m=names.match(/(?:14|20)\d{2}/);return m?Number(m[0]):1405;}
function parseSourceDate(value){const s=clean(value);if(!s)return null;let m=s.match(/^(\d{6})$/);if(m){const year=fileYear(),yy=Number(s.slice(0,2)),month=Number(s.slice(2,4)),day=Number(s.slice(4,6)),jy=Math.floor(year/100)*100+yy;return {display:`${jy}/${String(month).padStart(2,'0')}/${String(day).padStart(2,'0')}`,timestamp:Date.UTC(jy,month-1,day)};}m=s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);if(m){const mm=Number(m[1]),dd=Number(m[2]),yy=Number(m[3]);return {display:`${String(mm).padStart(2,'0')}/${String(dd).padStart(2,'0')}/${yy}`,timestamp:Date.UTC(yy,mm-1,dd)};}const d=new Date(s);if(!Number.isNaN(d.getTime()))return {display:d.toLocaleDateString('fa-IR'),timestamp:d.getTime()};return {display:s,timestamp:0};}
function displayDate(v){return parseSourceDate(v)?.display||'—';}
function extractStoreName(address){const m=clean(address).match(/(سوپر\s*مارکت|سوپرمارکت|هایپر|سوپر|مارکت|فروشگاه|مواد\s*غذایی|ارزانسرای)\s*[^،,؛;]*/i);return m?m[0].trim():'';}
function formatMoney(v){return num(v).toLocaleString('fa-IR',{maximumFractionDigits:0})+' ریال';}
function productMapBuild(){const out=new Map();for(const r of productRows){const code=norm(first(r,['Item Code']));if(code)out.set(code,first(r,['Product  Name','Product Name','Item'])||first(r,['Item']));}return out;}

/* Excel loading */
async function loadManualFiles(){const jf=$('jpFile')?.files?.[0],sf=$('salesFile')?.files?.[0],mf=$('mappingFile')?.files?.[0];if(!jf||!sf){setMessage('برای بارگذاری دستی، فایل Journey Plan و گزارش فروش هر دو لازم هستند.');return;}const b=$('loadBtn');b.disabled=true;try{setMessage('در حال خواندن فایل‌ها...','در حال پردازش');const [jp,sa,pm]=await Promise.all([readExcel(jf,0),readExcel(sf,1),mf?readExcel(mf,0):Promise.resolve([])]);if(!jp.length||!sa.length)throw new Error('یکی از فایل‌ها خالی است یا ردیف عنوان ستون‌ها درست تشخیص داده نشده است.');jpRows=jp;salesRows=sa;productRows=pm;activeSourceYear=Number((sf.name.match(/(14\d{2})/)||[])[1]||1405);buildData();renderAll();setMessage(`بارگذاری موفق بود. ${stores.length.toLocaleString('fa-IR')} مشتری پردازش شد.`,'آماده');}catch(e){console.error(e);setMessage('خطا در بارگذاری: '+e.message,'خطا');}finally{b.disabled=false;}}

/* Data model: JP owns customer identity; TPA only enriches orders. */
function buildData(){
 const pmap=productMapBuild();
 const jpMap=new Map(),lookup=new Map();
 for(const row of jpRows){
  const custI=first(row,['CustomerI']),newCode=first(row,['New Customer Code']),oracle=first(row,['Oracle Code']);
  const id=custI||newCode||oracle;if(!id)continue;
  const lat=num(row['Latitude']),lng=num(row['Longitude']);
  const visitDays=DAY_KEYS.filter(d=>d!=='Friday'&&dayIsSet(row[d]));
  const s={id,code:newCode||custI||oracle,name:first(row,['Customer Name','DescriptionAR','DescriptionE'])||'فروشگاه بدون نام',address:first(row,['Address',' Address']),storeName:extractStoreName(first(row,['Address',' Address'])),area:first(row,['Area Name','TerritoryNo']),routeName:first(row,['Route Name','RouteID','Route Id']),customerClass:first(row,['Customer Class','Outlet Class']),salesman:first(row,['Salesman Name']),salesmanCode:first(row,['Salesman Code']),lat,lng,hasCoords:coordOK(lat,lng),inJP:true,visitDays,today:visitDays.includes(todayKey()),orders:[],products:[],hasSales:false};
  jpMap.set(norm(id),s);
  for(const k of [custI,newCode,oracle])for(const v of variants(k))if(v)lookup.set(v,s);
 }
 const orderGroups=new Map(),tpaOnly=new Map();
 for(const row of salesRows){
  const status=statusText(row);if(isExcludedStatus(status))continue;
  const customer=first(row,['کد مشتری','Customer Code']),account=first(row,['Account Number']);
  const match=[...variants(customer),...variants(account)].map(k=>lookup.get(k)).find(Boolean)||null;
  const customerId=customer||account;if(!customerId)continue;
  let store=match;
  if(!store){const k=norm(customerId);if(!tpaOnly.has(k))tpaOnly.set(k,{id:customerId,code:customerId,name:first(row,['نام مشتری'])||'مشتری گزارش فروش',address:first(row,['آدرس']),storeName:extractStoreName(first(row,['آدرس'])),area:'',routeName:first(row,['نام مسیر']),customerClass:'',salesman:'',salesmanCode:'',lat:0,lng:0,hasCoords:false,inJP:false,visitDays:[],today:false,orders:[],products:[],hasSales:true});store=tpaOnly.get(k);}else store=match;
  const orderNo=first(row,['شماره سفارش','Order Number','Order No','OrderNumber'])||'بدون شماره';
  const orderDate=first(row,['تاریخ','Date','Order Date']),invoiceDate=first(row,['invoicedate','Invoice Date']),deliveryDate=first(row,['Target Deilvery Date','Target Delivery Date','TargetDeliveryDate']);
  const itemCode=first(row,['کد کالا','Item Code']);
  const itemName=pmap.get(norm(itemCode))||first(row,['نام کالا','Item','Product Name'])||itemCode||'محصول نامشخص';
  const quantity=num(first(row,['تعداد','Quantity']));
  const unitPrice=num(first(row,['قیمت واحد','Unit Price']));
  const amount=unitPrice*quantity;
  const groupKey=norm(customerId)+'|'+orderNo;
  if(!orderGroups.has(groupKey))orderGroups.set(groupKey,{orderNo,status,orderDate,invoiceDate,deliveryDate,amount:0,products:[],customerId,store});
  const order=orderGroups.get(groupKey);
  if(!order.status)order.status=status;
  if(!order.orderDate)order.orderDate=orderDate;
  if(!order.invoiceDate)order.invoiceDate=invoiceDate;
  if(!order.deliveryDate)order.deliveryDate=deliveryDate;
  order.amount+=amount;
  order.products.push({code:itemCode,name:itemName,quantity,unitPrice,amount});
 }
 for(const order of orderGroups.values()){
  order.products=order.products.filter(p=>p.code||p.name);
  order.orderDateDisplay=displayDate(order.orderDate);order.invoiceDateDisplay=displayDate(order.invoiceDate);order.deliveryDateDisplay=displayDate(order.deliveryDate);order.isOverdue=!order.invoiceDate&&!!order.deliveryDate&&dateSortValue(order.deliveryDate)>0&&dateSortValue(order.deliveryDate)<Date.now();order.store.orders.push(order);
 }
 stores=[...jpMap.values(),...tpaOnly.values()];
 for(const s of stores){
  s.orders.sort((a,b)=>dateSortValue(b.orderDate)-dateSortValue(a.orderDate));
  s.products=s.orders.flatMap(o=>o.products);
  s.hasSales=s.orders.length>0;
  s.openOrders=s.orders.filter(o=>!o.invoiceDate);
  s.invoicedTotal=s.orders.filter(o=>!!o.invoiceDate).reduce((n,o)=>n+o.amount,0);
  s.nonInvoicedTotal=s.orders.filter(o=>!o.invoiceDate).reduce((n,o)=>n+o.amount,0);
  s.lastPurchase=s.orders[0]?.orderDate||'';
  const recentOrderNos=new Set(s.orders.slice(0,3).map(o=>o.orderNo));
  const recentCodes=new Set(s.orders.filter(o=>recentOrderNos.has(o.orderNo)).flatMap(o=>o.products.map(p=>norm(p.code))));
  const allCodes=new Set(s.orders.flatMap(o=>o.products.map(p=>norm(p.code))));
  s.inactiveProducts=[...allCodes].filter(code=>code&&!recentCodes.has(code)).map(code=>{const p=s.orders.flatMap(o=>o.products).find(x=>norm(x.code)===code);return p?{code:p.code,name:p.name}:null;}).filter(Boolean);
 }
 saveOfflineSnapshot();
}

/* Leaflet map */
function initMap(){if(map)return;if(!window.L){setMessage('کتابخانه نقشه بارگذاری نشده است.');return;}map=L.map('map',{zoomControl:false}).setView([35.7,51.4],10);L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',{maxZoom:19,attribution:'© OpenStreetMap'}).addTo(map);L.control.zoom({position:'bottomright'}).addTo(map);}
function markerColor(s){if(!s.inJP)return '#2583e8';return s.hasSales?'#20a464':'#ed9b27';}
function renderMarkers(list){initMap();if(!map)return;markers.forEach(m=>map.removeLayer(m));markers=[];list.filter(s=>s.hasCoords).forEach(s=>{const m=L.circleMarker([s.lat,s.lng],{radius:7,color:'#fff',weight:1,fillColor:markerColor(s),fillOpacity:.95}).addTo(map);m.bindTooltip(esc(s.name));m.on('click',()=>openStore(s.id));markers.push(m);});}
function fillSelect(id,values,placeholder){const el=$(id);if(!el)return;const old=el.value;const vals=[...new Set(values.filter(Boolean))].sort((a,b)=>a.localeCompare(b,'fa'));el.innerHTML=`<option value="">${placeholder}</option>`+vals.map(v=>`<option value="${esc(v)}">${esc(v)}</option>`).join('');if(vals.includes(old))el.value=old;}
function populateFilters(){fillSelect('areaFilter',stores.map(s=>s.area),'همه مناطق');fillSelect('routeFilter',stores.map(s=>s.routeName),'همه مسیرها');fillSelect('personFilter',stores.map(s=>s.salesman),'همه ویزیتورها');fillSelect('classFilter',stores.map(s=>s.customerClass),'همه کلاس‌ها');}
function applyFilters(){
 const q=norm($('search')?.value||'');
 const type=$('typeFilter')?.value||'all',customerType=$('customerTypeFilter')?.value||'all',gps=$('gpsFilter')?.value||'all';
 const area=$('areaFilter')?.value||'',route=$('routeFilter')?.value||'',person=$('personFilter')?.value||'',cls=$('classFilter')?.value||'',day=$('dayFilter')?.value||'';
 const list=stores.filter(s=>{
  if(q&&!norm([s.name,s.code,s.id,s.address].join(' ')).includes(q))return false;
  if(type==='sales'&&!s.hasSales)return false;
  if(type==='nosales'&&s.hasSales)return false;
  if(type==='salesonly'&&!s.inJP)return false;
  if(customerType==='jp'&&!s.inJP)return false;
  if(customerType==='tpa'&&s.inJP)return false;
  if(gps==='yes'&&!s.hasCoords)return false;
  if(gps==='no'&&s.hasCoords)return false;
  if(area&&s.area!==area)return false;
  if(route&&s.routeName!==route)return false;
  if(person&&s.salesman!==person)return false;
  if(cls&&s.customerClass!==cls)return false;
  if(day&&(!s.inJP||!s.visitDays.includes(day)))return false;
  return true;
 });
 renderMarkers(list);renderStoreList(list);
}
function clearFilters(){['search','typeFilter','customerTypeFilter','gpsFilter','areaFilter','routeFilter','personFilter','classFilter','dayFilter'].forEach(id=>{const e=$(id);if(e)e.value='';});applyFilters();}
function renderStoreList(list){const box=$('storeList');if(!box)return;box.innerHTML=list.slice(0,500).map(s=>`<article class="store" data-store="${esc(s.id)}"><span class="storeinfo"><b>${esc(s.name)}</b><small>${esc(s.code||s.id)} · ${esc(s.routeName||s.area||'بدون مسیر')}</small><small>${esc(s.address||'')}</small></span><span class="storebadge ${s.hasSales?'sold':s.inJP?'nosale':'tpa'}">${s.hasSales?'فروش‌دار':s.inJP?'بدون فروش':'فقط TPA'}</span></article>`).join('')||'<div class="empty">موردی پیدا نشد.</div>';box.querySelectorAll('[data-store]').forEach(el=>el.addEventListener('click',()=>openStore(el.dataset.store)));}
/* Store details */
function openStore(id){const s=stores.find(x=>String(x.id)===String(id));if(!s)return;const modal=$('modal'),body=$('modalBody');if(!modal||!body)return;
 body.innerHTML=`
 <div class="modal-head"><div><h2>${esc(s.name)}</h2><small>${esc(s.code||s.id)}</small></div><button type="button" id="closeModal">×</button></div>
 <div class="store-details">
  <div class="detail-grid">
   <div><small>نام فروشگاه</small><b>${esc(s.storeName||s.name)}</b></div>
   <div><small>کلاس مشتری</small><b>${esc(s.customerClass||'—')}</b></div>
   <div><small>مسیر</small><b>${esc(s.routeName||'—')}</b></div>
   <div><small>ویزیتور</small><b>${esc(s.salesman||'—')}</b></div>
  </div>
  <div class="order-meta"><b>آدرس:</b> ${esc(s.address||'—')}</div>
  <div class="order-meta"><b>وضعیت:</b> ${s.inJP?'موجود در Journey Plan':'فقط در گزارش TPA'} · ${s.hasSales?'دارای سابقه فروش':'بدون سابقه فروش'}</div>
  <div class="order-meta"><b>آخرین سفارش:</b> ${esc(displayDate(s.lastPurchase)||'—')}</div>
  <div class="order-meta"><b>جمع فاکتورهای صادرشده:</b> ${formatMoney(s.invoicedTotal||0)}</div>
  <div class="order-meta"><b>جمع سفارش‌های فاکتورنشده:</b> ${formatMoney(s.nonInvoicedTotal||0)}</div>
  <div class="order-meta"><b>تعداد سفارش‌های باز:</b> ${s.openOrders.length.toLocaleString('fa-IR')}</div>
  <h3>تحلیل سبد خرید</h3>
  <div class="order-meta"><b>تعداد محصولات متمایز:</b> ${new Set(s.products.map(p=>norm(p.code)||p.name)).size.toLocaleString('fa-IR')}</div>
  ${s.inactiveProducts?.length?`<div class="order-meta"><b>محصولات غیرفعال (در سه سفارش اخیر خریداری نشده‌اند):</b><br>${s.inactiveProducts.map(p=>esc(p.name)).join('، ')}</div>`:'<div class="order-meta">محصول غیرفعال شناسایی نشد.</div>'}
  <h3>سفارش‌ها</h3>
  ${s.orders.length?s.orders.slice(0,30).map(o=>`<article class="order-card">
   <div class="order-title"><b>سفارش ${esc(o.orderNo)}</b><span>${esc(o.status||'—')}</span></div>
   <div class="order-meta">تاریخ سفارش: ${esc(o.orderDateDisplay||'—')} · تاریخ فاکتور: ${esc(o.invoiceDateDisplay||'—')}</div>
   <div class="order-meta">تاریخ تحویل هدف: ${esc(o.deliveryDateDisplay||'—')} ${o.isOverdue?'<b class="warning">· سررسید گذشته</b>':''}</div>
   <div class="order-meta">مبلغ سفارش: ${formatMoney(o.amount)}</div>
   <div class="order-products">${o.products.map(p=>`<div>${esc(p.name)} · تعداد ${num(p.quantity).toLocaleString('fa-IR')} · ${formatMoney(p.amount)}</div>`).join('')}</div>
  </article>`).join(''):'<div class="empty">سفارشی برای این مشتری ثبت نشده است.</div>'}
  <div class="detail-actions">
   ${s.hasCoords?`<a target="_blank" rel="noopener" href="https://www.google.com/maps/dir/?api=1&destination=${s.lat},${s.lng}">مسیریابی Google Maps</a><a target="_blank" rel="noopener" href="https://neshan.org/maps/?lat=${s.lat}&lng=${s.lng}&zoom=16">نمایش در نشان</a>`:''}
   ${s.hasCoords?`<button type="button" id="addToPlanBtn">افزودن به برنامه Auditor</button>`:''}
  </div>
 </div>`;
 modal.classList.add('show');
 $('closeModal')?.addEventListener('click',closeModal);
 $('addToPlanBtn')?.addEventListener('click',()=>{addStoreToPlan(s.id);});
}
function closeModal(){$('modal')?.classList.remove('show');}

/* Google Drive */
function driveRequest(url,token,options={}){return fetch(url,{...options,headers:{...(options.headers||{}),Authorization:'Bearer '+token}}).then(async r=>{if(!r.ok)throw new Error(`Google Drive: ${r.status} ${await r.text()}`);return r;});}
function getDriveToken(){return new Promise((resolve,reject)=>{
 if(!window.google?.accounts?.oauth2){reject(new Error('کتابخانه ورود Google بارگذاری نشده است.'));return;}
 if(!driveTokenClient)driveTokenClient=google.accounts.oauth2.initTokenClient({client_id:GOOGLE_CLIENT_ID,scope:'https://www.googleapis.com/auth/drive.readonly',callback:response=>{if(response.error){reject(new Error(response.error));return;}lastAccessToken=response.access_token;resolve(response.access_token);}});
 if(lastAccessToken){resolve(lastAccessToken);return;}
 driveTokenClient.requestAccessToken({prompt:'consent'});
 });}
async function listDriveFiles(token){const q=encodeURIComponent(`'${DRIVE_FOLDER_ID}' in parents and trashed=false`);const url=`https://www.googleapis.com/drive/v3/files?q=${q}&pageSize=100&orderBy=modifiedTime desc&fields=files(id,name,mimeType,modifiedTime,size)`;const data=await(await driveRequest(url,token)).json();return data.files||[];}
function newestFile(files,prefix){return files.filter(f=>f.name.toLowerCase().startsWith(prefix.toLowerCase())).sort((a,b)=>new Date(b.modifiedTime)-new Date(a.modifiedTime))[0]||null;}
async function downloadDriveExcel(file,token,headerRow){const isSheet=file.mimeType==='application/vnd.google-apps.spreadsheet';const url=isSheet?`https://www.googleapis.com/drive/v3/files/${encodeURIComponent(file.id)}/export?mimeType=${encodeURIComponent('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}`:`https://www.googleapis.com/drive/v3/files/${encodeURIComponent(file.id)}?alt=media`;const buf=await(await driveRequest(url,token)).arrayBuffer();const wb=XLSX.read(buf,{type:'array',cellDates:true});return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{range:headerRow,defval:'',raw:false});}
async function syncFromDrive(){const b=$('driveBtn');if(b)b.disabled=true;try{setMessage('در حال اتصال به Google Drive...','در حال اتصال');const token=await getDriveToken();const files=await listDriveFiles(token);const jf=newestFile(files,'JP_'),sf=newestFile(files,'TPA_Alborz'),mf=newestFile(files,'Product Mapping');if(!jf||!sf)throw new Error('فایل Journey Plan یا TPA_Alborz در پوشه تنظیم‌شده پیدا نشد. نام فایل‌ها و دسترسی پوشه را بررسی کن.');setMessage('در حال دریافت سه فایل اصلی...');const [jp,sa,pm]=await Promise.all([downloadDriveExcel(jf,token,0),downloadDriveExcel(sf,token,1),mf?downloadDriveExcel(mf,token,0):Promise.resolve([])]);if(!jp.length||!sa.length)throw new Error('یکی از فایل‌های دریافت‌شده خالی است؛ داده قبلی حفظ شد.');jpRows=jp;salesRows=sa;productRows=pm;activeSourceYear=Number((sf.name.match(/(14\d{2})/)||[])[1]||1405);buildData();renderAll();const stamp=new Date().toLocaleString('fa-IR');if($('lastSync'))$('lastSync').textContent='آخرین همگام‌سازی: '+stamp;setMessage(`همگام‌سازی موفق\nJP: ${jf.name}\nTPA: ${sf.name}\nMapping: ${mf?.name||'یافت نشد'}\nتعداد مشتری: ${stores.length.toLocaleString('fa-IR')}`,'همگام‌سازی موفق');}catch(e){console.error(e);setMessage('همگام‌سازی انجام نشد؛ داده قبلی دست‌نخورده باقی ماند.\n'+e.message,'خطا');}finally{if(b)b.disabled=false;}}

/* Auditor visit plan */
function loadSavedPlan(){try{const p=JSON.parse(localStorage.getItem(PLAN_KEY)||'[]');auditPlan=Array.isArray(p)?p:[];}catch(e){auditPlan=[];}renderPlan();}
function savePlan(){localStorage.setItem(PLAN_KEY,JSON.stringify(auditPlan));renderPlan();}
function addStoreToPlan(id){const s=stores.find(x=>String(x.id)===String(id));if(!s)return;if(!s.hasCoords){alert('فروشگاه بدون GPS به فهرست انتخاب Auditor اضافه نمی‌شود.');return;}const date=clean($('planDate')?.value),day=$('planDay')?.value||'Saturday';if(!date){alert('تاریخ بازدید را وارد کن.');return;}if(auditPlan.some(p=>norm(p.customerId)===norm(s.code||s.id)&&p.visitDate===date&&p.visitDay===day)){alert('این فروشگاه برای همین روز و تاریخ قبلاً ثبت شده است.');return;}auditPlan.push({visitDate:date,visitDay:day,customerId:s.code||s.id,storeName:s.name,customerName:s.name,address:s.address,routeName:s.routeName,latitude:s.lat,longitude:s.lng,priority:'Medium',selectionReason:'',auditorComment:'',selectionSource:'Internal',visitStatus:'Pending'});savePlan();}
function renderPlan(){const box=$('planList');if(!box)return;const day=$('planDay')?.value||'all';const shown=auditPlan.filter(p=>day==='all'||p.visitDay===day);if(!shown.length){box.innerHTML='<div class="empty">برای این روز هنوز بازدیدی ثبت نشده است.</div>';return;}box.innerHTML=shown.map(p=>{const i=auditPlan.indexOf(p);return `<article class="planrow"><b>${esc(p.storeName||p.customerName||p.customerId)}</b><small>کد: ${esc(p.customerId)} · ${esc(p.visitDate)} · ${esc(p.visitDay)}<br>مسیر: ${esc(p.routeName||'—')} · منبع: ${esc(p.selectionSource||'Internal')}</small><div class="plan-controls"><select data-status="${i}"><option value="Pending" ${p.visitStatus==='Pending'?'selected':''}>در انتظار</option><option value="Completed" ${p.visitStatus==='Completed'?'selected':''}>انجام شد</option><option value="Not completed" ${p.visitStatus==='Not completed'?'selected':''}>انجام نشد</option></select><button type="button" data-edit="${i}">ویرایش</button><button type="button" data-delete="${i}" class="secondary">حذف</button></div></article>`;}).join('');box.querySelectorAll('[data-status]').forEach(el=>el.addEventListener('change',()=>{auditPlan[+el.dataset.status].visitStatus=el.value;savePlan();}));box.querySelectorAll('[data-edit]').forEach(el=>el.addEventListener('click',()=>editPlanRow(+el.dataset.edit)));box.querySelectorAll('[data-delete]').forEach(el=>el.addEventListener('click',()=>removePlanRow(+el.dataset.delete)));}
function editPlanRow(i){const p=auditPlan[i];if(!p)return;const reason=prompt('دلیل انتخاب:',p.selectionReason||'');if(reason===null)return;p.selectionReason=reason;const comment=prompt('یادداشت Auditor:',p.auditorComment||'');if(comment!==null)p.auditorComment=comment;const priority=prompt('اولویت: High / Medium / Low',p.priority||'Medium');if(priority!==null&&['High','Medium','Low'].includes(priority))p.priority=priority;savePlan();}
function removePlanRow(i){if(!confirm('این بازدید از برنامه حذف شود؟'))return;auditPlan.splice(i,1);savePlan();}
const PLAN_HEADERS=['Visit Date','Customer ID','Store Name','Customer Name','Address','Route Name','Latitude','Longitude','Priority','Selection Reason','Auditor Comment','Selection Source','Visit Day','Visit Status'];
function exportAuditPlan(){const rows=auditPlan.map(p=>({'Visit Date':p.visitDate,'Customer ID':p.customerId,'Store Name':p.storeName,'Customer Name':p.customerName,'Address':p.address,'Route Name':p.routeName,'Latitude':p.latitude,'Longitude':p.longitude,'Priority':p.priority,'Selection Reason':p.selectionReason,'Auditor Comment':p.auditorComment,'Selection Source':p.selectionSource,'Visit Day':p.visitDay,'Visit Status':p.visitStatus}));const wb=XLSX.utils.book_new(),ws=XLSX.utils.json_to_sheet(rows,{header:PLAN_HEADERS});ws['!cols']=PLAN_HEADERS.map(()=>({wch:22}));XLSX.utils.book_append_sheet(wb,ws,'Audit Visit Plan');XLSX.writeFile(wb,'Audit_Visit_Plan.xlsx');}
async function importAuditPlan(){const f=$('importPlanFile')?.files?.[0];if(!f){alert('فایل اکسل را انتخاب کن.');return;}try{const rows=await readExcel(f,0),valid=[],bad=[],seen=new Set();rows.forEach((r,i)=>{const id=first(r,['Customer ID','CustomerID']),date=first(r,['Visit Date','VisitDate']);if(!id||!date){bad.push(i+2);return;}const k=norm(id)+'|'+date;if(seen.has(k)){bad.push(i+2);return;}seen.add(k);const s=stores.find(x=>variants(x.id).includes(norm(id))||variants(x.code).includes(norm(id)));if(!s||!s.hasCoords){bad.push(i+2);return;}valid.push({visitDate:date,visitDay:first(r,['Visit Day'])||$('planDay').value,customerId:id,storeName:first(r,['Store Name'])||s.name,customerName:first(r,['Customer Name'])||s.name,address:first(r,['Address'])||s.address,routeName:first(r,['Route Name'])||s.routeName,latitude:first(r,['Latitude'])||s.lat,longitude:first(r,['Longitude'])||s.lng,priority:first(r,['Priority'])||'Medium',selectionReason:first(r,['Selection Reason']),auditorComment:first(r,['Auditor Comment']),selectionSource:'External',visitStatus:'Pending'});});pendingExternal=valid;const box=$('importReport');box.innerHTML=`ردیف معتبر: ${valid.length} · ردیف نامعتبر/تکراری/بدون GPS: ${bad.length}<br>${valid.length?'<button type="button" id="acceptExternalBtn" class="green">تأیید و افزودن به برنامه</button>':'چیزی برای وارد کردن وجود ندارد.'}`;if(valid.length)$('acceptExternalBtn').addEventListener('click',()=>{let added=0;for(const p of pendingExternal){if(!auditPlan.some(x=>norm(x.customerId)===norm(p.customerId)&&x.visitDate===p.visitDate&&x.visitDay===p.visitDay)){auditPlan.push(p);added++;}}pendingExternal=[];savePlan();box.textContent=`${added} بازدید به برنامه اضافه شد.`;});}catch(e){$('importReport').textContent='خطا در فایل: '+e.message;}}

/* Analysis */
function renderAnalysis(){if(!$('noGpsCount'))return;$('noGpsCount').textContent=stores.filter(s=>!s.hasCoords).length.toLocaleString('fa-IR');$('openOrderCount').textContent=stores.reduce((n,s)=>n+s.openOrders.length,0).toLocaleString('fa-IR');$('noPurchaseCount').textContent=stores.filter(s=>!s.hasSales).length.toLocaleString('fa-IR');}
function showAnalysisList(kind){let list=stores;if(kind==='nogps')list=stores.filter(s=>!s.hasCoords);if(kind==='nopurchase')list=stores.filter(s=>!s.hasSales);if(kind==='open')list=stores.filter(s=>s.openOrders.length);$('analysisResults').innerHTML=list.length?list.slice(0,250).map(s=>`<article class="planrow"><b>${esc(s.name)}</b><small>${esc(s.code)} · ${esc(s.address||'')}</small></article>`).join(''):'<div class="empty">موردی پیدا نشد.</div>';}
function findNearby(){if(!navigator.geolocation){alert('موقعیت‌یابی در این مرورگر پشتیبانی نمی‌شود.');return;}navigator.geolocation.getCurrentPosition(pos=>{const {latitude,longitude}=pos.coords;const d=s=>{const R=6371,a=(s.lat-latitude)*Math.PI/180,b=(s.lng-longitude)*Math.PI/180;const x=Math.sin(a/2)**2+Math.cos(latitude*Math.PI/180)*Math.cos(s.lat*Math.PI/180)*Math.sin(b/2)**2;return 2*R*Math.asin(Math.sqrt(x));};const near=stores.filter(s=>s.hasCoords).map(s=>({...s,distance:d(s)})).filter(s=>s.distance<=5).sort((a,b)=>a.distance-b.distance);$('analysisResults').innerHTML=near.length?near.map(s=>`<article class="planrow"><b>${esc(s.name)}</b><small>${s.distance.toFixed(2)} کیلومتر · ${esc(s.code)}</small></article>`).join(''):'<div class="empty">در شعاع ۵ کیلومتری فروشگاهی پیدا نشد.</div>';showPage('analysisPage');},()=>alert('دسترسی موقعیت مکانی ممکن نشد.'),{enableHighAccuracy:true,timeout:12000});}

/* Pages, theme, offline cache */
function showPage(id){document.querySelectorAll('.page').forEach(p=>p.classList.toggle('active',p.id===id));document.querySelectorAll('.bottom-nav [data-page]').forEach(b=>b.classList.toggle('active',b.dataset.page===id));if(id==='mapPage')setTimeout(()=>map?.invalidateSize(),150);}
function saveOfflineSnapshot(){try{localStorage.setItem(SNAPSHOT_KEY,JSON.stringify({savedAt:new Date().toISOString(),jpRows,salesRows,productRows}));if($('lastSync'))$('lastSync').textContent='آخرین ذخیره: '+new Date().toLocaleString('fa-IR');}catch(e){console.warn('Offline snapshot failed',e);}}
function restoreOfflineSnapshot(){try{const x=JSON.parse(localStorage.getItem(SNAPSHOT_KEY)||'null');if(!x||!Array.isArray(x.jpRows)||!Array.isArray(x.salesRows))return false;jpRows=x.jpRows;salesRows=x.salesRows;productRows=x.productRows||[];buildData();if($('lastSync'))$('lastSync').textContent='داده بازیابی‌شده از حافظه: '+new Date(x.savedAt).toLocaleString('fa-IR');return true;}catch(e){console.warn(e);return false;}}
function applyTheme(){document.body.classList.toggle('dark',localStorage.getItem('tpa_theme')==='dark');}
function registerOfflineShell(){if('serviceWorker'in navigator&&location.protocol.startsWith('http'))navigator.serviceWorker.register('./sw.js').catch(e=>console.warn('SW registration',e));}

/* Events */
function bindEvents(){
 $('loadBtn')?.addEventListener('click',loadManualFiles);$('driveBtn')?.addEventListener('click',syncFromDrive);$('closeModal')?.addEventListener('click',closeModal);$('modal')?.addEventListener('click',e=>{if(e.target.id==='modal')closeModal();});
 $('search')?.addEventListener('input',applyFilters);['typeFilter','customerTypeFilter','gpsFilter'].forEach(id=>$(id)?.addEventListener('change',applyFilters));FILTER_IDS.forEach(id=>$(id)?.addEventListener('change',applyFilters));$('clearFiltersBtn')?.addEventListener('click',clearFilters);
 document.querySelectorAll('.bottom-nav [data-page]').forEach(b=>b.addEventListener('click',()=>showPage(b.dataset.page)));
 $('planSearch')?.addEventListener('input',()=>{const q=norm($('planSearch').value);const hits=stores.filter(s=>s.hasCoords&&q&&norm([s.code,s.name,s.address].join(' ')).includes(q)).slice(0,15);$('planSearchResults').innerHTML=hits.map(s=>`<div class="store"><span class="storeinfo"><b>${esc(s.name)}</b><small>${esc(s.code)} · ${esc(s.routeName||s.area||'')}</small></span><button type="button" class="smallbtn green" data-add="${esc(s.id)}">افزودن</button></div>`).join('');$('planSearchResults').querySelectorAll('[data-add]').forEach(b=>b.addEventListener('click',()=>addStoreToPlan(b.dataset.add)));});
 $('planDay')?.addEventListener('change',renderPlan);$('exportPlanBtn')?.addEventListener('click',exportAuditPlan);$('importPlanBtn')?.addEventListener('click',importAuditPlan);$('nearbyBtn')?.addEventListener('click',findNearby);$('noGpsBtn')?.addEventListener('click',()=>showAnalysisList('nogps'));$('noPurchaseBtn')?.addEventListener('click',()=>showAnalysisList('nopurchase'));$('openOrdersBtn')?.addEventListener('click',()=>showAnalysisList('open'));
 $('themeBtn')?.addEventListener('click',()=>{const dark=!document.body.classList.contains('dark');document.body.classList.toggle('dark',dark);localStorage.setItem('tpa_theme',dark?'dark':'light');});
 $('langBtn')?.addEventListener('click',()=>alert('رابط انگلیسی هنوز پیاده‌سازی نشده است؛ نسخه فعلی فارسی است.'));
 $('clearLocalBtn')?.addEventListener('click',()=>{if(confirm('فقط داده‌های ذخیره‌شده این دستگاه و برنامه بازدید پاک شوند؟')){localStorage.removeItem(SNAPSHOT_KEY);localStorage.removeItem(PLAN_KEY);auditPlan=[];renderPlan();alert('داده‌های محلی پاک شدند.');}});
}

window.addEventListener('error',e=>{console.error(e.error||e.message);if($('message'))setMessage('خطای برنامه: '+e.message,'خطا');});
window.addEventListener('unhandledrejection',e=>{console.error(e.reason);if($('message'))setMessage('خطای عملیات: '+(e.reason?.message||e.reason),'خطا');});
bindEvents();loadSavedPlan();applyTheme();registerOfflineShell();
if(!restoreOfflineSnapshot())renderAll();
