'use strict';
/* TPA Navigation — aligned with index.html IDs */
const CLIENT_ID='1098731892951-2e7jipvf33evn0a69tt0tje94oor54lm.apps.googleusercontent.com';
const DRIVE_FOLDER='1sgTsZu2zzSBi-pz2M9pVBHtjCtQELYhk';
const DB_NAME='tpa-navigation-db', DB_VERSION=1, STORE='data';
const PLAN_KEY='tpa_auditor_plan_v3';
let jpRows=[], tpaRows=[], mappingRows=[], stores=[], markers=[];
let map=null, tokenClient=null, accessToken=null, plan=[], externalRows=[];
const $=id=>document.getElementById(id);
const DAYS=['Saturday','Sunday','Monday','Tuesday','Wednesday','Thursday'];
const DAY_FA={Saturday:'شنبه',Sunday:'یکشنبه',Monday:'دوشنبه',Tuesday:'سه‌شنبه',Wednesday:'چهارشنبه',Thursday:'پنجشنبه'};
const clean=v=>String(v??'').trim();
const codeVariants=v=>{const s=norm(v);return [...new Set([s,...s.split('|').map(x=>x.trim()).filter(Boolean)])].filter(Boolean);};
const digits=v=>clean(v).replace(/[۰-۹]/g,d=>'۰۱۲۳۴۵۶۷۸۹'.indexOf(d)).replace(/[٠-٩]/g,d=>'٠١٢٣٤٥٦٧٨٩'.indexOf(d));
const norm=v=>digits(v).replace(/\s+/g,'').replace(/\.0$/,'');
const number=v=>{const n=Number(digits(v).replace(/,/g,'').replace(/[^\d.-]/g,''));return Number.isFinite(n)?n:0;};
const first=(r,keys)=>{for(const k of keys)if(r&&r[k]!==undefined&&clean(r[k])!=='')return clean(r[k]);return '';};
const esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
function message(s){$('message').textContent=s;}
function status(s){$('status').textContent=s;}
function validGPS(a,b){return Number.isFinite(a)&&Number.isFinite(b)&&a>=-90&&a<=90&&b>=-180&&b<=180&&(a!==0||b!==0);}
function parseExcel(file,range){return new Promise((resolve,reject)=>{const fr=new FileReader();fr.onload=()=>{try{const wb=XLSX.read(fr.result,{type:'array',cellDates:true});resolve(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{range,defval:'',raw:false}));}catch(e){reject(e);}};fr.onerror=()=>reject(fr.error);fr.readAsArrayBuffer(file);});}
function validate(jp,tpa,pm){if(!jp.length||!jp[0]||!Object.hasOwn(jp[0],'CustomerI'))throw Error('فایل JP معتبر نیست یا ردیف عنوان آن اشتباه است؛ ستون CustomerI پیدا نشد.');if(!tpa.length||!tpa[0]||(!Object.hasOwn(tpa[0],'کد مشتری')&&!Object.hasOwn(tpa[0],'Customer Code')))throw Error('فایل TPA معتبر نیست یا ردیف عنوان آن اشتباه است؛ ستون کد مشتری پیدا نشد.');if(!tpa[0].Status&&!Object.hasOwn(tpa[0],'Status')&&!Object.hasOwn(tpa[0],'وضعیت سفارش'))throw Error('ستون Status در TPA پیدا نشد.');if(pm.length&&!Object.hasOwn(pm[0],'Item Code'))throw Error('ساختار Product Mapping معتبر نیست.');}
function parseJalaliYYMMDD(v,yearHint){const s=digits(v);if(!/^\d{6}$/.test(s))return null;const jy=Math.floor(yearHint/100)*100+Number(s.slice(0,2)),jm=Number(s.slice(2,4)),jd=Number(s.slice(4,6));if(jm<1||jm>12||jd<1||jd>(jm<=6?31:jm<=11?30:30))return null;let gy=jy-979,jdn=365*gy+Math.floor(gy/33)*8+Math.floor((gy%33+3)/4)+78+jd+(jm<7?(jm-1)*31:(jm-7)*30+186);let g=1600+400*Math.floor(jdn/146097);jdn%=146097;let leap=true;if(jdn>=36525){jdn--;g+=100*Math.floor(jdn/36524);jdn%=36524;if(jdn>=365)jdn++;else leap=false;}g+=4*Math.floor(jdn/1461);jdn%=1461;if(jdn>=366){leap=false;jdn--;g+=Math.floor(jdn/365);jdn%=365;}const md=[31,leap?29:28,31,30,31,30,31,31,30,31,30,31];let m=0;while(m<12&&jdn>=md[m])jdn-=md[m++];return Date.UTC(g,m,jdn+1);}
function dateInfo(v,yearHint=(window.sourceYear||1405)){const s=clean(v);if(!s)return {ts:0,label:'—'};let m=s.match(/^(\d{6})$/);if(m){const ts=parseJalaliYYMMDD(s,yearHint);return {ts:ts||0,label:s};}m=s.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/);if(m){const d=new Date(Date.UTC(+m[3],+m[1]-1,+m[2]));return {ts:d.getTime(),label:d.toLocaleDateString('fa-IR-u-ca-persian')};}const d=new Date(s);return Number.isNaN(d.getTime())?{ts:0,label:s}:{ts:d.getTime(),label:d.toLocaleDateString('fa-IR-u-ca-persian')};}
function sourceYear(name){const m=String(name||'').match(/(14\d{2})/);return m?+m[1]:1405;}
function buildData(){
 const pm=new Map(mappingRows.map(r=>[norm(r['Item Code']),first(r,['Product Name','Product  Name','Item Name','Item'])]));
 const jpByCode=new Map(), jp=[];
 for(const r of jpRows){const id=first(r,['CustomerI']);if(!id)continue;const lat=number(r.Latitude),lng=number(r.Longitude);const days=DAYS.filter(d=>{const v=clean(r[d]).toLowerCase();return v!==''&&!['0','no','false','-'].includes(v);});const s={id,name:first(r,['Customer Name','CustomerName','Name','نام مشتری','نام فروشگاه'])||id,address:first(r,['Address',' Address','Customer Address','آدرس']),customerClass:first(r,['Customer Class','CustomerClass','Outlet Class']),salesman:first(r,['Salesman Name','Salesman','نام فروشنده']),route:first(r,['Route Name']),routeId:first(r,['RouteID','Route Id']),area:first(r,['Area Name','Area','ناحیه']),lat,lng,gps:validGPS(lat,lng),days,source:'JP',orders:[],hasSales:false};jp.push(s);for(const code of [id,r['New Customer Code'],r['Oracle Code'],r['Account Number']])for(const v of codeVariants(code))jpByCode.set(v,s);}
 const groups=new Map();for(const r of tpaRows){const st=first(r,['Status','وضعیت سفارش']);if(st.toLowerCase()==='closed/expired')continue;const code=first(r,['کد مشتری','Customer Code']);const order=first(r,['Order Number','Order No','OrderNumber','شماره سفارش']);if(!norm(code)||!norm(order))continue;const customer=codeVariants(code).map(v=>jpByCode.get(v)).find(Boolean);const key=(customer?.id||norm(code))+'|'+norm(order);if(!groups.has(key))groups.set(key,{customer,code,order,rows:[],date:first(r,['Date','Order Date','تاریخ']),target:first(r,['Target Delivery Date','Target Deilvery Date']),invoice:first(r,['Invoice Date','invoicedate']),status:st});groups.get(key).rows.push(r);}
 const tpaOnly=new Map();for(const g of groups.values()){const lines=g.rows.map(r=>{const code=first(r,['Item Code','ItemCode','کد کالا']),qty=number(first(r,['Quantity','Qty','Order Quantity','تعداد'])),price=number(first(r,['Unit Price','Price','قیمت واحد']));return {code,name:pm.get(norm(code))||first(r,['Item Name','Item Description','Description','نام کالا'])||code,qty,price,amount:qty*price};});const o={number:g.order,date:g.date,target:g.target,invoice:g.invoice,status:g.status,open:!clean(g.invoice),lines,total:lines.reduce((a,x)=>a+x.amount,0)};if(g.customer){g.customer.orders.push(o);g.customer.hasSales=true;}else{const k=norm(g.code)||'unknown'+tpaOnly.size;if(!tpaOnly.has(k))tpaOnly.set(k,{id:g.code||k,name:first(g.rows[0],['Customer Name','نام مشتری'])||g.code,address:first(g.rows[0],['Address','آدرس']),customerClass:'',salesman:'',route:'',routeId:'',area:'',lat:0,lng:0,gps:false,days:[],source:'TPA',orders:[],hasSales:true});tpaOnly.get(k).orders.push(o);}}
 stores=[...jp,...tpaOnly.values()];for(const s of stores){s.orders.sort((a,b)=>dateInfo(b.date,window.sourceYear||1405).ts-dateInfo(a.date,window.sourceYear||1405).ts);s.orderCount=s.orders.length;s.total=s.orders.reduce((a,o)=>a+o.total,0);const recent=new Set(s.orders.slice(0,3).flatMap(o=>o.lines.map(x=>norm(x.code)||x.name)));const all=new Map();for(const o of s.orders)for(const x of o.lines){const k=norm(x.code)||x.name;if(!all.has(k))all.set(k,{...x,qtyAll:0});all.get(k).qtyAll+=x.qty;}s.inactive=[...all.values()].filter(x=>!recent.has(norm(x.code)||x.name));}
}
function db(){return new Promise((res,rej)=>{const q=indexedDB.open(DB_NAME,DB_VERSION);q.onupgradeneeded=()=>{if(!q.result.objectStoreNames.contains(STORE))q.result.createObjectStore(STORE,{keyPath:'id'});};q.onsuccess=()=>res(q.result);q.onerror=()=>rej(q.error);});}
async function saveLocal(){const d=await db();return new Promise((res,rej)=>{const tx=d.transaction(STORE,'readwrite');tx.objectStore(STORE).put({id:'source',jpRows,tpaRows,mappingRows,year:window.sourceYear||1405,updated:Date.now()});tx.oncomplete=()=>{d.close();res();};tx.onerror=()=>{d.close();rej(tx.error);};});}
async function restoreLocal(){try{const d=await db();const data=await new Promise((res,rej)=>{const q=d.transaction(STORE).objectStore(STORE).get('source');q.onsuccess=()=>res(q.result);q.onerror=()=>rej(q.error);});d.close();if(!data)return false;jpRows=data.jpRows;tpaRows=data.tpaRows;mappingRows=data.mappingRows;window.sourceYear=data.year;buildData();renderAll();message('اطلاعات ذخیره‌شده دستگاه بازیابی شد.');return true;}catch(e){console.warn(e);return false;}}
function valuesOf(key){return [...new Set(stores.filter(s=>s.source==='JP').map(s=>s[key]).filter(Boolean))].sort((a,b)=>String(a).localeCompare(String(b),'fa'));}
function selected(id){return [...($(id)?.selectedOptions||[])].map(o=>o.value);}
function fillMulti(id,values){const el=$(id),old=selected(id);if(!el)return;el.innerHTML=values.map(v=>`<option value="${esc(v)}">${esc(v)}</option>`).join('');for(const o of el.options)o.selected=old.includes(o.value);}
function todayEnglish(){return ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'][new Date().getDay()];}
function filters(){return {kind:$('customerTypeFilter').value,type:$('typeFilter').value,gps:$('gpsFilter').value,area:selected('areaFilter'),route:selected('routeFilter'),person:selected('personFilter'),cls:selected('classFilter'),day:selected('dayFilter')};}
function filtered(){const f=filters();return stores.filter(s=>(f.kind==='all'||(f.kind==='jp'&&s.source==='JP')||(f.kind==='salesOnly'&&s.source==='TPA'))&&(f.type==='all'||(f.type==='sales'&&s.hasSales)||(f.type==='nosales'&&!s.hasSales))&&(f.gps==='all'||(f.gps==='yes'&&s.gps)||(f.gps==='no'&&!s.gps))&&(s.source==='TPA'||(!f.area.length||f.area.includes(s.area))&&(!f.route.length||f.route.includes(s.route))&&(!f.person.length||f.person.includes(s.salesman))&&(!f.cls.length||f.cls.includes(s.customerClass))&&(!f.day.length||f.day.some(d=>s.days.includes(d)||(d==='Today'&&s.days.includes(todayEnglish()))))));}
function activeCount(){const f=filters();return Object.values(f).reduce((n,v)=>n+(Array.isArray(v)?v.length:(v==='all'?0:1)),0);}
function initMap(){if(!window.L)return;map=L.map('map',{zoomControl:false}).setView([35.7,51.4],10);L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',{maxZoom:19,attribution:'© OpenStreetMap'}).addTo(map);L.control.zoom({position:'bottomright'}).addTo(map);}
function drawMap(list){if(!map)return;markers.forEach(m=>map.removeLayer(m));markers=[];for(const s of list){if(!s.gps)continue;const color=s.source==='TPA'?'#2878e5':s.hasSales?'#19a56b':'#e79a2e';const m=L.circleMarker([s.lat,s.lng],{radius:8,color:'#fff',weight:2,fillColor:color,fillOpacity:1}).addTo(map);m.bindTooltip(esc(s.name));m.on('click',()=>showDetails(s.id));markers.push(m);}}
function renderList(){const q=norm($('search').value);const list=filtered().filter(s=>!q||[s.id,s.name,s.address].some(x=>norm(x).includes(q)));$('resultCount').textContent=list.length.toLocaleString('fa-IR')+' فروشگاه';$('storeList').innerHTML=list.map(s=>`<button class="store-row" data-id="${esc(s.id)}"><span class="dot ${s.source==='TPA'?'blue-dot':s.hasSales?'green-dot':'orange-dot'}"></span><span><b>${esc(s.name)}</b><small>${esc(s.address||'آدرس ثبت نشده')}</small></span><small>${s.hasSales?'دارای سفارش':'بدون سفارش'}${s.gps?'':' · بدون GPS'}</small></button>`).join('')||'<div class="empty">موردی پیدا نشد.</div>';document.querySelectorAll('[data-id]').forEach(b=>b.addEventListener('click',()=>showDetails(b.dataset.id)));}
function renderStats(){const f=filtered();$('total').textContent=stores.length.toLocaleString('fa-IR');$('withSales').textContent=stores.filter(s=>s.hasSales).length.toLocaleString('fa-IR');$('withoutSales').textContent=stores.filter(s=>s.source==='JP'&&!s.hasSales).length.toLocaleString('fa-IR');$('noGpsCount').textContent=stores.filter(s=>s.source==='JP'&&!s.gps).length.toLocaleString('fa-IR');$('openOrderCount').textContent=stores.reduce((n,s)=>n+s.orders.filter(o=>o.open).length,0).toLocaleString('fa-IR');$('noPurchaseCount').textContent=stores.filter(s=>s.source==='JP'&&!s.hasSales).length.toLocaleString('fa-IR');$('activeFilterCount').textContent=activeCount().toLocaleString('fa-IR');}
function renderAll(){for(const [id,key] of [['areaFilter','area'],['routeFilter','route'],['personFilter','salesman'],['classFilter','customerClass']])fillMulti(id,valuesOf(key));fillMulti('dayFilter',['Today',...DAYS]);renderList();drawMap(filtered());renderStats();renderPlan();}
function showDetails(id){
 const s=stores.find(x=>String(x.id)===String(id));if(!s)return;
 const orderHtml=s.orders.map(o=>{const lines=o.lines.map(l=>`<div>${esc(l.name)} — ${l.qty} × ${l.price.toLocaleString('fa-IR')}</div>`).join('');return `<article class="order-item"><b>${esc(o.number||'سفارش')}</b><p>تاریخ: ${dateInfo(o.date).label} · تحویل هدف: ${dateInfo(o.target).label} · فاکتور: ${dateInfo(o.invoice).label}</p>${lines}<b>جمع: ${o.total.toLocaleString('fa-IR')}</b> · ${o.open?'باز':'فاکتور شده'}</article>`;}).join('');
 $('modalTitle').textContent=s.name;
 $('modalContent').innerHTML=`<div class="detail-block"><p>${esc(s.address||'آدرس ثبت نشده')}</p><div class="detail-grid"><span>کد مشتری: <b>${esc(s.id)}</b></span><span>کلاس: <b>${esc(s.customerClass||'—')}</b></span><span>ویزیتور: <b>${esc(s.salesman||'—')}</b></span><span>Route: <b>${esc(s.route||s.routeId||'—')}</b></span><span>تعداد سفارش: <b>${s.orders.length}</b></span><span>جمع مبلغ: <b>${s.total.toLocaleString('fa-IR')}</b></span></div><h3>سفارش‌ها</h3>${orderHtml||'<p>سفارشی ثبت نشده است.</p>'}<h3>کالاهای غیرفعال</h3><p>${s.inactive.map(x=>esc(x.name)).join('، ')||'موردی نیست'}</p><button class="primary-button full-button" id="detailAddPlan">افزودن به برنامه Auditor</button>${s.gps?'<button class="secondary-button full-button" id="detailRoute">مسیریابی</button>':''}</div>`;
 $('modal').style.display='flex';$('detailAddPlan').onclick=()=>{addToPlan(s);closeModal();};if(s.gps)$('detailRoute').onclick=()=>window.open(`https://www.google.com/maps/dir/?api=1&destination=${s.lat},${s.lng}`,'_blank','noopener');
}
function closeModal(){$('modal').style.display='none';}
function loadPlan(){try{plan=JSON.parse(localStorage.getItem(PLAN_KEY)||'[]');if(!Array.isArray(plan))plan=[];}catch{plan=[];}}
function savePlan(){localStorage.setItem(PLAN_KEY,JSON.stringify(plan));renderPlan();}
function selectedDate(){const raw=clean($('planDate').value);if(!raw)return new Date().toISOString().slice(0,10);return raw;}
function addToPlan(s){if(!s||!s.gps){message('فروشگاه فاقد GPS است و نمی‌تواند به برنامه Auditor اضافه شود.');return;}const date=selectedDate();if(plan.some(x=>x.customerId===s.id&&x.visitDate===date)){message('این مشتری برای این تاریخ در برنامه وجود دارد.');return;}plan.push({visitDate:date,day:$('planDay').value,customerId:s.id,storeName:s.name,customerName:s.name,address:s.address,routeName:s.route,latitude:s.lat,longitude:s.lng,priority:'Normal',selectionReason:'Manual selection',auditorComment:'',selectionSource:'Internal',status:'Pending'});savePlan();message('فروشگاه به برنامه اضافه شد.');}
function renderPlan(){const root=$('planList');if(!root)return;root.innerHTML=plan.map((x,i)=>`<article class="plan-item"><div><b>${esc(x.storeName)}</b><small>${esc(x.customerId)} · ${esc(x.visitDate)} · ${esc(x.routeName||'بدون مسیر')}</small></div><select data-plan-status="${i}"><option value="Pending" ${x.status==='Pending'?'selected':''}>در انتظار</option><option value="Completed" ${x.status==='Completed'?'selected':''}>انجام شد</option><option value="Not completed" ${x.status==='Not completed'?'selected':''}>انجام نشد</option></select><select data-plan-priority="${i}"><option ${x.priority==='Normal'?'selected':''} value="Normal">عادی</option><option ${x.priority==='High'?'selected':''} value="High">زیاد</option><option ${x.priority==='Urgent'?'selected':''} value="Urgent">فوری</option></select><input data-plan-comment="${i}" value="${esc(x.auditorComment||'')}" placeholder="یادداشت Auditor"><button type="button" data-plan-up="${i}">↑</button><button type="button" data-plan-down="${i}">↓</button><button type="button" data-plan-delete="${i}">حذف</button></article>`).join('')||'<div class="empty">هنوز بازدیدی ثبت نشده است.</div>';root.querySelectorAll('[data-plan-status]').forEach(e=>e.onchange=()=>{plan[+e.dataset.planStatus].status=e.value;localStorage.setItem(PLAN_KEY,JSON.stringify(plan));});root.querySelectorAll('[data-plan-priority]').forEach(e=>e.onchange=()=>{plan[+e.dataset.planPriority].priority=e.value;localStorage.setItem(PLAN_KEY,JSON.stringify(plan));});root.querySelectorAll('[data-plan-comment]').forEach(e=>e.onchange=()=>{plan[+e.dataset.planComment].auditorComment=e.value;localStorage.setItem(PLAN_KEY,JSON.stringify(plan));});root.querySelectorAll('[data-plan-delete]').forEach(e=>e.onclick=()=>{plan.splice(+e.dataset.planDelete,1);savePlan();});root.querySelectorAll('[data-plan-up]').forEach(e=>e.onclick=()=>movePlan(+e.dataset.planUp,-1));root.querySelectorAll('[data-plan-down]').forEach(e=>e.onclick=()=>movePlan(+e.dataset.planDown,1));}
function movePlan(i,d){const j=i+d;if(j<0||j>=plan.length)return;[plan[i],plan[j]]=[plan[j],plan[i]];savePlan();}
function renderPlanSearch(){const q=norm($('planSearch').value),root=$('planSearchResults');if(!q){root.innerHTML='';return;}const list=stores.filter(s=>s.source==='JP'&&s.gps&&[s.id,s.name,s.address].some(v=>norm(v).includes(q))).slice(0,12);root.innerHTML=list.map((s,i)=>`<button class="plan-search-result" data-add-store="${i}"><b>${esc(s.name)}</b><small>${esc(s.id)} · ${esc(s.address)}</small></button>`).join('')||'<p class="empty">موردی پیدا نشد.</p>';root.querySelectorAll('[data-add-store]').forEach(b=>b.onclick=()=>addToPlan(list[+b.dataset.addStore]));}
function exportPlan(){const rows=plan.map(x=>({'Visit Date':x.visitDate,'Customer ID':x.customerId,'Store Name':x.storeName,'Customer Name':x.customerName,'Address':x.address,'Route Name':x.routeName,'Latitude':x.latitude,'Longitude':x.longitude,'Priority':x.priority,'Selection Reason':x.selectionReason,'Auditor Comment':x.auditorComment,'Selection Source':x.selectionSource,'Status':x.status}));const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),'Audit Visit Plan');XLSX.writeFile(wb,'Audit_Visit_Plan.xlsx');}
async function importPlan(){const f=$('importPlanFile').files[0];if(!f){message('ابتدا فایل Excel را انتخاب کن.');return;}try{const rows=await parseExcel(f,0);const required=['Visit Date','Customer ID','Store Name'];const missing=required.filter(k=>!Object.hasOwn(rows[0]||{},k));if(missing.length)throw Error('ستون‌های ضروری وجود ندارند: '+missing.join('، '));const seen=new Set();externalRows=rows.map((r,i)=>{const id=clean(r['Customer ID']),date=clean(r['Visit Date']),key=id+'|'+date,store=stores.find(s=>s.source==='JP'&&norm(s.id)===norm(id));const errors=[];if(!store)errors.push('کد مشتری در JP پیدا نشد');if(!/^\d{4}[-/]\d{1,2}[-/]\d{1,2}$/.test(date)&&!/^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/.test(date))errors.push('فرمت تاریخ معتبر نیست');if(seen.has(key)||plan.some(p=>norm(p.customerId)===norm(id)&&p.visitDate===date))errors.push('ردیف تکراری است');if(!store?.gps)errors.push('مشتری GPS معتبر ندارد');seen.add(key);return {row:{visitDate:date,day:r.Day||$('planDay').value,customerId:id,storeName:r['Store Name']||store?.name||'',customerName:r['Customer Name']||store?.name||'',address:r.Address||store?.address||'',routeName:r['Route Name']||store?.route||'',latitude:number(r.Latitude)||store?.lat,longitude:number(r.Longitude)||store?.lng,priority:r.Priority||'Normal',selectionReason:r['Selection Reason']||'External suggestion',auditorComment:r['Auditor Comment']||'',selectionSource:'External',status:'Pending'},errors,accepted:false,index:i};});renderExternal();}catch(e){$('importReport').textContent='خطا: '+e.message;}}
function renderExternal(){const root=$('importReport');root.innerHTML='<h3>بررسی پیشنهادها</h3>'+externalRows.map((x,i)=>`<div class="import-row"><label><input type="checkbox" data-external-accept="${i}" ${x.errors.length?'disabled':''}><b>${esc(x.row.storeName)}</b></label><p>${esc(x.row.customerId)} · ${esc(x.row.visitDate)}</p><small>${x.errors.length?esc(x.errors.join('؛ ')):'معتبر؛ برای ورود انتخاب کن.'}</small></div>`).join('')+'<button type="button" id="acceptExternal" class="primary-button full-button">ثبت پیشنهادهای انتخاب‌شده</button>';root.querySelector('#acceptExternal').onclick=()=>{let n=0;root.querySelectorAll('[data-external-accept]:checked').forEach(el=>{const x=externalRows[+el.dataset.externalAccept];if(x&&!x.errors.length){plan.push(x.row);n++;}});savePlan();root.insertAdjacentHTML('beforeend',`<p>${n} ردیف اضافه شد.</p>`);};}
function bindDrive(){if(!window.google?.accounts?.oauth2)return;tokenClient=google.accounts.oauth2.initTokenClient({client_id:CLIENT_ID,scope:'https://www.googleapis.com/auth/drive.readonly',callback:r=>{if(r.error){message('ورود Google ناموفق بود: '+r.error);return;}accessToken=r.access_token;syncDrive();}});}
function connectDrive(){if(!tokenClient)bindDrive();if(!tokenClient){message('کتابخانه ورود Google هنوز بارگذاری نشده است؛ چند لحظه بعد دوباره بزن.');return;}tokenClient.requestAccessToken({prompt:'consent'});}
async function driveFiles(){const q=encodeURIComponent(`'${DRIVE_FOLDER}' in parents and trashed=false`);const r=await fetch(`https://www.googleapis.com/drive/v3/files?q=${q}&fields=files(id,name,mimeType,modifiedTime)&pageSize=100&orderBy=modifiedTime desc`,{headers:{Authorization:'Bearer '+accessToken}});if(!r.ok)throw Error('فهرست Drive دریافت نشد: '+r.status);return (await r.json()).files||[];}
async function downloadDrive(f){const r=await fetch(`https://www.googleapis.com/drive/v3/files/${f.id}?alt=media`,{headers:{Authorization:'Bearer '+accessToken}});if(!r.ok)throw Error('دانلود '+f.name+' ناموفق بود.');return r.arrayBuffer();}
function latest(files,prefix){return files.filter(f=>f.name.toLowerCase().startsWith(prefix.toLowerCase())&&f.name.toLowerCase().endsWith('.xlsx')).sort((a,b)=>new Date(b.modifiedTime)-new Date(a.modifiedTime))[0];}
async function syncDrive(){try{message('در حال دریافت آخرین فایل‌های Drive...');const fs=await driveFiles(),jf=latest(fs,'JP_'),tf=latest(fs,'TPA_Alborz'),mf=latest(fs,'Product Mapping');if(!jf||!tf)throw Error('فایل JP_ یا TPA_Alborz در پوشه پیدا نشد.');const [jb,tb,mb]=await Promise.all([downloadDrive(jf),downloadDrive(tf),mf?downloadDrive(mf):Promise.resolve(null)]);const parse=b=>{const w=XLSX.read(b,{type:'array'});return XLSX.utils.sheet_to_json(w.Sheets[w.SheetNames[0]],{range:0,defval:'',raw:false});};const jpW=XLSX.read(jb,{type:'array'}),tW=XLSX.read(tb,{type:'array'});const jp=XLSX.utils.sheet_to_json(jpW.Sheets[jpW.SheetNames[0]],{range:1,defval:'',raw:false});const tp=XLSX.utils.sheet_to_json(tW.Sheets[tW.SheetNames[0]],{range:3,defval:'',raw:false});let pm=[];if(mb){const mw=XLSX.read(mb,{type:'array'});pm=XLSX.utils.sheet_to_json(mw.Sheets[mw.SheetNames[0]],{range:0,defval:'',raw:false});}validate(jp,tp,pm);const old=[jpRows,tpaRows,mappingRows];try{jpRows=jp;tpaRows=tp;mappingRows=pm;window.sourceYear=sourceYear(tf.name);buildData();await saveLocal();}catch(e){[jpRows,tpaRows,mappingRows]=old;buildData();throw e;}renderAll();$('lastSync').textContent='آخرین همگام‌سازی: '+new Date().toLocaleString('fa-IR');message('همگام‌سازی با موفقیت انجام شد.');status('همگام');}catch(e){console.error(e);message('همگام‌سازی ناموفق بود؛ اطلاعات قبلی حفظ شد. '+e.message);}}
async function loadManual(){const jf=$('jpFile').files[0],tf=$('salesFile').files[0],mf=$('mappingFile').files[0];if(!jf||!tf){message('فایل Journey Plan و TPA را هر دو انتخاب کن.');return;}const b=$('loadBtn');b.disabled=true;try{message('در حال پردازش فایل‌ها...');const [jp,tp,pm]=await Promise.all([parseExcel(jf,1),parseExcel(tf,3),mf?parseExcel(mf,0):Promise.resolve([])]);validate(jp,tp,pm);const old=[jpRows,tpaRows,mappingRows];try{jpRows=jp;tpaRows=tp;mappingRows=pm;window.sourceYear=sourceYear(tf.name);buildData();await saveLocal();}catch(e){[jpRows,tpaRows,mappingRows]=old;buildData();throw e;}renderAll();$('lastSync').textContent='آخرین بارگذاری: '+new Date().toLocaleString('fa-IR');message('بارگذاری انجام شد.');status('آماده');}catch(e){message('خطا: '+e.message);}finally{b.disabled=false;}}
function clearFilters(){for(const id of ['customerTypeFilter','typeFilter','gpsFilter'])$(id).value='all';for(const id of ['areaFilter','routeFilter','personFilter','classFilter','dayFilter'])for(const o of $(id).options)o.selected=false;renderAll();}
function showPage(id){document.querySelectorAll('.page').forEach(p=>p.classList.toggle('active',p.id===id));document.querySelectorAll('.bottom-nav button').forEach(b=>b.classList.toggle('active',b.dataset.page===id));if(id==='mapPage'&&map)setTimeout(()=>map.invalidateSize(),150);}
function setLanguage(lang){document.documentElement.lang=lang;document.documentElement.dir=lang==='fa'?'rtl':'ltr';const dict={fa:{map:'نقشه',plan:'برنامه من',analysis:'تحلیل',settings:'تنظیمات',clear:'پاک‌کردن فیلترها',storeTitle:'نقشه فروشگاه‌ها',filterTitle:'فیلتر فروشگاه‌ها',results:'فروشگاه‌ها',search:'جستجوی نام، کد یا آدرس فروشگاه',load:'پردازش فایل‌های انتخاب‌شده',drive:'اتصال و دریافت آخرین فایل‌ها',export:'خروجی Excel برنامه بازدید'},en:{map:'Map',plan:'My Plan',analysis:'Insights',settings:'Settings',clear:'Clear filters',storeTitle:'Store map',filterTitle:'Store filters',results:'Stores',search:'Search store name, code or address',load:'Process selected files',drive:'Connect and sync latest files',export:'Export visit plan to Excel'}};const d=dict[lang];const nav=document.querySelectorAll('.bottom-nav button');[d.map,d.plan,d.analysis,d.settings].forEach((t,i)=>{if(nav[i])nav[i].lastChild.textContent=t;});$('clearFiltersBtn').textContent=d.clear;$('search').placeholder=d.search;$('loadBtn').textContent=d.load;$('driveBtn').textContent=d.drive;$('exportPlanBtn').textContent=d.export;document.querySelector('#mapPage .map-heading h2').textContent=d.storeTitle;document.querySelector('#mapPage .filter-panel .section-heading h2').textContent=d.filterTitle;document.querySelector('#mapPage .store-panel .section-heading h2').textContent=d.results;localStorage.setItem('tpa-language',lang);}
function bind(){document.querySelectorAll('.bottom-nav [data-page]').forEach(b=>b.onclick=()=>showPage(b.dataset.page));$('search').oninput=renderList;$('clearFiltersBtn').onclick=clearFilters;for(const id of ['customerTypeFilter','typeFilter','gpsFilter','areaFilter','routeFilter','personFilter','classFilter','dayFilter'])$(id).onchange=()=>{renderList();drawMap(filtered());renderStats();};$('loadBtn').onclick=loadManual;$('driveBtn').onclick=connectDrive;$('exportPlanBtn').onclick=exportPlan;$('importPlanBtn').onclick=importPlan;$('planSearch').oninput=renderPlanSearch;$('closeModal').onclick=closeModal;$('modal').onclick=e=>{if(e.target===$('modal'))closeModal();};$('themeBtn').onclick=()=>{const d=document.body.classList.contains('dark')?'light':'dark';document.body.classList.toggle('dark',d==='dark');localStorage.setItem('tpa-theme',d);};$('langBtn').onclick=()=>setLanguage(document.documentElement.lang==='fa'?'en':'fa');$('clearLocalBtn').onclick=async()=>{if(!confirm('داده‌های ذخیره‌شده و برنامه بازدید این دستگاه پاک شود؟'))return;try{const d=await db();await new Promise((res,rej)=>{const tx=d.transaction(STORE,'readwrite');tx.objectStore(STORE).clear();tx.oncomplete=res;tx.onerror=()=>rej(tx.error);});d.close();localStorage.removeItem(PLAN_KEY);plan=[];renderPlan();message('داده‌های محلی پاک شد.');}catch(e){message('پاک‌سازی ناموفق بود: '+e.message);}};$('nearbyBtn').onclick=()=>{if(!navigator.geolocation)return; navigator.geolocation.getCurrentPosition(p=>{const near=stores.filter(s=>s.gps).map(s=>({...s,distance:Math.hypot(s.lat-p.coords.latitude,s.lng-p.coords.longitude)})).sort((a,b)=>a.distance-b.distance).slice(0,30);renderAnalysis(near,'فروشگاه‌های نزدیک');},()=>message('دریافت موقعیت ناموفق بود.'));};$('noGpsBtn').onclick=()=>renderAnalysis(stores.filter(s=>s.source==='JP'&&!s.gps),'مشتریان بدون GPS');$('noPurchaseBtn').onclick=()=>renderAnalysis(stores.filter(s=>s.source==='JP'&&!s.hasSales),'بدون سابقه سفارش');$('openOrdersBtn').onclick=()=>renderAnalysis(stores.filter(s=>s.orders.some(o=>o.open)),'سفارش‌های بدون فاکتور');}
function renderAnalysis(list,title){$('analysisResults').innerHTML=`<h3>${esc(title)} (${list.length})</h3>`+list.map(s=>`<button class="store-row" data-analysis-id="${esc(s.id)}"><b>${esc(s.name)}</b><small>${esc(s.address||'')}</small></button>`).join('');$('analysisResults').querySelectorAll('[data-analysis-id]').forEach(b=>b.onclick=()=>showDetails(b.dataset.analysisId));}
async function init(){loadPlan();bind();initMap();bindDrive();window.addEventListener('load',bindDrive,{once:true});const theme=localStorage.getItem('tpa-theme');if(theme==='dark')document.body.classList.add('dark');setLanguage(localStorage.getItem('tpa-language')||'fa');const ok=await restoreLocal();if(!ok)message('برای شروع فایل‌ها را دستی بارگذاری کن یا به Google Drive متصل شو.');if('serviceWorker'in navigator&&location.protocol!=='file:')navigator.serviceWorker.register('./sw.js').catch(console.warn);window.addEventListener('online',()=>status('آنلاین'));window.addEventListener('offline',()=>status('آفلاین'));}
document.addEventListener('DOMContentLoaded',init);

/* Integrated specification completion layer: kept inside app.js; no extra script file required. */
/* TPA Navigation — specification completion layer
   Loaded after app.js. Keeps source data immutable and extends analysis/planning UI. */
(function(){
'use strict';
const $x=id=>document.getElementById(id);
const txt=v=>String(v??'').trim();
const normX=v=>txt(v).replace(/[۰-۹]/g,d=>'۰۱۲۳۴۵۶۷۸۹'.indexOf(d)).replace(/[٠-٩]/g,d=>'٠١٢٣٤٥٦٧٨٩'.indexOf(d)).replace(/\s+/g,'').replace(/\.0$/,'');
const escX=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const moneyX=n=>Number(n||0).toLocaleString('fa-IR')+' ریال';
const dayNamesX=['Saturday','Sunday','Monday','Tuesday','Wednesday','Thursday'];
const faDaysX={Saturday:'شنبه',Sunday:'یکشنبه',Monday:'دوشنبه',Tuesday:'سه‌شنبه',Wednesday:'چهارشنبه',Thursday:'پنجشنبه'};
const todayX=()=>['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'][new Date().getDay()];
const localTodayKeyX=()=>{const d=new Date();return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;};
let presentedHistoryX=[],auditMapX=null,auditMarkersX=[];
const validCoordX=(a,b)=>Number.isFinite(a)&&Number.isFinite(b)&&a>=-90&&a<=90&&b>=-180&&b<=180&&(a!==0||b!==0);
const toDateKeyX=d=>{if(!d)return '';const z=new Date(d);return Number.isNaN(z.getTime())?'':z.toISOString().slice(0,10);};
function jalaliToGregorianX(jy,jm,jd){
 let jy2=jy-979,jdn=365*jy2+Math.floor(jy2/33)*8+Math.floor((jy2%33+3)/4)+78+jd+(jm<7?(jm-1)*31:(jm-7)*30+186);
 let gy=1600+400*Math.floor(jdn/146097);jdn%=146097;let leap=true;
 if(jdn>=36525){jdn--;gy+=100*Math.floor(jdn/36524);jdn%=36524;if(jdn>=365)jdn++;else leap=false;}
 gy+=4*Math.floor(jdn/1461);jdn%=1461;if(jdn>=366){leap=false;jdn--;gy+=Math.floor(jdn/365);jdn%=365;}
 const md=[31,leap?29:28,31,30,31,30,31,31,30,31,30,31];let m=0;while(m<12&&jdn>=md[m])jdn-=md[m++];return new Date(Date.UTC(gy,m,jdn+1));
}
function jalaliKeyX(v,yearHint){if(typeof v==='number'&&v>20000&&v<80000)return toDateKeyX(new Date(Date.UTC(1899,11,30)+v*86400000));const s=normX(v);let m=s.match(/^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$/);if(m){const jy=+m[1],jm=+m[2],jd=+m[3];if(jy>1300&&jm>=1&&jm<=12&&jd>=1&&jd<=31)return toDateKeyX(jalaliToGregorianX(jy,jm,jd));}
 m=s.match(/^(\d{6})$/);if(m){const jy=Math.floor((yearHint||1405)/100)*100+Number(s.slice(0,2)),jm=+s.slice(2,4),jd=+s.slice(4,6);if(jm>=1&&jm<=12&&jd>=1&&jd<=(jm<=6?31:jm<=11?30:30))return toDateKeyX(jalaliToGregorianX(jy,jm,jd));}
 m=s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);if(m)return toDateKeyX(new Date(Date.UTC(+m[1],+m[2]-1,+m[3])));
 m=s.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/);if(m)return toDateKeyX(new Date(Date.UTC(+m[3],+m[1]-1,+m[2])));
 const d=new Date(v);return Number.isNaN(d.getTime())?'':toDateKeyX(d);
}
function displayJalaliX(v){const key=jalaliKeyX(v,window.sourceYear||1405);if(!key)return txt(v)||'—';const d=new Date(key+'T00:00:00Z');return d.toLocaleDateString('fa-IR-u-ca-persian',{year:'numeric',month:'2-digit',day:'2-digit',timeZone:'UTC'});}
function normalizeVisitDateX(v){const key=jalaliKeyX(v,window.sourceYear||1405);return key||'';}
function statusLabelX(s){const map={'Invoiced':'فاکتور شده','Closed/Undelivered':'بسته / تحویل‌نشده','Confirmed-NotAllocated':'تأیید شده / تخصیص نیافته','Confirm (Shipped)':'تأیید و ارسال شده','Closed/expired':'بسته / منقضی'};return map[s]||s||'نامشخص';}
function orderIsInvoicedX(o){return !!txt(o.invoice);}
function orderIsOverdueX(o){if(txt(o.invoice))return false;const target=jalaliKeyX(o.target,window.sourceYear||1405);return !!target&&target<localTodayKeyX();}
function getStoreByIdX(id){return stores.find(s=>String(s.id)===String(id));}
function refreshFinancialModelX(){
 for(const s of stores){
  const valid=s.orders.filter(o=>o.status!=='Closed/expired');
  s.confirmedTotal=valid.filter(orderIsInvoicedX).reduce((a,o)=>a+Number(o.total||0),0);
  s.uninvoicedTotal=valid.filter(o=>!orderIsInvoicedX(o)).reduce((a,o)=>a+Number(o.total||0),0);
  s.openOrders=valid.filter(o=>!txt(o.invoice));
  s.overdueOrders=valid.filter(orderIsOverdueX);
  s.lastValidOrder=valid.slice().sort((a,b)=>(jalaliKeyX(b.date,window.sourceYear||1405)||'').localeCompare(jalaliKeyX(a.date,window.sourceYear||1405)||''))[0]||null;
  const byProduct=new Map();
  for(const o of valid){const d=jalaliKeyX(o.date,window.sourceYear||1405);for(const l of o.lines){const k=normX(l.code)||normX(l.name);if(!k)continue;if(!byProduct.has(k))byProduct.set(k,{code:l.code,name:l.name,orders:new Set(),qty:0,amount:0,firstDate:d,lastDate:d,firstRaw:o.date,lastRaw:o.date});const p=byProduct.get(k);p.orders.add(o.number||`${d}-${p.orders.size}`);p.qty+=Number(l.qty||0);p.amount+=Number(l.amount||0);if(d&&(!p.firstDate||d<p.firstDate)){p.firstDate=d;p.firstRaw=o.date;}if(d&&(!p.lastDate||d>p.lastDate)){p.lastDate=d;p.lastRaw=o.date;}}}
  const latest=valid.slice().sort((a,b)=>(jalaliKeyX(b.date,window.sourceYear||1405)||'').localeCompare(jalaliKeyX(a.date,window.sourceYear||1405)||'')).slice(0,3);
  const latestCodes=new Set(latest.flatMap(o=>o.lines.map(l=>normX(l.code)||normX(l.name))));
  s.basket=[...byProduct.values()].map(p=>({...p,frequency:p.orders.size,active:latestCodes.has(normX(p.code)||normX(p.name))})).sort((a,b)=>b.frequency-a.frequency||b.qty-a.qty);
  s.inactive=s.basket.filter(p=>!p.active);
 }
}
function extractStoreNameX(address){const a=txt(address);if(!a)return '';const pattern=/(سوپر\s?مارکت|سوپرمارکت|هایپر\s?مارکت|هایپرمارکت|مارکت|فروشگاه|مواد\s?غذایی|ارزانسرای)\s*[\u0600-\u06FF\w\-]*/g;const matches=[...a.matchAll(pattern)];if(!matches.length)return '';const m=matches[matches.length-1];const value=a.slice(m.index).trim();return value.length>=4?value:'';}
function storeSearchTextX(s){return [s.id,s.name,s.customerName,s.address,s.salesman,s.route,s.routeId,s.area].map(normX).join('|');}
function applyEnhancedFiltersX(){
 const f=filters();
 const selectedDays=f.day.filter(d=>d&&d!=='__all__');
 let list=stores.filter(s=>{
  if(f.kind==='jp'&&s.source!=='JP')return false;
  if(f.kind==='salesOnly'&&s.source!=='TPA')return false;
  if(f.type==='sales'&&!s.hasSales)return false;
  if(f.type==='nosales'&&s.hasSales)return false;
  const hist=$x('presentedFilter')?.value||'all';
  const wasPresented=presentedHistoryX.some(h=>normX(h.customerId)===normX(s.id));
  if(hist==='new'&&wasPresented)return false;
  if(hist==='presented'&&!wasPresented)return false;
  if(f.gps==='yes'&&!s.gps)return false;
  if(f.gps==='no'&&s.gps)return false;
  // JP-only dimensions must not suppress TPA-only records.
  if(s.source==='JP'){
   if(f.area.length&&!f.area.includes(s.area))return false;
   if(f.route.length&&!f.route.includes(s.route||s.routeId))return false;
   if(f.person.length&&!f.person.includes(s.salesman))return false;
   if(f.cls.length&&!f.cls.includes(s.customerClass))return false;
   if(selectedDays.length&&!selectedDays.some(d=>s.days.includes(d)||(d==='Today'&&s.days.includes(todayX()))))return false;
  }
  return true;
 });
 const q=normX($x('search')?.value);
 if(q)list=list.filter(s=>storeSearchTextX(s).includes(q));
 return list;
}
function updateEnhancedFilterOptionsX(){
 const current=filters();
 const dimensions=[
  ['areaFilter','area',s=>s.area],
  ['routeFilter','route',s=>s.route||s.routeId],
  ['personFilter','person',s=>s.salesman],
  ['classFilter','cls',s=>s.customerClass]
 ];
 const selectedDays=current.day.filter(d=>d&&d!=='__all__');
 const matchesBase=(s,skipKey)=>{
  if(s.source!=='JP')return false;
  if(current.kind==='salesOnly')return false;
  if(current.type==='sales'&&!s.hasSales)return false;
  if(current.type==='nosales'&&s.hasSales)return false;
  if(current.gps==='yes'&&!s.gps)return false;
  if(current.gps==='no'&&s.gps)return false;
  if(skipKey!=='area'&&current.area.length&&!current.area.includes(s.area))return false;
  if(skipKey!=='route'&&current.route.length&&!current.route.includes(s.route||s.routeId))return false;
  if(skipKey!=='person'&&current.person.length&&!current.person.includes(s.salesman))return false;
  if(skipKey!=='cls'&&current.cls.length&&!current.cls.includes(s.customerClass))return false;
  if(selectedDays.length&&!selectedDays.some(d=>s.days.includes(d)||(d==='Today'&&s.days.includes(todayX()))))return false;
  return true;
 };
 for(const [id,key,getValue] of dimensions){
  const el=$x(id);if(!el)continue;
  const selectedNow=[...el.selectedOptions].map(o=>o.value);
  const vals=[...new Set(stores.filter(s=>matchesBase(s,key)).map(getValue).filter(Boolean))]
   .sort((a,b)=>String(a).localeCompare(String(b),'fa'));
  el.innerHTML=vals.map(v=>`<option value="${escX(v)}">${escX(v)}</option>`).join('');
  for(const option of el.options)option.selected=selectedNow.includes(option.value);
 }
 const dayEl=$x('dayFilter');
 if(dayEl){
  const selectedNow=[...dayEl.selectedOptions].map(o=>o.value).filter(v=>v!=='__all__');
  dayEl.innerHTML='<option value="__all__">همه روزها</option>'+['Today',...dayNamesX].map(d=>`<option value="${d}">${d==='Today'?'امروز':faDaysX[d]}</option>`).join('');
  for(const option of dayEl.options)option.selected=selectedNow.includes(option.value);
  if(!selectedNow.length)dayEl.options[0].selected=true;
 }
}
function renderList(){
 const list=applyEnhancedFiltersX();$x('resultCount').textContent=list.length.toLocaleString('fa-IR')+' فروشگاه';
 $x('storeList').innerHTML=list.map(s=>`<button class="store-row" data-id="${escX(s.id)}"><span class="dot ${s.source==='TPA'?'blue-dot':s.hasSales?'green-dot':'orange-dot'}"></span><span><b>${escX(s.name||s.customerName||s.id)}</b><small>${escX(s.address||'آدرس ثبت نشده')}</small></span><small>${s.hasSales?'دارای سفارش':'بدون سفارش'}${s.gps?'':' · بدون GPS'}</small></button>`).join('')||'<div class="empty">موردی پیدا نشد.</div>';
 $x('storeList').querySelectorAll('[data-id]').forEach(b=>b.addEventListener('click',()=>showDetails(b.dataset.id)));
}
function renderAll(){refreshFinancialModelX();updateEnhancedFilterOptionsX();renderList();drawMap(applyEnhancedFiltersX());renderStats();renderPlan();renderAnalysisSummaryX();renderQualityPanelX();}
function renderAnalysisSummaryX(){if(!$x('noGpsCount'))return;$x('noGpsCount').textContent=stores.filter(s=>s.source==='JP'&&!s.gps).length.toLocaleString('fa-IR');$x('openOrderCount').textContent=stores.reduce((n,s)=>n+(s.openOrders||[]).length,0).toLocaleString('fa-IR');$x('noPurchaseCount').textContent=stores.filter(s=>s.source==='JP'&&!s.hasSales).length.toLocaleString('fa-IR');}
function renderEnhancedDetailsX(s){
 const orders=s.orders.slice().sort((a,b)=>(jalaliKeyX(b.date,window.sourceYear||1405)||'').localeCompare(jalaliKeyX(a.date,window.sourceYear||1405)||''));
 const cards=orders.map(o=>`<article class="order-card"><div class="order-head"><b>${escX(o.number||'سفارش بدون شماره')}</b><span class="pill">${escX(statusLabelX(o.status))}</span></div><div class="order-meta">تاریخ سفارش: ${displayJalaliX(o.date)}<br>تحویل هدف: ${displayJalaliX(o.target)}<br>تاریخ فاکتور: ${txt(o.invoice)?displayJalaliX(o.invoice):'صادر نشده'}<br>${orderIsOverdueX(o)?'<b class="overdue-label">معوق</b>':''}</div><b>مبلغ سفارش: ${moneyX(o.total)}</b> · تعداد: ${Number(o.lines.reduce((n,l)=>n+Number(l.qty||0),0)).toLocaleString('fa-IR')}<details><summary>مشاهده محصولات (${o.lines.length})</summary>${o.lines.map(l=>`<div class="product"><b>${escX(l.name)}</b><small>کد: ${escX(l.code)} · تعداد: ${Number(l.qty||0).toLocaleString('fa-IR')} · مبلغ: ${moneyX(l.amount)}${l.mappingMissing?' · بدون مپینگ':''}</small></div>`).join('')}</details></article>`).join('');
 const basket=(s.basket||[]).map(p=>`<div class="basket-row"><div><b>${escX(p.name)}</b><small>کد کالا: ${escX(p.code||'—')} · دفعات خرید: ${p.frequency} · تعداد کل: ${p.qty}</small><small>اولین خرید: ${p.firstRaw?displayJalaliX(p.firstRaw):'—'} · آخرین خرید: ${p.lastRaw?displayJalaliX(p.lastRaw):'—'}</small></div><span class="basket-state ${p.active?'active':'inactive'}">${p.active?'فعال':'غیرفعال'}</span></div>`).join('');
 $x('modalTitle').textContent=s.name||s.customerName||s.id;
 $x('modalContent').innerHTML=`<div class="detail-block">
 <div class="detail-hero"><div><span class="detail-kicker">${s.source==='JP'?'JOURNEY PLAN':'TPA ONLY'}</span><h3>${escX(s.name||s.customerName||s.id)}</h3><p>${escX(s.address||'آدرس ثبت نشده')}</p></div><span class="store-status ${s.source==='TPA'?'tpa':s.hasSales?'sales':'no-sales'}">${s.source==='TPA'?'فقط TPA':s.hasSales?'دارای سفارش':'بدون سفارش'}</span></div>
 <div class="detail-grid">
 <div><small>کد مشتری</small><b>${escX(s.id)}</b></div><div><small>نام مشتری</small><b>${escX(s.customerName||s.name||'—')}</b></div>
 <div><small>Customer Class</small><b>${escX(s.customerClass||'—')}</b></div><div><small>فروشنده</small><b>${escX(s.salesman||'—')}</b></div>
 <div><small>Route Name</small><b>${escX(s.route||'در JP موجود نیست')}</b></div><div><small>RouteID</small><b>${escX(s.routeId||'—')}</b></div>
 <div><small>Area</small><b>${escX(s.area||'—')}</b></div><div><small>روزهای JP</small><b>${s.days?.map(d=>faDaysX[d]).join('، ')||'ثبت نشده'}</b></div>
 <div class="wide"><small>GPS ${s.gpsSource==='Auditor correction'?'· اصلاح‌شده توسط Auditor':''}</small><b>${s.gps?`${Number(s.lat).toFixed(6)}, ${Number(s.lng).toFixed(6)}`:'ثبت نشده'}</b></div></div>
 <div class="finance-grid"><div><small>مجموع خرید فاکتورشده</small><b>${moneyX(s.confirmedTotal)}</b></div><div><small>مبلغ سفارش‌های فاکتورنشده</small><b>${moneyX(s.uninvoicedTotal)}</b></div></div>
 <div class="detail-actions"><button class="secondary-button" id="editGPSButtonX">${s.gps?'اصلاح موقعیت GPS':'ثبت موقعیت GPS'}</button>${s.gps?`<a target="_blank" rel="noopener" href="https://www.google.com/maps/dir/?api=1&destination=${s.lat},${s.lng}">مسیریابی Google Maps</a><a target="_blank" rel="noopener" href="https://neshan.org/maps/share/${s.lat},${s.lng}">مسیریابی نشان</a>`:''}</div>
 <section id="gpsEditPanelX" class="gps-editor" hidden><div class="section-heading"><h3>انتخاب موقعیت روی نقشه</h3><span class="hint">روی نقشه بزن یا نشانگر را جابه‌جا کن.</span></div><div id="gpsPickerMapX" role="application" aria-label="انتخاب مختصات GPS"></div><p id="gpsPickerCoordinatesX" class="hint">هنوز نقطه‌ای انتخاب نشده است.</p><div class="gps-editor-actions"><button id="gpsPickerCurrentX" class="secondary-button">استفاده از GPS دستگاه</button><button id="gpsPickerSaveX" class="primary-button">ذخیره مختصات</button><button id="gpsPickerCancelX" class="tool-button">انصراف</button></div><p class="hint">اصلاح فقط روی همین دستگاه ذخیره می‌شود و فایل اصلی JP تغییر نمی‌کند.</p></section>
 <h3>تاریخچه سفارش‌ها (${orders.length})</h3>${cards||'<p>سفارشی ثبت نشده است.</p>'}<h3>تحلیل سبد خرید</h3><p>محصول در صورت حضور در یکی از سه سفارش معتبر اخیر، فعال محسوب می‌شود.</p><div class="basket-list">${basket||'<p>محصولی برای تحلیل وجود ندارد.</p>'}</div><button class="primary-button full-button" id="detailAddPlan">افزودن به برنامه Auditor</button></div>`;
 $x('modal').style.display='flex';if(localStorage.getItem('tpa-language')==='en')translateTreeX($x('modalContent'),'en');
 $x('detailAddPlan').onclick=()=>{addToPlan(s);closeModal();};
 $x('editGPSButtonX').onclick=()=>openGPSPickerX(s);
}
function showDetails(id){const s=getStoreByIdX(id);if(!s)return;renderEnhancedDetailsX(s);}
function activeCount(){const f=filters();return Object.entries(f).reduce((n,[k,v])=>n+(Array.isArray(v)?v.filter(x=>x!=='__all__').length:(v==='all'?0:1)),0)+(($x('presentedFilter')?.value||'all')==='all'?0:1);}
function renderStats(){const all=stores;$x('total').textContent=all.length.toLocaleString('fa-IR');$x('withSales').textContent=all.filter(s=>s.hasSales).length.toLocaleString('fa-IR');$x('withoutSales').textContent=all.filter(s=>s.source==='JP'&&!s.hasSales).length.toLocaleString('fa-IR');$x('noGpsCount').textContent=all.filter(s=>s.source==='JP'&&!s.gps).length.toLocaleString('fa-IR');$x('openOrderCount').textContent=all.reduce((n,s)=>n+(s.openOrders||[]).length,0).toLocaleString('fa-IR');$x('noPurchaseCount').textContent=all.filter(s=>s.source==='JP'&&!s.hasSales).length.toLocaleString('fa-IR');$x('activeFilterCount').textContent=activeCount().toLocaleString('fa-IR');}

const GPS_OVERRIDE_KEY_X='tpa_navigation_gps_overrides_v1';
let gpsPickerMapX=null,gpsPickerMarkerX=null,gpsPickerTargetX=null,gpsPickerPointX=null;
function readGPSOverridesX(){try{const v=JSON.parse(localStorage.getItem(GPS_OVERRIDE_KEY_X)||'{}');return v&&typeof v==='object'?v:{};}catch(e){console.warn(e);return {};}}
function applyGPSOverridesX(list){const saved=readGPSOverridesX();for(const store of list){const x=saved[normX(store.id)];if(x&&validCoordX(Number(x.lat),Number(x.lng))){store.lat=Number(x.lat);store.lng=Number(x.lng);store.gps=true;store.gpsSource='Auditor correction';}}}
function saveGPSOverrideX(store,lat,lng){if(!store||!validCoordX(lat,lng))throw Error('مختصات انتخاب‌شده معتبر نیست.');const all=readGPSOverridesX();all[normX(store.id)]={lat:Number(lat.toFixed(6)),lng:Number(lng.toFixed(6)),updatedAt:new Date().toISOString(),source:'Auditor'};localStorage.setItem(GPS_OVERRIDE_KEY_X,JSON.stringify(all));store.lat=Number(lat.toFixed(6));store.lng=Number(lng.toFixed(6));store.gps=true;store.gpsSource='Auditor correction';}
function closeGPSPickerX(){if(gpsPickerMapX){gpsPickerMapX.off();gpsPickerMapX.remove();gpsPickerMapX=null;}gpsPickerMarkerX=null;gpsPickerTargetX=null;gpsPickerPointX=null;}
function openGPSPickerX(store){
 if(!store)return;gpsPickerTargetX=store;
 const panel=$x('gpsEditPanelX'),node=$x('gpsPickerMapX');if(!panel||!node)return;
 panel.hidden=false;
 const center=store.gps?[store.lat,store.lng]:[35.7,51.4];
 setTimeout(()=>{
  if(!window.L){message('کتابخانه نقشه بارگذاری نشده است.');return;}
  closeGPSPickerX();
  gpsPickerTargetX=store;
  gpsPickerMapX=L.map(node,{zoomControl:true}).setView(center,store.gps?16:5);
  L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',{maxZoom:19,attribution:'© OpenStreetMap'}).addTo(gpsPickerMapX);
  const setPoint=ll=>{gpsPickerPointX={lat:ll.lat,lng:ll.lng};if(!gpsPickerMarkerX)gpsPickerMarkerX=L.marker(ll,{draggable:true}).addTo(gpsPickerMapX);else gpsPickerMarkerX.setLatLng(ll);gpsPickerMarkerX.off('dragend').on('dragend',e=>setPoint(e.target.getLatLng()));const label=$x('gpsPickerCoordinatesX');if(label)label.textContent=`عرض: ${ll.lat.toFixed(6)} · طول: ${ll.lng.toFixed(6)}`;};
  if(store.gps)setPoint({lat:store.lat,lng:store.lng});
  gpsPickerMapX.on('click',e=>setPoint(e.latlng));gpsPickerMapX.invalidateSize();
  const current=$x('gpsPickerCurrentX');if(current)current.onclick=()=>{if(!navigator.geolocation){message('موقعیت‌یابی پشتیبانی نمی‌شود.');return;}current.disabled=true;navigator.geolocation.getCurrentPosition(p=>{const ll={lat:p.coords.latitude,lng:p.coords.longitude};gpsPickerMapX.setView(ll,17);setPoint(ll);current.disabled=false;},()=>{message('دریافت GPS ناموفق بود؛ نقطه را دستی روی نقشه انتخاب کن.');current.disabled=false;},{enableHighAccuracy:true,timeout:15000});};
  const save=$x('gpsPickerSaveX');if(save)save.onclick=()=>{if(!gpsPickerPointX){message('ابتدا روی نقشه نقطه‌ای انتخاب کن.');return;}try{saveGPSOverrideX(gpsPickerTargetX,gpsPickerPointX.lat,gpsPickerPointX.lng);const target=gpsPickerTargetX;closeGPSPickerX();panel.hidden=true;renderAll();showDetails(target.id);message('مختصات اصلاح‌شده در همین دستگاه ذخیره شد.');}catch(e){message(e.message);}};
  const cancel=$x('gpsPickerCancelX');if(cancel)cancel.onclick=()=>{closeGPSPickerX();panel.hidden=true;};
 },120);
}

function buildData(){
 const pm=new Map();
 for(const r of mappingRows){
  const code=normX(r['Item Code']);
  if(code&&!pm.has(code))pm.set(code,txt(r['Product Name']||r['Product  Name']||r['Item']||r['Item Name']));
 }

 // Build an alias index without silently assigning ambiguous aliases to the last JP row.
 const jp=[],aliasOwners=new Map(),jpById=new Map();
 const aliasColumns=['CustomerI','New Customer Code','Oracle Code','Account Number'];
 for(const r of jpRows){
  const id=txt(r.CustomerI);
  const idKey=normX(id);
  if(!idKey)continue;
  const lat=number(r.Latitude),lng=number(r.Longitude);
  const days=dayNamesX.filter(d=>{
   const v=txt(r[d]).toLowerCase();
   return v!==''&&!['0','0.0','no','false','-'].includes(v);
  });
  const address=txt(r.Address||r[' Address']);
  const customerName=txt(r['Customer Name']);
  const storeName=extractStoreNameX(address);
  const s={
   id,name:storeName||customerName||id,storeName,customerName,address,
   customerClass:txt(r['Outlet Class']),salesman:txt(r['Salesman Name']),
   salesmanCode:txt(r['Salesman Code']),route:txt(r['Route Name']),
   routeId:txt(r.RouteID),area:txt(r['Area Name']),territory:txt(r.TerritoryNo),
   lat,lng,gps:validCoordX(lat,lng),days,source:'JP',orders:[],hasSales:false
  };
  jp.push(s);jpById.set(idKey,s);
  for(const col of aliasColumns){
   const raw=txt(r[col]);
   if(!raw)continue;
   const variants=[...new Set([normX(raw),...String(raw).split('|').map(normX)].filter(Boolean))];
   for(const key of variants){
    if(!aliasOwners.has(key))aliasOwners.set(key,new Set());
    aliasOwners.get(key).add(s);
   }
  }
 }
 const jpByCode=new Map(),ambiguousAliases=[];
 for(const [key,owners] of aliasOwners){
  if(owners.size===1)jpByCode.set(key,[...owners][0]);
  else ambiguousAliases.push(key);
 }

 const groups=new Map();
 let expiredRows=0,skippedRows=0,matchedRows=0,unmatchedRows=0;
 for(const r of tpaRows){
  const status=txt(r.Status||r['وضعیت سفارش']);
  if(status.toLowerCase()==='closed/expired'){expiredRows++;continue;}
  const rawCode=txt(r['کد مشتری']||r['Customer Code']);
  const rawOrder=txt(r['شماره سفارش']||r['Order Number']||r['Order No']||r.OrderNumber);
  const code=normX(rawCode),orderKey=normX(rawOrder);
  if(!code||!orderKey){skippedRows++;continue;}
  const variants=[...new Set([code,...String(rawCode).split('|').map(normX)].filter(Boolean))];
  const customer=variants.map(k=>jpByCode.get(k)).find(Boolean)||null;
  if(customer)matchedRows++;else unmatchedRows++;
  const canonical=customer?normX(customer.id):code;
  const key=canonical+'|'+orderKey;
  if(!groups.has(key)){
   groups.set(key,{
    customer,code:rawCode,order:rawOrder,rows:[],
    date:txt(r['تاریخ']||r.Date||r['Order Date']),
    target:txt(r['Target Deilvery Date']||r['Target Delivery Date']),
    invoice:txt(r.invoicedate||r['Invoice Date']),
    status
   });
  }
  const g=groups.get(key);
  g.rows.push(r);
  if(!g.date)g.date=txt(r['تاریخ']||r.Date||r['Order Date']);
  if(!g.target)g.target=txt(r['Target Deilvery Date']||r['Target Delivery Date']);
  if(!g.invoice)g.invoice=txt(r.invoicedate||r['Invoice Date']);
  if(!g.status)g.status=status;
 }

 const tpaOnly=new Map(),unmappedCodes=new Set();
 for(const g of groups.values()){
  const lines=g.rows.map(r=>{
   const code=txt(r['کد کالا']||r['Item Code']||r.ItemCode);
   const codeKey=normX(code);
   const mapped=!!(codeKey&&pm.has(codeKey));
   if(codeKey&&!mapped)unmappedCodes.add(codeKey);
   const qty=number(r['تعداد']??r.Quantity??r.Qty??r['Order Quantity']);
   const price=number(r['قیمت واحد']??r['Unit Price']??r.Price);
   return {
    code,name:mapped?pm.get(codeKey):(txt(r['نام کالا']||r['Item Name']||r['Item Description']||r.Description)||code),
    mapped,mappingMissing:!!codeKey&&!mapped,qty,price,amount:qty*price
   };
  });
  const o={
   number:g.order,date:g.date,target:g.target,invoice:g.invoice,status:g.status,
   open:!txt(g.invoice),lines,total:lines.reduce((n,l)=>n+l.amount,0),
   routeTPA:txt(g.rows[0]['نام مسیر'])
  };
  if(g.customer){
   g.customer.orders.push(o);g.customer.hasSales=true;
  }else{
   const k=normX(g.code)||'unknown-'+tpaOnly.size;
   if(!tpaOnly.has(k)){
    const firstRow=g.rows[0],address=txt(firstRow['آدرس']),customerName=txt(firstRow['نام مشتری']),storeName=extractStoreNameX(address);
    tpaOnly.set(k,{
     id:g.code||k,name:storeName||customerName||g.code,storeName,customerName,address,
     customerClass:txt(firstRow['کلاس مشتری']),salesman:'',route:txt(firstRow['نام مسیر']),
     routeId:'',area:'',lat:0,lng:0,gps:false,days:[],source:'TPA',orders:[],hasSales:true
    });
   }
   tpaOnly.get(k).orders.push(o);
  }
 }
 stores=[...jp,...tpaOnly.values()];
 applyGPSOverridesX(stores);
 for(const s of stores){
  s.orders.sort((a,b)=>(jalaliKeyX(b.date,window.sourceYear||1405)||'').localeCompare(jalaliKeyX(a.date,window.sourceYear||1405)||''));
  s.orderCount=s.orders.length;
  s.total=s.orders.reduce((n,o)=>n+o.total,0);
 }
 window.dataQualityX={
  jpCustomers:jp.length,tpaOnlyCustomers:tpaOnly.size,orderGroups:groups.size,
  matchedOrderGroups:[...groups.values()].filter(g=>!!g.customer).length,
  unmatchedOrderGroups:[...groups.values()].filter(g=>!g.customer).length,
  expiredRows,skippedRows,ambiguousAliasCount:ambiguousAliases.length,
  unmappedCodes:[...unmappedCodes].sort()
 };
 refreshFinancialModelX();
}
function dateRangeOptionsX(){const now=new Date(),monthStart=new Date(now.getFullYear(),now.getMonth(),1),prevStart=new Date(now.getFullYear(),now.getMonth()-1,1),prevEnd=new Date(now.getFullYear(),now.getMonth(),0);return {current:[toDateKeyX(monthStart),toDateKeyX(now)],previous:[toDateKeyX(prevStart),toDateKeyX(prevEnd)]};}
function addAnalysisControlsX(){const page=$x('analysisPage');if(!page||$x('analysisDateFrom'))return;const section=document.createElement('section');section.className='panel';section.id='analysisControls';section.innerHTML=`<h2>تحلیل آخرین سفارش معتبر در بازه</h2><div class="analysis-controls"><label>از تاریخ (شمسی)<input id="analysisDateFrom" placeholder="1405/01/01"></label><label>تا تاریخ (شمسی)<input id="analysisDateTo" placeholder="1405/12/29"></label><label>گروه سفارش<select id="analysisGroup"><option value="all">همه گروه‌ها</option><option value="no-purchase">بدون خرید</option><option value="Invoiced">Invoiced</option><option value="Closed/Undelivered">Closed/Undelivered</option><option value="not-delivered">تأییدشده / ارسال‌شده</option></select></label><label>نوع مشتری<select id="analysisCustomerType"><option value="all">همه</option><option value="jp">فقط JP</option><option value="tpa">فقط TPA-only</option></select></label><button id="runOrderAnalysis" class="primary-button full-button">اجرای تحلیل</button></div><div id="orderAnalysisResults"></div>`;page.insertBefore(section,page.children[1]);$x('runOrderAnalysis').onclick=runOrderAnalysisX;}
function runOrderAnalysisX(){const fromRaw=txt($x('analysisDateFrom').value),toRaw=txt($x('analysisDateTo').value),from=jalaliKeyX(fromRaw,window.sourceYear||1405),to=jalaliKeyX(toRaw,window.sourceYear||1405);if((fromRaw&&!from)||(toRaw&&!to)){message('فرمت تاریخ تحلیل معتبر نیست؛ تاریخ شمسی یا میلادی وارد کن.');return;}if(from&&to&&from>to){message('تاریخ شروع نباید بعد از تاریخ پایان باشد.');return;}const group=$x('analysisGroup').value,type=$x('analysisCustomerType').value;const rows=[];for(const s of stores){if(type==='jp'&&s.source!=='JP')continue;if(type==='tpa'&&s.source!=='TPA')continue;const valid=s.orders.filter(o=>o.status!=='Closed/expired').filter(o=>{const d=jalaliKeyX(o.date,window.sourceYear||1405);return (!from||d>=from)&&(!to||d<=to);}).sort((a,b)=>(jalaliKeyX(b.date,window.sourceYear||1405)||'').localeCompare(jalaliKeyX(a.date,window.sourceYear||1405)||''));const last=valid[0]||null;let g='no-purchase';if(last){if(last.status==='Invoiced')g='Invoiced';else if(last.status==='Closed/Undelivered')g='Closed/Undelivered';else if(['Confirmed-NotAllocated','Confirm (Shipped)'].includes(last.status))g='not-delivered';else g=last.status;}if(group!=='all'&&group!==g)continue;rows.push({s,last,g});}const root=$x('orderAnalysisResults');root.innerHTML=`<p>تعداد فروشگاه‌های منطبق: <b>${rows.length.toLocaleString('fa-IR')}</b></p>`+rows.map(x=>`<button class="store-row" data-analysis-store="${escX(x.s.id)}"><span><b>${escX(x.s.name)}</b><small>${escX(x.s.id)} · ${escX(x.last?statusLabelX(x.last.status):'بدون خرید در بازه')} · ${x.last?displayJalaliX(x.last.date):'—'}</small></span></button>`).join('');root.querySelectorAll('[data-analysis-store]').forEach(b=>b.onclick=()=>showDetails(b.dataset.analysisStore));}
function ensureAuditMapX(){if($x('auditMap')||!$x('planPage'))return;const section=document.createElement('section');section.className='panel';section.innerHTML='<div class="section-heading"><h2>نقشه بازدیدهای Auditor</h2><span class="hint">سبز: انجام شد · قرمز: انجام نشد · خاکستری: در انتظار</span></div><div id="auditMap"></div>';const planList=$x('planList')?.closest('.panel');$x('planPage').insertBefore(section,planList||null);if(window.L){auditMapX=L.map('auditMap',{zoomControl:false}).setView([35.7,51.4],10);L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',{maxZoom:19,attribution:'© OpenStreetMap'}).addTo(auditMapX);L.control.zoom({position:'bottomright'}).addTo(auditMapX);}}
function showPage(id){document.querySelectorAll('.page').forEach(p=>p.classList.toggle('active',p.id===id));document.querySelectorAll('.bottom-nav button').forEach(b=>b.classList.toggle('active',b.dataset.page===id));if(id==='mapPage'&&map)setTimeout(()=>map.invalidateSize(),150);if(id==='planPage'&&auditMapX)setTimeout(()=>auditMapX.invalidateSize(),150);}
function renderPlan(){const root=$x('planList');if(!root)return;ensureAuditMapX();const sorted=plan.slice().sort((a,b)=>(normalizeVisitDateX(a.visitDate)||'').localeCompare(normalizeVisitDateX(b.visitDate)||''));const groups=new Map();for(const p of sorted){const k=p.visitDate||'بدون تاریخ';if(!groups.has(k))groups.set(k,[]);groups.get(k).push(p);}root.innerHTML=[...groups.entries()].map(([date,items])=>{const done=items.filter(p=>p.status==='Completed').length,notDone=items.filter(p=>p.status==='Not completed').length,pending=items.length-done-notDone;return `<section class="week-day"><h3>${escX(date)} <small>(${items.length} بازدید · ${done} انجام‌شده · ${notDone} انجام‌نشده · ${pending} در انتظار)</small></h3>${items.map(p=>{const i=plan.indexOf(p);return `<article class="plan-item"><div><b>${escX(p.storeName||p.customerName)}</b><small>${escX(p.customerId)} · ${escX(p.routeName||'بدون مسیر')} · ${p.selectionSource==='External'?'پیشنهاد مدیر':'انتخاب داخلی'}</small></div><select data-plan-status="${i}"><option value="Pending" ${p.status==='Pending'?'selected':''}>در انتظار</option><option value="Completed" ${p.status==='Completed'?'selected':''}>انجام شد</option><option value="Not completed" ${p.status==='Not completed'?'selected':''}>انجام نشد</option></select><select data-plan-priority="${i}"><option value="High" ${p.priority==='High'?'selected':''}>زیاد</option><option value="Medium" ${!p.priority||p.priority==='Medium'?'selected':''}>متوسط</option><option value="Low" ${p.priority==='Low'?'selected':''}>کم</option></select><input data-plan-date="${i}" value="${escX(p.visitDate||'')}" placeholder="تاریخ بازدید"><input data-plan-comment="${i}" value="${escX(p.auditorComment||'')}" placeholder="یادداشت Auditor"><button type="button" data-plan-up="${i}">↑</button><button type="button" data-plan-down="${i}">↓</button><button type="button" data-plan-delete="${i}">حذف</button></article>`;}).join('')}</section>`;}).join('')||'<div class="empty">هنوز بازدیدی ثبت نشده است.</div>';root.querySelectorAll('[data-plan-status]').forEach(e=>e.onchange=()=>{plan[+e.dataset.planStatus].status=e.value;savePlan();});root.querySelectorAll('[data-plan-priority]').forEach(e=>e.onchange=()=>{plan[+e.dataset.planPriority].priority=e.value;savePlan();});root.querySelectorAll('[data-plan-date]').forEach(e=>e.onchange=()=>{const i=+e.dataset.planDate,date=normalizeVisitDateX(e.value);if(!date){message('تاریخ بازدید معتبر نیست.');renderPlan();return;}if(plan.some((p,j)=>j!==i&&normX(p.customerId)===normX(plan[i].customerId)&&normalizeVisitDateX(p.visitDate)===date)){message('این فروشگاه برای تاریخ جدید از قبل در برنامه است.');renderPlan();return;}plan[i].visitDate=date;savePlan();});root.querySelectorAll('[data-plan-comment]').forEach(e=>e.onchange=()=>{plan[+e.dataset.planComment].auditorComment=e.value;localStorage.setItem(PLAN_KEY,JSON.stringify(plan));});root.querySelectorAll('[data-plan-delete]').forEach(e=>e.onclick=()=>{plan.splice(+e.dataset.planDelete,1);savePlan();});root.querySelectorAll('[data-plan-up]').forEach(e=>e.onclick=()=>movePlan(+e.dataset.planUp,-1));root.querySelectorAll('[data-plan-down]').forEach(e=>e.onclick=()=>movePlan(+e.dataset.planDown,1));if(auditMapX){auditMarkersX.forEach(m=>auditMapX.removeLayer(m));auditMarkersX=[];const withGps=plan.filter(p=>validCoordX(Number(p.latitude),Number(p.longitude)));for(const p of withGps){const color=p.status==='Completed'?'#15966b':p.status==='Not completed'?'#d94c4c':'#8796a5';const marker=L.circleMarker([Number(p.latitude),Number(p.longitude)],{radius:9,color:'#fff',weight:2,fillColor:color,fillOpacity:1}).addTo(auditMapX);marker.bindTooltip(`${p.storeName||p.customerId} · ${p.status}`);marker.on('click',()=>{const s=getStoreByIdX(p.customerId);if(s)showDetails(s.id);});auditMarkersX.push(marker);}if(withGps.length){const bounds=L.latLngBounds(withGps.map(p=>[Number(p.latitude),Number(p.longitude)]));auditMapX.fitBounds(bounds,{padding:[25,25],maxZoom:14});}setTimeout(()=>auditMapX.invalidateSize(),120);}}

function addToPlan(s){if(!s||!s.gps){message('فروشگاه فاقد GPS معتبر است و به برنامه اضافه نمی‌شود.');return;}const raw=txt($x('planDate').value);const date=normalizeVisitDateX(raw)||localTodayKeyX();if(plan.some(p=>normX(p.customerId)===normX(s.id)&&normalizeVisitDateX(p.visitDate)===date)){message('این فروشگاه برای این تاریخ قبلاً در برنامه است.');return;}plan.push({visitDate:date,day:$x('planDay').value,customerId:s.id,storeName:s.name,customerName:s.customerName||s.name,address:s.address,routeName:s.route||'',latitude:s.lat,longitude:s.lng,priority:'Medium',selectionReason:'Manual selection',auditorComment:'',selectionSource:'Internal',status:'Pending'});savePlan();message('فروشگاه به برنامه Auditor اضافه شد.');}
function renderPlanSearch(){const q=normX($x('planSearch').value),root=$x('planSearchResults');if(!q){root.innerHTML='';return;}const list=stores.filter(s=>s.gps&&storeSearchTextX(s).includes(q)).slice(0,20);root.innerHTML=list.map((s,i)=>`<button class="plan-search-result" data-add-store="${i}"><b>${escX(s.name)}</b><small>${escX(s.id)} · ${escX(s.address||'')}</small></button>`).join('')||'<p class="empty">موردی پیدا نشد.</p>';root.querySelectorAll('[data-add-store]').forEach(b=>b.onclick=()=>addToPlan(list[+b.dataset.addStore]));}
function exportPlan(){const headers=['Visit Date','Customer ID','Store Name','Customer Name','Address','Route Name','Latitude','Longitude','Priority','Selection Reason','Auditor Comment','Selection Source'];const rows=plan.map(p=>({'Visit Date':normalizeVisitDateX(p.visitDate)||p.visitDate,'Customer ID':p.customerId,'Store Name':p.storeName,'Customer Name':p.customerName,'Address':p.address,'Route Name':p.routeName,'Latitude':p.latitude,'Longitude':p.longitude,'Priority':p.priority||'Medium','Selection Reason':p.selectionReason||'','Auditor Comment':p.auditorComment||'','Selection Source':p.selectionSource==='External'?'External':'Internal'}));const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows,{header:headers}), 'Audit Visit Plan');XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([['راهنما','Visit Date و Customer ID الزامی هستند. Selection Source را برنامه تعیین می‌کند. ردیف‌های واردشده از مدیر به‌صورت پیشنهاد External نمایش داده می‌شوند.']]),'Instructions');XLSX.writeFile(wb,'Audit_Visit_Plan.xlsx');}
function renderExternal(){const root=$x('importReport');if(!externalRows.length){root.innerHTML='<p>پیشنهاد واردشده‌ای وجود ندارد.</p>';return;}root.innerHTML='<h3>پیشنهادهای فایل مدیر</h3>'+externalRows.map((x,i)=>`<div class="import-row"><label><input type="checkbox" data-external-accept="${i}" ${x.errors.length?'disabled':''}><b>${escX(x.row.storeName||x.row.customerId)}</b></label><p>${escX(x.row.customerId)} · ${escX(x.row.visitDate)}</p><small>${x.errors.length?escX(x.errors.join('؛ ')):'معتبر؛ برای پذیرش انتخاب کن.'}</small></div>`).join('')+'<button type="button" id="acceptExternal" class="primary-button full-button">افزودن پیشنهادهای انتخاب‌شده</button>';root.querySelector('#acceptExternal').onclick=()=>{let n=0;root.querySelectorAll('[data-external-accept]:checked').forEach(el=>{const x=externalRows[+el.dataset.externalAccept];if(x&&!x.errors.length){plan.push({...x.row,selectionSource:'External'});n++;}});savePlan();message(`${n} پیشنهاد خارجی به برنامه افزوده شد.`);externalRows=[];renderExternal();};}
async function importPlan(){const f=$x('importPlanFile').files[0];if(!f){message('ابتدا فایل Excel را انتخاب کن.');return;}try{const wb=XLSX.read(await f.arrayBuffer(),{type:'array'});const sheet=wb.Sheets['Audit Visit Plan']||wb.Sheets[wb.SheetNames[0]];const rows=XLSX.utils.sheet_to_json(sheet,{defval:'',raw:false});const required=['Visit Date','Customer ID'];const missing=required.filter(k=>!Object.prototype.hasOwnProperty.call(rows[0]||{},k));if(missing.length)throw Error('ستون ضروری وجود ندارد: '+missing.join('، '));const seen=new Set();externalRows=rows.map(r=>{const id=txt(r['Customer ID']),dateRaw=txt(r['Visit Date']),date=normalizeVisitDateX(dateRaw),key=normX(id)+'|'+date,store=stores.find(s=>[normX(s.id),...normX(s.id).split('|')].includes(normX(id)));const errors=[];if(!store)errors.push('کد مشتری در داده‌های فعلی پیدا نشد');if(!date)errors.push('تاریخ معتبر نیست');if(seen.has(key)||plan.some(p=>normX(p.customerId)===normX(id)&&normalizeVisitDateX(p.visitDate)===date))errors.push('ردیف تکراری است');if(!store?.gps)errors.push('مشتری GPS معتبر ندارد');seen.add(key);const lat=Number(r.Latitude),lng=Number(r.Longitude);if(store&&validCoordX(lat,lng)&&(Math.abs(lat-store.lat)>.00001||Math.abs(lng-store.lng)>.00001))errors.push('مختصات فایل با مختصات JP متفاوت است؛ مختصات JP حفظ می‌شود');return {row:{visitDate:date,day:r.Day||$x('planDay').value,customerId:id,storeName:r['Store Name']||store?.name||'',customerName:r['Customer Name']||store?.customerName||'',address:store?.address||r.Address||'',routeName:store?.route||r['Route Name']||'',latitude:store?.lat,longitude:store?.lng,priority:['High','Medium','Low'].includes(r.Priority)?r.Priority:'Medium',selectionReason:r['Selection Reason']||'External suggestion',auditorComment:r['Auditor Comment']||'',selectionSource:'External',status:'Pending'},errors};});renderExternal();}catch(e){$x('importReport').textContent='خطا در بررسی فایل: '+e.message;}}
function clearFilters(){for(const id of ['customerTypeFilter','typeFilter','gpsFilter'])$x(id).value='all';for(const id of ['areaFilter','routeFilter','personFilter','classFilter','dayFilter'])for(const o of $x(id).options)o.selected=false;$x('search').value='';renderAll();}
function bindEnhancedX(){
 document.querySelectorAll('.bottom-nav [data-page]').forEach(b=>b.onclick=()=>showPage(b.dataset.page));
 for(const id of ['customerTypeFilter','typeFilter','gpsFilter','areaFilter','routeFilter','personFilter','classFilter','dayFilter'])if($x(id))$x(id).onchange=()=>{updateEnhancedFilterOptionsX();renderList();drawMap(applyEnhancedFiltersX());renderStats();};
 if($x('search'))$x('search').oninput=renderList;if($x('clearFiltersBtn'))$x('clearFiltersBtn').onclick=clearFilters;
 if($x('importPlanBtn'))$x('importPlanBtn').onclick=importPlan;if($x('loadBtn'))$x('loadBtn').onclick=loadManualEnhancedX;if($x('exportPlanBtn'))$x('exportPlanBtn').onclick=exportPlan;if($x('planSearch'))$x('planSearch').oninput=renderPlanSearch;
 if($x('nearbyBtn'))$x('nearbyBtn').onclick=()=>{if(!navigator.geolocation){message('موقعیت‌یابی در این مرورگر پشتیبانی نمی‌شود.');return;}navigator.geolocation.getCurrentPosition(p=>{const near=stores.filter(s=>s.gps).map(s=>({...s,distance:6371000*2*Math.asin(Math.sqrt(Math.sin((s.lat-p.coords.latitude)*Math.PI/360)**2+Math.cos(p.coords.latitude*Math.PI/180)*Math.cos(s.lat*Math.PI/180)*Math.sin((s.lng-p.coords.longitude)*Math.PI/360)**2))})).sort((a,b)=>a.distance-b.distance);renderAnalysis(near.slice(0,50),'فروشگاه‌های نزدیک — فاصله مستقیم');},e=>message('دریافت موقعیت ناموفق بود: '+e.message),{enableHighAccuracy:true,timeout:15000});};
 if($x('noGpsBtn'))$x('noGpsBtn').onclick=()=>renderAnalysis(stores.filter(s=>s.source==='JP'&&!s.gps),'مشتریان بدون GPS');
 if($x('noPurchaseBtn'))$x('noPurchaseBtn').onclick=()=>renderAnalysis(stores.filter(s=>s.source==='JP'&&!s.hasSales),'بدون سابقه سفارش');
 if($x('openOrdersBtn'))$x('openOrdersBtn').onclick=()=>renderAnalysis(stores.filter(s=>(s.openOrders||[]).length),'سفارش‌های بدون فاکتور');
}
function renderAnalysis(list,title){$x('analysisResults').innerHTML=`<h3>${escX(title)} (${list.length.toLocaleString('fa-IR')})</h3>`+list.map(s=>`<button class="store-row" data-analysis-id="${escX(s.id)}"><span><b>${escX(s.name)}</b><small>${escX(s.address||'')}</small></span></button>`).join('');$x('analysisResults').querySelectorAll('[data-analysis-id]').forEach(b=>b.onclick=()=>showDetails(b.dataset.analysisId));}
async function loadManualEnhancedX(){const jf=$x('jpFile')?.files?.[0],tf=$x('salesFile')?.files?.[0],mf=$x('mappingFile')?.files?.[0];if(!jf||!tf){message('فایل Journey Plan و TPA را هر دو انتخاب کن.');return;}const btn=$x('loadBtn');if(btn)btn.disabled=true;const old=[jpRows,tpaRows,mappingRows,window.sourceYear];try{message('در حال پردازش فایل‌ها...');const [jp,tp]=await Promise.all([parseExcel(jf,1),parseExcel(tf,3)]);if(jp.length<100||tp.length<100)throw Error('تعداد ردیف‌ها غیرعادی کم است یا فایل ناقص است.');let pm=mappingRows;if(mf){try{const parsed=await parseExcel(mf,0);if(!parsed.length||!Object.prototype.hasOwnProperty.call(parsed[0]||{},'Item Code'))throw Error('ساختار Product Mapping معتبر نیست.');pm=parsed;}catch(e){message('Product Mapping انتخاب‌شده معتبر نیست؛ نسخه قبلی آن حفظ شد. '+e.message);}}validate(jp,tp,pm);jpRows=jp;tpaRows=tp;mappingRows=pm;window.sourceYear=sourceYear(tf.name);buildData();await saveLocal();renderAll();$x('lastSync').textContent='آخرین بارگذاری: '+new Date().toLocaleString('fa-IR')+' · '+jf.name+' · '+tf.name;message(`بارگذاری موفق؛ ${stores.length.toLocaleString('fa-IR')} فروشگاه پردازش شد.`);status('آماده');}
 catch(e){[jpRows,tpaRows,mappingRows,window.sourceYear]=old;buildData();message('بارگذاری ناموفق بود؛ داده قبلی حفظ شد. '+e.message);}
 finally{if(btn)btn.disabled=false;}}
function filenameDateX(name){const m=String(name||'').match(/(14\d{6})/);return m?m[1]:'';}
function latest(files,prefix){return files.filter(f=>f.name.toLowerCase().startsWith(prefix.toLowerCase())&&f.name.toLowerCase().endsWith('.xlsx')).sort((a,b)=>{const da=filenameDateX(a.name),db=filenameDateX(b.name);if(da!==db)return db.localeCompare(da);return new Date(b.modifiedTime)-new Date(a.modifiedTime);})[0];}
async function syncDrive(){
 try{
  message('در حال بررسی نسخه‌های Google Drive...');const fs=await driveFiles();
  const candidates={jp:fs.filter(f=>/^JP_.*\.xlsx$/i.test(f.name)).sort((a,b)=>filenameDateX(b.name).localeCompare(filenameDateX(a.name))||new Date(b.modifiedTime)-new Date(a.modifiedTime)),tpa:fs.filter(f=>/^TPA_Alborz.*\.xlsx$/i.test(f.name)).sort((a,b)=>filenameDateX(b.name).localeCompare(filenameDateX(a.name))||new Date(b.modifiedTime)-new Date(a.modifiedTime)),pm:fs.filter(f=>/^Product Mapping.*\.xlsx$/i.test(f.name)).sort((a,b)=>new Date(b.modifiedTime)-new Date(a.modifiedTime))};
  if(!candidates.jp.length||!candidates.tpa.length)throw Error('فایل JP_ یا TPA_Alborz در پوشه Drive پیدا نشد.');
  let accepted=null,lastError=null;
  // Try newest-to-oldest pairs; a corrupt newest file must never replace a valid local dataset.
  for(const jf of candidates.jp.slice(0,5)){
   for(const tf of candidates.tpa.slice(0,5)){
    try{
     const [jb,tb,mb]=await Promise.all([downloadDrive(jf),downloadDrive(tf),candidates.pm[0]?downloadDrive(candidates.pm[0]):Promise.resolve(null)]);
     const jw=XLSX.read(jb,{type:'array'}),tw=XLSX.read(tb,{type:'array'});
     const jp=XLSX.utils.sheet_to_json(jw.Sheets[jw.SheetNames[0]],{range:1,defval:'',raw:false});
     const tp=XLSX.utils.sheet_to_json(tw.Sheets[tw.SheetNames[0]],{range:3,defval:'',raw:false});
     let pm=mappingRows;if(mb){try{const mw=XLSX.read(mb,{type:'array'});const parsed=XLSX.utils.sheet_to_json(mw.Sheets[mw.SheetNames[0]],{range:0,defval:'',raw:false});if(!parsed.length||!Object.prototype.hasOwnProperty.call(parsed[0]||{},'Item Code'))throw Error('Product Mapping ساختار معتبر ندارد.');pm=parsed;}catch(mappingError){console.warn('Product Mapping جدید نامعتبر است؛ نسخه قبلی حفظ می‌شود.',mappingError);pm=mappingRows;}}
     validate(jp,tp,pm);if(jp.length<100||tp.length<100)throw Error('تعداد ردیف‌ها غیرعادی کم است.');
     accepted={jf,tf,jp,tp,pm};break;
    }catch(e){lastError=e;}
   }
   if(accepted)break;
  }
  if(!accepted)throw Error('هیچ ترکیب معتبر از فایل‌های Drive پیدا نشد. '+(lastError?.message||''));
  const old=[jpRows,tpaRows,mappingRows,window.sourceYear];
  try{jpRows=accepted.jp;tpaRows=accepted.tp;mappingRows=accepted.pm;window.sourceYear=Number((accepted.tf.name.match(/(14\d{2})/)||[])[1]||1405);buildData();await saveLocal();}
  catch(e){[jpRows,tpaRows,mappingRows,window.sourceYear]=old;buildData();throw e;}
  renderAll();$x('lastSync').textContent=`آخرین همگام‌سازی: ${new Date().toLocaleString('fa-IR')} · ${accepted.jf.name} · ${accepted.tf.name}`;message(`همگام‌سازی موفق؛ ${stores.length.toLocaleString('fa-IR')} فروشگاه بارگذاری شد.`);status('همگام');
 }catch(e){console.error(e);message('همگام‌سازی انجام نشد؛ اطلاعات معتبر قبلی حفظ شد. '+e.message);}
}
async function saveLocal(){const d=await db();return new Promise((resolve,reject)=>{const tx=d.transaction(STORE,'readwrite'),os=tx.objectStore(STORE),now=Date.now();const get=os.get('source');get.onsuccess=()=>{if(get.result)os.put({...get.result,id:'backup:'+now,backupAt:now});os.put({id:'source',jpRows,tpaRows,mappingRows,year:window.sourceYear||1405,updated:now});const all=os.getAll();all.onsuccess=()=>{const backups=all.result.filter(x=>String(x.id).startsWith('backup:')).sort((a,b)=>b.backupAt-a.backupAt);for(const old of backups.slice(5))os.delete(old.id);};};tx.oncomplete=()=>{d.close();resolve();};tx.onerror=()=>{d.close();reject(tx.error);};});}
async function listBackupsX(){const d=await db();const all=await new Promise((res,rej)=>{const q=d.transaction(STORE).objectStore(STORE).getAll();q.onsuccess=()=>res(q.result||[]);q.onerror=()=>rej(q.error);});d.close();return all.filter(x=>String(x.id).startsWith('backup:')).sort((a,b)=>b.backupAt-a.backupAt);}
async function restoreBackupX(id){if(!confirm('نسخه پشتیبان انتخاب‌شده جایگزین داده فعلی شود؟ ابتدا از داده فعلی نسخه پشتیبان ساخته می‌شود.'))return;const d=await db();const data=await new Promise((res,rej)=>{const q=d.transaction(STORE).objectStore(STORE).get(id);q.onsuccess=()=>res(q.result);q.onerror=()=>rej(q.error);});d.close();if(!data)return message('نسخه پشتیبان پیدا نشد.');const old=[jpRows,tpaRows,mappingRows,window.sourceYear];try{jpRows=data.jpRows;tpaRows=data.tpaRows;mappingRows=data.mappingRows;window.sourceYear=data.year;await saveLocal();buildData();renderAll();message('نسخه پشتیبان بازیابی شد.');}catch(e){[jpRows,tpaRows,mappingRows,window.sourceYear]=old;buildData();message('بازیابی ناموفق بود: '+e.message);}}
async function uploadJsonDriveX(name,text,existingId){const boundary='-------TPAJsonBoundary'+Date.now();const metadata={name,mimeType:'application/json',...(existingId?{}:{parents:[DRIVE_FOLDER]})};const body=new Blob([`--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(metadata)}\r\n--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n`,text,`\r\n--${boundary}--`]);const url=existingId?`https://www.googleapis.com/upload/drive/v3/files/${existingId}?uploadType=multipart`:'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart';const res=await fetch(url,{method:existingId?'PATCH':'POST',headers:{Authorization:'Bearer '+accessToken,'Content-Type':'multipart/related; boundary='+boundary},body});if(!res.ok)throw Error('ذخیره پشتیبان Drive ناموفق بود: '+res.status);return res.json();}
async function backupToDriveX(){if(!accessToken){message('ابتدا به Google Drive متصل شو.');return;}try{const payload={schemaVersion:1,createdAt:new Date().toISOString(),sourceYear:window.sourceYear||1405,jpRows,tpaRows,mappingRows,plan,presentedHistory:presentedHistoryX};const files=await driveFiles(),old=files.find(f=>f.name==='TPA_Navigation_Backup.json');if(old){const previous=await downloadDrive(old);const archive=`TPA_Navigation_Backup_Archive_${new Date().toISOString().replace(/[:.]/g,'-')}.json`;await uploadJsonDriveX(archive,new TextDecoder().decode(previous),null);await uploadJsonDriveX('TPA_Navigation_Backup.json',JSON.stringify(payload),old.id);}else await uploadJsonDriveX('TPA_Navigation_Backup.json',JSON.stringify(payload),null);message('نسخه پشتیبان داده‌ها و برنامه در Google Drive ذخیره شد.');}catch(e){message('پشتیبان‌گیری Drive ناموفق بود: '+e.message);}}
async function restoreBackupFromDriveX(){if(!accessToken){message('ابتدا به Google Drive متصل شو.');return;}if(!confirm('داده‌های محلی با آخرین نسخه پشتیبان Drive جایگزین شود؟ ابتدا نسخه فعلی محلی نگهداری می‌شود.'))return;try{const f=(await driveFiles()).find(x=>x.name==='TPA_Navigation_Backup.json');if(!f)throw Error('فایل پشتیبان در Drive پیدا نشد.');const bytes=await downloadDrive(f),payload=JSON.parse(new TextDecoder().decode(bytes));if(payload.schemaVersion!==1||!Array.isArray(payload.jpRows)||!Array.isArray(payload.tpaRows))throw Error('ساختار پشتیبان معتبر نیست.');const old=[jpRows,tpaRows,mappingRows,window.sourceYear,plan,presentedHistoryX];try{jpRows=payload.jpRows;tpaRows=payload.tpaRows;mappingRows=payload.mappingRows||[];window.sourceYear=payload.sourceYear||1405;plan=Array.isArray(payload.plan)?payload.plan:[];presentedHistoryX=Array.isArray(payload.presentedHistory)?payload.presentedHistory:[];validate(jpRows,tpaRows,mappingRows);buildData();await saveLocal();await savePresentedHistoryX();localStorage.setItem(PLAN_KEY,JSON.stringify(plan));renderAll();message('نسخه پشتیبان Drive بازیابی شد.');}catch(e){[jpRows,tpaRows,mappingRows,window.sourceYear,plan,presentedHistoryX]=old;buildData();throw e;}}catch(e){message('بازیابی از Drive ناموفق بود: '+e.message);}}
function addBackupControlsX(){const settings=$x('settingsPage');if(!settings||$x('backupPanel'))return;const panel=document.createElement('section');panel.className='panel';panel.id='backupPanel';panel.innerHTML='<h2>پشتیبان‌گیری و بازیابی</h2><p>قبل از هر جایگزینی داده، نسخه قبلی به‌صورت محلی نگهداری می‌شود. پنج نسخه اخیر قابل بازیابی است. نسخه Drive شامل فایل‌های منبع، برنامه Auditor و تاریخچه ارائه است.</p><button id="backupDriveBtn" class="primary-button full-button">پشتیبان‌گیری در Google Drive</button><button id="restoreDriveBtn" class="secondary-button full-button">بازیابی از Google Drive</button><button id="refreshBackups" class="secondary-button full-button">نمایش نسخه‌های پشتیبان محلی</button><div id="backupList"></div>';
 settings.appendChild(panel);$x('backupDriveBtn').onclick=backupToDriveX;$x('restoreDriveBtn').onclick=restoreBackupFromDriveX;$x('refreshBackups').onclick=async()=>{const items=await listBackupsX();$x('backupList').innerHTML=items.map(x=>`<div class="backup-row"><span>${new Date(x.backupAt).toLocaleString('fa-IR')}</span><button type="button" data-restore-backup="${escX(x.id)}">بازیابی</button></div>`).join('')||'<p>نسخه پشتیبانی موجود نیست.</p>';$x('backupList').querySelectorAll('[data-restore-backup]').forEach(b=>b.onclick=()=>restoreBackupX(b.dataset.restoreBackup));};}
async function loadPresentedHistoryX(){try{const d=await db();const record=await new Promise((res,rej)=>{const q=d.transaction(STORE).objectStore(STORE).get('presented-history');q.onsuccess=()=>res(q.result);q.onerror=()=>rej(q.error);});d.close();presentedHistoryX=record?.rows||[];}catch(e){console.warn('History restore failed',e);presentedHistoryX=[];}}
async function savePresentedHistoryX(){const d=await db();return new Promise((res,rej)=>{const tx=d.transaction(STORE,'readwrite');tx.objectStore(STORE).put({id:'presented-history',rows:presentedHistoryX,updated:Date.now()});tx.oncomplete=()=>{d.close();res();};tx.onerror=()=>{d.close();rej(tx.error);};});}
function historyKeyX(p){return normX(p.customerId)+'|'+normalizeVisitDateX(p.visitDate);}
async function markPlanPresentedX(){if(!plan.length){message('برنامه خالی است.');return;}const ok=confirm('فقط در صورتی تأیید کن که فایل Audit_Visit_Plan.xlsx را واقعاً برای مدیر ارسال یا در محل مورد توافق بارگذاری کرده‌ای. آیا ارسال واقعی انجام شده است؟');if(!ok)return;const submittedAt=new Date().toISOString();let added=0;for(const p of plan){const key=historyKeyX(p);if(!presentedHistoryX.some(h=>h.key===key)){presentedHistoryX.push({key,visitDate:normalizeVisitDateX(p.visitDate)||p.visitDate,customerId:p.customerId,storeName:p.storeName,customerName:p.customerName,address:p.address,routeName:p.routeName,latitude:p.latitude,longitude:p.longitude,priority:p.priority,selectionReason:p.selectionReason,auditorComment:p.auditorComment,selectionSource:p.selectionSource,submittedAt});added++;}}try{await savePresentedHistoryX();message(`${added} بازدید در سابقه ارائه‌شده ثبت شد.`);renderList();renderPlan();}catch(e){message('ثبت سابقه ناموفق بود؛ وضعیت ارسال در دستگاه ثبت نشد: '+e.message);}}
function exportPresentedHistoryX(){const rows=presentedHistoryX.map(h=>({'Visit Date':h.visitDate,'Customer ID':h.customerId,'Store Name':h.storeName,'Customer Name':h.customerName,'Address':h.address,'Route Name':h.routeName,'Latitude':h.latitude,'Longitude':h.longitude,'Priority':h.priority,'Selection Reason':h.selectionReason,'Auditor Comment':h.auditorComment,'Selection Source':h.selectionSource,'Submitted At':h.submittedAt}));const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),'Audit Visit History');XLSX.writeFile(wb,'Audit_Visit_History.xlsx');}
function bindDriveEnhancedX(){if(!window.google?.accounts?.oauth2)return;tokenClient=google.accounts.oauth2.initTokenClient({client_id:CLIENT_ID,scope:'https://www.googleapis.com/auth/drive.readonly https://www.googleapis.com/auth/drive.file',callback:async r=>{if(r.error){console.error('Google OAuth error',r);message('ورود Google ناموفق بود: '+r.error+(r.error_description?' — '+r.error_description:''));return;}if(!r.access_token){message('Google مجوز دسترسی برنگرداند. دوباره تلاش کن.');return;}accessToken=r.access_token;await syncDrive();}});}
function connectDriveEnhancedX(){tokenClient=null;bindDriveEnhancedX();if(!tokenClient){message('کتابخانه Google بارگذاری نشده است.');return;}tokenClient.requestAccessToken({prompt:'consent'});}
async function uploadDriveFileX(name,bytes,existingId){const boundary='-------TPAFormBoundary'+Math.random().toString(16).slice(2);const metadata={name,mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',...(existingId?{}:{parents:[DRIVE_FOLDER]})};const body=new Blob([`--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(metadata)}\r\n--${boundary}\r\nContent-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n`,bytes,`\r\n--${boundary}--`]);const url=existingId?`https://www.googleapis.com/upload/drive/v3/files/${existingId}?uploadType=multipart`:'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart';const method=existingId?'PATCH':'POST';const res=await fetch(url,{method,headers:{Authorization:'Bearer '+accessToken,'Content-Type':'multipart/related; boundary='+boundary},body});if(!res.ok)throw Error('ذخیره فایل در Drive ناموفق بود: '+res.status+' '+(await res.text()).slice(0,250));return res.json();}
async function syncHistoryToDriveX(){if(!accessToken){message('ابتدا به Google Drive متصل شو تا مجوز ذخیره تاریخچه دریافت شود.');return;}try{const rows=presentedHistoryX.map(h=>({'Visit Date':h.visitDate,'Customer ID':h.customerId,'Store Name':h.storeName,'Customer Name':h.customerName,'Address':h.address,'Route Name':h.routeName,'Latitude':h.latitude,'Longitude':h.longitude,'Priority':h.priority,'Selection Reason':h.selectionReason,'Auditor Comment':h.auditorComment,'Selection Source':h.selectionSource,'Submitted At':h.submittedAt}));const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),'Audit Visit History');const bytes=XLSX.write(wb,{bookType:'xlsx',type:'array'});const files=await driveFiles();const old=files.find(f=>f.name==='Audit_Visit_History.xlsx');if(old){const archiveName=`Audit_Visit_History_Archive_${new Date().toISOString().replace(/[:.]/g,'-')}.xlsx`;const oldBytes=await downloadDrive(old);await uploadDriveFileX(archiveName,oldBytes,null);await uploadDriveFileX('Audit_Visit_History.xlsx',bytes,old.id);}else await uploadDriveFileX('Audit_Visit_History.xlsx',bytes,null);message('تاریخچه در Google Drive ذخیره شد و نسخه قبلی، در صورت وجود، آرشیو شد.');}catch(e){message('تاریخچه در Drive ذخیره نشد؛ نسخه محلی حفظ شده است. '+e.message);}}
function addHistoryControlsX(){const mapPage=$x('mapPage');const filterPanel=mapPage?.querySelector('.filter-panel .filter-grid');if(filterPanel&&!$x('presentedFilter')){const label=document.createElement('label');label.textContent='سابقه ارائه به مدیر';label.innerHTML='سابقه ارائه به مدیر<select id="presentedFilter"><option value="all">همه</option><option value="new">هرگز ارائه نشده</option><option value="presented">قبلاً ارائه شده</option></select>';filterPanel.appendChild(label);$x('presentedFilter').onchange=()=>{renderList();drawMap(applyEnhancedFiltersX());renderStats();};}
 const planPage=$x('planPage');if(planPage&&!$x('markPlanPresented')){const box=document.createElement('section');box.className='panel';box.innerHTML='<h2>ارائه برنامه به مدیر</h2><p>پس از ارسال واقعی فایل خروجی برای مدیر، دکمه ثبت را بزن. پیش‌نویس یا خروجی‌گرفتن به‌تنهایی سابقه ارائه ایجاد نمی‌کند.</p><button id="markPlanPresented" class="primary-button full-button">ثبت اینکه برنامه واقعاً ارائه شده است</button><button id="exportHistoryBtn" class="secondary-button full-button">خروجی تاریخچه ارائه‌ها</button><button id="saveHistoryDriveBtn" class="secondary-button full-button">ذخیره تاریخچه در Google Drive</button><div id="historyCount" class="message"></div>';planPage.appendChild(box);$x('markPlanPresented').onclick=markPlanPresentedX;$x('exportHistoryBtn').onclick=exportPresentedHistoryX;$x('saveHistoryDriveBtn').onclick=syncHistoryToDriveX;}}
function renderHistoryCountX(){if($x('historyCount'))$x('historyCount').textContent=`تعداد بازدیدهای ثبت‌شده در تاریخچه ارائه: ${presentedHistoryX.length.toLocaleString('fa-IR')}`;}
function getDataQualityX(){
 const q=window.dataQualityX||{};
 const mappingCounts=new Map();
 for(const r of mappingRows){const c=normX(r['Item Code']);if(c)mappingCounts.set(c,(mappingCounts.get(c)||0)+1);}
 const duplicateMappingCodes=[...mappingCounts].filter(([k,n])=>n>1).map(([k])=>k);
 const missingItems=(q.unmappedCodes||[]).slice();
 const jpCodes=new Set();
 for(const r of jpRows)for(const v of [r.CustomerI,r['New Customer Code'],r['Oracle Code'],r['Account Number']]){
  const n=normX(v);if(n){jpCodes.add(n);n.split('|').forEach(x=>jpCodes.add(x));}
 }
 const unmatched=[...new Set((tpaRows||[]).filter(r=>txt(r.Status).toLowerCase()!=='closed/expired').map(r=>normX(r['کد مشتری'])).filter(Boolean))]
  .filter(c=>![c,...c.split('|')].some(k=>jpCodes.has(k)));
 return {
  missingItems,duplicateMappingCodes,unmatchedCustomerCodes:unmatched,
  emptyGPS:stores.filter(s=>s.source==='JP'&&!s.gps).length,
  matchedOrderGroups:q.matchedOrderGroups||0,unmatchedOrderGroups:q.unmatchedOrderGroups||0,
  orderGroups:q.orderGroups||0,ambiguousAliasCount:q.ambiguousAliasCount||0,
  expiredRows:q.expiredRows||0,skippedRows:q.skippedRows||0
 };
}
function addQualityPanelX(){const page=$x('analysisPage');if(!page||$x('dataQualityPanel'))return;const panel=document.createElement('section');panel.className='panel';panel.id='dataQualityPanel';panel.innerHTML='<h2>کنترل کیفیت داده‌ها</h2><div id="dataQualityContent"></div>';page.appendChild(panel);}
function renderQualityPanelX(){
 addQualityPanelX();
 const q=getDataQualityX(),root=$x('dataQualityContent');if(!root)return;
 root.innerHTML=`<div class="quality-grid">
  <div><b>${q.missingItems.length.toLocaleString('fa-IR')}</b><small>کد کالا بدون تطبیق در Product Mapping</small></div>
  <div><b>${q.duplicateMappingCodes.length.toLocaleString('fa-IR')}</b><small>کد تکراری در Product Mapping</small></div>
  <div><b>${q.unmatchedOrderGroups.toLocaleString('fa-IR')}</b><small>گروه سفارش بدون تطبیق با JP</small></div>
  <div><b>${q.ambiguousAliasCount.toLocaleString('fa-IR')}</b><small>شناسه جایگزین مبهم</small></div>
  <div><b>${q.emptyGPS.toLocaleString('fa-IR')}</b><small>مشتری JP بدون GPS معتبر</small></div>
  <div><b>${q.expiredRows.toLocaleString('fa-IR')}</b><small>ردیف Closed/expired حذف‌شده</small></div>
 </div>
 <p class="hint">گروه سفارش: ${q.orderGroups.toLocaleString('fa-IR')} · تطبیق‌شده با JP: ${q.matchedOrderGroups.toLocaleString('fa-IR')} · ردیف‌های فاقد شناسه یا شماره سفارش: ${q.skippedRows.toLocaleString('fa-IR')}</p>
 ${q.missingItems.length?`<details><summary>کدهای کالای نامنطبق</summary><p>${q.missingItems.map(escX).join('، ')}</p></details>`:''}
 ${q.duplicateMappingCodes.length?`<details><summary>کدهای تکراری Product Mapping</summary><p>${q.duplicateMappingCodes.map(escX).join('، ')}</p></details>`:''}
 ${q.unmatchedCustomerCodes.length?`<details><summary>کدهای مشتری بدون تطبیق</summary><p>${q.unmatchedCustomerCodes.map(escX).join('، ')}</p></details>`:''}`;
}
function haversineX(a,b){const r=6371000,p=Math.PI/180,dLat=(b.lat-a.lat)*p,dLng=(b.lng-a.lng)*p;const h=Math.sin(dLat/2)**2+Math.cos(a.lat*p)*Math.cos(b.lat*p)*Math.sin(dLng/2)**2;return 2*r*Math.asin(Math.sqrt(h));}
function addRouteSuggestionControlX(){
 const root=$x('planPage');if(!root||$x('suggestRouteOrder'))return;
 const panel=document.createElement('section');panel.className='panel';
 panel.innerHTML='<h2>ترتیب پیشنهادی توقف‌ها</h2><p>در صورت دسترسی به GPS، نقطه شروع از موقعیت فعلی گرفته می‌شود. اگر GPS در دسترس نبود، مختصات شروع را دستی وارد کن. این پیشنهاد بر اساس فاصله مستقیم است، نه مسیر جاده‌ای یا زمان رانندگی.</p><div class="analysis-controls"><label>عرض جغرافیایی شروع<input id="routeStartLat" inputmode="decimal" placeholder="Latitude"></label><label>طول جغرافیایی شروع<input id="routeStartLng" inputmode="decimal" placeholder="Longitude"></label></div><button id="suggestRouteOrder" class="secondary-button full-button">پیشنهاد ترتیب توقف‌ها</button><div id="routeSuggestionInfo" class="message"></div>';
 root.insertBefore(panel,root.querySelectorAll('.panel')[2]||null);
 $x('suggestRouteOrder').onclick=()=>{
  const day=normalizeVisitDateX($x('planDate').value)||localTodayKeyX();
  const items=plan.filter(p=>normalizeVisitDateX(p.visitDate)===day&&validCoordX(Number(p.latitude),Number(p.longitude)));
  if(!items.length){$x('routeSuggestionInfo').textContent='برای این تاریخ توقف GPSدار ثبت نشده است.';return;}
  const manual={lat:Number($x('routeStartLat').value),lng:Number($x('routeStartLng').value)};
  const useStart=start=>{if(!validCoordX(start.lat,start.lng)){$x('routeSuggestionInfo').textContent='مختصات شروع دستی معتبر نیست.';return;}let current={lat:start.lat,lng:start.lng};const remaining=items.slice(),ordered=[];while(remaining.length){remaining.sort((a,b)=>haversineX(current,{lat:Number(a.latitude),lng:Number(a.longitude)})-haversineX(current,{lat:Number(b.latitude),lng:Number(b.longitude)}));const next=remaining.shift();ordered.push(next);current={lat:Number(next.latitude),lng:Number(next.longitude)};}const other=plan.filter(p=>normalizeVisitDateX(p.visitDate)!==day);plan=[...other,...ordered];savePlan();let total=0,prev={lat:start.lat,lng:start.lng};for(const p of ordered){total+=haversineX(prev,{lat:Number(p.latitude),lng:Number(p.longitude)});prev={lat:Number(p.latitude),lng:Number(p.longitude)};}$x('routeSuggestionInfo').textContent=`ترتیب پیشنهادی اعمال شد: ${ordered.length} توقف؛ فاصله مستقیم تقریبی ${(total/1000).toFixed(1)} کیلومتر. این مقدار مسافت جاده‌ای نیست.`;};
  if(navigator.geolocation){navigator.geolocation.getCurrentPosition(pos=>useStart({lat:pos.coords.latitude,lng:pos.coords.longitude}),()=>{if(validCoordX(manual.lat,manual.lng))useStart(manual);else $x('routeSuggestionInfo').textContent='GPS دریافت نشد؛ مختصات شروع را دستی وارد کن.';},{enableHighAccuracy:true,timeout:12000});}
  else if(validCoordX(manual.lat,manual.lng))useStart(manual);else $x('routeSuggestionInfo').textContent='موقعیت‌یابی پشتیبانی نمی‌شود؛ مختصات شروع را دستی وارد کن.';
 };
}
const translationsX={
'TPA Navigation':'TPA Navigation','نقشه فروشگاه‌ها و برنامه بازدید Auditor':'Store map and Auditor visit planner','آماده':'Ready','نقشه فروشگاه‌ها':'Store map','فیلتر فروشگاه‌ها':'Store filters','فروشگاه‌ها':'Stores','فیلتر فروشگاه‌ها':'Store filters','پاک‌کردن فیلترها':'Clear filters','فیلتر فعال:':'Active filters:','جستجوی نام، کد یا آدرس فروشگاه':'Search store, customer, code or address','نوع مشتری':'Customer type','همه مشتریان':'All customers','همه':'All','فقط TPA':'TPA only','سابقه سفارش':'Order history','دارای سفارش معتبر':'With valid orders','بدون سفارش':'No orders','وضعیت GPS':'GPS status','GPS موجود':'GPS available','بدون GPS':'No GPS','منطقه (چند انتخاب)':'Area (multi-select)','ویزیتور (چند انتخاب)':'Salesperson (multi-select)','Customer Class (چند انتخاب)':'Customer Class (multi-select)','روز بازدید (چند انتخاب)':'Visit day (multi-select)','روز بازدید':'Visit day','برنامه بازدید من':'My visit plan','برنامه مستقل Auditor؛ فروشگاه‌های دارای GPS را انتخاب کن، برنامه را ویرایش کن و خروجی Excel بگیر.':'Independent Auditor plan. Select GPS-enabled stores, edit your plan and export to Excel.','افزودن بازدید':'Add visit','فروشگاه‌های JP':'JP stores','تاریخ بازدید':'Visit date','جستجوی فروشگاه':'Search stores','نام یا کد مشتری':'Store name or customer code','فهرست برنامه':'Visit plan','وضعیت هر بازدید مستقل است':'Each visit has an independent status','خروجی Excel برنامه بازدید':'Export visit plan to Excel','دریافت برنامه پیشنهادی مدیر':'Import manager suggestions','بررسی فایل و نمایش پیشنهادها':'Review file and show suggestions','تحلیل و بررسی':'Insights','نمای سریع از کیفیت داده و وضعیت سفارش‌ها.':'Overview of data quality and order status.','مشتری بدون GPS':'Stores without GPS','سفارش بدون فاکتور':'Uninvoiced orders','بدون سابقه سفارش':'No purchase history','فهرست‌های تحلیلی':'Analysis lists','فروشگاه‌های نزدیک من':'Stores near me','سفارش‌های بدون فاکتور':'Uninvoiced orders','تنظیمات و همگام‌سازی':'Settings and sync','دریافت نسخه‌های جدید فایل‌ها از پوشه Google Drive یا بارگذاری دستی.':'Get the latest files from Google Drive or upload them manually.','اتصال به Google Drive':'Connect to Google Drive','اتصال و دریافت آخرین فایل‌ها':'Connect and sync latest files','بارگذاری دستی فایل‌ها':'Manual file upload','Journey Plan':'Journey Plan','گزارش فروش TPA':'TPA sales report','پردازش فایل‌های انتخاب‌شده':'Process selected files','خلاصه داده‌های بارگذاری‌شده':'Loaded data summary','کل مشتریان':'Total stores','دارای سفارش':'With orders','JP بدون سفارش':'JP without orders','آخرین همگام‌سازی: ثبت نشده':'Last sync: Not recorded','تنظیمات دستگاه':'Device settings','تغییر حالت روشن / تاریک':'Toggle light / dark mode','English / فارسی':'English / فارسی','پاک‌کردن داده‌های محلی این دستگاه':'Clear local data on this device','جزئیات فروشگاه':'Store details','در انتظار':'Pending','انجام شد':'Completed','انجام نشد':'Not completed','افزودن به برنامه Auditor':'Add to Auditor plan','مجموع خرید فاکتورشده':'Confirmed invoiced purchases','مبلغ سفارش‌های فاکتورنشده':'Uninvoiced order value','تاریخچه سفارش‌ها':'Order history','تحلیل سبد خرید':'Basket analysis','فعال':'Active','غیرفعال':'Inactive','صادر نشده':'Not invoiced','بدون مسیر':'No route','ثبت نشده':'Not recorded','ندارد':'None','در JP موجود نیست':'Not available in JP','مشتری بدون GPS قابل افزودن نیست':'Stores without GPS cannot be added','موردی پیدا نشد.':'No results found.','هنوز بازدیدی ثبت نشده است.':'No visits have been added yet.','سفارشی ثبت نشده است.':'No orders found.','موردی نیست':'None','تاریخ سفارش:':'Order date:','تحویل هدف:':'Target delivery:','تاریخ فاکتور:':'Invoice date:','مبلغ سفارش:':'Order amount:','تعداد:':'Quantity:','مشاهده محصولات':'View products','کد کالا:':'Item code:','دفعات خرید:':'Purchase frequency:','تعداد کل:':'Total quantity:','اولین خرید:':'First purchase:','آخرین خرید:':'Last purchase:','کد مشتری:':'Customer ID:','نام مشتری:':'Customer name:','فروشنده:':'Salesperson:','روزهای JP:':'JP visit days:','منبع:':'Source:','فقط TPA':'TPA only','فاکتور شده':'Invoiced','بسته / تحویل‌نشده':'Closed / undelivered','تأیید شده / تخصیص نیافته':'Confirmed / not allocated','تأیید و ارسال شده':'Confirmed and shipped','بسته / منقضی':'Closed / expired','Google Maps':'Google Maps','نشان':'Neshan','پیشنهاد ترتیب توقف‌ها':'Suggested stop order','کنترل کیفیت داده‌ها':'Data quality checks','کد کالا بدون تطبیق در Product Mapping':'Item codes not mapped in Product Mapping','کد مشتری TPA بدون تطبیق مستقیم با CustomerI':'TPA customer codes not directly matched to CustomerI','مشتری JP بدون GPS معتبر':'JP stores without valid GPS','ارائه برنامه به مدیر':'Submit plan to manager','خروجی تاریخچه ارائه‌ها':'Export submission history','ذخیره تاریخچه در Google Drive':'Save history to Google Drive','نمایش نسخه‌های پشتیبان':'Show backups','بازیابی':'Restore','همه گروه‌ها':'All groups','فقط JP':'JP only','همه گروه‌ها':'All groups','بدون خرید':'No purchase','تأییدشده / ارسال‌شده':'Confirmed / shipped','اجرای تحلیل':'Run analysis','فیلتر سابقه ارائه به مدیر':'Manager submission history','همه روزها':'All days','امروز':'Today','هرگز ارائه نشده':'Never submitted','قبلاً ارائه شده':'Previously submitted','● داده‌های محلی':'● Local data','JP با سابقه سفارش':'JP with order history','JP بدون سابقه سفارش':'JP without order history','فقط TPA':'TPA only','شنبه':'Saturday','یکشنبه':'Sunday','دوشنبه':'Monday','سه‌شنبه':'Tuesday','چهارشنبه':'Wednesday','پنجشنبه':'Thursday','Today':'Today','Saturday':'Saturday','Sunday':'Sunday','Monday':'Monday','Tuesday':'Tuesday','Wednesday':'Wednesday','Thursday':'Thursday','High':'High','Medium':'Medium','Low':'Low','زیاد':'High','متوسط':'Medium','کم':'Low','عادی':'Normal','اولویت':'Priority','یادداشت Auditor':'Auditor comment','حذف':'Delete','افزودن به برنامه Auditor':'Add to Auditor plan','در حال دریافت آخرین فایل‌های Drive...':'Downloading latest Drive files...','بارگذاری انجام شد.':'Files loaded successfully.','اتصال و دریافت آخرین فایل‌ها':'Connect and sync latest files','فایل Journey Plan و TPA را هر دو انتخاب کن.':'Select both the Journey Plan and TPA files.','برای این تاریخ توقف GPSدار ثبت نشده است.':'No GPS-enabled stops are planned for this date.','فروشگاه‌های نزدیک':'Nearby stores','مشتریان بدون GPS':'Stores without GPS','سفارش‌های بدون فاکتور':'Uninvoiced orders','سابقه ارائه به مدیر':'Manager submission history','فیلتر فعال':'Active filters','مسیر':'Route','روزهای ویزیت':'Visit days','کلاس مشتری':'Customer class','تعداد سفارش':'Order count','مجموع مبلغ':'Total amount','تعداد فروشگاه‌های منطبق:':'Matching stores:','گروه سفارش':'Order group','نوع مشتری':'Customer type','از تاریخ (شمسی)':'From date (Jalali)','تا تاریخ (شمسی)':'To date (Jalali)','راهنما':'Instructions','در JP':'In JP','دارای سفارش':'With orders','بدون سفارش':'No orders','فروشگاه‌های نزدیک — فاصله مستقیم':'Nearby stores — straight-line distance','در حال بررسی نسخه‌های Google Drive...':'Checking Google Drive versions...','فایل JP_ یا TPA_Alborz در پوشه Drive پیدا نشد.':'JP_ or TPA_Alborz file was not found in the Drive folder.','تعداد ردیف‌ها غیرعادی کم است.':'The row count is unexpectedly low.','همگام‌سازی موفق؛':'Sync successful;','همگام‌سازی انجام نشد؛ اطلاعات معتبر قبلی حفظ شد.':'Sync failed; previous valid data was preserved.','ابتدا فایل Excel را انتخاب کن.':'Select an Excel file first.','کد مشتری در داده‌های فعلی پیدا نشد':'Customer ID was not found in the current data','تاریخ معتبر نیست':'Invalid date','ردیف تکراری است':'Duplicate row','مشتری GPS معتبر ندارد':'Customer has no valid GPS','این فروشگاه برای این تاریخ قبلاً در برنامه است.':'This store is already in the plan for this date.','فروشگاه فاقد GPS معتبر است و به برنامه اضافه نمی‌شود.':'This store has no valid GPS and cannot be added to the plan.','برای این تاریخ توقف GPSدار ثبت نشده است.':'No GPS-enabled stops are in the plan for this date.','دریافت موقعیت ناموفق بود:':'Could not get your location:','فاصله مستقیم تقریبی':'Approximate straight-line distance','این مقدار مسافت جاده‌ای نیست.':'This is not road distance.','در حال پردازش فایل‌ها...':'Processing files...','برای شروع فایل‌ها را دستی بارگذاری کن یا به Google Drive متصل شو.':'Upload the files manually or connect to Google Drive to get started.'
};
function translateTreeX(root,lang){if(!root)return;const map=translationsX;const reverse=new Map(Object.entries(map).map(([fa,en])=>[en,fa]));const walker=document.createTreeWalker(root,NodeFilter.SHOW_TEXT);let node;while((node=walker.nextNode())){const raw=node.nodeValue,trim=raw.trim();if(!trim)continue;let replacement=lang==='en'?map[trim]:reverse.get(trim);if(!replacement&&lang==='en'){replacement=trim;for(const [fa,en] of Object.entries(map))if(replacement.includes(fa))replacement=replacement.replaceAll(fa,en);}else if(!replacement&&lang==='fa'){replacement=trim;for(const [en,fa] of reverse)if(replacement.includes(en))replacement=replacement.replaceAll(en,fa);}if(replacement!==undefined)node.nodeValue=raw.replace(trim,replacement);}}
function setLanguageEnhancedX(lang){const use=lang==='en'?'en':'fa';document.documentElement.lang=use;document.documentElement.dir=use==='fa'?'rtl':'ltr';localStorage.setItem('tpa-language',use);translateTreeX(document.body,use);if($x('search'))$x('search').placeholder=use==='en'?'Search store, customer, code or address':'جستجوی نام، کد یا آدرس فروشگاه';if($x('planSearch'))$x('planSearch').placeholder=use==='en'?'Store name or customer code':'نام یا کد مشتری';if($x('planDate'))$x('planDate').placeholder=use==='en'?'YYYY-MM-DD or Jalali date':'مثلاً 1405/08/20';if($x('langBtn'))$x('langBtn').textContent=use==='en'?'فارسی':'English';}
function initEnhancementsX(){
 if(typeof buildData==='function'){};addAnalysisControlsX();addQualityPanelX();addRouteSuggestionControlX();addHistoryControlsX();addBackupControlsX();bindDriveEnhancedX();if($x('driveBtn'))$x('driveBtn').onclick=connectDriveEnhancedX;
 const oldMessage=window.message;window.addEventListener('online',()=>status('آنلاین'));window.addEventListener('offline',()=>status('آفلاین'));
 bindEnhancedX();if($x('langBtn'))$x('langBtn').onclick=()=>setLanguageEnhancedX(document.documentElement.lang==='fa'?'en':'fa');setLanguageEnhancedX(localStorage.getItem('tpa-language')||'fa');
 const oldRenderAll=window.renderAll;
 // Rebuild enriched fields after the base loader or IndexedDB restore has populated source rows.
 const watch=()=>{if(Array.isArray(stores)&&stores.length){refreshFinancialModelX();updateEnhancedFilterOptionsX();renderList();drawMap(applyEnhancedFiltersX());renderStats();renderPlan();renderAnalysisSummaryX();renderQualityPanelX();renderHistoryCountX();if(document.documentElement.lang==='en')translateTreeX(document.body,'en');}};
 loadPresentedHistoryX().then(()=>{renderList();renderHistoryCountX();});if(jpRows.length||tpaRows.length)watch();renderQualityPanelX();
 const loadButton=$x('loadBtn');if(loadButton)loadButton.onclick=async()=>{await loadManualEnhancedX();watch();};
 const driveButton=$x('driveBtn');if(driveButton){const original=driveButton.onclick;driveButton.onclick=async function(...args){if(original)await original.apply(this,args);setTimeout(watch,1500);};}
}
// Publish overrides into the classic-script global environment used by app.js.
Object.assign(window,{buildData,renderAll,renderList,renderStats,showPage,showDetails,loadManual:loadManualEnhancedX,renderPlan,addToPlan,renderPlanSearch,exportPlan,importPlan,renderExternal,clearFilters,renderAnalysis,activeCount,syncDrive,latest,saveLocal,message:messageEnhancedX,setLanguage:setLanguageEnhancedX,bindDrive:bindDriveEnhancedX,connectDrive:connectDriveEnhancedX});
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',initEnhancementsX,{once:true});else initEnhancementsX();
})();


/* Fixed viewport workspace, on-demand panels, live location and Drive diagnostics v2 */
(function(){
'use strict';
const byId=id=>document.getElementById(id);
let activePanel=null, activePlaceholder=null, userMarker=null, userAccuracy=null;
function ensureWorkspace(){
 const modal=byId('workspaceModal'); if(!modal||byId('workspaceInitialized'))return;
 const mark=document.createElement('i');mark.id='workspaceInitialized';mark.hidden=true;document.body.appendChild(mark);
 const content=byId('workspaceContent');
 function close(){
  if(activePanel&&activePlaceholder&&activePlaceholder.parentNode){activePlaceholder.parentNode.insertBefore(activePanel,activePlaceholder);activePlaceholder.remove();}
  activePanel=null;activePlaceholder=null;content.innerHTML='';modal.classList.remove('open');modal.setAttribute('aria-hidden','true');document.body.classList.remove('workspace-open');
 }
 function open(panel,title){if(!panel)return;close();activePanel=panel;activePlaceholder=document.createComment('workspace-panel-origin');panel.parentNode.insertBefore(activePlaceholder,panel);content.appendChild(panel);byId('workspaceTitle').textContent=title||panel.querySelector('h2')?.textContent||'بخش';modal.classList.add('open');modal.setAttribute('aria-hidden','false');document.body.classList.add('workspace-open');content.scrollTop=0;}
 byId('closeWorkspace').onclick=close;modal.addEventListener('click',e=>{if(e.target===modal)close();});document.addEventListener('keydown',e=>{if(e.key==='Escape'&&modal.classList.contains('open'))close();});
 function addLaunchers(pageId,panels,host){const page=byId(pageId);if(!page)return;const area=document.createElement('div');area.className='workspace-launchers';area.setAttribute('aria-label','بخش‌های قابل نمایش');panels.forEach(panel=>{if(!panel)return;const heading=panel.querySelector('h2')?.textContent?.trim()||'نمایش بخش';const b=document.createElement('button');b.type='button';b.className='workspace-launcher';b.textContent=heading;b.onclick=()=>open(panel,heading);area.appendChild(b);panel.classList.add('workspace-source-panel');});(host||page).appendChild(area);}
 const map=byId('mapPage');
 if(map){const filter=map.querySelector('.filter-panel'), stores=map.querySelector('.store-panel');
  if(filter)filter.classList.add('workspace-source-panel');if(stores)stores.classList.add('workspace-source-panel');
  byId('openFiltersBtn').onclick=()=>open(filter,'فیلتر فروشگاه‌ها');byId('openStoresBtn').onclick=()=>open(stores,'فهرست فروشگاه‌ها');
 }
 addLaunchers('planPage',[...byId('planPage').children].filter(x=>x.classList.contains('panel')&&!x.classList.contains('page-intro')));
 addLaunchers('analysisPage',[...byId('analysisPage').children].filter(x=>!x.classList.contains('page-intro')));
 addLaunchers('settingsPage',[...byId('settingsPage').children].filter(x=>!x.classList.contains('page-intro')));
}
function locateUser(){
 if(!navigator.geolocation){message('موقعیت‌یابی در این مرورگر پشتیبانی نمی‌شود.');return;}
 const btn=byId('locateMeBtn');if(btn){btn.disabled=true;btn.textContent='در حال یافتن...';}
 navigator.geolocation.getCurrentPosition(pos=>{
  const lat=pos.coords.latitude,lng=pos.coords.longitude,acc=pos.coords.accuracy||0;
  if(!map||!window.L){message('نقشه هنوز آماده نیست؛ دوباره تلاش کن.');return;}
  if(userMarker)map.removeLayer(userMarker);if(userAccuracy)map.removeLayer(userAccuracy);
  userMarker=L.circleMarker([lat,lng],{radius:9,color:'#fff',weight:3,fillColor:'#1769e0',fillOpacity:1}).addTo(map).bindTooltip('موقعیت فعلی من',{permanent:false});
  userAccuracy=L.circle([lat,lng],{radius:Math.min(acc,1000),color:'#1769e0',weight:1,fillColor:'#1769e0',fillOpacity:.10}).addTo(map);
  map.setView([lat,lng],Math.max(map.getZoom(),15),{animate:true});
  message(`موقعیت فعلی نمایش داده شد · دقت تقریبی ${Math.round(acc)} متر`);
 },err=>{
  const reasons={1:'اجازه دسترسی به موقعیت مکانی داده نشد. در تنظیمات Safari، Location را برای این سایت مجاز کن.',2:'موقعیت مکانی فعلاً در دسترس نیست. GPS و خدمات مکان‌یابی دستگاه را بررسی کن.',3:'دریافت موقعیت بیش از حد طول کشید؛ دوباره تلاش کن.'};
  message(reasons[err.code]||('دریافت موقعیت ناموفق بود: '+err.message));
 },{enableHighAccuracy:true,timeout:20000,maximumAge:0});
 if(btn){const restore=()=>{btn.disabled=false;btn.textContent='موقعیت من';};setTimeout(restore,21000);}
}
function installFixedLayout(){
 ensureWorkspace();
 const locate=byId('locateMeBtn');if(locate)locate.onclick=locateUser;
 document.querySelectorAll('.bottom-nav [data-page]').forEach(b=>b.addEventListener('click',()=>{const id=b.dataset.page;document.body.dataset.activePage=id;setTimeout(()=>{if(id==='mapPage'&&window.map)map.invalidateSize();},180);}));
 document.body.dataset.activePage='mapPage';
}
// Detailed Drive API errors distinguish OAuth, folder access, API enablement and file permissions.
async function readDriveError(res,context){let detail='';try{const j=await res.clone().json();detail=j?.error?.message||j?.error_description||'';if(j?.error?.errors?.[0]?.reason)detail+=' ('+j.error.errors[0].reason+')';}catch{try{detail=(await res.clone().text()).slice(0,300);}catch{}}
 const statusText={400:'درخواست Drive نامعتبر است',401:'مجوز ورود منقضی یا نامعتبر است؛ دوباره متصل شو',403:'دسترسی رد شد؛ مجوز Drive API یا دسترسی پوشه را بررسی کن',404:'پوشه یا فایل پیدا نشد یا برای این حساب قابل مشاهده نیست',429:'محدودیت موقت درخواست‌های Google Drive'}[res.status]||('خطای HTTP '+res.status);
 return `${context}: ${statusText}${detail?' — '+detail:''}`;
}
window.driveFiles=async function(){
 if(!window.accessToken&&!accessToken)throw Error('ابتدا به Google Drive متصل شو.');
 const q=encodeURIComponent(`'${DRIVE_FOLDER}' in parents and trashed=false`);
 const url=`https://www.googleapis.com/drive/v3/files?q=${q}&fields=files(id,name,mimeType,modifiedTime)&pageSize=100&orderBy=modifiedTime%20desc&supportsAllDrives=true&includeItemsFromAllDrives=true`;
 const r=await fetch(url,{headers:{Authorization:'Bearer '+accessToken}});if(!r.ok)throw Error(await readDriveError(r,'دریافت فهرست فایل‌های پوشه Google Drive'));
 const data=await r.json();return data.files||[];
};
window.downloadDrive=async function(f){
 const r=await fetch(`https://www.googleapis.com/drive/v3/files/${encodeURIComponent(f.id)}?alt=media&supportsAllDrives=true`,{headers:{Authorization:'Bearer '+accessToken}});
 if(!r.ok)throw Error(await readDriveError(r,`دریافت فایل «${f.name}»`));
 const bytes=await r.arrayBuffer();if(!bytes.byteLength)throw Error(`فایل «${f.name}» خالی دریافت شد.`);return bytes;
};
// Replace the token request handler while preserving the existing OAuth client and scope.
window.connectDriveEnhancedX=window.connectDriveEnhancedX||undefined;
const driveButton=byId('driveBtn');
if(driveButton){driveButton.addEventListener('click',()=>{setTimeout(()=>{const t=byId('message')?.textContent||'';if(/کتابخانه Google بارگذاری نشده/.test(t))message('کتابخانه ورود Google بارگذاری نشد. اتصال اینترنت و مسدود نبودن accounts.google.com را بررسی کن؛ سپس صفحه را دوباره باز کن.');},1200);});}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',installFixedLayout,{once:true});else installFixedLayout();
})();
