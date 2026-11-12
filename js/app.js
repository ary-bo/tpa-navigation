
"use strict";

const GOOGLE_CLIENT_ID =
"1098731892951-2e7jipvf33evn0a69tt0tje94oor54lm.apps.googleusercontent.com";

const DRIVE_FOLDER_ID =
"1sgTsZu2zzSBi-pz2M9pVBHtjCtQELYhk";

let jpRows=[];
let salesRows=[];
let productRows=[];
let stores=[];
let map=null;
let markers=[];
let driveTokenClient=null;
let lastAccessToken=null;

const $=id=>document.getElementById(id);

// اتصال دکمه Drive از ابتدای اجرای برنامه
const driveButtonElement=$("driveBtn");
if(driveButtonElement){
 driveButtonElement.addEventListener("click",()=>{
  setMessage("درخواست اتصال به Google Drive ثبت شد...","در حال اتصال");
  syncFromDrive();
 });
}

window.addEventListener("error",event=>{
 const box=$("message");
 if(box){
  box.textContent="خطای اجرایی برنامه: "+
   (event.message||"خطای نامشخص")+
   ". لطفاً صفحه را Refresh کن و در صورت تکرار، تصویر خطا را بفرست.";
 }
});

function clean(v){
 return String(v??"").trim();
}

function norm(v){
 return clean(v)
  .replace(/\s/g,"")
  .replace(/[۰-۹]/g,d=>"۰۱۲۳۴۵۶۷۸۹".indexOf(d))
  .replace(/[٠-٩]/g,d=>"٠١٢٣٤٥٦٧٨٩".indexOf(d));
}

function keyVariants(v){
 const s=norm(v);
 if(!s)return[];
 const result=[s];
 if(s.includes("|"))result.push(s.split("|").pop());
 return [...new Set(result)];
}

function first(row,keys){
 for(const key of keys){
  if(row[key]!==undefined&&clean(row[key])!==""){
   return clean(row[key]);
  }
 }
 return "";
}

function numeric(v){
 const n=Number(String(v??"").replace(/,/g,""));
 return Number.isFinite(n)?n:0;
}

function validCoord(lat,lng){
 return lat>10&&lat<60&&lng>20&&lng<70;
}

function escapeHtml(value){
 return String(value??"").replace(/[&<>"']/g,c=>({
  "&":"&amp;",
  "<":"&lt;",
  ">":"&gt;",
  '"':"&quot;",
  "'":"&#39;"
 }[c]));
}

function setMessage(text,status){
 $("message").textContent=text;
 if(status)$("status").textContent=status;
}

function readExcel(file,headerRow){
 return new Promise((resolve,reject)=>{
  const reader=new FileReader();

  reader.onload=e=>{
   try{
    const workbook=XLSX.read(e.target.result,{type:"array"});
    const sheet=workbook.Sheets[workbook.SheetNames[0]];

    resolve(XLSX.utils.sheet_to_json(sheet,{
     range:headerRow,
     defval:""
    }));
   }catch(error){
    reject(error);
   }
  };

  reader.onerror=()=>reject(new Error("خواندن فایل ممکن نشد."));
  reader.readAsArrayBuffer(file);
 });
}

/* بارگذاری دستی فایل‌ها */
async function loadManualFiles(){
 const jpFile=$("jpFile").files[0];
 const salesFile=$("salesFile").files[0];
 const mappingFile=$("mappingFile").files[0];

 if(!jpFile||!salesFile){
  setMessage("لطفاً هر دو فایل Journey Plan و Sales را انتخاب کن.");
  return;
 }

 $("loadBtn").disabled=true;

 try{
  setMessage("در حال خواندن فایل‌ها...","در حال پردازش");

  jpRows=await readExcel(jpFile,1);
  salesRows=await readExcel(salesFile,3);
  productRows=mappingFile?await readExcel(mappingFile,0):[];

  buildData();
  renderAll();

  setMessage(
   "اطلاعات با موفقیت بارگذاری شد. تعداد فروشگاه‌ها: "+
   stores.length.toLocaleString("fa-IR"),
   "آماده"
  );

 }catch(error){
  console.error(error);
  setMessage("خطا در خواندن فایل: "+error.message,"خطا");

 }finally{
  $("loadBtn").disabled=false;
 }
}

/* نگاشت کد کالا به نام محصول */
function buildProductMap(){
 const result={};

 for(const row of productRows){
  const code=norm(first(row,["Item Code"]));
  const name=first(row,["Product Name","Product  Name","Item"]);

  if(code&&name)result[code]=name;
 }

 return result;
}

/* تطبیق اطلاعات فروش با مشتریان */
function buildDataCore(){
 const productMap=buildProductMap();
 const salesByCustomer=new Map();

 for(const row of salesRows){
  const customerCode=first(row,["کد مشتری"]);
  const accountNumber=first(row,["Account Number"]);

  const keys=[...new Set([
   ...keyVariants(customerCode),
   ...keyVariants(accountNumber)
  ])];

  const itemCode=first(row,["کد کالا"]);
  const rawName=first(row,["نام کالا"]);

  const item={
   code:itemCode,
   name:productMap[norm(itemCode)]||rawName||itemCode||"محصول نامشخص",
   date:first(row,["invoicedate","تاریخ"]),
   quantity:first(row,["تعداد"])
  };

  for(const key of keys){
   if(!salesByCustomer.has(key)){
    salesByCustomer.set(key,[]);
   }

   if(item.code||rawName){
    salesByCustomer.get(key).push(item);
   }
  }
 }

 const jpMap=new Map();

 for(const row of jpRows){
  const customerI=first(row,["CustomerI"]);
  const newCode=first(row,["New Customer Code"]);
  const oracleCode=first(row,["Oracle Code"]);

  const keys=[...new Set([
   ...keyVariants(customerI),
   ...keyVariants(newCode),
   ...keyVariants(oracleCode)
  ])];

  const products=[];
  const seen=new Set();

  for(const key of keys){
   for(const item of salesByCustomer.get(key)||[]){
    const itemId=[item.code,item.name,item.date].join("|");

    if(!seen.has(itemId)){
     seen.add(itemId);
     products.push(item);
    }
   }
  }

  const id=customerI||newCode||oracleCode;
  if(!id)continue;

  const lat=numeric(row["Latitude"]);
  const lng=numeric(row["Longitude"]);

  jpMap.set(id,{
   id,
   code:newCode||customerI,
   name:first(row,["Customer Name","DescriptionAR","DescriptionE"])||"فروشگاه بدون نام",
   address:first(row,[" Address","Address"]),
   area:first(row,["Area Name"]),
   territory:first(row,["TerritoryNo"]),
   salesman:first(row,["Salesman Name"]),
   salesmanCode:first(row,["Salesman Code"]),
   lat,
   lng,
   hasCoords:validCoord(lat,lng),
   hasSales:products.length>0,
   inJP:true,
   today:isTodayVisit(row),
   products
  });
 }

 const salesOnly=new Map();

 for(const row of salesRows){
  const customerCode=first(row,["کد مشتری"]);
  const accountNumber=first(row,["Account Number"]);

  const keys=[...new Set([
   ...keyVariants(customerCode),
   ...keyVariants(accountNumber)
  ])];

  const matched=keys.some(key=>{
   for(const store of jpMap.values()){
    const storeKeys=[
     ...keyVariants(store.id),
     ...keyVariants(store.code)
    ];

    if(storeKeys.includes(key))return true;
   }

   return false;
  });

  if(matched)continue;

  const id=customerCode||accountNumber;
  if(!id)continue;

  if(!salesOnly.has(id)){
   salesOnly.set(id,{
    id,
    code:customerCode||accountNumber,
    name:first(row,["نام مشتری"])||"مشتری گزارش فروش",
    address:first(row,["آدرس"]),
    area:first(row,["گروه مشتری"]),
    territory:"",
    salesman:first(row,["نام فروشنده"]),
    salesmanCode:first(row,["کد فروشنده"]),
    lat:0,
    lng:0,
    hasCoords:false,
    hasSales:true,
    inJP:false,
    today:false,
    products:[]
   });
  }

  const itemCode=first(row,["کد کالا"]);
  const itemName=productMap[norm(itemCode)]||
   first(row,["نام کالا"])||itemCode||"محصول نامشخص";

  salesOnly.get(id).products.push({
   code:itemCode,
   name:itemName,
   date:first(row,["invoicedate","تاریخ"]),
   quantity:first(row,["تعداد"])
  });
 }

 stores=[...jpMap.values(),...salesOnly.values()];

 for(const store of stores){
  const unique=new Map();

  for(const p of store.products){
   unique.set([p.code,p.name,p.date].join("|"),p);
  }

  store.products=[...unique.values()];
 }
}

function isTodayVisit(row){
 const days=[
  "Sunday","Monday","Tuesday","Wednesday",
  "Thursday","Friday","Saturday"
 ];

 const today=days[new Date().getDay()];
 const value=clean(row[today]);

 return value!==""&&value!=="0"&&value!=="0.0";
}

/* نقشه */
function initMap(){
 if(map)return;

 map=L.map("map").setView([35.84,50.94],10);

 L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png",{
  attribution:"© OpenStreetMap"
 }).addTo(map);
}

function storeColor(store){
 if(!store.inJP)return"#3478d4";
 return store.hasSales?"#15966b":"#e69a27";
}

function renderMap(list){
 initMap();

 markers.forEach(marker=>map.removeLayer(marker));
 markers=[];

 const points=[];

 list.forEach(store=>{
  if(!store.hasCoords)return;

  const marker=L.circleMarker([store.lat,store.lng],{
   radius:7,
   color:"#fff",
   weight:1.5,
   fillColor:storeColor(store),
   fillOpacity:1
  }).addTo(map);

  marker.bindTooltip(escapeHtml(store.name));
  marker.on("click",()=>showStore(store.id));

  markers.push(marker);
  points.push([store.lat,store.lng]);
 });

 if(points.length===1){
  map.setView(points[0],14);
 }else if(points.length>1){
  map.fitBounds(points,{
   padding:[20,20],
   maxZoom:13
  });
 }

 setTimeout(()=>map.invalidateSize(),150);
}

/* فیلترها */
function fillFilters(){
 const areas=[
  ...new Set(stores.map(s=>s.area).filter(Boolean))
 ].sort();

 const people=[
  ...new Set(stores.map(s=>s.salesman).filter(Boolean))
 ].sort();

 $("areaFilter").innerHTML=
  '<option value="all">همه مناطق</option>'+
  areas.map(x=>
   `<option value="${escapeHtml(x)}">${escapeHtml(x)}</option>`
  ).join("");

 $("personFilter").innerHTML=
  '<option value="all">همه ویزیتورها</option>'+
  people.map(x=>
   `<option value="${escapeHtml(x)}">${escapeHtml(x)}</option>`
  ).join("");
}

function filteredStores(){
 const q=norm($("search").value).toLowerCase();
 const type=$("typeFilter").value;
 const area=$("areaFilter").value;
 const person=$("personFilter").value;
 const day=$("dayFilter").value;

 return stores.filter(s=>{
  const hay=norm([
   s.name,s.code,s.address,s.area,s.salesman
  ].join(" ")).toLowerCase();

  if(q&&!hay.includes(q))return false;
  if(type==="sales"&&!s.hasSales)return false;
  if(type==="nosales"&&(!s.inJP||s.hasSales))return false;
  if(type==="salesOnly"&&s.inJP)return false;
  if(area!=="all"&&s.area!==area)return false;
  if(person!=="all"&&s.salesman!==person)return false;
  if(day==="today"&&!s.today)return false;

  if([
   "Saturday","Sunday","Monday",
   "Tuesday","Wednesday","Thursday"
  ].includes(day)&&!s.visitDays?.includes(day)){
   return false;
  }

  return true;
 });
}

function renderList(list){
 $("resultCount").textContent=
  list.length.toLocaleString("fa-IR")+" فروشگاه";

 const shown=list.slice(0,300);

 if(!shown.length){
  $("storeList").innerHTML=
   '<div class="empty">فروشگاهی پیدا نشد.</div>';
  return;
 }

 $("storeList").innerHTML=shown.map(s=>{
  const index=stores.indexOf(s);

  return `
   <div class="store" data-store-index="${index}">
    <i class="pin" style="background:${storeColor(s)}"></i>
    <div class="storeinfo">
     <b>${escapeHtml(s.name)}</b>
     <small>
      ${escapeHtml(s.code)} · ${escapeHtml(s.area||"بدون منطقه")}<br>
      ${s.hasSales?"دارای سابقه فروش":"بدون سابقه فروش"} ·
      ${s.hasCoords?"دارای مختصات":"بدون مختصات"}
     </small>
    </div>
    <span class="arrow">‹</span>
   </div>`;
 }).join("");

 $("storeList").querySelectorAll("[data-store-index]").forEach(el=>{
  el.addEventListener("click",()=>{
   const store=stores[Number(el.dataset.storeIndex)];
   if(store)showStore(store.id);
  });
 });
}

function renderAll(){
 $("total").textContent=stores.length.toLocaleString("fa-IR");

 $("withSales").textContent=
  stores.filter(s=>s.hasSales).length.toLocaleString("fa-IR");

 $("withoutSales").textContent=
  stores.filter(s=>s.inJP&&!s.hasSales).length.toLocaleString("fa-IR");

 fillFilters();
 applyFilters();
}

function applyFilters(){
 const list=filteredStores();
 renderList(list);
 renderMap(list);
}

/* جزئیات فروشگاه */
function showStore(id){
 const store=stores.find(s=>s.id===id);
 if(!store)return;

 $("modalTitle").textContent=store.name;

 const productHtml=store.products.length?
  store.products.map(p=>`
   <div class="product">
    <b>${escapeHtml(p.name)}</b>
    <small>
     کد کالا: ${escapeHtml(p.code||"—")} ·
     تاریخ: ${escapeHtml(p.date||"—")} ·
     تعداد: ${escapeHtml(p.quantity||"—")}
    </small>
   </div>
  `).join("")
  :'<div class="empty">برای این فروشگاه محصولی در گزارش فروش پیدا نشد.</div>';

 const mapsUrl=store.hasCoords?
  `https://www.google.com/maps/dir/?api=1&destination=${store.lat},${store.lng}`:
  `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(store.address||store.name)}`;

 $("modalContent").innerHTML=`
  <div style="font-size:12px;line-height:2;color:#617487;margin:12px 0">
   کد مشتری: ${escapeHtml(store.code||"—")}<br>
   منطقه: ${escapeHtml(store.area||"—")}<br>
   قلمرو: ${escapeHtml(store.territory||"—")}<br>
   ویزیتور: ${escapeHtml(store.salesman||"—")}<br>
   آدرس: ${escapeHtml(store.address||"—")}<br>
   وضعیت: ${store.inJP?"در Journey Plan":"فقط در گزارش فروش"} ·
   ${store.hasSales?"دارای فروش":"بدون فروش"}
  </div>

  <div class="section-title">
   محصولات ثبت‌شده (${store.products.length})
  </div>

  ${productHtml}

  <div class="actions">
   <a href="${mapsUrl}" target="_blank" rel="noopener">
    مسیریابی با Google Maps
   </a>
   ${store.hasCoords?
    `<a href="https://www.google.com/maps?q=${store.lat},${store.lng}" target="_blank" rel="noopener">نمایش موقعیت</a>`:
    ""
   }
  </div>
 `;

 $("modal").style.display="flex";
}

function closeModal(){
 $("modal").style.display="none";
}

/* ورود و مجوز Google */
function getDriveToken(){
 return new Promise((resolve,reject)=>{
  if(!window.google?.accounts?.oauth2){
   reject(new Error(
    "کتابخانه ورود Google هنوز بارگذاری نشده است. "+
    "اتصال اینترنت را بررسی کن، صفحه را یک‌بار Refresh کن "+
    "و دوباره دکمه سبز را بزن."
   ));
   return;
  }

  if(!driveTokenClient){
   driveTokenClient=google.accounts.oauth2.initTokenClient({
    client_id:GOOGLE_CLIENT_ID,
    scope:"https://www.googleapis.com/auth/drive.readonly",

    callback:response=>{
     if(response.error){
      reject(new Error(
       response.error_description||response.error
      ));
     }else if(response.access_token){
      lastAccessToken=response.access_token;
      resolve(response.access_token);
     }else{
      reject(new Error(
       "Google مجوز دسترسی برنگرداند. دوباره تلاش کن."
      ));
     }
    },

    error_callback:error=>{
     reject(new Error(
      error.message||"ورود گوگل لغو شد."
     ));
    }
   });
  }

  driveTokenClient.requestAccessToken({
   prompt:lastAccessToken?"":"consent"
  });
 });
}

async function driveRequest(url,token){
 const response=await fetch(url,{
  headers:{Authorization:"Bearer "+token}
 });

 if(!response.ok){
  const detail=await response.text();
  let readable=detail;

  try{
   const parsed=JSON.parse(detail);
   readable=parsed.error?.message||detail;
  }catch(e){}

  throw new Error(
   "Google Drive API ("+response.status+"): "+
   readable.slice(0,1000)
  );
 }

 return response;
}

/* تاریخ داخل نام فایل */
function dateFromFilename(name){
 const match=name.match(/(\d{8})/);
 return match?match[1]:"00000000";
}

/* پیدا کردن جدیدترین فایل */
function newestFile(files,prefix){
 return files
  .filter(f=>f.name.toLowerCase().startsWith(prefix.toLowerCase()))
  .sort((a,b)=>{
   const dateCompare=
    dateFromFilename(b.name).localeCompare(dateFromFilename(a.name));

   if(dateCompare!==0)return dateCompare;

   return String(b.modifiedTime||"").localeCompare(
    String(a.modifiedTime||"")
   );
  })[0]||null;
}

/* دریافت فهرست فایل‌های پوشه */
async function listDriveFiles(token){
 let all=[];
 let pageToken="";

 do{
  const params=new URLSearchParams({
   q:`'${DRIVE_FOLDER_ID}' in parents and trashed=false`,
   fields:"nextPageToken,files(id,name,mimeType,modifiedTime)",
   pageSize:"1000",
   supportsAllDrives:"true",
   includeItemsFromAllDrives:"true"
  });

  if(pageToken)params.set("pageToken",pageToken);

  const response=await driveRequest(
   "https://www.googleapis.com/drive/v3/files?"+params.toString(),
   token
  );

  const data=await response.json();

  all=all.concat(data.files||[]);
  pageToken=data.nextPageToken||"";

 }while(pageToken);

 return all;
}

/*
 دانلود فایل:
 Google Sheets با Export
 Excel معمولی با alt=media
*/
async function downloadDriveExcel(file,token,headerRow){
 let url;

 if(file.mimeType==="application/vnd.google-apps.spreadsheet"){
  url=
   "https://www.googleapis.com/drive/v3/files/"+
   encodeURIComponent(file.id)+
   "/export?mimeType="+
   encodeURIComponent(
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
   );

 }else{
  url=
   "https://www.googleapis.com/drive/v3/files/"+
   encodeURIComponent(file.id)+"?alt=media";
 }

 const response=await driveRequest(url,token);
 const buffer=await response.arrayBuffer();

 const workbook=XLSX.read(buffer,{type:"array"});
 const sheet=workbook.Sheets[workbook.SheetNames[0]];

 return XLSX.utils.sheet_to_json(sheet,{
  range:headerRow,
  defval:""
 });
}

/* همگام‌سازی */
async function syncFromDrive(){
 const button=$("driveBtn");
 button.disabled=true;

 try{
  setMessage(
   "در حال درخواست مجوز Google Drive...",
   "در حال اتصال"
  );

  const token=await getDriveToken();

  setMessage("در حال جستجو در پوشه Navigation...");

  const files=await listDriveFiles(token);

  const jpFile=newestFile(files,"JP_");
  const salesFile=newestFile(files,"TPA_Alborz");
  const mappingFile=newestFile(files,"Product Mapping");

  if(!jpFile){
   throw new Error(
    "فایل Journey Plan پیدا نشد.\n"+
    "فایل‌های دیده‌شده:\n"+
    (files.map(f=>f.name+" ("+f.mimeType+")").join("\n")||
     "هیچ فایلی دیده نشد")
   );
  }

  if(!salesFile){
   throw new Error(
    "فایل گزارش فروش پیدا نشد.\n"+
    "فایل‌های دیده‌شده:\n"+
    (files.map(f=>f.name+" ("+f.mimeType+")").join("\n")||
     "هیچ فایلی دیده نشد")
   );
  }

  setMessage(
   "فایل‌های پیدا شده:\n"+
   jpFile.name+" — "+jpFile.mimeType+"\n"+
   salesFile.name+" — "+salesFile.mimeType+"\n"+
   (mappingFile?
    mappingFile.name+" — "+mappingFile.mimeType:
    "Product Mapping پیدا نشد")+
   "\n\nدر حال دریافت فایل‌ها..."
  );

  jpRows=await downloadDriveExcel(jpFile,token,1);
  salesRows=await downloadDriveExcel(salesFile,token,3);

  productRows=mappingFile?
   await downloadDriveExcel(mappingFile,token,0):[];

  buildData();
  renderAll();

  setMessage(
   "همگام‌سازی موفق بود.\n"+
   "Journey Plan: "+jpFile.name+"\n"+
   "Sales: "+salesFile.name+"\n"+
   "Product Mapping: "+(mappingFile?mappingFile.name:"پیدا نشد")+
   "\nتعداد فروشگاه‌ها: "+stores.length.toLocaleString("fa-IR"),
   "همگام‌سازی موفق"
  );

 }catch(error){
  console.error(error);

  setMessage(
   "دریافت اطلاعات انجام نشد:\n"+
   error.message+
   "\n\nاگر خطای 401 یا 403 مشاهده شد، جزئیات آن را برای بررسی بفرست.",
   "خطا در اتصال"
  );

 }finally{
  button.disabled=false;
 }
}

/* ===== برنامه هفتگی Auditor، تبادل Excel و ذخیره محلی ===== */

const PLAN_STORAGE_KEY="tpa_auditor_visit_plan_v1";

let auditPlan=[];
let pendingExternal=[];
let lastSyncLabel="";
let currentLanguage="fa";

function loadSavedPlan(){
 try{
  auditPlan=JSON.parse(
   localStorage.getItem(PLAN_STORAGE_KEY)||"[]"
  );
 }catch(e){
  auditPlan=[];
 }

 renderPlan();
}

function savePlan(){
 try{
  localStorage.setItem(
   PLAN_STORAGE_KEY,
   JSON.stringify(auditPlan)
  );
 }catch(e){
  console.warn("Plan storage unavailable",e);
 }

 renderPlan();
}

function storeNameFromAddress(address){
 const s=clean(address);

 const m=s.match(
  /(?:سوپر\s*مارکت|سوپرمارکت|هایپر|سوپر|مارکت|فروشگاه|مواد\s*غذایی|ارزانسرای)\s*[^،,؛;]*/i
 );

 return m?m[0].trim():"";
}

function jalaliDateText(value){
 if(value===null||value===undefined||value==="")return "";

 if(typeof value==="string"){
  const s=value.trim();
  const short=s.match(/^(\d{2})(\d{2})(\d{2})$/);

  if(short){
   return short[1]+"/"+short[2]+"/"+short[3];
  }

  return s;
 }

 if(typeof value==="number"&&value>20000&&value<80000){
  const d=new Date(
   Date.UTC(1899,11,30)+value*86400000
  );

  return d.toISOString().slice(0,10);
 }

 return String(value);
}

function enrichStoreFields(){
 const jpById=new Map();

 for(const row of jpRows){
  const id=first(row,[
   "CustomerI","New Customer Code","Oracle Code"
  ]);

  if(id)jpById.set(id,row);
 }

 const statusKeys=[
  "Status","Order Status","وضعیت سفارش","OrderStatus"
 ];

 const orderNoKeys=[
  "Order Number","Order No","OrderNumber",
  "شماره سفارش","شماره فاکتور"
 ];

 const dateKeys=[
  "Date","Order Date","تاریخ سفارش"
 ];

 const invoiceKeys=[
  "Invoicedate","Invoice Date","تاریخ صدور فاکتور"
 ];

 const deliveryKeys=[
  "Target Delivery Date","TargetDeliveryDate",
  "تاریخ تحویل مورد انتظار"
 ];

 const grouped=new Map();

 for(const row of salesRows){
  const status=first(row,statusKeys).trim().toLowerCase();

  if(status==="closex"||status==="expired")continue;

  const cust=first(row,[
   "کد مشتری","Customer Code","Account Number"
  ]);

  if(!cust)continue;

  const orderNo=first(row,orderNoKeys);

  const entry={
   orderNo,
   status,
   orderDate:first(row,dateKeys),
   invoiceDate:first(row,invoiceKeys),
   deliveryDate:first(row,deliveryKeys),
   product:first(row,["نام کالا","Item","Product Name"]),
   itemCode:first(row,["کد کالا","Item Code"]),
   quantity:first(row,["تعداد","Quantity"])
  };

  if(!grouped.has(cust))grouped.set(cust,[]);
  grouped.get(cust).push(entry);
 }

 for(const s of stores){
  const raw=jpById.get(s.id);

  s.routeName=raw?
   first(raw,["Route Name"]):
   s.area;

  s.customerClass=raw?
   first(raw,["Customer Class"]):
   "";

  s.storeName=storeNameFromAddress(s.address)||s.name;
  s.orders=[];

  const keys=[
   ...keyVariants(s.id),
   ...keyVariants(s.code)
  ];

  for(const [code,orders] of grouped){
   if(keys.some(k=>keyVariants(code).includes(k))){
    s.orders.push(...orders);
   }
  }

  s.orders=s.orders.filter(o=>
   o.orderNo||o.product||o.itemCode
  );

  s.openOrders=s.orders.filter(o=>!o.invoiceDate);

  s.lastPurchase=s.orders
   .map(o=>o.invoiceDate||o.orderDate)
   .filter(Boolean)
   .sort()
   .pop()||"";

  s.visitDays=raw?
   [
    "Saturday","Sunday","Monday",
    "Tuesday","Wednesday","Thursday"
   ].filter(d=>{
    const v=clean(raw[d]);
    return v!==""&&v!=="0"&&v!=="0.0";
   }):
   [];
 }
}

function buildData(){
 const excluded=new Set(["closex","expired"]);

 salesRows=salesRows.filter(r=>
  !excluded.has(
   first(r,[
    "Status","Order Status","وضعیت سفارش","OrderStatus"
   ]).toLowerCase()
  )
 );

 buildDataCore();
 enrichStoreFields();
 saveOfflineSnapshot();
}

function saveOfflineSnapshot(){
 try{
  const snapshot={
   savedAt:new Date().toISOString(),
   stores,
   jpRows,
   salesRows,
   productRows
  };

  localStorage.setItem(
   "tpa_navigation_snapshot_v1",
   JSON.stringify(snapshot)
  );

  $("lastSync").textContent=
   "آخرین همگام‌سازی: "+new Date().toLocaleString();

 }catch(e){
  console.warn("Offline cache could not be saved",e);
 }
}

function restoreOfflineSnapshot(){
 try{
  const raw=localStorage.getItem(
   "tpa_navigation_snapshot_v1"
  );

  if(!raw)return false;

  const data=JSON.parse(raw);

  if(!Array.isArray(data.stores)||!data.stores.length){
   return false;
  }

  stores=data.stores;
  jpRows=data.jpRows||[];
  salesRows=data.salesRows||[];
  productRows=data.productRows||[];

  $("lastSync").textContent=
   "آخرین همگام‌سازی: "+
   (data.savedAt?
    new Date(data.savedAt).toLocaleString():
    "نامشخص")+
   " (ذخیره محلی)";

  renderAll();
  return true;

 }catch(e){
  return false;
 }
}

function addStoreToPlan(id){
 const s=stores.find(x=>String(x.id)===String(id));
 if(!s)return;

 const date=clean($("planDate")?.value)||"";
 const day=clean($("planDay")?.value)||"Saturday";

 const exists=auditPlan.some(p=>
  String(p.customerId)===String(s.code||s.id)&&
  p.visitDate===date&&
  p.visitDay===day
 );

 if(exists){
  alert("این فروشگاه از قبل در برنامه همین روز وجود دارد.");
  return;
 }

 auditPlan.push({
  visitDate:date,
  visitDay:day,
  customerId:s.code||s.id,
  storeName:s.storeName||s.name,
  customerName:s.name,
  address:s.address,
  routeName:s.routeName||s.area,
  latitude:s.lat||"",
  longitude:s.lng||"",
  priority:"Medium",
  selectionReason:"",
  auditorComment:"",
  selectionSource:"Internal",
  approved:true
 });

 savePlan();
 closeModal();
}

function renderPlan(){
 const box=$("planList");
 if(!box)return;

 if(!auditPlan.length){
  box.innerHTML=
   '<div class="empty">هنوز فروشگاهی به برنامه اضافه نشده است.</div>';
  return;
 }

 box.innerHTML=auditPlan.map((p,i)=>`
  <div class="planrow">
   <strong>
    ${escapeHtml(p.storeName||p.customerName||p.customerId)}
   </strong>
   <small>
    ${escapeHtml(p.visitDay||"")} ·
    ${escapeHtml(p.visitDate||"بدون تاریخ")} ·
    کد: ${escapeHtml(p.customerId)}<br>
    منبع: <b>${escapeHtml(p.selectionSource||"Internal")}</b> ·
    اولویت: ${escapeHtml(p.priority||"Medium")}<br>
    ${escapeHtml(p.selectionReason||"بدون دلیل ثبت‌شده")}
   </small>
   <div class="rowactions">
    <button class="smallbtn secondary" onclick="editPlanRow(${i})">ویرایش</button>
    <button class="smallbtn" onclick="removePlanRow(${i})">حذف</button>
   </div>
  </div>
 `).join("");
}

function editPlanRow(i){
 const p=auditPlan[i];
 if(!p)return;

 const reason=prompt(
  "دلیل انتخاب فروشگاه:",
  p.selectionReason||""
 );

 if(reason===null)return;

 p.selectionReason=reason;

 const priority=prompt(
  "اولویت: High / Medium / Low",
  p.priority||"Medium"
 );

 if(priority!==null&&
  ["High","Medium","Low"].includes(priority)){
  p.priority=priority;
 }

 savePlan();
}

function removePlanRow(i){
 auditPlan.splice(i,1);
 savePlan();
}

function planRowsForExcel(){
 return auditPlan.map(p=>({
  "Visit Date":p.visitDate||"",
  "Customer ID":p.customerId||"",
  "Store Name":p.storeName||"",
  "Customer Name":p.customerName||"",
  "Address":p.address||"",
  "Route Name":p.routeName||"",
  "Latitude":p.latitude??"",
  "Longitude":p.longitude??"",
  "Priority":p.priority||"Medium",
  "Selection Reason":p.selectionReason||"",
  "Auditor Comment":p.auditorComment||"",
  "Selection Source":
   p.selectionSource==="External"?"External":"Internal"
 }));
}

function exportAuditPlan(){
 if(!window.XLSX){
  alert("کتابخانه Excel در دسترس نیست.");
  return;
 }

 const wb=XLSX.utils.book_new();
 const rows=planRowsForExcel();

 const headers=[
  "Visit Date","Customer ID","Store Name","Customer Name",
  "Address","Route Name","Latitude","Longitude","Priority",
  "Selection Reason","Auditor Comment","Selection Source"
 ];

 const ws=XLSX.utils.json_to_sheet(rows,{
  header:headers
 });

 ws["!cols"]=[
  16,16,24,26,42,22,14,14,14,30,34,18
 ].map(w=>({wch:w}));

 XLSX.utils.book_append_sheet(
  wb,ws,"Audit Visit Plan"
 );

 const ins=[
  ["Field","Instructions"],
  ["Required fields","Visit Date and Customer ID"],
  ["Selection Source","Internal = selected in app; External = imported from uploaded workbook"],
  ["Import review","External rows are presented to Auditor for approval before becoming part of the active plan"]
 ];

 XLSX.utils.book_append_sheet(
  wb,
  XLSX.utils.aoa_to_sheet(ins),
  "Instructions"
 );

 XLSX.writeFile(wb,"Audit_Visit_Plan.xlsx");
}

async function importAuditPlan(){
 const file=$("importPlanFile").files[0];

 if(!file){
  alert("ابتدا فایل اکسل را انتخاب کن.");
  return;
 }

 try{
  const rows=await readExcel(file,0);
  const valid=[];
  const invalid=[];
  const seen=new Set();

  rows.forEach((r,i)=>{
   const id=first(r,[
    "Customer ID","CustomerID","کد مشتری"
   ]);

   const date=first(r,[
    "Visit Date","VisitDate","تاریخ بازدید"
   ]);

   if(!id||!date){
    invalid.push(i+2);
    return;
   }

   const k=String(id)+"|"+String(date);

   if(seen.has(k)){
    invalid.push(i+2);
    return;
   }

   seen.add(k);

   const matched=stores.find(s=>
    keyVariants(s.id).includes(norm(id))||
    keyVariants(s.code).includes(norm(id))
   );

   if(!matched){
    invalid.push(i+2);
    return;
   }

   valid.push({
    visitDate:date,
    visitDay:first(r,["Visit Day"])||$("planDay").value,
    customerId:id,
    storeName:first(r,["Store Name"])||
     matched.storeName||matched.name,
    customerName:first(r,["Customer Name"])||matched.name,
    address:first(r,["Address"])||matched.address,
    routeName:first(r,["Route Name"])||
     matched.routeName||matched.area,
    latitude:first(r,["Latitude"])||matched.lat||"",
    longitude:first(r,["Longitude"])||matched.lng||"",
    priority:first(r,["Priority"])||"Medium",
    selectionReason:first(r,["Selection Reason"]),
    auditorComment:first(r,["Auditor Comment"]),
    selectionSource:"External",
    approved:false
   });
  });

  pendingExternal=valid;

  $("importReport").innerHTML=`
   تعداد ردیف‌های معتبر: <b>${valid.length}</b><br>
   ردیف‌های نامعتبر، تکراری یا با کد ناشناخته:
   <b>${invalid.length}</b><br>
   ${valid.length?"پیشنهادها آماده بررسی هستند. برای افزودن آن‌ها تأیید کن.":""}
   ${valid.length?
    '<button class="smallbtn green" id="acceptExternalBtn">تأیید و افزودن پیشنهادهای مدیر</button>':
    ""
   }
  `;

  if(valid.length){
   $("acceptExternalBtn").onclick=()=>{
    for(const p of pendingExternal){
     const duplicate=auditPlan.some(x=>
      String(x.customerId)===String(p.customerId)&&
      String(x.visitDate)===String(p.visitDate)
     );

     if(!duplicate)auditPlan.push(p);
    }

    pendingExternal=[];
    savePlan();

    $("importReport").textContent=
     "پیشنهادهای غیرتکراری به برنامه اضافه شدند.";
   };
  }

 }catch(e){
  $("importReport").textContent=
   "خطا در خواندن فایل: "+e.message;
 }
}

function showPage(id){
 document.querySelectorAll(".page").forEach(p=>
  p.classList.toggle("active",p.id===id)
 );

 document.querySelectorAll(".appnav button").forEach(b=>
  b.classList.toggle("active",b.dataset.page===id)
 );

 if(id==="mapPage"){
  setTimeout(()=>map?.invalidateSize(),100);
 }
}

function renderAnalysis(){
 if(!$("noGpsCount"))return;

 $("noGpsCount").textContent=
  stores.filter(s=>!s.hasCoords).length.toLocaleString("fa-IR");

 $("openOrderCount").textContent=
  stores.reduce((n,s)=>n+(s.openOrders?.length||0),0).toLocaleString("fa-IR");

 $("noPurchaseCount").textContent=
  stores.filter(s=>!s.hasSales).length.toLocaleString("fa-IR");
}

function showAnalysisList(kind){
 let list=stores;

 if(kind==="nogps"){
  list=stores.filter(s=>!s.hasCoords);
 }

 if(kind==="nopurchase"){
  list=stores.filter(s=>!s.hasSales);
 }

 if(kind==="open"){
  list=stores.filter(s=>(s.openOrders||[]).length);
 }

 $("analysisResults").innerHTML=list.length?
  list.slice(0,150).map(s=>`
   <div class="planrow">
    <strong>${escapeHtml(s.storeName||s.name)}</strong>
    <small>
     ${escapeHtml(s.code)} · ${escapeHtml(s.address||"")}
    </small>
   </div>
  `).join("")
  :'<div class="empty">موردی پیدا نشد.</div>';
}

function findNearby(){
 if(!navigator.geolocation){
  alert("موقعیت‌یابی در این مرورگر پشتیبانی نمی‌شود.");
  return;
 }

 navigator.geolocation.getCurrentPosition(pos=>{
  const {latitude,longitude}=pos.coords;

  const dist=(a,b)=>{
   const R=6371;
   const d1=(a.lat-latitude)*Math.PI/180;
   const d2=(a.lng-longitude)*Math.PI/180;

   const x=
    Math.sin(d1/2)**2+
    Math.cos(latitude*Math.PI/180)*
    Math.cos(a.lat*Math.PI/180)*
    Math.sin(d2/2)**2;

   return 2*R*Math.asin(Math.sqrt(x));
  };

  const near=stores
   .filter(s=>s.hasCoords)
   .map(s=>({...s,distance:dist(s)}))
   .filter(s=>s.distance<=5)
   .sort((a,b)=>a.distance-b.distance);

  $("analysisResults").innerHTML=near.length?
   near.map(s=>`
    <div class="planrow">
     <strong>${escapeHtml(s.storeName||s.name)}</strong>
     <small>
      ${s.distance.toFixed(2)} کیلومتر ·
      ${escapeHtml(s.code)}
     </small>
    </div>
   `).join("")
   :'<div class="empty">در شعاع ۵ کیلومتری فروشگاهی پیدا نشد.</div>';

  showPage("analysisPage");

 },()=>{
  alert("دسترسی موقعیت مکانی ممکن نشد. مجوز Location را بررسی کن.");
 },{
  enableHighAccuracy:true,
  timeout:12000
 });
}

function applyTheme(){
 if(localStorage.getItem("tpa_theme")==="dark"){
  document.body.classList.add("dark");
 }
}

function registerOfflineShell(){
 if("serviceWorker" in navigator&&location.protocol.startsWith("http")){
  const sw=
   `const CACHE='tpa-nav-shell-v1';`+
   `const ASSETS=[location.pathname];`+
   `self.addEventListener('install',e=>e.waitUntil(caches.open(CACHE).then(c=>c.addAll(ASSETS))));`+
   `self.addEventListener('fetch',e=>{if(e.request.method==='GET'&&new URL(e.request.url).origin===location.origin)e.respondWith(caches.match(e.request).then(r=>r||fetch(e.request).then(x=>{const copy=x.clone();caches.open(CACHE).then(c=>c.put(e.request,copy));return x}).catch(()=>r))) });`;

  const blob=new Blob([sw],{
   type:"text/javascript"
  });

  navigator.serviceWorker.register("./sw.js").catch(()=>{});
 }
}

/* رویدادها */
$("loadBtn").addEventListener("click",loadManualFiles);
$("closeModal").addEventListener("click",closeModal);

$("modal").addEventListener("click",event=>{
 if(event.target.id==="modal")closeModal();
});

[
 "search","typeFilter","areaFilter",
 "personFilter","dayFilter"
].forEach(id=>{
 $(id).addEventListener(
  id==="search"?"input":"change",
  applyFilters
 );
});

/* رویدادهای بخش برنامه بازدید و تحلیل */
document.querySelectorAll(".appnav button").forEach(b=>
 b.addEventListener("click",()=>showPage(b.dataset.page))
);

$("planSearch").addEventListener("input",()=>{
 const q=norm($("planSearch").value);

 const results=stores
  .filter(s=>
   q&&norm([s.code,s.name,s.address].join(" ")).includes(q)
  )
  .slice(0,12);

 $("planSearchResults").innerHTML=results.map(s=>`
  <div class="store">
   <div class="storeinfo">
    <b>${escapeHtml(s.storeName||s.name)}</b>
    <small>
     ${escapeHtml(s.code)} ·
     ${escapeHtml(s.routeName||s.area||"")}
    </small>
   </div>
   <button class="smallbtn green"
    onclick="addStoreToPlan('${'${escapeHtml(s.id)}'}')">
    افزودن
   </button>
  </div>
 `).join("");
});

$("exportPlanBtn").addEventListener("click",exportAuditPlan);
$("importPlanBtn").addEventListener("click",importAuditPlan);
$("nearbyBtn").addEventListener("click",findNearby);

$("noGpsBtn").addEventListener(
 "click",()=>showAnalysisList("nogps")
);

$("noPurchaseBtn").addEventListener(
 "click",()=>showAnalysisList("nopurchase")
);

$("openOrdersBtn").addEventListener(
 "click",()=>showAnalysisList("open")
);

$("themeBtn").addEventListener("click",()=>{
 document.body.classList.toggle("dark");

 localStorage.setItem(
  "tpa_theme",
  document.body.classList.contains("dark")?"dark":"light"
 );
});

$("langBtn").addEventListener("click",()=>{
 alert(
  "نسخه فعلی با رابط فارسی ارائه شده است؛ آماده‌سازی کامل ترجمه انگلیسی نیازمند بازبینی همه متن‌ها و پیام‌هاست."
 );
});

$("clearLocalBtn").addEventListener("click",()=>{
 if(confirm("داده‌های محلی و برنامه ذخیره‌شده پاک شوند؟")){
  localStorage.removeItem("tpa_navigation_snapshot_v1");
  localStorage.removeItem(PLAN_STORAGE_KEY);

  auditPlan=[];
  renderPlan();

  alert("داده‌های محلی پاک شدند.");
 }
});

const originalRenderAll=renderAll;

renderAll=function(){
 originalRenderAll();
 renderAnalysis();
 renderPlan();
};

loadSavedPlan();
applyTheme();
registerOfflineShell();

if(!stores.length){
 restoreOfflineSnapshot();
}
