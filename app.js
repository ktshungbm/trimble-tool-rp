/**
 * Paint Approval Tool v10.0
 * ─────────────────────────────────────
 * 2 file Excel:
 *   File 1 (Ban hành)  → 🟢 Xanh lá  #00FF00
 *   File 2 (RFI)       → 🟠 Màu cam  #FFBB00
 *   Còn lại            → ⚪ Màu xám  #808080
 *
 * Dựa trên v9 đã chứng minh hoạt động.
 * Chiến lược: tô xám hết → hiện+màu ban hành → hiện+màu rfi
 * ─────────────────────────────────────
 */

var COLOR_GREEN  = "#00FF00";
var COLOR_ORANGE = "#FFBB00";
var COLOR_GRAY   = "#808080";
var RETRY_MAX   = 7;
var RETRY_DELAY = 2000;
var BATCH_CVT   = 500;
var BATCH_CLR   = 300;
var PAINT_DELAY = 150;

var _api = null;
var _guidsGreen = [];  // file 1
var _guidsOrange= [];  // file 2

/* ═══ UI ═══ */
function log(m,t){var e=document.getElementById("log");if(!e){console.log(m);return;}var s=document.createElement("span");if(t)s.className=t;s.textContent=m+"\n";e.appendChild(s);e.scrollTop=e.scrollHeight;console.log("["+(t||"")+"] "+m);}
function clearLog(){var e=document.getElementById("log");if(e)e.innerHTML="";}
function setStat(id,v){var e=document.getElementById(id);if(e)e.textContent=(v!=null)?v:"—";}
function setProgress(p){var w=document.getElementById("progWrap"),b=document.getElementById("progBar");if(!w||!b)return;if(p<=0){w.classList.remove("on");b.style.width="0%";return;}w.classList.add("on");b.style.width=Math.min(p,100)+"%";}
function lockUI(y){["applyBtn","resetBtn"].forEach(function(id){var e=document.getElementById(id);if(e)e.disabled=y;});}
function sleep(ms){return new Promise(function(r){setTimeout(r,ms);});}
function pad2(n){return String(n).padStart(2,"0");}
function fmtN(n){return typeof n==="number"?n.toLocaleString():String(n);}
function checkApplyBtn(){document.getElementById("applyBtn").disabled=(!_guidsGreen.length&&!_guidsOrange.length);}

/* ═══ UUID ↔ IFC GUID ═══ */
var B64="0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz_$";
function to64(n,d){var r=[];for(var i=0;i<d;i++){r.push(B64.charAt(n%64));n=Math.floor(n/64);}return r.reverse().join("");}
function from64(s){var r=0;for(var i=0;i<s.length;i++){var x=B64.indexOf(s.charAt(i));if(x<0)return-1;r=r*64+x;}return r;}
function uuid2ifc(u){if(!u)return null;var h=String(u).replace(/-/g,"").toLowerCase();if(h.length!==32||!/^[0-9a-f]{32}$/.test(h))return null;var n=[parseInt(h.substr(0,2),16)];for(var i=0;i<5;i++)n.push(parseInt(h.substr(2+i*6,6),16));var r=to64(n[0],2);for(var i=1;i<6;i++)r+=to64(n[i],4);return r;}
function ifc2uuid(c){if(!c||c.length!==22)return null;var p=[from64(c.substr(0,2))];for(var i=0;i<5;i++)p.push(from64(c.substr(2+i*4,4)));if(p.some(function(x){return x<0;}))return null;var h=p[0].toString(16).padStart(2,"0");for(var i=1;i<6;i++)h+=p[i].toString(16).padStart(6,"0");return h.substr(0,8)+"-"+h.substr(8,4)+"-"+h.substr(12,4)+"-"+h.substr(16,4)+"-"+h.substr(20,12);}
function detectFmt(g){if(!g)return"x";var s=String(g).trim();if(s.length===36&&/^[0-9a-f]{8}-/i.test(s))return"uuid";if(s.length===32&&/^[0-9a-f]{32}$/i.test(s))return"nd";if(s.length===22)return"ifc";return"x";}

/* ═══ API ═══ */
async function getAPI(){if(_api)return _api;_api=await TrimbleConnectWorkspace.connect(window.parent,function(e,d){console.log("[T]",e,d);});log("Đã kết nối Trimble API.","ok");return _api;}

/* ═══ Excel ═══ */
function readWB(f){return new Promise(function(ok,no){var r=new FileReader();r.onload=function(e){try{ok(XLSX.read(e.target.result,{type:"array"}));}catch(err){no(err);}};r.onerror=no;r.readAsArrayBuffer(f);});}

function extractGuids(wb,label){
  if(!wb||!wb.SheetNames||!wb.SheetNames.length)throw new Error("Excel không có sheet.");
  var sn=wb.SheetNames[0];
  var rows=XLSX.utils.sheet_to_json(wb.Sheets[sn],{defval:""});
  if(!rows.length)throw new Error("Sheet trống.");
  var keys=Object.keys(rows[0]);
  
  var gk=keys.find(function(k){
    var up=k.trim().toUpperCase();
    return up==="GUID"||up==="GLOBALID"||up==="IFCGUID"||up==="TEKLA_GUID"||up==="ID"||up==="ID_PART";
  });
  
  if(!gk){
    var maxHits = 0;
    for(var i=0; i<keys.length; i++){
      var k = keys[i];
      var hits = 0;
      for(var r=0; r<Math.min(rows.length, 10); r++){
        var val = String(rows[r][k]||"").trim();
        if(val.toUpperCase().startsWith("ID")) val = val.substring(2);
        if(val.length===22 || val.length===32 || val.length===36) hits++;
      }
      if(hits > maxHits){ maxHits = hits; gk = k; }
    }
  }

  if(!gk){gk=keys[0];log('  ⚠ Không thấy cột GUID, dùng cột đầu: "'+gk+'"',"warn");}
  else { log('  ℹ Đã chọn cột: "'+gk+'"',"info"); }
  
  var seen={},out=[];
  rows.forEach(function(r){var g=String(r[gk]||"").trim();if(g&&!seen[g]){seen[g]=true;out.push(g);}});
  log('  ['+label+'] Sheet "'+sn+'": '+out.length+' dòng data',"info");
  return out;
}

/* ═══ Model ═══ */
async function getModelIds(){
  var api=await getAPI();
  for(var a=1;a<=RETRY_MAX;a++){
    var raw;
    try{raw=await api.viewer.getObjects();}catch(e){if(a<RETRY_MAX){await sleep(RETRY_DELAY);continue;}throw e;}
    if(!Array.isArray(raw)||!raw.length){if(a<RETRY_MAX){await sleep(RETRY_DELAY);continue;}throw new Error("Viewer trống.");}
    var total=0,mids=[];
    raw.forEach(function(g){if(!g||!g.modelId)return;if(mids.indexOf(g.modelId)===-1)mids.push(g.modelId);if(Array.isArray(g.objects))total+=g.objects.length;else if(Array.isArray(g.objectRuntimeIds))total+=g.objectRuntimeIds.length;else if(Array.isArray(g.ids))total+=g.ids.length;});
    if(!mids.length){if(a<RETRY_MAX){await sleep(RETRY_DELAY);continue;}throw new Error("Không thấy modelId.");}
    return{modelIds:mids,total:total};
  }
}

/* ═══ Convert GUIDs → runtimeIds ═══ */
function flat(v){if(v==null)return[];if(typeof v==="number")return[v];if(Array.isArray(v)){var o=[];v.forEach(function(x){if(typeof x==="number")o.push(x);else if(Array.isArray(x))x.forEach(function(y){if(typeof y==="number")o.push(y);});});return o;}return[];}

async function batchConvert(api,mid,guids){
  var out=[];
  for(var i=0;i<guids.length;i+=BATCH_CVT){
    var c=guids.slice(i,i+BATCH_CVT);var r;
    try{r=await api.viewer.convertToObjectRuntimeIds(mid,c);}catch(e){for(var k=0;k<c.length;k++)out.push(null);continue;}
    if(!Array.isArray(r)){for(var k=0;k<c.length;k++)out.push(null);continue;}
    out=out.concat(r);
  }
  return out;
}

/** Thử nhiều format, bao gồm cắt bỏ prefix "ID" */
async function convertAll(api,modelIds,guids,label){
  var result=new Map();
  if(!guids.length) return result;

  var uuids=[],ifcs=[],others=[];
  guids.forEach(function(g){
    var gStr = String(g).trim();
    if(!gStr) return;
    var gNoID = gStr.toUpperCase().startsWith("ID") ? gStr.substring(2) : gStr;
    
    others.push(gStr);
    if(gNoID !== gStr) others.push(gNoID);
    
    var checkPush = function(val){
      var len=val.length;
      if(len===36||len===32) uuids.push(val);
      else if(len===22) ifcs.push(val);
    };
    checkPush(gStr);
    if(gNoID !== gStr) checkPush(gNoID);
  });
  
  function unique(arr){var u={};arr.forEach(function(x){u[x]=1;});return Object.keys(u);}
  uuids = unique(uuids);
  ifcs = unique(ifcs);
  others = unique(others);

  var u2i=uuids.map(uuid2ifc).filter(Boolean);
  var i2u=ifcs.map(ifc2uuid).filter(Boolean);

  for(var mi=0;mi<modelIds.length;mi++){
    var mid=modelIds[mi];var all=[];
    async function tryL(list,lbl){
      if(!list.length)return;
      var conv=await batchConvert(api,mid,list);var hit=0;
      for(var i=0;i<list.length;i++){var ids=flat(conv[i]);if(ids.length){hit++;all=all.concat(ids);}}
      if(hit>0) log("  ["+label+"/"+lbl+"] "+hit+" GUIDs matched","ok");
    }
    await tryL(uuids,"UUID");await tryL(ifcs,"IFC");await tryL(u2i,"U→I");await tryL(i2u,"I→U");await tryL(others,"RAW");
    if(all.length){var u={};all.forEach(function(id){u[id]=1;});result.set(mid,Object.keys(u).map(Number));}
  }
  return result;
}

/* ═══ Paint batch ═══ */
async function paintBatch(api,mid,ids,state){
  for(var i=0;i<ids.length;i+=BATCH_CLR){
    var chunk=ids.slice(i,i+BATCH_CLR);
    try{await api.viewer.setObjectState({modelObjectIds:[{modelId:mid,objectRuntimeIds:chunk}]},state);}catch(e){}
    if(i+BATCH_CLR<ids.length)await sleep(PAINT_DELAY);
  }
}

/* ═══════════════════════════════════════
   MAIN — v10
   1. Reset
   2. Hiện và tô XÁM tất cả
   3. Hiện + tô XANH LÁ (Ban hành)
   4. Hiện + tô CAM (RFI)
   5. Auto Save View
═══════════════════════════════════════ */
async function applyColors(){
  lockUI(true);clearLog();setProgress(5);
  try{
    if(!_guidsGreen.length&&!_guidsOrange.length) throw new Error("Chưa có file nào.");
    var api=await getAPI();

    // 1. Reset
    log("Reset...","info");
    try{await api.viewer.setObjectState(undefined,{color:"reset",visible:"reset"});}catch(e){}
    await sleep(500);
    setProgress(10);

    // 2. Get models
    var mi=await getModelIds();
    setStat("s-total",fmtN(mi.total));
    setProgress(18);

    // 3. Convert GUIDs
    log("Map GUIDs...","info");
    var greenMap=await convertAll(api,mi.modelIds,_guidsGreen,"🟢");
    var orangeMap =await convertAll(api,mi.modelIds,_guidsOrange, "🟠");

    var greenTotal=0,orangeTotal=0;
    greenMap.forEach(function(ids){greenTotal+=ids.length;});
    orangeMap.forEach(function(ids){orangeTotal+=ids.length;});
    setStat("s-green",fmtN(greenTotal));
    setStat("s-orange",fmtN(orangeTotal));

    if(greenTotal===0&&orangeTotal===0){
      log("✗ Không match object nào!","err");
      setProgress(0);lockUI(false);checkApplyBtn();return;
    }
    setProgress(35);

    // 4. Hiện và tô màu XÁM cho TẤT CẢ
    log("Tô xám toàn bộ model...","info");
    try{await api.viewer.setObjectState(undefined,{visible:true, color:COLOR_GRAY});}catch(e){}
    await sleep(800);
    setProgress(42);

    // 5. Hiện + tô XANH LÁ (Ban hành)
    if(greenTotal>0){
      log("━━━ Tô XANH LÁ (Ban hành): "+fmtN(greenTotal)+" ━━━","info");
      for(var i=0;i<mi.modelIds.length;i++){
        var mid=mi.modelIds[i];
        var ids=greenMap.get(mid);
        if(!ids||!ids.length)continue;
        await paintBatch(api,mid,ids,{visible:true,color:COLOR_GREEN});
        log("  ▪ "+fmtN(ids.length)+" objects xanh lá","ok");
      }
    }
    setProgress(58);
    await sleep(300);

    // 6. Hiện + tô MÀU CAM (RFI)
    if(orangeTotal>0){
      log("━━━ Tô MÀU CAM (RFI): "+fmtN(orangeTotal)+" ━━━","info");
      for(var i=0;i<mi.modelIds.length;i++){
        var mid=mi.modelIds[i];
        var ids=orangeMap.get(mid);
        if(!ids||!ids.length)continue;
        await paintBatch(api,mid,ids,{visible:true,color:COLOR_ORANGE});
        log("  ▪ "+fmtN(ids.length)+" objects màu cam","ok");
      }
    }
    setProgress(75);
    await sleep(300);

    setProgress(100);

    log("","info");
    log("✓ HOÀN TẤT!","ok");
    if(greenTotal) log("  🟢 Ban hành: "+fmtN(greenTotal)+" cấu kiện","ok");
    if(orangeTotal)log("  🟠 RFI: "+fmtN(orangeTotal)+" cấu kiện","ok");
    log("  Còn lại: màu xám","info");
    
    setTimeout(async function(){
      setProgress(0);
      log("Đang tự động lưu View...","info");
      await saveView();
    },1500);

  }catch(err){
    log("✗ "+(err&&err.message?err.message:String(err)),"err");setProgress(0);
  }finally{lockUI(false);checkApplyBtn();}
}

/* ═══ Reset ═══ */
async function resetViewer(){
  lockUI(true);clearLog();setProgress(10);
  try{var api=await getAPI();try{await api.viewer.setObjectState(undefined,{color:"reset",visible:"reset"});}catch(e){}await api.viewer.reset();
  setStat("s-total","—");setStat("s-green","—");setStat("s-orange","—");
  setProgress(100);log("✓ Reset OK.","ok");setTimeout(function(){setProgress(0);},1000);}
  catch(e){log("✗ "+(e&&e.message?e.message:String(e)),"err");setProgress(0);}
  finally{lockUI(false);checkApplyBtn();}
}

/* ═══ Save View ═══ */
async function saveView(){
  try{
    var api=await getAPI();
    var n=new Date();
    var name="Auto Update "+n.getFullYear()+"-"+pad2(n.getMonth()+1)+"-"+pad2(n.getDate())+" "+pad2(n.getHours())+":"+pad2(n.getMinutes());
    
    var c=await api.view.createView({
      name:name,
      description:"Paint Approval Tool v10.0 | Auto Save"
    });
    
    if(!c||!c.id)throw new Error("Không nhận được ID của View.");
    
    await api.view.updateView({id:c.id});
    await api.view.selectView(c.id);
    
    log('✓ Đã tự động lưu View: "'+name+'"',"ok");
  }catch(e){
    log("✗ Lỗi lưu view: "+(e&&e.message?e.message:String(e)),"err");
  }
}

/* ═══ File Events ═══ */
async function handleFile(fileInput,fnameId,label,setGuids){
  var f=fileInput.files&&fileInput.files[0];
  if(!f)return;
  document.getElementById(fnameId).textContent=f.name;
  log('Đang đọc ['+label+'] "'+f.name+'"...',"info");
  try{
    var wb=await readWB(f);
    var guids=extractGuids(wb,label);
    setGuids(guids);
    checkApplyBtn();
    if(guids.length>0) log('  ✓ '+guids.length+' GUID sẵn sàng.',"ok");
    else log("  ⚠ Không thấy GUID.","warn");
  }catch(e){
    log("  ✗ "+(e&&e.message?e.message:String(e)),"err");
    setGuids([]);checkApplyBtn();
  }
}

document.getElementById("file1").addEventListener("change",function(){
  handleFile(this,"fname1","Ban hành",function(g){_guidsGreen=g;});
});
document.getElementById("file2").addEventListener("change",function(){
  handleFile(this,"fname2","RFI",function(g){_guidsOrange=g;});
});

document.getElementById("applyBtn").addEventListener("click",applyColors);
document.getElementById("resetBtn").addEventListener("click",resetViewer);
