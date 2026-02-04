(function(){
"use strict";
if(document.getElementById('aop-panel')){document.getElementById('aop-panel').style.display='block';return;}

var items=[],running=false,stopFlag=false;

var css=document.createElement('style');
css.textContent=`
#aop-panel{position:fixed;top:0;right:0;width:460px;height:100vh;background:#f5f5f0;z-index:99999;box-shadow:-4px 0 20px rgba(0,0,0,.15);overflow-y:auto;font-family:sans-serif;font-size:14px;color:#2a2a2a;display:block}
#aop-panel *{box-sizing:border-box;margin:0;padding:0}
.aop-hd{background:#fff;border-bottom:3px solid #e8a840;padding:12px 16px;display:flex;justify-content:space-between;align-items:center;position:sticky;top:0;z-index:1}
.aop-hd h3{font-size:16px}
.aop-x{background:none;border:none;font-size:22px;cursor:pointer;color:#888;padding:2px 8px}
.aop-x:hover{color:#333}
.aop-bd{padding:16px}
.aop-c{background:#fff;border-radius:8px;padding:16px;margin-bottom:12px;box-shadow:0 1px 4px rgba(0,0,0,.06)}
.aop-c h4{font-size:14px;font-weight:700;margin-bottom:10px;padding-bottom:6px;border-bottom:1px solid #eee}
.aop-up{border:2px dashed #ccc;border-radius:8px;padding:24px;text-align:center;cursor:pointer}
.aop-up:hover{border-color:#e8a840;background:#fffbf0}
.aop-up input{display:none}
.aop-tb{width:100%;border-collapse:collapse;font-size:12px;margin-top:8px}
.aop-tb th{background:#f5f5f0;padding:6px 8px;text-align:left;font-weight:600;border-bottom:2px solid #ddd}
.aop-tb td{padding:5px 8px;border-bottom:1px solid #eee;word-break:break-all}
.aop-log{background:#1a1a1a;color:#ddd;font-family:monospace;font-size:12px;padding:10px;border-radius:6px;max-height:260px;overflow-y:auto;white-space:pre-wrap;line-height:1.5}
.aop-btn{padding:10px 20px;border:none;border-radius:6px;font-size:14px;font-weight:600;cursor:pointer;width:100%}
.aop-go{background:#e8a840;color:#fff}.aop-go:hover{background:#d49530}
.aop-stop{background:#e74c3c;color:#fff}.aop-stop:hover{background:#c0392b}
.aop-res{margin-top:6px;padding:8px;border-radius:6px;font-size:13px}
.aop-ok{background:#d4edda;color:#155724}.aop-ng{background:#f8d7da;color:#721c24}
`;
document.head.appendChild(css);

var panel=document.createElement('div');
panel.id='aop-panel';
panel.innerHTML=`
<div class="aop-hd"><h3>\u{1F4E6} imos \u81EA\u52D5\u4E0B\u55AE</h3><button class="aop-x" id="aopClose">\u2715</button></div>
<div class="aop-bd">
 <div class="aop-c"><h4>\u{1F4C4} \u4E0A\u50B3\u8A02\u55AE Excel</h4>
  <div class="aop-up" id="aopDropZone"><p>\u9EDE\u64CA\u9078\u64C7 \u6216 \u62D6\u653E Excel \u6A94\u6848</p><p style="font-size:12px;color:#888;margin-top:4px">\u652F\u63F4\u7D93\u92B7\u5546\u4E0B\u55AE\u9801\u9762\u532F\u51FA\u7684 .xlsx</p><input type="file" id="aopFile" accept=".xlsx,.xls"></div>
  <div id="aopPreview" style="display:none"></div>
 </div>
 <div class="aop-c" id="aopCtrl" style="display:none"><h4>\u2699\uFE0F \u8A2D\u5B9A</h4>
  <div style="display:flex;align-items:center;gap:12px;margin-bottom:12px">
   <label>\u901F\u5EA6:</label>
   <select id="aopSpeed" style="padding:4px 8px;border:1px solid #ccc;border-radius:4px">
    <option value="1000">\u5FEB\u901F</option><option value="1500" selected>\u6B63\u5E38</option><option value="2500">\u6162\u901F</option>
   </select>
  </div>
  <button class="aop-btn aop-go" id="aopStart">\u25B6 \u958B\u59CB\u81EA\u52D5\u4E0B\u55AE</button>
 </div>
 <div class="aop-c" id="aopLogBox" style="display:none"><h4>\u{1F4CB} \u57F7\u884C\u8A18\u9304</h4>
  <div id="aopProg" style="margin-bottom:8px"><div style="background:#eee;border-radius:4px;height:8px"><div id="aopProgBar" style="background:#e8a840;height:8px;border-radius:4px;width:0%;transition:width .3s"></div></div><div id="aopProgTxt" style="font-size:12px;color:#888;margin-top:4px"></div></div>
  <div class="aop-log" id="aopLog"></div>
 </div>
 <div class="aop-c" id="aopResult" style="display:none"><h4>\u{1F4CA} \u7D50\u679C</h4><div id="aopResTxt"></div>
  <button class="aop-btn" id="aopCartBtn" style="background:#27ae60;color:#fff;margin-top:10px;display:none">\u{1F6D2} \u524D\u5F80\u8CFC\u7269\u8ECA\u78BA\u8A8D</button>
  <button class="aop-btn" id="aopReset" style="background:#6c757d;color:#fff;margin-top:8px">\u{1F504} \u91CD\u65B0\u4E0A\u50B3</button>
 </div>
</div>`;
document.body.appendChild(panel);

var $=function(id){return document.getElementById(id);};
function lg(msg,type){
  var log=$('aopLog');
  var t=new Date().toLocaleTimeString();
  var color=type==='ok'?'#4caf50':type==='e'?'#f44336':type==='w'?'#ff9800':'#ddd';
  log.innerHTML+='<span style="color:'+color+'">['+t+'] '+msg+'</span>\n';
  log.scrollTop=log.scrollHeight;
}
function prog(cur,total){
  var pct=total?Math.round(cur/total*100):0;
  $('aopProgBar').style.width=pct+'%';
  $('aopProgTxt').textContent=cur+'/'+total+' ('+pct+'%)';
}
function sleep(ms){return new Promise(function(r){setTimeout(r,ms)});}

function normName(n){
  return n.replace(/\s*\$\d+.*$/,'').replace(/\(預購中\)/g,'').replace(/\(延伸\)/g,'')
    .replace(/\([^)]*色\)/g,'').replace(/\([^)]*鋼[^)]*\)/g,'').replace(/\s*五顆/g,'')
    .replace(/\s+/g,' ').trim();
}

function extractKW(name){
  var m=name.match(/for\s+(.+?)(?:\s*\$|\s*\(|$)/i);
  if(m){
    var kw=m[1].replace(/\s*\$\d+.*$/,'').replace(/\([^)]*\)/g,'').trim();
    kw=kw.replace(/Samsung/gi,'SAMSUNG');
    return kw;
  }
  return name.replace(/\([^)]*\)/g,'').replace(/\$\d+.*$/,'').trim().substring(0,20);
}

// DOM search: use native imos form submit
function domSearch(keyword){
  return new Promise(function(resolve){
    var form=document.getElementById('form_listProductCondition');
    var resultDiv=document.getElementById('div_listProductResult');
    if(!form||!resultDiv){resolve([]);return;}
    form.attr1.value=keyword;
    resultDiv.innerHTML='';
    document.getElementById('btn_listProduct').click();
    var attempts=0;
    var check=setInterval(function(){
      attempts++;
      var rows=resultDiv.querySelectorAll('tr');
      if(rows.length>1||attempts>30){
        clearInterval(check);
        var products=[];
        rows.forEach(function(tr,i){
          if(i===0)return;
          var cells=tr.querySelectorAll('td');
          if(cells.length<6)return;
          var nameLink=cells[2]?cells[2].querySelector('a'):null;
          var qtyInput=tr.querySelector('input[id^="input_cartQty_"]');
          var addBtn=tr.querySelector('input[value="\u52A0\u5165\u8CFC\u7269\u8ECA"]');
          if(nameLink&&qtyInput&&addBtn){
            products.push({id:qtyInput.id.replace('input_cartQty_',''),name:nameLink.textContent.trim(),qtyInput:qtyInput,addBtn:addBtn});
          }
        });
        resolve(products);
      }
    },300);
  });
}

// DOM add to cart: set qty input + click button
function domAddToCart(product,qty){
  return new Promise(function(resolve){
    product.qtyInput.value=qty;
    product.addBtn.click();
    setTimeout(function(){resolve(true);},800);
  });
}

function findMatch(excelName,products){
  var en=normName(excelName).toLowerCase();
  // Tier 1: substring
  for(var i=0;i<products.length;i++){
    var pn=normName(products[i].name).toLowerCase();
    if(pn.indexOf(en)!==-1||en.indexOf(pn)!==-1)return products[i];
  }
  // Tier 2: token scoring
  var parts=en.split(/[\s,/]+/).filter(function(w){return w.length>1;});
  var best=null,bestScore=0;
  for(var i=0;i<products.length;i++){
    var pn=normName(products[i].name).toLowerCase();
    var score=0,total=0;
    for(var j=0;j<parts.length;j++){
      if(parts[j].length<=1)continue;
      total++;
      if(pn.indexOf(parts[j])!==-1)score++;
    }
    var pct=total?score/total:0;
    if(pct>bestScore){bestScore=pct;best=products[i];}
  }
  if(bestScore>=0.6)return best;
  return null;
}

function parseExcel(file){
  return new Promise(function(resolve,reject){
    var reader=new FileReader();
    reader.onload=function(e){
      try{
        var wb=XLSX.read(e.target.result,{type:'array'});
        var ws=wb.Sheets[wb.SheetNames[0]];
        var data=XLSX.utils.sheet_to_json(ws);
        var parsed=[];
        data.forEach(function(row){
          var name=row['\u5546\u54C1\u540D\u7A31']||'';
          var qty=parseInt(row['\u6578\u91CF'])||0;
          var kw=row['\u641C\u5C0B\u95DC\u9375\u5B57']||'';
          if(name&&qty>0)parsed.push({name:name.trim(),qty:qty,kw:kw.trim()});
        });
        resolve(parsed);
      }catch(err){reject(err);}
    };
    reader.onerror=function(){reject(new Error('\u8B80\u53D6\u5931\u6557'));};
    reader.readAsArrayBuffer(file);
  });
}

function showPreview(list){
  var html='<table class="aop-tb"><tr><th>#</th><th>\u5546\u54C1\u540D\u7A31</th><th>\u6578\u91CF</th><th>\u95DC\u9375\u5B57</th></tr>';
  list.forEach(function(it,i){
    html+='<tr><td>'+(i+1)+'</td><td>'+it.name.substring(0,50)+'</td><td>'+it.qty+'</td><td>'+it.kw+'</td></tr>';
  });
  html+='</table>';
  $('aopPreview').innerHTML=html;
  $('aopPreview').style.display='block';
  $('aopCtrl').style.display='block';
}

async function startAuto(){
  if(running)return;
  running=true;stopFlag=false;
  $('aopLogBox').style.display='block';
  $('aopResult').style.display='none';
  $('aopLog').innerHTML='';
  $('aopStart').textContent='\u23F9 \u505C\u6B62';
  $('aopStart').className='aop-btn aop-stop';

  var speed=parseInt($('aopSpeed').value)||1500;
  var okCount=0,ngCount=0,results=[];

  lg('\u958B\u59CB\u81EA\u52D5\u4E0B\u55AE\uFF0C\u5171 '+items.length+' \u7B46\u5546\u54C1');
  lg('\u901F\u5EA6: '+speed+'ms / \u7B46');

  var groups={};
  items.forEach(function(it){
    var kw=it.kw||extractKW(it.name);
    if(!groups[kw])groups[kw]=[];
    groups[kw].push(it);
  });
  var kwList=Object.keys(groups);
  lg('\u5171 '+kwList.length+' \u7D44\u95DC\u9375\u5B57');
  lg('\u2500'.repeat(30));

  var itemIdx=0;
  for(var g=0;g<kwList.length;g++){
    if(stopFlag)break;
    var kw=kwList[g];
    var gitems=groups[kw];
    lg('\u{1F50D} \u641C\u5C0B: "'+kw+'" ('+gitems.length+'\u9805)');

    var products=await domSearch(kw);
    lg('  \u2192 '+products.length+' \u9805\u7D50\u679C');
    await sleep(500);

    if(!products.length){
      var fbKw=extractKW(gitems[0].name);
      if(fbKw&&fbKw!==kw){
        lg('  \u26A0\uFE0F \u539F\u95DC\u9375\u5B57\u7121\u7D50\u679C\uFF0C\u5617\u8A66: "'+fbKw+'"','w');
        await sleep(500);
        products=await domSearch(fbKw);
        lg('  \u2192 '+products.length+' \u9805\u7D50\u679C');
      }
    }

    if(!products.length){
      gitems.forEach(function(it){
        itemIdx++;prog(itemIdx,items.length);
        lg('  \u274C '+it.name.substring(0,40)+' \u2192 \u641C\u5C0B\u7121\u7D50\u679C','e');
        ngCount++;results.push({name:it.name,status:'ng',reason:'\u641C\u5C0B\u7121\u7D50\u679C'});
      });
      await sleep(speed);
      continue;
    }

    for(var j=0;j<gitems.length;j++){
      if(stopFlag)break;
      itemIdx++;prog(itemIdx,items.length);
      var it=gitems[j];
      lg('  \u{1F4CC} '+it.name.substring(0,45));
      var match=findMatch(it.name,products);
      if(match){
        await domAddToCart(match,it.qty);
        lg('  \u2705 \u5DF2\u52A0\u5165 x'+it.qty+' (ID:'+match.id+')','ok');
        okCount++;results.push({name:it.name,status:'ok',pid:match.id});
      }else{
        lg('  \u274C \u7121\u5339\u914D\u5546\u54C1','e');
        ngCount++;results.push({name:it.name,status:'ng',reason:'\u7121\u5339\u914D\u5546\u54C1'});
      }
      await sleep(speed);
    }
  }

  lg('\u2500'.repeat(30));
  lg('\u5B8C\u6210\uFF01\u6210\u529F: '+okCount+' / \u5931\u6557: '+ngCount);

  $('aopResult').style.display='block';
  var rhtml='';
  results.forEach(function(r){
    if(r.status==='ok') rhtml+='<div class="aop-res aop-ok">\u2705 '+r.name.substring(0,50)+'</div>';
    else rhtml+='<div class="aop-res aop-ng">\u274C '+r.name.substring(0,50)+' \u2014 '+r.reason+'</div>';
  });
  $('aopResTxt').innerHTML=rhtml;
  if(okCount>0)$('aopCartBtn').style.display='block';

  $('aopStart').textContent='\u25B6 \u91CD\u65B0\u57F7\u884C';
  $('aopStart').className='aop-btn aop-go';
  running=false;
}

// Event bindings
$('aopClose').onclick=function(){$('aop-panel').style.display='none';};
$('aopDropZone').onclick=function(){$('aopFile').click();};
$('aopDropZone').ondragover=function(e){e.preventDefault();this.style.borderColor='#e8a840';};
$('aopDropZone').ondragleave=function(){this.style.borderColor='#ccc';};
$('aopDropZone').ondrop=function(e){
  e.preventDefault();this.style.borderColor='#ccc';
  if(e.dataTransfer.files.length)handleFile(e.dataTransfer.files[0]);
};
$('aopFile').onchange=function(){if(this.files.length)handleFile(this.files[0]);};

function handleFile(file){
  if(!file.name.match(/\.xlsx?$/i)){alert('\u8ACB\u4E0A\u50B3 .xlsx \u6A94\u6848');return;}
  parseExcel(file).then(function(parsed){
    if(!parsed.length){alert('\u672A\u627E\u5230\u6709\u6548\u5546\u54C1');return;}
    items=parsed;
    showPreview(items);
  }).catch(function(e){alert('\u89E3\u6790\u5931\u6557: '+e.message);});
}

$('aopStart').onclick=function(){
  if(running){stopFlag=true;lg('\u23F9 \u4F7F\u7528\u8005\u505C\u6B62','w');return;}
  if(!items.length){alert('\u8ACB\u5148\u4E0A\u50B3 Excel');return;}
  startAuto();
};

$('aopCartBtn').onclick=function(){window.location.href='/frontend/initFastOrder.shtml';};

$('aopReset').onclick=function(){
  items=[];running=false;stopFlag=false;
  $('aopPreview').style.display='none';
  $('aopCtrl').style.display='none';
  $('aopLogBox').style.display='none';
  $('aopResult').style.display='none';
  $('aopFile').value='';
  $('aopStart').textContent='\u25B6 \u958B\u59CB\u81EA\u52D5\u4E0B\u55AE';
  $('aopStart').className='aop-btn aop-go';
};

if(typeof XLSX==='undefined'){
  var s=document.createElement('script');
  s.src='https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js';
  s.onerror=function(){alert('SheetJS CDN \u8F09\u5165\u5931\u6557');};
  document.head.appendChild(s);
}

console.log('[imos Auto Order v2] DOM-based, ready');
})();
