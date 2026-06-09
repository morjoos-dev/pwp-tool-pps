let appConfig=null;
const DEFAULT_PROTECTED_TERMS=["PDF","Technical ID","ID"];
const DEFAULT_FIXED_TRANSLATIONS={"Datos generales":"General data","Nombre":"Name","Texto de ayuda":"Help text","Visibilidad":"Visibility","Concepto de negocio":"Business concept","Datos avanzados":"Advanced data","Datos campos":"Field data"};
const MAX_DEEP_LEVEL=8;

window.pwpShowTab=function(targetId){
  document.querySelectorAll(".tab-button").forEach(b=>b.classList.remove("active"));
  document.querySelectorAll(".tab-panel").forEach(p=>p.classList.remove("active"));
  const panel=document.getElementById(targetId);
  const btn=document.querySelector(`[data-target="${targetId}"]`);
  if(panel) panel.classList.add("active");
  if(btn) btn.classList.add("active");
  try{sessionStorage.setItem("pwpToolActiveTab",targetId);}catch(_){ }
};

Office.onReady((info)=>{
  if(info.host===Office.HostType.PowerPoint){
    bindEvents();restoreSessionUrl();restoreTranslatorSettings();restoreActiveTab();log("PwP Tool PPS cargado correctamente.");
  }else{log("Este complemento está pensado para PowerPoint.");}
});

function bindEvents(){
  document.querySelectorAll(".tab-button").forEach(btn=>btn.addEventListener("click",()=>window.pwpShowTab(btn.dataset.target)));
  document.getElementById("loadConfigFromUrl").onclick=loadConfigFromUrlInput;
  document.getElementById("clearConfigUrl").onclick=clearConfigUrl;
  document.getElementById("configFile").onchange=loadConfigFromFileInput;
  document.getElementById("clearLog").onclick=clearLog;
  document.getElementById("formatBoldByColon").onclick=formatBoldByColonAllSlides;
  document.getElementById("formatCurrentSlide").onclick=formatBoldByColonCurrentSlide;
  document.getElementById("testTranslator").onclick=testTranslator;
  document.getElementById("translateCurrentSlide").onclick=translateCurrentSlide;
  document.getElementById("translateAllSlides").onclick=translateAllSlides;
  ["translatorEndpoint","sourceLang","targetLang","maxTexts","delayMs"].forEach(id=>document.getElementById(id).addEventListener("change",saveTranslatorSettings));
}
function restoreActiveTab(){try{const t=sessionStorage.getItem("pwpToolActiveTab");if(t) window.pwpShowTab(t);}catch(_){}}
function log(m){const box=document.getElementById("log");const line=`[${new Date().toLocaleTimeString()}] ${m}`;if(box){box.textContent+=line+"\n";box.scrollTop=box.scrollHeight;}console.log(line)}
function clearLog(){document.getElementById("log").textContent=""}
function sleep(ms){return new Promise(r=>setTimeout(r,ms))}
function restoreSessionUrl(){const u=sessionStorage.getItem("pwpToolConfigUrl");if(u){document.getElementById("privateConfigUrl").value=u;log("URL restaurada desde la sesión.")}}
function saveTranslatorSettings(){sessionStorage.setItem("pwpToolTranslatorSettings",JSON.stringify(getTranslatorSettings()))}
function restoreTranslatorSettings(){const raw=sessionStorage.getItem("pwpToolTranslatorSettings");if(!raw)return;try{const s=JSON.parse(raw);if(s.endpoint)translatorEndpoint.value=s.endpoint;if(s.source)sourceLang.value=s.source;if(s.target)targetLang.value=s.target;if(s.maxTexts)maxTexts.value=s.maxTexts;if(s.delayMs)delayMs.value=s.delayMs}catch(_){}}
function getTranslatorSettings(){return{endpoint:translatorEndpoint.value.trim().replace(/\/$/,""),source:sourceLang.value.trim()||"auto",target:targetLang.value.trim()||"en",apiKey:apiKey.value.trim(),maxTexts:Math.max(1,parseInt(maxTexts.value||"10",10)),delayMs:Math.max(0,parseInt(delayMs.value||"700",10))}}

async function loadConfigFromUrlInput(){const url=privateConfigUrl.value.trim();if(!url){log("Introduce una URL privada de configuración.");return}try{sessionStorage.setItem("pwpToolConfigUrl",url);log("Intentando cargar configuración desde URL...");appConfig=await loadJsonFromUrl(url);log("Configuración cargada correctamente.");log(JSON.stringify(appConfig,null,2))}catch(e){log("Error cargando configuración desde URL.");log(e.message||String(e));log("Si la URL devuelve visor HTML o está bloqueada, usa archivo JSON.")}}
function clearConfigUrl(){sessionStorage.removeItem("pwpToolConfigUrl");privateConfigUrl.value="";log("URL olvidada de la sesión.")}
async function loadJsonFromUrl(url){const res=await fetch(url,{method:"GET",credentials:"include"});if(!res.ok)throw new Error(`HTTP ${res.status}: no se pudo cargar el archivo.`);const txt=await res.text();try{return JSON.parse(txt)}catch{throw new Error("La respuesta no es JSON válido.")}}
function loadConfigFromFileInput(ev){const file=ev.target.files[0];if(!file)return;const reader=new FileReader();reader.onload=()=>{try{appConfig=JSON.parse(reader.result);log("Configuración cargada desde archivo local.");log(JSON.stringify(appConfig,null,2))}catch(e){log("Error leyendo el JSON seleccionado.");log(e.message||String(e))}};reader.onerror=()=>log("No se pudo leer el archivo seleccionado.");reader.readAsText(file)}

async function testTranslator(){const s=getTranslatorSettings();saveTranslatorSettings();try{log("Probando traductor...");const t=await translateText("Prueba de conexión",s);log(`Traductor OK: ${t}`)}catch(e){log("Error probando traductor.");log(e.message||String(e))}}

async function translateCurrentSlide(){log("Iniciando traducción profunda de diapositiva actual...");const s=getTranslatorSettings();saveTranslatorSettings();await PowerPoint.run(async ctx=>{let selected;try{selected=ctx.presentation.getSelectedSlides();selected.load("items");await ctx.sync()}catch(e){log("No se pudo obtener diapositiva actual. Usa traducir presentación con Máx. elementos bajo.");return}if(!selected.items||!selected.items.length){log("No hay diapositiva seleccionada.");return}const slide=selected.items[0];const r=await translateSlideDeep(ctx,slide,"actual",s,s.maxTexts);await ctx.sync();log(`Diapositiva actual traducida: ${r.translatedCount} elemento/s. Detectados: ${r.detectedCount}.`)})}
async function translateAllSlides(){log("Iniciando traducción profunda de presentación...");const s=getTranslatorSettings();saveTranslatorSettings();await PowerPoint.run(async ctx=>{const slides=ctx.presentation.slides;slides.load("items");await ctx.sync();let processed=0,translated=0,detected=0;for(let i=0;i<slides.items.length;i++){if(processed>=s.maxTexts){log(`Límite alcanzado: ${s.maxTexts} elemento/s.`);break}const r=await translateSlideDeep(ctx,slides.items[i],i+1,s,s.maxTexts-processed);processed+=r.processedCount;translated+=r.translatedCount;detected+=r.detectedCount;await ctx.sync();log(`OK diapositiva ${i+1}: ${r.translatedCount}/${r.detectedCount} elemento/s.`)}log(`Traducción finalizada. Traducidos: ${translated}. Detectados: ${detected}.`)})}
async function translateSlideDeep(ctx,slide,slideNumber,s,remainingLimit){const targets=[];await collectTextTargetsFromShapeCollection(ctx,slide.shapes,targets,`slide ${slideNumber}`,0);await ctx.sync();let processedCount=0,translatedCount=0;const limit=remainingLimit||s.maxTexts;for(const item of targets){if(processedCount>=limit)break;const original=getTargetText(item);if(!shouldTranslate(original))continue;try{log(`Traduciendo ${item.path}...`);const translated=await translatePreservingStructure(normalizePowerPointText(original),s);if(translated&&translated.trim()){setTargetText(item,translated);translatedCount++}processedCount++;await ctx.sync();await sleep(s.delayMs)}catch(e){log(`Error traduciendo ${item.path}.`);log(e.message||String(e));break}}return{processedCount,translatedCount,detectedCount:targets.length}}

async function collectTextTargetsFromShapeCollection(ctx,shapeCollection,targets,path,depth){
  if(depth>MAX_DEEP_LEVEL){log(`Profundidad máxima alcanzada en ${path}`);return}
  try{shapeCollection.load("items/id,items/name,items/type");await ctx.sync()}catch(e){log(`No se pudo leer colección: ${path}`);return}
  for(const shape of shapeCollection.items){
    const shapeName=safeShapeName(shape);
    const shapePath=`${path} > ${shapeName}`;
    try{const tf=shape.getTextFrameOrNullObject();tf.load("isNullObject,hasText,textRange/text");targets.push({kind:"textFrame",textFrame:tf,path:shapePath})}catch(_){ }
    if(isGroupShape(shape)){
      try{log(`Grupo detectado: ${shapePath}`);await collectTextTargetsFromShapeCollection(ctx,shape.group.shapes,targets,`${shapePath} > group`,depth+1)}catch(e){log(`No se pudo entrar en grupo: ${shapePath}`)}
    }
    if(isTableShape(shape)){
      try{await collectTableTargets(ctx,shape,targets,shapePath)}catch(e){log(`No se pudo leer tabla: ${shapePath}`)}
    }
  }
  try{await ctx.sync()}catch(_){ }
}
function safeShapeName(shape){try{return shape.name||shape.id||"shape"}catch{return"shape"}}
function shapeTypeText(shape){try{return String(shape.type||"").toLowerCase()}catch{return""}}
function isGroupShape(shape){const t=shapeTypeText(shape);return t.includes("group")}
function isTableShape(shape){const t=shapeTypeText(shape);return t.includes("table")}
async function collectTableTargets(ctx,shape,targets,path){
  if(typeof shape.getTable!=="function")return;
  const table=shape.getTable();table.load("rowCount,columnCount");await ctx.sync();
  for(let r=0;r<table.rowCount;r++)for(let c=0;c<table.columnCount;c++){const cell=table.getCellOrNullObject(r,c);cell.load("isNullObject,text");targets.push({kind:"tableCell",cell,path:`${path} > cell ${r+1},${c+1}`})}
}
function getTargetText(item){try{if(item.kind==="textFrame"){if(item.textFrame.isNullObject||!item.textFrame.hasText)return"";return item.textFrame.textRange.text||""}if(item.kind==="tableCell"){if(item.cell.isNullObject)return"";return item.cell.text||""}}catch{return""}return""}
function setTargetText(item,text){if(item.kind==="textFrame")item.textFrame.textRange.text=text;else if(item.kind==="tableCell")item.cell.text=text}

async function translatePreservingStructure(text,s){const parts=text.split(/(\n+)/);const out=[];for(const part of parts){if(/^\n+$/.test(part)||!part.trim()){out.push(part);continue}const leading=part.match(/^\s*/)[0],trailing=part.match(/\s*$/)[0],clean=part.trim();const fixed=getFixedTranslation(clean);if(fixed){out.push(leading+fixed+trailing);continue}out.push(leading+await translateLabelValueOrText(clean,s)+trailing);await sleep(150)}return out.join("")}
async function translateLabelValueOrText(line,s){const idx=line.indexOf(":");if(idx>0&&idx<=50){const label=line.slice(0,idx).trim(),value=line.slice(idx+1).trim();const fixedLabel=getFixedTranslation(label);if(fixedLabel){if(!value)return `${fixedLabel}:`;return `${fixedLabel}: ${await translateProtectedText(value,s)}`}}return translateProtectedText(line,s)}
async function translateProtectedText(text,s){if(!shouldTranslate(text))return text;const fixed=getFixedTranslation(text);if(fixed)return fixed;const pr=protectTerms(text);const chunks=splitTextIntoChunks(pr.text,900);const out=[];for(const c of chunks){out.push(await translateText(c,s));await sleep(200)}return restoreProtectedTerms(out.join(" "),pr.map)}
async function translateText(text,s){const body={q:text,source:s.source,target:s.target,format:"text"};if(s.apiKey)body.api_key=s.apiKey;const res=await fetch(`${s.endpoint}/translate`,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(body)});const raw=await res.text();if(!res.ok)throw new Error(`HTTP ${res.status}: ${raw}`);let data;try{data=JSON.parse(raw)}catch{throw new Error(`Respuesta no JSON: ${raw}`)}if(!data.translatedText)throw new Error(`Respuesta inesperada: ${JSON.stringify(data)}`);return data.translatedText}
function getProtectedTerms(){return(appConfig&&Array.isArray(appConfig.protectedTerms))?appConfig.protectedTerms:DEFAULT_PROTECTED_TERMS}
function getFixedTranslations(){return(appConfig&&appConfig.fixedTranslations)?appConfig.fixedTranslations:DEFAULT_FIXED_TRANSLATIONS}
function getFixedTranslation(text){const clean=text.replace(/\s+/g," ").trim();return getFixedTranslations()[clean]||null}
function protectTerms(text){let result=text;const map={};[...getProtectedTerms()].sort((a,b)=>b.length-a.length).forEach((term,i)=>{const token=`__PWP_TERM_${i}__`;const re=new RegExp(escapeRegExp(term),"g");if(re.test(result)){result=result.replace(re,token);map[token]=term}});return{text:result,map}}
function restoreProtectedTerms(text,map){let result=text;Object.keys(map).forEach(token=>{result=result.replace(new RegExp(escapeRegExp(token),"g"),map[token])});return result}
function escapeRegExp(text){return text.replace(/[.*+?^${}()|[\]\\]/g,"\\$&")}
function normalizePowerPointText(text){return(text||"").replace(/\u000B/g,"\n").replace(/\r\n/g,"\n").replace(/\r/g,"\n").replace(/&nbsp;/g," ").replace(/&gt;/g,">").replace(/&lt;/g,"<").replace(/&amp;/g,"&").replace(/&quot;/g,"\"").trim()}
function shouldTranslate(text){if(!text||!text.trim()||text.trim().length<2)return false;if(/^[\d\s.,:;\/\-–—()]+$/.test(text.trim()))return false;if(/^https?:\/\//i.test(text.trim()))return false;return true}
function splitTextIntoChunks(text,maxLength){const clean=text.trim();if(clean.length<=maxLength)return[clean];const sentences=clean.match(/[^.!?;:]+[.!?;:]?/g)||[clean];const chunks=[];let current="";for(const sentence of sentences){const ss=sentence.trim();const candidate=current?`${current} ${ss}`:ss;if(candidate.length<=maxLength)current=candidate;else{if(current)chunks.push(current);current=ss}}if(current)chunks.push(current);return chunks}

async function formatBoldByColonAllSlides(){log("Iniciando formateo profundo de toda la presentación...");await PowerPoint.run(async ctx=>{const slides=ctx.presentation.slides;slides.load("items");await ctx.sync();let fs=0,fr=0,detected=0;for(let i=0;i<slides.items.length;i++){const r=await formatSlideDeep(ctx,slides.items[i],i+1);fs+=r.formattedShapes;fr+=r.formattedRanges;detected+=r.detectedCount;await ctx.sync();log(`OK diapositiva ${i+1}: ${r.formattedShapes}/${r.detectedCount} caja/s, ${r.formattedRanges} rango/s.`)}log(`Formateo terminado. Cajas: ${fs}. Rangos: ${fr}. Detectadas: ${detected}.`)})}
async function formatBoldByColonCurrentSlide(){log("Intentando formateo profundo de diapositiva actual...");await PowerPoint.run(async ctx=>{const selected=ctx.presentation.getSelectedSlides();selected.load("items");await ctx.sync();if(!selected.items||!selected.items.length){log("No hay diapositiva seleccionada.");return}const r=await formatSlideDeep(ctx,selected.items[0],"actual");await ctx.sync();log(`Diapositiva actual formateada: ${r.formattedShapes}/${r.detectedCount} caja/s, ${r.formattedRanges} rango/s.`)})}
async function formatSlideDeep(ctx,slide,slideNumber){const targets=[];await collectTextTargetsFromShapeCollection(ctx,slide.shapes,targets,`slide ${slideNumber}`,0);await ctx.sync();let formattedShapes=0,formattedRanges=0;for(const item of targets){if(item.kind!=="textFrame")continue;const text=getTargetText(item);if(!text.trim())continue;try{const tr=item.textFrame.textRange;tr.font.bold=false;for(const rg of getBoldRangesByColonAndLineBreak(text)){const sub=tr.getSubstring(rg.start,rg.length);sub.font.bold=true;formattedRanges++}formattedShapes++}catch(e){log(`No se pudo formatear ${item.path}`)}}return{formattedShapes,formattedRanges,detectedCount:targets.length}}
function getBoldRangesByColonAndLineBreak(text){const ranges=[];let lineStart=0;while(lineStart<text.length){const info=findNextLineEnd(text,lineStart);const line=text.slice(lineStart,info.index);processLineForBoldRanges(line,lineStart,ranges);lineStart=info.index+info.separatorLength;if(info.separatorLength===0&&info.index>=text.length)break}return ranges}
function processLineForBoldRanges(line,abs,ranges){if(!line||!line.trim())return;if(/^[\d\s.,:;\/\-–—()]+$/.test(line.trim()))return;const first=findFirstNonSpaceIndex(line),last=findLastNonSpaceIndex(line);if(first===-1||last===-1)return;const colon=line.indexOf(":");if(colon>=0)ranges.push({start:abs+first,length:colon-first+1});else ranges.push({start:abs+first,length:last-first+1})}
function findNextLineEnd(text,startIndex){let nearestIndex=text.length,separatorLength=0;for(let i=startIndex;i<text.length;i++){const ch=text[i];if(ch==="\r"){nearestIndex=i;separatorLength=text[i+1]==="\n"?2:1;break}if(ch==="\n"||ch==="\u000B"){nearestIndex=i;separatorLength=1;break}}return{index:nearestIndex,separatorLength}}
function findFirstNonSpaceIndex(text){for(let i=0;i<text.length;i++)if(!/\s/.test(text[i]))return i;return-1}
function findLastNonSpaceIndex(text){for(let i=text.length-1;i>=0;i--)if(!/\s/.test(text[i]))return i;return-1}
