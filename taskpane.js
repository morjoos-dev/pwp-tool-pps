let appConfig = null;

const DEFAULT_PROTECTED_TERMS = ["POPS", "DAR", "PDF", "Technical ID", "ID"];
const DEFAULT_FIXED_TRANSLATIONS = {
  "Datos generales": "General data",
  "Nombre": "Name",
  "Texto de ayuda": "Help text",
  "Visibilidad": "Visibility",
  "Concepto de negocio": "Business concept",
  "Datos avanzados": "Advanced data",
  "Datos campos": "Field data"
};

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    bindEvents();
    restoreSessionUrl();
    restoreTranslatorSettings();
    log("PwP Tool PPS cargado correctamente.");
  } else {
    log("Este complemento está pensado para PowerPoint.");
  }
});

function bindEvents() {
  bindTabs();

  document.getElementById("loadConfigFromUrl").onclick = loadConfigFromUrlInput;
  document.getElementById("clearConfigUrl").onclick = clearConfigUrl;
  document.getElementById("configFile").onchange = loadConfigFromFileInput;
  document.getElementById("clearLog").onclick = clearLog;

  document.getElementById("formatBoldByColon").onclick = formatBoldByColonAllSlides;
  document.getElementById("formatCurrentSlide").onclick = formatBoldByColonCurrentSlide;

  document.getElementById("testTranslator").onclick = testTranslator;
  document.getElementById("translateCurrentSlide").onclick = translateCurrentSlide;
  document.getElementById("translateAllSlides").onclick = translateAllSlides;

  ["translatorEndpoint", "sourceLang", "targetLang", "maxTexts", "delayMs"].forEach((id) => {
    document.getElementById(id).addEventListener("change", saveTranslatorSettings);
  });
}

function bindTabs() {
  const buttons = document.querySelectorAll(".tab-button");
  const panels = document.querySelectorAll(".tab-panel");

  buttons.forEach((button) => {
    button.addEventListener("click", () => {
      const targetId = button.getAttribute("data-target");
      buttons.forEach((item) => item.classList.remove("active"));
      panels.forEach((panel) => panel.classList.remove("active"));
      button.classList.add("active");
      const target = document.getElementById(targetId);
      if (target) target.classList.add("active");
      sessionStorage.setItem("pwpToolActiveTab", targetId);
    });
  });

  const savedTab = sessionStorage.getItem("pwpToolActiveTab");
  if (savedTab) {
    const savedButton = document.querySelector(`[data-target="${savedTab}"]`);
    if (savedButton) savedButton.click();
  }
}

function log(message) {
  const logBox = document.getElementById("log");
  const line = `[${new Date().toLocaleTimeString()}] ${message}`;
  if (logBox) {
    logBox.textContent += line + "\n";
    logBox.scrollTop = logBox.scrollHeight;
  }
  console.log(line);
}

function clearLog() {
  document.getElementById("log").textContent = "";
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function restoreSessionUrl() {
  const savedUrl = sessionStorage.getItem("pwpToolConfigUrl");
  if (savedUrl) {
    document.getElementById("privateConfigUrl").value = savedUrl;
    log("URL restaurada desde la sesión.");
  }
}

function saveTranslatorSettings() {
  const settings = getTranslatorSettings();
  sessionStorage.setItem("pwpToolTranslatorSettings", JSON.stringify(settings));
}

function restoreTranslatorSettings() {
  const raw = sessionStorage.getItem("pwpToolTranslatorSettings");
  if (!raw) return;
  try {
    const s = JSON.parse(raw);
    if (s.endpoint) document.getElementById("translatorEndpoint").value = s.endpoint;
    if (s.source) document.getElementById("sourceLang").value = s.source;
    if (s.target) document.getElementById("targetLang").value = s.target;
    if (s.maxTexts) document.getElementById("maxTexts").value = s.maxTexts;
    if (s.delayMs) document.getElementById("delayMs").value = s.delayMs;
  } catch (_) {}
}

function getTranslatorSettings() {
  return {
    endpoint: document.getElementById("translatorEndpoint").value.trim().replace(/\/$/, ""),
    source: document.getElementById("sourceLang").value.trim() || "auto",
    target: document.getElementById("targetLang").value.trim() || "en",
    apiKey: document.getElementById("apiKey").value.trim(),
    maxTexts: Math.max(1, parseInt(document.getElementById("maxTexts").value || "10", 10)),
    delayMs: Math.max(0, parseInt(document.getElementById("delayMs").value || "700", 10))
  };
}

async function loadConfigFromUrlInput() {
  const url = document.getElementById("privateConfigUrl").value.trim();
  if (!url) {
    log("Introduce una URL privada de configuración.");
    return;
  }
  try {
    sessionStorage.setItem("pwpToolConfigUrl", url);
    log("Intentando cargar configuración desde URL...");
    appConfig = await loadJsonFromUrl(url);
    log("Configuración cargada correctamente.");
    log(JSON.stringify(appConfig, null, 2));
  } catch (error) {
    log("Error cargando configuración desde URL.");
    log(error.message || String(error));
    log("Si la URL devuelve visor HTML o está bloqueada, usa archivo JSON.");
  }
}

function clearConfigUrl() {
  sessionStorage.removeItem("pwpToolConfigUrl");
  document.getElementById("privateConfigUrl").value = "";
  log("URL olvidada de la sesión.");
}

async function loadJsonFromUrl(url) {
  const response = await fetch(url, { method: "GET", credentials: "include" });
  if (!response.ok) throw new Error(`HTTP ${response.status}: no se pudo cargar el archivo.`);
  const text = await response.text();
  try {
    return JSON.parse(text);
  } catch {
    throw new Error("La respuesta no es JSON válido.");
  }
}

function loadConfigFromFileInput(event) {
  const file = event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = () => {
    try {
      appConfig = JSON.parse(reader.result);
      log("Configuración cargada desde archivo local.");
      log(JSON.stringify(appConfig, null, 2));
    } catch (error) {
      log("Error leyendo el JSON seleccionado.");
      log(error.message || String(error));
    }
  };
  reader.onerror = () => log("No se pudo leer el archivo seleccionado.");
  reader.readAsText(file);
}

async function testTranslator() {
  const s = getTranslatorSettings();
  saveTranslatorSettings();
  try {
    log("Probando traductor...");
    const translated = await translateText("Prueba de conexión", s);
    log(`Traductor OK: ${translated}`);
  } catch (error) {
    log("Error probando traductor.");
    log(error.message || String(error));
  }
}

async function translateCurrentSlide() {
  log("Iniciando traducción de diapositiva actual...");
  const settings = getTranslatorSettings();
  saveTranslatorSettings();
  await PowerPoint.run(async (context) => {
    let selectedSlides;
    try {
      selectedSlides = context.presentation.getSelectedSlides();
      selectedSlides.load("items");
      await context.sync();
    } catch (error) {
      log("No se pudo obtener diapositiva actual. Usa traducir presentación con Máx. cajas bajo.");
      return;
    }
    if (!selectedSlides.items || selectedSlides.items.length === 0) {
      log("No hay diapositiva seleccionada.");
      return;
    }
    const slide = selectedSlides.items[0];
    slide.shapes.load("items");
    await context.sync();
    const result = await translateSlideShapes(context, slide, "actual", settings);
    await context.sync();
    log(`Diapositiva actual traducida: ${result.translatedCount} caja/s.`);
  });
}

async function translateAllSlides() {
  log("Iniciando traducción de presentación...");
  const settings = getTranslatorSettings();
  saveTranslatorSettings();
  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();
    for (const slide of slides.items) slide.shapes.load("items");
    await context.sync();
    let processed = 0;
    let translated = 0;
    for (let i = 0; i < slides.items.length; i++) {
      if (processed >= settings.maxTexts) {
        log(`Límite alcanzado: ${settings.maxTexts} caja/s.`);
        break;
      }
      const result = await translateSlideShapes(context, slides.items[i], i + 1, settings, settings.maxTexts - processed);
      processed += result.processedCount;
      translated += result.translatedCount;
      await context.sync();
      log(`OK diapositiva ${i + 1}: ${result.translatedCount} caja/s traducida/s.`);
    }
    log(`Traducción finalizada. Cajas traducidas: ${translated}`);
  });
}

async function translateSlideShapes(context, slide, slideNumber, settings, remainingLimit) {
  const targets = [];
  for (const shape of slide.shapes.items) {
    const textFrame = shape.getTextFrameOrNullObject();
    textFrame.load("isNullObject,hasText,textRange/text");
    targets.push({ textFrame });
  }
  await context.sync();
  let processedCount = 0;
  let translatedCount = 0;
  const limit = remainingLimit || settings.maxTexts;
  for (const item of targets) {
    if (processedCount >= limit) break;
    if (item.textFrame.isNullObject || !item.textFrame.hasText) continue;
    const original = normalizePowerPointText(item.textFrame.textRange.text || "");
    if (!shouldTranslate(original)) continue;
    try {
      log(`Traduciendo diapositiva ${slideNumber}, caja ${processedCount + 1}...`);
      const translated = await translatePreservingStructure(original, settings);
      if (translated && translated.trim()) {
        item.textFrame.textRange.text = translated;
        translatedCount++;
      }
      processedCount++;
      await context.sync();
      await sleep(settings.delayMs);
    } catch (error) {
      log(`Error traduciendo caja en diapositiva ${slideNumber}.`);
      log(error.message || String(error));
      break;
    }
  }
  return { processedCount, translatedCount };
}

async function translatePreservingStructure(text, settings) {
  const parts = text.split(/(\n+)/);
  const output = [];
  for (const part of parts) {
    if (/^\n+$/.test(part) || !part.trim()) {
      output.push(part);
      continue;
    }
    const leading = part.match(/^\s*/)[0];
    const trailing = part.match(/\s*$/)[0];
    const clean = part.trim();
    const fixed = getFixedTranslation(clean);
    if (fixed) {
      output.push(leading + fixed + trailing);
      continue;
    }
    output.push(leading + await translateLabelValueOrText(clean, settings) + trailing);
    await sleep(150);
  }
  return output.join("");
}

async function translateLabelValueOrText(line, settings) {
  const colonIndex = line.indexOf(":");
  if (colonIndex > 0 && colonIndex <= 50) {
    const label = line.slice(0, colonIndex).trim();
    const value = line.slice(colonIndex + 1).trim();
    const fixedLabel = getFixedTranslation(label);
    if (fixedLabel) {
      if (!value) return `${fixedLabel}:`;
      return `${fixedLabel}: ${await translateProtectedText(value, settings)}`;
    }
  }
  return translateProtectedText(line, settings);
}

async function translateProtectedText(text, settings) {
  if (!shouldTranslate(text)) return text;
  const fixed = getFixedTranslation(text);
  if (fixed) return fixed;
  const protectedResult = protectTerms(text);
  const chunks = splitTextIntoChunks(protectedResult.text, 900);
  const translatedChunks = [];
  for (const chunk of chunks) {
    translatedChunks.push(await translateText(chunk, settings));
    await sleep(200);
  }
  return restoreProtectedTerms(translatedChunks.join(" "), protectedResult.map);
}

async function translateText(text, settings) {
  const body = { q: text, source: settings.source, target: settings.target, format: "text" };
  if (settings.apiKey) body.api_key = settings.apiKey;
  const response = await fetch(`${settings.endpoint}/translate`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body)
  });
  const raw = await response.text();
  if (!response.ok) throw new Error(`HTTP ${response.status}: ${raw}`);
  let data;
  try {
    data = JSON.parse(raw);
  } catch {
    throw new Error(`Respuesta no JSON: ${raw}`);
  }
  if (!data.translatedText) throw new Error(`Respuesta inesperada: ${JSON.stringify(data)}`);
  return data.translatedText;
}

function getProtectedTerms() {
  return (appConfig && Array.isArray(appConfig.protectedTerms)) ? appConfig.protectedTerms : DEFAULT_PROTECTED_TERMS;
}
function getFixedTranslations() {
  return (appConfig && appConfig.fixedTranslations) ? appConfig.fixedTranslations : DEFAULT_FIXED_TRANSLATIONS;
}
function getFixedTranslation(text) {
  const clean = text.replace(/\s+/g, " ").trim();
  return getFixedTranslations()[clean] || null;
}
function protectTerms(text) {
  let result = text;
  const map = {};
  [...getProtectedTerms()].sort((a, b) => b.length - a.length).forEach((term, index) => {
    const token = `__PWP_TERM_${index}__`;
    const regex = new RegExp(escapeRegExp(term), "g");
    if (regex.test(result)) {
      result = result.replace(regex, token);
      map[token] = term;
    }
  });
  return { text: result, map };
}
function restoreProtectedTerms(text, map) {
  let result = text;
  Object.keys(map).forEach((token) => {
    result = result.replace(new RegExp(escapeRegExp(token), "g"), map[token]);
  });
  return result;
}
function escapeRegExp(text) {
  return text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
function normalizePowerPointText(text) {
  return (text || "")
    .replace(/\u000B/g, "\n")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .replace(/&nbsp;/g, " ")
    .replace(/&gt;/g, ">")
    .replace(/&lt;/g, "<")
    .replace(/&amp;/g, "&")
    .replace(/&quot;/g, "\"")
    .trim();
}
function shouldTranslate(text) {
  if (!text || !text.trim() || text.trim().length < 2) return false;
  if (/^[\d\s.,:;\/\-–—()]+$/.test(text.trim())) return false;
  if (/^https?:\/\//i.test(text.trim())) return false;
  return true;
}
function splitTextIntoChunks(text, maxLength) {
  const clean = text.trim();
  if (clean.length <= maxLength) return [clean];
  const sentences = clean.match(/[^.!?;:]+[.!?;:]?/g) || [clean];
  const chunks = [];
  let current = "";
  for (const sentence of sentences) {
    const s = sentence.trim();
    const candidate = current ? `${current} ${s}` : s;
    if (candidate.length <= maxLength) current = candidate;
    else {
      if (current) chunks.push(current);
      current = s;
    }
  }
  if (current) chunks.push(current);
  return chunks;
}

async function formatBoldByColonAllSlides() {
  log("Iniciando formateo de toda la presentación...");
  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();
    for (const slide of slides.items) slide.shapes.load("items");
    await context.sync();
    let formattedShapes = 0;
    let formattedRanges = 0;
    for (let i = 0; i < slides.items.length; i++) {
      const result = await formatSlideShapes(context, slides.items[i], i + 1);
      formattedShapes += result.formattedShapes;
      formattedRanges += result.formattedRanges;
      await context.sync();
      log(`OK diapositiva ${i + 1}: ${result.formattedShapes} caja/s, ${result.formattedRanges} rango/s.`);
    }
    log(`Formateo terminado. Cajas: ${formattedShapes}. Rangos: ${formattedRanges}.`);
  });
}

async function formatBoldByColonCurrentSlide() {
  log("Intentando formatear diapositiva actual...");
  await PowerPoint.run(async (context) => {
    const selectedSlides = context.presentation.getSelectedSlides();
    selectedSlides.load("items");
    await context.sync();
    if (!selectedSlides.items || selectedSlides.items.length === 0) {
      log("No hay diapositiva seleccionada.");
      return;
    }
    const slide = selectedSlides.items[0];
    slide.shapes.load("items");
    await context.sync();
    const result = await formatSlideShapes(context, slide, "actual");
    await context.sync();
    log(`Diapositiva actual formateada: ${result.formattedShapes} caja/s, ${result.formattedRanges} rango/s.`);
  });
}

async function formatSlideShapes(context, slide, slideNumber) {
  const targets = [];
  for (const shape of slide.shapes.items) {
    const textFrame = shape.getTextFrameOrNullObject();
    textFrame.load("isNullObject,hasText,textRange/text");
    targets.push({ textFrame });
  }
  await context.sync();
  let formattedShapes = 0;
  let formattedRanges = 0;
  for (const item of targets) {
    if (item.textFrame.isNullObject || !item.textFrame.hasText) continue;
    const textRange = item.textFrame.textRange;
    const text = textRange.text || "";
    if (!text.trim()) continue;
    textRange.font.bold = false;
    const boldRanges = getBoldRangesByColonAndLineBreak(text);
    for (const rangeInfo of boldRanges) {
      const subRange = textRange.getSubstring(rangeInfo.start, rangeInfo.length);
      subRange.font.bold = true;
      formattedRanges++;
    }
    formattedShapes++;
  }
  return { formattedShapes, formattedRanges };
}

function getBoldRangesByColonAndLineBreak(text) {
  const ranges = [];
  let lineStart = 0;
  while (lineStart < text.length) {
    const info = findNextLineEnd(text, lineStart);
    const line = text.slice(lineStart, info.index);
    processLineForBoldRanges(line, lineStart, ranges);
    lineStart = info.index + info.separatorLength;
    if (info.separatorLength === 0 && info.index >= text.length) break;
  }
  return ranges;
}

function processLineForBoldRanges(line, absoluteLineStart, ranges) {
  if (!line || !line.trim()) return;
  if (/^[\d\s.,:;\/\-–—()]+$/.test(line.trim())) return;
  const first = findFirstNonSpaceIndex(line);
  const last = findLastNonSpaceIndex(line);
  if (first === -1 || last === -1) return;
  const colon = line.indexOf(":");
  if (colon >= 0) ranges.push({ start: absoluteLineStart + first, length: colon - first + 1 });
  else ranges.push({ start: absoluteLineStart + first, length: last - first + 1 });
}

function findNextLineEnd(text, startIndex) {
  let nearestIndex = text.length;
  let separatorLength = 0;
  for (let i = startIndex; i < text.length; i++) {
    const ch = text[i];
    if (ch === "\r") {
      nearestIndex = i;
      separatorLength = text[i + 1] === "\n" ? 2 : 1;
      break;
    }
    if (ch === "\n" || ch === "\u000B") {
      nearestIndex = i;
      separatorLength = 1;
      break;
    }
  }
  return { index: nearestIndex, separatorLength };
}

function findFirstNonSpaceIndex(text) {
  for (let i = 0; i < text.length; i++) if (!/\s/.test(text[i])) return i;
  return -1;
}

function findLastNonSpaceIndex(text) {
  for (let i = text.length - 1; i >= 0; i--) if (!/\s/.test(text[i])) return i;
  return -1;
}
