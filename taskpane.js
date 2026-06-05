let appConfig = null;

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    bindEvents();
    restoreSessionUrl();
    log("PwP Tool PPS cargado correctamente.");
  } else {
    log("Este complemento está pensado para PowerPoint.");
  }
});

function bindEvents() {
  document.getElementById("loadConfigFromUrl").onclick = loadConfigFromUrlInput;
  document.getElementById("clearConfigUrl").onclick = clearConfigUrl;
  document.getElementById("configFile").onchange = loadConfigFromFileInput;
  document.getElementById("clearLog").onclick = clearLog;

  document.getElementById("formatBoldByColon").onclick = formatBoldByColonAllSlides;
  document.getElementById("formatCurrentSlide").onclick = formatBoldByColonCurrentSlide;

  document.getElementById("translateCurrentSlide").onclick = () => log("Pendiente: traducir diapositiva actual.");
  document.getElementById("translateAllSlides").onclick = () => log("Pendiente: traducir presentación.");
  document.getElementById("translateTables").onclick = () => log("Pendiente: traducir tablas.");
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
  const logBox = document.getElementById("log");
  if (logBox) {
    logBox.textContent = "";
  }
}

function restoreSessionUrl() {
  const savedUrl = sessionStorage.getItem("pwpToolConfigUrl");

  if (savedUrl) {
    document.getElementById("privateConfigUrl").value = savedUrl;
    log("URL restaurada desde la sesión.");
  }
}

async function loadConfigFromUrlInput() {
  const input = document.getElementById("privateConfigUrl");
  const url = input.value.trim();

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
    log("Si la URL devuelve un visor HTML o está bloqueada, usa la opción de archivo JSON.");
  }
}

function clearConfigUrl() {
  sessionStorage.removeItem("pwpToolConfigUrl");
  document.getElementById("privateConfigUrl").value = "";
  log("URL olvidada de la sesión.");
}

async function loadJsonFromUrl(url) {
  const response = await fetch(url, {
    method: "GET",
    credentials: "include"
  });

  if (!response.ok) {
    throw new Error(`HTTP ${response.status}: no se pudo cargar el archivo.`);
  }

  const text = await response.text();

  try {
    return JSON.parse(text);
  } catch (error) {
    throw new Error("La respuesta no es JSON válido. Puede que se haya devuelto un visor HTML en vez del archivo.");
  }
}

function loadConfigFromFileInput(event) {
  const file = event.target.files[0];

  if (!file) {
    return;
  }

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

  reader.onerror = () => {
    log("No se pudo leer el archivo seleccionado.");
  };

  reader.readAsText(file);
}

async function formatBoldByColonAllSlides() {
  log("Iniciando formateo de toda la presentación...");

  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    log(`Diapositivas detectadas: ${slides.items.length}`);

    for (const slide of slides.items) {
      slide.shapes.load("items");
    }

    await context.sync();

    let formattedShapes = 0;
    let formattedRanges = 0;

    for (let slideIndex = 0; slideIndex < slides.items.length; slideIndex++) {
      const slide = slides.items[slideIndex];
      const slideNumber = slideIndex + 1;

      const result = await formatSlideShapes(context, slide, slideNumber);
      formattedShapes += result.formattedShapes;
      formattedRanges += result.formattedRanges;

      await context.sync();
      log(`OK diapositiva ${slideNumber}: ${result.formattedShapes} caja/s, ${result.formattedRanges} rango/s.`);
    }

    log("Formateo terminado.");
    log(`Cajas formateadas: ${formattedShapes}`);
    log(`Rangos en negrita: ${formattedRanges}`);
  });
}

async function formatBoldByColonCurrentSlide() {
  log("Intentando formatear diapositiva actual...");

  await PowerPoint.run(async (context) => {
    let selectedSlides;

    try {
      selectedSlides = context.presentation.getSelectedSlides();
      selectedSlides.load("items");
      await context.sync();
    } catch (error) {
      log("No se pudo obtener la diapositiva actual con esta versión de PowerPoint.");
      log("Usa el botón de formatear toda la presentación.");
      return;
    }

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
    if (item.textFrame.isNullObject) continue;
    if (!item.textFrame.hasText) continue;

    const textRange = item.textFrame.textRange;
    const text = textRange.text || "";

    if (!text.trim()) continue;

    try {
      textRange.font.bold = false;

      const boldRanges = getBoldRangesByColonAndLineBreak(text);

      for (const rangeInfo of boldRanges) {
        if (rangeInfo.length <= 0) continue;
        const subRange = textRange.getSubstring(rangeInfo.start, rangeInfo.length);
        subRange.font.bold = true;
        formattedRanges++;
      }

      formattedShapes++;
    } catch (error) {
      log(`Error formateando una caja en diapositiva ${slideNumber}.`);
      console.error(error);
    }
  }

  return { formattedShapes, formattedRanges };
}

function getBoldRangesByColonAndLineBreak(text) {
  const ranges = [];
  let lineStart = 0;

  while (lineStart < text.length) {
    const lineEndInfo = findNextLineEnd(text, lineStart);
    const lineEnd = lineEndInfo.index;
    const line = text.slice(lineStart, lineEnd);

    processLineForBoldRanges(line, lineStart, ranges);

    lineStart = lineEnd + lineEndInfo.separatorLength;

    if (lineEndInfo.separatorLength === 0 && lineEnd >= text.length) {
      break;
    }
  }

  return ranges;
}

function processLineForBoldRanges(line, absoluteLineStart, ranges) {
  if (!line || line.trim() === "") return;

  const cleanLine = line.trim();

  if (/^[\d\s.,:;\/\-–—()]+$/.test(cleanLine)) {
    return;
  }

  const firstNonSpaceRelative = findFirstNonSpaceIndex(line);
  const lastNonSpaceRelative = findLastNonSpaceIndex(line);

  if (firstNonSpaceRelative === -1 || lastNonSpaceRelative === -1) {
    return;
  }

  const colonRelative = line.indexOf(":");

  if (colonRelative >= 0) {
    const start = absoluteLineStart + firstNonSpaceRelative;
    const length = colonRelative - firstNonSpaceRelative + 1;

    if (length > 0) {
      ranges.push({ start, length });
    }

    return;
  }

  const start = absoluteLineStart + firstNonSpaceRelative;
  const length = lastNonSpaceRelative - firstNonSpaceRelative + 1;

  if (length > 0) {
    ranges.push({ start, length });
  }
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

  return {
    index: nearestIndex,
    separatorLength
  };
}

function findFirstNonSpaceIndex(text) {
  for (let i = 0; i < text.length; i++) {
    if (!/\s/.test(text[i])) return i;
  }

  return -1;
}

function findLastNonSpaceIndex(text) {
  for (let i = text.length - 1; i >= 0; i--) {
    if (!/\s/.test(text[i])) return i;
  }

  return -1;
}
