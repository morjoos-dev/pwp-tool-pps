let appConfig = null;

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    bindEvents();
    log("PwP Tool PPS cargado correctamente.");
    restoreSessionUrl();
  }
});

function bindEvents() {
  document.getElementById("loadConfigFromUrl").onclick = loadConfigFromUrlInput;
  document.getElementById("clearConfigUrl").onclick = clearConfigUrl;
  document.getElementById("configFile").onchange = loadConfigFromFileInput;

  document.getElementById("formatBoldByColon").onclick = formatBoldByColon;
  document.getElementById("formatCurrentSlide").onclick = () => log("Pendiente: formatear diapositiva actual.");
  document.getElementById("translateCurrentSlide").onclick = () => log("Pendiente: traducir diapositiva actual.");
  document.getElementById("translateAllSlides").onclick = () => log("Pendiente: traducir presentación.");
  document.getElementById("translateTables").onclick = () => log("Pendiente: traducir tablas.");
}

function log(message) {
  const logBox = document.getElementById("log");
  const line = `[${new Date().toLocaleTimeString()}] ${message}`;

  if (logBox) {
    logBox.textContent += line + "\n";
  }

  console.log(line);
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