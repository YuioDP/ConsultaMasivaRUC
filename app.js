(() => {
  "use strict";

  // Fuente primaria: jsDelivr. Respaldo: GitHub Raw.
  const DATA_SOURCES = [
    {
      name: "jsDelivr",
      base: "https://cdn.jsdelivr.net/gh/alb3rt0ru1z/tribio-padron-ruc@main",
    },
    {
      name: "GitHub Raw",
      base: "https://raw.githubusercontent.com/alb3rt0ru1z/tribio-padron-ruc/main",
    },
  ];
  const OFFICIAL_SOURCE = "https://www.sunat.gob.pe/descargaPRR/mrc137_padron_reducido.html";
  const MIRROR_SOURCE = "https://github.com/alb3rt0ru1z/tribio-padron-ruc";
  const PREFIX_LENGTH = 5;
  const VALID_RUC_PREFIXES = new Set(["10", "15", "16", "17", "20"]);
  const MAX_FILE_SIZE = 25 * 1024 * 1024;
  const CONCURRENCY = 2;

  const el = (id) => document.getElementById(id);
  const state = {
    mode: "file",
    file: null,
    entries: [],
    output: null,
    manifest: null,
    summary: null,
    sourceLabel: "",
    busy: false,
    parseVersion: 0,
  };

  const input = el("fileInput");
  const dropzone = el("dropzone");
  const startButton = el("startButton");
  const progressPanel = el("progressPanel");
  const progressFill = el("progressFill");
  const errorPanel = el("errorPanel");

  function formatBytes(bytes) {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
  }

  function normalizeText(value) {
    return String(value ?? "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toUpperCase()
      .replace(/[^A-Z0-9]/g, "");
  }

  function cellToString(value) {
    if (value === null || value === undefined) return "";
    if (typeof value === "number" && Number.isFinite(value)) return String(Math.trunc(value));
    return String(value).trim();
  }

  function extractRuc(value) {
    let raw = cellToString(value).trim();
    if (!raw) return { ruc: "", original: "" };

    if (/^[+-]?(?:\d+\.?\d*|\.\d+)[eE][+-]?\d+$/.test(raw)) {
      const number = Number(raw);
      if (Number.isFinite(number)) raw = String(Math.trunc(number));
    }
    raw = raw.replace(/\.0+$/, "");
    const digits = raw.replace(/\D/g, "");
    return { ruc: digits.length === 11 ? digits : "", original: raw };
  }

  function isRucHeader(value) {
    return ["RUC", "NUMERORUC", "NRORUC", "RUCNUMERO", "NUMERODERUC"].includes(normalizeText(value));
  }

  function parseTextList(text) {
    const lines = String(text ?? "").replace(/^\uFEFF/, "").split(/\r?\n/);
    const entries = [];

    lines.forEach((rawLine, lineIndex) => {
      const line = rawLine.trim();
      if (!line) return;
      if (isRucHeader(line)) return;

      // Cada elemento se analiza por separado. Así, si una persona pega una
      // línea con varios RUC y uno tiene una longitud incorrecta, ese elemento
      // también queda en la lista como «Formato inválido» y no se pierde.
      const tokens = line.split(/[\s,;|]+/).map((part) => part.trim()).filter(Boolean);
      tokens.forEach((token) => {
        const cleaned = token.replace(/^["'([{]+|["')\]}.,]+$/g, "");
        if (!cleaned || isRucHeader(cleaned)) return;
        const parsed = extractRuc(cleaned);
        entries.push({ ...parsed, line: lineIndex + 1 });
      });
    });

    return entries;
  }

  async function parseTxt(file) {
    return parseTextList(await file.text());
  }

  async function parseSpreadsheet(file) {
    await ensureSheetJS();
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array", cellText: true, cellDates: false });
    const firstSheetName = workbook.SheetNames[0];
    if (!firstSheetName) throw new Error("El archivo Excel no contiene hojas.");

    const sheet = workbook.Sheets[firstSheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false, blankrows: false });
    if (!rows.length) throw new Error("La primera hoja del Excel está vacía.");

    let rucIndex = 0;
    let startRow = 0;
    const headerIndex = rows[0].findIndex(isRucHeader);
    if (headerIndex >= 0) {
      rucIndex = headerIndex;
      startRow = 1;
    }

    const entries = [];
    for (let rowIndex = startRow; rowIndex < rows.length; rowIndex++) {
      const value = rows[rowIndex]?.[rucIndex] ?? "";
      if (cellToString(value).trim() === "") continue;
      const parsed = extractRuc(value);
      entries.push({ ...parsed, row: rowIndex + 1 });
    }

    if (!entries.length) throw new Error("No se encontraron valores en la columna RUC (o en la primera columna) de la primera hoja.");
    return entries;
  }

  function setError(message) {
    errorPanel.textContent = message;
    errorPanel.classList.remove("hidden");
  }

  function clearError() {
    errorPanel.textContent = "";
    errorPanel.classList.add("hidden");
  }

  function clearResults() {
    state.output = null;
    state.summary = null;
    el("resultsSection").classList.add("hidden");
    progressPanel.classList.add("hidden");
  }

  function updateProgress(percent, title, detail) {
    progressPanel.classList.remove("hidden");
    progressFill.style.width = `${Math.max(0, Math.min(100, percent))}%`;
    el("progressPercent").textContent = `${Math.round(percent)}%`;
    el("progressTitle").textContent = title;
    el("progressDetail").textContent = detail;
  }

  function updateStartButton() {
    startButton.disabled = state.busy || !state.entries.length;
    if (!state.busy) startButton.querySelector("span").textContent = state.output ? "Volver a consultar" : "Consultar RUC";
  }

  function setBusy(busy) {
    state.busy = busy;
    startButton.disabled = busy || !state.entries.length;
    startButton.querySelector("span").textContent = busy ? "Consultando…" : (state.output ? "Volver a consultar" : "Consultar RUC");
  }

  function updateEntrySelection(entries, sourceLabel) {
    state.entries = entries;
    state.sourceLabel = sourceLabel;
    clearResults();
    clearError();
    updateStartButton();
  }

  function setMode(mode) {
    if (state.busy) return;
    state.mode = mode;
    state.file = null;
    state.parseVersion++;
    input.value = "";
    updateEntrySelection([], "");
    el("fileCard").classList.add("hidden");
    dropzone.classList.remove("hidden");
    el("fileModePanel").classList.toggle("hidden", mode !== "file");
    el("textModePanel").classList.toggle("hidden", mode !== "text");
    el("fileModeButton").classList.toggle("active", mode === "file");
    el("textModeButton").classList.toggle("active", mode === "text");
    el("fileModeButton").setAttribute("aria-selected", String(mode === "file"));
    el("textModeButton").setAttribute("aria-selected", String(mode === "text"));
    if (mode === "file") {
      el("fileModeButton").focus();
    } else {
      el("textModeButton").focus();
      el("rucTextInput").focus();
    }
  }

  async function setFile(file) {
    if (state.busy) return;
    clearError();
    state.file = file || null;
    const parseVersion = ++state.parseVersion;
    updateEntrySelection([], "");

    if (!file) {
      el("fileCard").classList.add("hidden");
      dropzone.classList.remove("hidden");
      input.value = "";
      return;
    }

    const extension = file.name.split(".").pop().toLowerCase();
    const supportedExtensions = new Set(["txt", "csv", "xls", "xlsx", "xlsm", "xlsb"]);
    if (!supportedExtensions.has(extension)) {
      state.file = null;
      setError("Formato no compatible. Selecciona un archivo TXT, CSV o Excel.");
      input.value = "";
      return;
    }
    if (file.size > MAX_FILE_SIZE) {
      state.file = null;
      setError(`El archivo supera el límite de ${formatBytes(MAX_FILE_SIZE)}. Divide la lista en archivos más pequeños.`);
      input.value = "";
      return;
    }

    try {
      const entries = extension === "txt" || extension === "csv"
        ? await parseTxt(file)
        : await parseSpreadsheet(file);
      if (parseVersion !== state.parseVersion) return;
      if (!entries.length) throw new Error("No se encontraron RUC ni valores para consultar en el archivo.");

      updateEntrySelection(entries, file.name);
      el("fileName").textContent = file.name;
      el("fileMeta").textContent = `${entries.length.toLocaleString("es-PE")} filas para consultar · ${formatBytes(file.size)}`;
      el("fileIcon").textContent = extension.toUpperCase().slice(0, 4);
      el("fileCard").classList.remove("hidden");
      dropzone.classList.add("hidden");
    } catch (error) {
      if (parseVersion !== state.parseVersion) return;
      state.file = null;
      el("fileCard").classList.add("hidden");
      dropzone.classList.remove("hidden");
      input.value = "";
      setError(error.message || "No se pudo leer el archivo.");
    }
  }

  function preparePastedList() {
    if (state.busy) return;
    clearError();
    const text = el("rucTextInput").value;
    const entries = parseTextList(text);
    if (!entries.length) {
      updateEntrySelection([], "");
      setError("Pega o escribe al menos un RUC antes de preparar la lista.");
      el("pasteCount").textContent = "0 registros detectados";
      return;
    }
    updateEntrySelection(entries, "Lista pegada o escrita en la página");
    el("pasteCount").textContent = `${entries.length.toLocaleString("es-PE")} registros preparados`;
  }

  function dateLabel(value) {
    const date = new Date(value);
    if (!Number.isFinite(date.getTime())) return "Fecha no informada";
    return new Intl.DateTimeFormat("es-PE", {
      timeZone: "America/Lima",
      day: "2-digit", month: "short", year: "numeric",
      hour: "2-digit", minute: "2-digit",
    }).format(date);
  }

  function updateSourceNotice(manifest) {
    const notice = el("sourceNotice");
    const stamp = manifest.updated_at || "";
    const parsed = new Date(stamp);
    if (!Number.isFinite(parsed.getTime())) {
      el("sourceFreshness").textContent = "El distribuidor no informó la fecha de generación de esta copia.";
      notice.classList.add("stale");
      return;
    }

    const ageDays = Math.max(0, Math.floor((Date.now() - parsed.getTime()) / 86400000));
    const base = `Copia de datos generada: ${dateLabel(stamp)} · ${Number(manifest.total_records || 0).toLocaleString("es-PE")} registros.`;
    if (ageDays > 3) {
      el("sourceFreshness").textContent = `${base} Aviso: esta copia tiene ${ageDays} días; podría no incluir cambios recientes de SUNAT.`;
      notice.classList.add("stale");
    } else {
      el("sourceFreshness").textContent = `${base} La copia tiene ${ageDays} días de antigüedad.`;
      notice.classList.remove("stale");
    }
  }

  async function fetchJsonFromSources(path, options = {}) {
    const {
      cache = "force-cache",
      version = "",
      missingIsNull = false,
      timeoutMs = 30000,
    } = options;
    const failures = [];
    let allResponsesWere404 = true;

    for (const source of DATA_SOURCES) {
      const suffix = version ? `?v=${encodeURIComponent(version)}` : "";
      const url = `${source.base}/${path}${suffix}`;
      const controller = new AbortController();
      const timeout = setTimeout(() => controller.abort(), timeoutMs);

      try {
        const response = await fetch(url, {
          cache,
          signal: controller.signal,
          headers: { Accept: "application/json" },
        });

        if (response.status === 404) {
          failures.push(`${source.name}: HTTP 404`);
          continue;
        }
        allResponsesWere404 = false;
        if (!response.ok) {
          let responseHint = "";
          try { responseHint = (await response.text()).replace(/\s+/g, " ").slice(0, 140); } catch { /* Sin cuerpo disponible. */ }
          failures.push(`${source.name}: HTTP ${response.status}${responseHint ? ` (${responseHint})` : ""}`);
          continue;
        }

        try {
          return await response.json();
        } catch {
          failures.push(`${source.name}: respuesta JSON inválida`);
        }
      } catch (error) {
        allResponsesWere404 = false;
        failures.push(`${source.name}: ${error.name === "AbortError" ? "tiempo de espera agotado" : (error.message || "error de conexión")}`);
      } finally {
        clearTimeout(timeout);
      }
    }

    if (missingIsNull && allResponsesWere404 && failures.length === DATA_SOURCES.length) return null;
    const detail = failures.length ? failures.join("; ") : "sin respuesta de los distribuidores";
    throw new Error(
      `No se pudo obtener ${path}. Se intentaron las fuentes disponibles (${detail}). ` +
      "Puede tratarse de una restricción temporal del CDN o de un problema de conexión. Vuelve a intentarlo en unos minutos.",
    );
  }

  async function fetchManifest() {
    const manifest = await fetchJsonFromSources("manifest.json", { cache: "no-store", timeoutMs: 20000 });
    if (!manifest || !manifest.updated_at || !manifest.columns) throw new Error("El manifiesto de datos tiene un formato inesperado.");
    return manifest;
  }

  async function fetchChunk(prefix, manifest) {
    const payload = await fetchJsonFromSources(`chunks/${prefix}.json`, {
      cache: "force-cache",
      version: manifest.updated_at,
      missingIsNull: true,
      timeoutMs: 30000,
    });
    if (payload === null) return null;
    if (!Array.isArray(payload.records) || !Array.isArray(payload.columns)) throw new Error(`El bloque ${prefix} no tiene la estructura de datos esperada.`);
    if (payload.prefix && String(payload.prefix) !== prefix) throw new Error(`El bloque recibido no corresponde al prefijo ${prefix}.`);
    return payload;
  }

  function columnIndex(columns, name, fallback) {
    const index = columns.findIndex((column) => normalizeText(column) === normalizeText(name));
    return index >= 0 ? index : fallback;
  }

  function recordToData(record, columns) {
    const ruc = columnIndex(columns, "ruc", 0);
    const name = columnIndex(columns, "razon_social", 1);
    const status = columnIndex(columns, "estado", 2);
    const condition = columnIndex(columns, "condicion", 3);
    const address = columnIndex(columns, "direccion", 6);
    const district = columnIndex(columns, "distrito", 9);
    const province = columnIndex(columns, "provincia", 8);
    const department = columnIndex(columns, "departamento", 7);
    const parts = [record[address], record[district], record[province], record[department]]
      .map((part) => part === null || part === undefined ? "" : String(part).trim())
      .filter((part, index, array) => part && array.indexOf(part) === index);

    return {
      ruc: String(record[ruc] ?? ""),
      razon_social: String(record[name] ?? ""),
      estado: String(record[status] ?? ""),
      condicion: String(record[condition] ?? ""),
      domicilio: parts.join(", "),
    };
  }

  async function loadMatches(entries, manifest) {
    const groups = new Map();
    for (const entry of entries) {
      if (!entry.ruc || !VALID_RUC_PREFIXES.has(entry.ruc.slice(0, 2))) continue;
      const prefix = entry.ruc.slice(0, PREFIX_LENGTH);
      if (!groups.has(prefix)) groups.set(prefix, new Set());
      groups.get(prefix).add(entry.ruc);
    }

    const prefixes = [...groups.keys()];
    const matches = new Map();
    let completed = 0;
    let nextIndex = 0;
    updateProgress(4, "Conectando con el padrón…", `${prefixes.length.toLocaleString("es-PE")} bloques de datos necesarios para esta lista.`);

    async function worker() {
      while (true) {
        const current = nextIndex++;
        if (current >= prefixes.length) return;
        const prefix = prefixes[current];
        const wanted = groups.get(prefix);
        const payload = await fetchChunk(prefix, manifest);

        if (payload) {
          const rucIndex = columnIndex(payload.columns, "ruc", 0);
          for (const record of payload.records) {
            const ruc = String(record[rucIndex] ?? "");
            if (wanted.has(ruc)) matches.set(ruc, recordToData(record, payload.columns));
          }
        }

        completed++;
        const percent = prefixes.length ? 5 + (completed / prefixes.length) * 90 : 95;
        updateProgress(percent, "Consultando registros…", `Bloques revisados: ${completed.toLocaleString("es-PE")} de ${prefixes.length.toLocaleString("es-PE")}. Coincidencias: ${matches.size.toLocaleString("es-PE")}.`);
      }
    }

    const workers = Array.from({ length: Math.min(CONCURRENCY, Math.max(prefixes.length, 1)) }, () => worker());
    await Promise.all(workers);
    return { matches, prefixCount: prefixes.length };
  }

  function classifyMatch(match) {
    const normalizedState = normalizeText(match.estado);
    if (normalizedState === "NUMINTERNOIDENTIF") return "Identificador interno";
    if (normalizedState === "BAJADEFINITIVA") return "BAJA DEFINITIVA";
    return "Encontrado";
  }

  function buildOutput(entries, matches) {
    return entries.map((entry) => {
      if (!entry.ruc) {
        return { ruc: entry.original, razon_social: "", estado: "", condicion: "", domicilio: "", resultado: "Formato inválido" };
      }
      const match = matches.get(entry.ruc);
      if (!match) {
        return { ruc: entry.ruc, razon_social: "", estado: "", condicion: "", domicilio: "", resultado: "No encontrado en el padrón" };
      }
      return {
        ruc: entry.ruc,
        razon_social: match.razon_social,
        estado: match.estado,
        condicion: match.condicion,
        domicilio: match.domicilio,
        resultado: classifyMatch(match),
      };
    });
  }

  function resultClass(result) {
    if (result === "Encontrado") return "status-ok";
    if (result === "BAJA DEFINITIVA") return "status-baja";
    if (result === "Identificador interno") return "status-internal";
    if (result === "Formato inválido") return "status-invalid";
    return "status-miss";
  }

  function renderPreview(rows) {
    const tbody = el("previewBody");
    tbody.innerHTML = "";
    rows.slice(0, 10).forEach((row) => {
      const tr = document.createElement("tr");
      tr.className = `preview-${resultClass(row.resultado).replace("status-", "")}`;
      const values = [row.ruc, row.razon_social || "—", row.estado || "—", row.condicion || "—", row.domicilio || "—"];
      values.forEach((value) => {
        const td = document.createElement("td");
        td.textContent = value;
        tr.appendChild(td);
      });
      const statusCell = document.createElement("td");
      const tag = document.createElement("span");
      tag.className = `status-tag ${resultClass(row.resultado)}`;
      tag.textContent = row.resultado;
      statusCell.appendChild(tag);
      tr.appendChild(statusCell);
      tbody.appendChild(tr);
    });
    el("previewCount").textContent = `${rows.length.toLocaleString("es-PE")} registros`;
  }

  function renderResults(rows) {
    const counts = {
      found: rows.filter((row) => row.resultado === "Encontrado").length,
      internal: rows.filter((row) => row.resultado === "Identificador interno").length,
      baja: rows.filter((row) => row.resultado === "BAJA DEFINITIVA").length,
      missing: rows.filter((row) => row.resultado === "No encontrado en el padrón").length,
      invalid: rows.filter((row) => row.resultado === "Formato inválido").length,
    };
    const occurrences = new Map();
    rows.forEach((row) => {
      if (row.ruc && row.resultado !== "Formato inválido") occurrences.set(row.ruc, (occurrences.get(row.ruc) || 0) + 1);
    });
    counts.duplicateRows = [...occurrences.values()].reduce((sum, count) => sum + Math.max(0, count - 1), 0);

    el("statTotal").textContent = rows.length.toLocaleString("es-PE");
    el("statFound").textContent = counts.found.toLocaleString("es-PE");
    el("statInternal").textContent = counts.internal.toLocaleString("es-PE");
    el("statBaja").textContent = counts.baja.toLocaleString("es-PE");
    el("statMissing").textContent = counts.missing.toLocaleString("es-PE");
    el("statInvalid").textContent = counts.invalid.toLocaleString("es-PE");
    renderPreview(rows);
    el("resultsSection").classList.remove("hidden");
    state.summary = counts;
    el("resultsSection").scrollIntoView({ behavior: "smooth", block: "start" });
  }

  async function runQuery() {
    if (state.busy || !state.entries.length) return;
    clearError();
    state.output = null;
    setBusy(true);
    el("resultsSection").classList.add("hidden");

    try {
      updateProgress(2, "Verificando la fuente…", "Leyendo los metadatos de la copia de datos.");
      const manifest = await fetchManifest();
      state.manifest = manifest;
      updateSourceNotice(manifest);

      const uniqueCount = new Set(state.entries.filter((entry) => entry.ruc).map((entry) => entry.ruc)).size;
      if (uniqueCount > 50000) throw new Error("La lista contiene más de 50.000 RUC distintos. Divide la lista en partes para evitar un consumo excesivo de memoria y solicitudes.");

      const { matches, prefixCount } = await loadMatches(state.entries, manifest);
      const rows = buildOutput(state.entries, matches);
      state.output = rows;
      updateProgress(100, "Consulta completada", `${rows.length.toLocaleString("es-PE")} filas procesadas y listas para exportar.`);
      renderResults(rows);
      el("progressTitle").textContent = "Consulta completada";
      el("progressDetail").textContent = `Se consultaron ${prefixCount.toLocaleString("es-PE")} bloques. El orden, los duplicados y las entradas no encontradas se conservaron.`;
      updateStartButton();
    } catch (error) {
      setError(`${error.message || "No se pudo completar la consulta."} No se ha generado ningún archivo para evitar clasificar como «no encontrados» datos que no se pudieron verificar. Revisa la conexión y vuelve a intentarlo.`);
      progressPanel.classList.add("hidden");
    } finally {
      setBusy(false);
    }
  }

  async function loadBrowserLibrary(globalName, sources) {
    if (window[globalName]) return window[globalName];
    const failures = [];
    for (const src of sources) {
      try {
        await new Promise((resolve, reject) => {
          const script = document.createElement("script");
          const timeout = window.setTimeout(() => {
            script.remove();
            reject(new Error("tiempo de espera agotado"));
          }, 20000);
          script.src = src;
          script.async = true;
          script.onload = () => { window.clearTimeout(timeout); resolve(); };
          script.onerror = () => { window.clearTimeout(timeout); script.remove(); reject(new Error("no se pudo cargar")); };
          document.head.appendChild(script);
        });
        if (window[globalName]) return window[globalName];
        failures.push(`${new URL(src).hostname}: la biblioteca no quedó disponible`);
      } catch (error) {
        failures.push(`${new URL(src).hostname}: ${error.message}`);
      }
    }
    throw new Error(`No se pudo cargar ${globalName} desde los distribuidores disponibles (${failures.join("; ")}). Comprueba tu conexión e inténtalo otra vez.`);
  }

  function ensureSheetJS() {
    return loadBrowserLibrary("XLSX", [
      "https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js",
      "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js",
      "https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js",
    ]);
  }

  function ensureExcelJS() {
    return loadBrowserLibrary("ExcelJS", [
      "https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js",
      "https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js",
      "https://unpkg.com/exceljs@4.4.0/dist/exceljs.min.js",
    ]);
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = filename;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 1500);
  }

  function rowColors(result) {
    const colors = {
      "Encontrado": { fill: "FFE8F5EE", accent: "FF23845B", tag: "FFD2EBDD" },
      "BAJA DEFINITIVA": { fill: "FFFFE8E8", accent: "FFB42332", tag: "FFFFD5D8" },
      "Identificador interno": { fill: "FFEAF2FF", accent: "FF285B9A", tag: "FFD9E8FF" },
      "No encontrado en el padrón": { fill: "FFFFF5DF", accent: "FF986500", tag: "FFFFE9B5" },
      "Formato inválido": { fill: "FFFFE7E7", accent: "FFB42332", tag: "FFFFD2D2" },
    };
    return colors[result] || colors["Encontrado"];
  }

  function applyCellBorder(cell, color = "FFE1E7EF") {
    cell.border = { bottom: { style: "thin", color: { argb: color } } };
  }

  async function downloadExcel() {
    if (!state.output) return;
    clearError();
    const downloadButton = el("downloadExcelButton");
    downloadButton.disabled = true;
    const previousLabel = downloadButton.textContent.trim();
    downloadButton.lastChild.textContent = " Preparando Excel…";

    try {
      await ensureExcelJS();
      const rows = state.output;
      const manifest = state.manifest || {};
      const summary = state.summary || {};
      const workbook = new ExcelJS.Workbook();
      workbook.creator = "Consulta Masiva RUC";
      workbook.subject = "Resultados derivados del Padrón Reducido del RUC de SUNAT";
      workbook.created = new Date();
      workbook.modified = new Date();

      const resultsSheet = workbook.addWorksheet("Resultados", {
        views: [{ state: "frozen", ySplit: 1, topLeftCell: "A2", activeCell: "A2" }],
      });
      resultsSheet.columns = [
        { header: "RUC / Número ingresado", key: "ruc", width: 21 },
        { header: "Razón social", key: "razon_social", width: 42 },
        { header: "Estado", key: "estado", width: 24 },
        { header: "Condición", key: "condicion", width: 22 },
        { header: "Domicilio fiscal según el padrón", key: "domicilio", width: 58 },
        { header: "Resultado", key: "resultado", width: 29 },
      ];
      resultsSheet.getRow(1).height = 28;
      resultsSheet.getRow(1).eachCell((cell) => {
        cell.font = { name: "Aptos", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF18364B" } };
        cell.alignment = { vertical: "middle", horizontal: "left", wrapText: true };
        cell.border = { bottom: { style: "medium", color: { argb: "FF55CDB8" } } };
      });

      rows.forEach((record) => {
        const row = resultsSheet.addRow({
          ruc: String(record.ruc ?? ""),
          razon_social: record.razon_social,
          estado: record.estado,
          condicion: record.condicion,
          domicilio: record.domicilio,
          resultado: record.resultado,
        });
        const colors = rowColors(record.resultado);
        row.height = 22;
        for (let columnNumber = 1; columnNumber <= 6; columnNumber++) {
          const cell = row.getCell(columnNumber);
          cell.font = { name: "Aptos", size: 10, color: { argb: "FF263746" }, bold: columnNumber === 6 };
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.fill } };
          cell.alignment = { vertical: "middle", wrapText: columnNumber === 2 || columnNumber === 5 };
          applyCellBorder(cell);
        }
        row.getCell(1).numFmt = "@";
        row.getCell(6).font = { name: "Aptos", size: 10, bold: true, color: { argb: colors.accent } };
        row.getCell(6).fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.tag } };
      });
      if (rows.length) resultsSheet.autoFilter = { from: "A1", to: `F${rows.length + 1}` };
      resultsSheet.properties.defaultRowHeight = 20;

      const summarySheet = workbook.addWorksheet("Resumen", {
        views: [{ state: "frozen", ySplit: 3, topLeftCell: "A4", activeCell: "A4" }],
      });
      summarySheet.columns = [{ width: 47 }, { width: 96 }];
      summarySheet.mergeCells("A1:B1");
      summarySheet.getCell("A1").value = "RESUMEN DE CONSULTA MASIVA RUC";
      summarySheet.getCell("A1").font = { name: "Aptos Display", size: 15, bold: true, color: { argb: "FFFFFFFF" } };
      summarySheet.getCell("A1").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF18364B" } };
      summarySheet.getCell("A1").alignment = { vertical: "middle" };
      summarySheet.getRow(1).height = 32;
      summarySheet.addRow([]);
      const summaryHeader = summarySheet.addRow(["Indicador", "Valor"]);
      summaryHeader.height = 24;
      summaryHeader.eachCell((cell) => {
        cell.font = { name: "Aptos", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF28566C" } };
        cell.alignment = { vertical: "middle" };
      });

      const summaryRows = [
        ["Total de filas procesadas", rows.length],
        ["Encontrados", summary.found || 0],
        ["Identificadores internos", summary.internal || 0],
        ["Baja definitiva", summary.baja || 0],
        ["No encontrados en el padrón", summary.missing || 0],
        ["Entradas con formato inválido", summary.invalid || 0],
        ["Filas duplicadas conservadas", summary.duplicateRows || 0],
        ["Forma de ingreso", state.sourceLabel || "No informada"],
        ["Fecha de generación de la copia consultada", manifest.updated_at ? dateLabel(manifest.updated_at) : "No informada"],
        ["Registros declarados en la copia", Number(manifest.total_records || 0)],
        ["Fuente original", "Padrón Reducido del RUC – SUNAT"],
        ["Página oficial de descarga", OFFICIAL_SOURCE],
        ["Distribución de la copia JSON", MIRROR_SOURCE],
        ["Observación", "Los datos son una copia del padrón, no una consulta en tiempo real. Un registro no encontrado no demuestra por sí solo que el RUC sea inexistente."],
        ["Domicilio fiscal", "Se conserva el dato de dirección que figure en el registro de la copia consultada, sin verificarlo de manera independiente."],
      ];
      summaryRows.forEach((values, index) => {
        const row = summarySheet.addRow(values);
        row.height = values[0] === "Observación" || values[0] === "Domicilio fiscal" ? 36 : 21;
        row.eachCell((cell) => {
          cell.font = { name: "Aptos", size: 10, color: { argb: "FF334155" } };
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: index % 2 === 0 ? "FFF1F5F9" : "FFFFFFFF" } };
          cell.alignment = { vertical: "middle", wrapText: true };
          applyCellBorder(cell);
        });
      });

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const date = new Date().toISOString().slice(0, 10);
      downloadBlob(blob, `consulta_masiva_ruc_${date}.xlsx`);
    } catch (error) {
      setError(`No se pudo generar el archivo Excel: ${error.message || "error desconocido"}. Vuelve a intentarlo.`);
    } finally {
      downloadButton.disabled = false;
      downloadButton.lastChild.textContent = previousLabel;
    }
  }

  function csvEscape(value) {
    let text = String(value ?? "");
    if (/^[=+@\-\t\r]/.test(text)) text = `'${text}`;
    return `"${text.replace(/"/g, '""')}"`;
  }

  function downloadCsv() {
    if (!state.output) return;
    clearError();
    const headers = ["RUC / Número ingresado", "Razón social", "Estado", "Condición", "Domicilio fiscal según el padrón", "Resultado"];
    const lines = [headers, ...state.output.map((row) => [row.ruc, row.razon_social, row.estado, row.condicion, row.domicilio, row.resultado])];
    const csv = "\uFEFFsep=;\r\n" + lines.map((line) => line.map(csvEscape).join(";")).join("\r\n");
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
    const date = new Date().toISOString().slice(0, 10);
    downloadBlob(blob, `consulta_masiva_ruc_${date}.csv`);
  }

  // Cambiar el modo de entrada.
  el("fileModeButton").addEventListener("click", () => setMode("file"));
  el("textModeButton").addEventListener("click", () => setMode("text"));

  input.addEventListener("change", (event) => {
    const file = event.target.files?.[0];
    if (file) setFile(file);
  });

  ["dragenter", "dragover"].forEach((name) => dropzone.addEventListener(name, (event) => {
    event.preventDefault();
    dropzone.classList.add("dragover");
  }));
  ["dragleave", "drop"].forEach((name) => dropzone.addEventListener(name, (event) => {
    event.preventDefault();
    dropzone.classList.remove("dragover");
  }));
  dropzone.addEventListener("drop", (event) => {
    const file = event.dataTransfer?.files?.[0];
    if (file) setFile(file);
  });
  dropzone.addEventListener("keydown", (event) => {
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      input.click();
    }
  });

  el("removeFile").addEventListener("click", () => setFile(null));
  el("prepareTextButton").addEventListener("click", preparePastedList);
  el("rucTextInput").addEventListener("input", () => {
    if (state.mode !== "text" || state.busy) return;
    state.entries = [];
    state.sourceLabel = "";
    clearResults();
    clearError();
    el("pasteCount").textContent = "Lista modificada; pulsa «Preparar lista».";
    updateStartButton();
  });
  startButton.addEventListener("click", runQuery);
  el("downloadExcelButton").addEventListener("click", downloadExcel);
  el("downloadCsvButton").addEventListener("click", downloadCsv);
})();
