(() => {
  "use strict";

  const DATA_BASE = "https://cdn.jsdelivr.net/gh/alb3rt0ru1z/tribio-padron-ruc@main";
  const MANIFEST_URL = `${DATA_BASE}/manifest.json`;
  const OFFICIAL_SOURCE = "https://www.sunat.gob.pe/descargaPRR/mrc137_padron_reducido.html";
  const MIRROR_SOURCE = "https://github.com/alb3rt0ru1z/tribio-padron-ruc";
  const PREFIX_LENGTH = 5;
  const VALID_RUC_PREFIXES = new Set(["10", "15", "16", "17", "20"]);
  const MAX_FILE_SIZE = 25 * 1024 * 1024;
  const CONCURRENCY = 3;

  const el = (id) => document.getElementById(id);
  const state = {
    file: null,
    entries: [],
    output: null,
    manifest: null,
    busy: false,
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

    // Recupera números que Excel pudo representar en notación científica.
    if (/^[+-]?(?:\d+\.?\d*|\.\d+)[eE][+-]?\d+$/.test(raw)) {
      const n = Number(raw);
      if (Number.isFinite(n)) raw = String(Math.trunc(n));
    }
    raw = raw.replace(/\.0+$/, "");

    const digits = raw.replace(/\D/g, "");
    return {
      ruc: digits.length === 11 ? digits : "",
      original: raw,
    };
  }

  function isRucHeader(value) {
    const normalized = normalizeText(value);
    return ["RUC", "NUMERORUC", "NRO RUC", "NRORUC", "RUCNUMERO"].includes(normalized);
  }

  async function parseTxt(file) {
    const text = await file.text();
    const lines = text.replace(/^\uFEFF/, "").split(/\r?\n/);
    const entries = [];

    for (let index = 0; index < lines.length; index++) {
      const line = lines[index].trim();
      if (!line) continue;
      if (index === 0 && /RUC/i.test(line) && !/\d{11}/.test(line)) continue;

      const matches = line.match(/(?<!\d)\d{11}(?!\d)/g);
      if (matches && matches.length) {
        for (const match of matches) {
          entries.push({ original: match, ruc: match, line: index + 1 });
        }
      } else {
        entries.push({ original: line, ruc: "", line: index + 1 });
      }
    }
    return entries;
  }

  async function parseSpreadsheet(file) {
    if (!window.XLSX) throw new Error("No se pudo cargar la biblioteca de Excel. Actualiza la página e inténtalo nuevamente.");
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

  function updateProgress(percent, title, detail) {
    progressPanel.classList.remove("hidden");
    progressFill.style.width = `${Math.max(0, Math.min(100, percent))}%`;
    el("progressPercent").textContent = `${Math.round(percent)}%`;
    el("progressTitle").textContent = title;
    el("progressDetail").textContent = detail;
  }

  function setBusy(busy) {
    state.busy = busy;
    startButton.disabled = busy || !state.entries.length;
    startButton.querySelector("span").textContent = busy ? "Consultando…" : "Consultar RUC";
    if (!busy) progressPanel.classList.add("hidden");
  }

  function setFile(file) {
    clearError();
    state.file = file;
    state.entries = [];
    state.output = null;
    startButton.disabled = true;
    el("resultsSection").classList.add("hidden");
    progressPanel.classList.add("hidden");
    if (!file) {
      el("fileCard").classList.add("hidden");
      dropzone.classList.remove("hidden");
      startButton.disabled = true;
      input.value = "";
      return;
    }

    const extension = file.name.split(".").pop().toLowerCase();
    const supportedExtensions = new Set(["txt", "csv", "xls", "xlsx", "xlsm", "xlsb"]);
    if (!supportedExtensions.has(extension)) {
      setError("Formato no compatible. Selecciona un archivo TXT, CSV o Excel.");
      setFile(null);
      return;
    }
    if (file.size > MAX_FILE_SIZE) {
      setError(`El archivo supera el límite de ${formatBytes(MAX_FILE_SIZE)}. Divide la lista en archivos más pequeños.`);
      setFile(null);
      return;
    }

    const parsePromise = extension === "txt" || extension === "csv"
      ? parseTxt(file)
      : parseSpreadsheet(file);

    parsePromise.then((entries) => {
      if (!entries.length) throw new Error("No se encontraron RUC ni valores para consultar en el archivo.");
      state.entries = entries;
      el("fileName").textContent = file.name;
      el("fileMeta").textContent = `${entries.length.toLocaleString("es-PE")} filas para consultar · ${formatBytes(file.size)}`;
      el("fileIcon").textContent = extension.toUpperCase().slice(0, 4);
      el("fileCard").classList.remove("hidden");
      dropzone.classList.add("hidden");
      startButton.disabled = false;
    }).catch((error) => {
      setFile(null);
      setError(error.message || "No se pudo leer el archivo.");
    });
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

  async function fetchJson(url) {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 60000);
    try {
      const response = await fetch(url, { cache: "no-store", signal: controller.signal });
      if (!response.ok) {
        const error = new Error(`La fuente de datos respondió HTTP ${response.status}.`);
        error.status = response.status;
        throw error;
      }
      return await response.json();
    } finally {
      clearTimeout(timeout);
    }
  }

  async function fetchManifest() {
    const manifest = await fetchJson(`${MANIFEST_URL}?t=${Date.now()}`);
    if (!manifest || !manifest.updated_at || !manifest.columns) {
      throw new Error("El manifiesto de datos tiene un formato inesperado.");
    }
    return manifest;
  }

  async function fetchChunk(prefix, manifest) {
    const url = `${DATA_BASE}/chunks/${prefix}.json?v=${encodeURIComponent(manifest.updated_at)}`;
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 60000);
    try {
      const response = await fetch(url, { cache: "force-cache", signal: controller.signal });
      if (response.status === 404) return null;
      if (!response.ok) throw new Error(`No se pudo descargar el bloque ${prefix} (HTTP ${response.status}).`);
      const payload = await response.json();
      if (!Array.isArray(payload.records) || !Array.isArray(payload.columns)) {
        throw new Error(`El bloque ${prefix} no tiene la estructura de datos esperada.`);
      }
      if (payload.prefix && String(payload.prefix) !== prefix) {
        throw new Error(`El bloque recibido no corresponde al prefijo ${prefix}.`);
      }
      return payload;
    } catch (error) {
      if (error.name === "AbortError") throw new Error(`La descarga del bloque ${prefix} tardó demasiado. Comprueba tu conexión e inténtalo nuevamente.`);
      throw error;
    } finally {
      clearTimeout(timeout);
    }
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

  function buildOutput(entries, matches) {
    return entries.map((entry) => {
      if (!entry.ruc) {
        return {
          ruc: entry.original,
          razon_social: "",
          estado: "",
          condicion: "",
          domicilio: "",
          resultado: "Formato inválido",
        };
      }
      const match = matches.get(entry.ruc);
      if (!match) {
        return {
          ruc: entry.ruc,
          razon_social: "",
          estado: "",
          condicion: "",
          domicilio: "",
          resultado: "No encontrado en el padrón",
        };
      }
      return {
        ruc: entry.ruc,
        razon_social: match.razon_social,
        estado: match.estado,
        condicion: match.condicion,
        domicilio: match.domicilio,
        resultado: "Encontrado",
      };
    });
  }

  function renderPreview(rows) {
    const tbody = el("previewBody");
    tbody.innerHTML = "";
    rows.slice(0, 8).forEach((row) => {
      const tr = document.createElement("tr");
      const cells = [row.ruc, row.razon_social || "—", row.estado || "—", row.condicion || "—"];
      cells.forEach((text) => {
        const td = document.createElement("td");
        td.textContent = text;
        tr.appendChild(td);
      });
      const tdStatus = document.createElement("td");
      const tag = document.createElement("span");
      tag.className = "status-tag " + (row.resultado === "Encontrado" ? "status-ok" : row.resultado === "Formato inválido" ? "status-invalid" : "status-miss");
      tag.textContent = row.resultado;
      tdStatus.appendChild(tag);
      tr.appendChild(tdStatus);
      tbody.appendChild(tr);
    });
    el("previewCount").textContent = `${rows.length.toLocaleString("es-PE")} registros`;
  }

  function renderResults(rows) {
    const found = rows.filter((row) => row.resultado === "Encontrado").length;
    const missing = rows.filter((row) => row.resultado === "No encontrado en el padrón").length;
    const invalid = rows.filter((row) => row.resultado === "Formato inválido").length;
    const counts = new Map();
    rows.forEach((row) => { if (row.ruc && row.resultado !== "Formato inválido") counts.set(row.ruc, (counts.get(row.ruc) || 0) + 1); });
    const duplicateRows = [...counts.values()].reduce((sum, count) => sum + Math.max(0, count - 1), 0);

    el("statTotal").textContent = rows.length.toLocaleString("es-PE");
    el("statFound").textContent = found.toLocaleString("es-PE");
    el("statMissing").textContent = missing.toLocaleString("es-PE");
    el("statInvalid").textContent = invalid.toLocaleString("es-PE");
    renderPreview(rows);
    el("resultsSection").classList.remove("hidden");
    state.summary = { found, missing, invalid, duplicateRows };
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
      if (uniqueCount > 50000) throw new Error("La lista contiene más de 50.000 RUC distintos. Divide el archivo en partes para evitar un consumo excesivo de memoria y solicitudes.");

      const { matches, prefixCount } = await loadMatches(state.entries, manifest);
      const rows = buildOutput(state.entries, matches);
      state.output = rows;
      updateProgress(100, "Consulta completada", `${rows.length.toLocaleString("es-PE")} filas procesadas y listas para exportar.`);
      renderResults(rows);
      el("progressTitle").textContent = "Consulta completada";
      el("progressDetail").textContent = `Se consultaron ${prefixCount.toLocaleString("es-PE")} bloques. El orden y los duplicados del archivo original se conservaron.`;
      startButton.querySelector("span").textContent = "Volver a consultar";
    } catch (error) {
      setError(`${error.message || "No se pudo completar la consulta."} No se ha generado un Excel para evitar marcar como «no encontrados» datos que no se pudieron verificar. Revisa la conexión y vuelve a intentarlo.`);
      progressPanel.classList.add("hidden");
    } finally {
      setBusy(false);
      if (state.output) startButton.querySelector("span").textContent = "Volver a consultar";
    }
  }

  function downloadWorkbook() {
    if (!state.output) return;
    if (!window.XLSX) {
      setError("No se pudo cargar la biblioteca necesaria para generar Excel. Comprueba tu conexión, actualiza la página e inténtalo nuevamente.");
      return;
    }
    clearError();
    const rows = state.output;
    const headers = ["RUC", "Razón social", "Estado", "Condición", "Domicilio fiscal", "Resultado"];
    const data = [headers, ...rows.map((row) => [row.ruc, row.razon_social, row.estado, row.condicion, row.domicilio, row.resultado])];
    const resultsSheet = XLSX.utils.aoa_to_sheet(data);
    resultsSheet["!cols"] = [
      { wch: 16 }, { wch: 42 }, { wch: 20 }, { wch: 20 }, { wch: 58 }, { wch: 27 },
    ];
    resultsSheet["!autofilter"] = { ref: `A1:F${Math.max(1, data.length)}` };

    const manifest = state.manifest || {};
    const summary = state.summary || { found: 0, missing: 0, invalid: 0, duplicateRows: 0 };
    const summaryRows = [
      ["RESUMEN DE CONSULTA MASIVA RUC", ""],
      ["Indicador", "Valor"],
      ["Total de filas procesadas", rows.length],
      ["RUC encontrados", summary.found],
      ["RUC no encontrados en el padrón", summary.missing],
      ["Entradas con formato inválido", summary.invalid],
      ["Filas duplicadas conservadas", summary.duplicateRows],
      ["Fecha de generación de la copia consultada", manifest.updated_at ? dateLabel(manifest.updated_at) : "No informada"],
      ["Registros declarados en la copia", Number(manifest.total_records || 0)],
      ["Fuente original", "Padrón Reducido del RUC – SUNAT"],
      ["Página oficial de descarga", OFFICIAL_SOURCE],
      ["Distribución de la copia JSON", MIRROR_SOURCE],
      ["Observación", "Los datos son una copia del padrón, no una consulta en tiempo real. Un RUC no encontrado no implica por sí solo que sea inexistente."],
    ];
    const summarySheet = XLSX.utils.aoa_to_sheet(summaryRows);
    summarySheet["!cols"] = [{ wch: 44 }, { wch: 92 }];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, resultsSheet, "Resultados");
    XLSX.utils.book_append_sheet(workbook, summarySheet, "Resumen");

    const date = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(workbook, `consulta_masiva_ruc_${date}.xlsx`, { compression: true });
  }

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
  startButton.addEventListener("click", runQuery);
  el("downloadButton").addEventListener("click", downloadWorkbook);
})();
