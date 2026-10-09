const DB_NAME = "correctivos-db-v1";
const DB_VERSION = 1;
const STORE_NAME = "state";
const LAST_NUMBER_KEY = "correctivos-last-number-used";
const cellColorOptions = ["", "#fff3bf", "#d3f9d8", "#d0ebff", "#ffe3f2", "#e5dbff", "#ffd8a8"];

const defectOptions = [
  "Extintor caducado.",
  "Hay un obstáculo.",
  "Extintor descargado.",
  "Extintor sin presión.",
  "Extintor en el suelo.",
  "Cristal armario roto o sin cristal.",
  "Sin señal.",
  "Señal caducada.",
  "Extintor en mal estado.",
];

const fields = ["edificio", "ubicacion", "fechaFabricacion", "cantidad", "numeroSerie", "observaciones"];
const statusFields = ["det", "cex", "rxt", "gru", "red", "mon"];

function statusValue(value) {
  if (normalizeHeader(value) === "ok") return "ok";
  return value === true || ["si", "true", "1"].includes(normalizeHeader(value));
}

function statusLabel(value) {
  return value === "ok" ? "OK" : value ? "Sí" : "No";
}

function setStatusButton(button, value, label) {
  value = statusValue(value);
  button.className = `statusButton ${value === "ok" ? "statusOk" : value ? "statusYes" : "statusNo"}`;
  button.textContent = statusLabel(value);
  button.dataset.status = String(value);
  button.setAttribute("aria-pressed", value === "ok" ? "mixed" : String(value));
  button.setAttribute("aria-label", `${label}: ${statusLabel(value)}`);
}

function bindStatusButton(button, readValue, writeValue) {
  let clickTimer;
  const commit = (value) => {
    clearTimeout(clickTimer);
    clickTimer = null;
    return writeValue(value);
  };
  button.onclick = (event) => {
    if (event.detail === 0) return commit(readValue() !== true);
    clearTimeout(clickTimer);
    clickTimer = setTimeout(() => commit(readValue() !== true), 350);
  };
  button.ondblclick = (event) => {
    event.preventDefault();
    return commit("ok");
  };
}

let records = [];
let currentPhotos = ["", ""];
let clients = [];
let activeClientId = null;
let importing = false;

const $ = (id) => document.getElementById(id);

function safeText(value) {
  return value === undefined || value === null ? "" : String(value);
}

function createId() {
  return `rec-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function normalizeDefects(defects) {
  return (Array.isArray(defects) ? defects : []).map((defect) => {
    if (defect === "Cristal del extintor ausente o roto.") return "Cristal armario roto o sin cristal.";
    return defect;
  });
}

function cleanRecord(record = {}) {
  return {
    id: record.id || createId(),
    ...Object.fromEntries(statusFields.map((key) => [key, statusValue(record[key])])),
    cliente: safeText(record.cliente),
    edificio: safeText(record.edificio ?? record.edificioCodigo),
    cantidad: safeText(record.cantidad),
    ubicacion: safeText(record.ubicacion),
    modelo: safeText(record.modelo),
    numeroSerie: safeText(record.numeroSerie),
    fechaFabricacion: safeText(record.fechaFabricacion),
    fechaProximoRetimbrado: safeText(record.fechaProximoRetimbrado),
    observaciones: safeText(record.observaciones),
    senal: safeText(record.senal),
    defectos: normalizeDefects(record.defectos),
    photos: Array.isArray(record.photos) ? [safeText(record.photos[0]), safeText(record.photos[1])] : ["", ""],
    visto: Boolean(record.visto),
    cellColors: {
      edificio: safeText(record.cellColors?.edificio),
      ubicacion: safeText(record.cellColors?.ubicacion),
      modelo: safeText(record.cellColors?.modelo),
    },
    origen: record.origen || "excel",
  };
}

function normalizeKeyPart(value) {
  return safeText(value).trim().toLowerCase().replace(/\s+/g, " ");
}

function recordKey(record) {
  return [
    normalizeKeyPart(record.cantidad),
    normalizeKeyPart(record.numeroSerie),
    normalizeKeyPart(record.ubicacion),
  ].join("|");
}

function excelCellToText(value) {
  if (value === undefined || value === null) return "";
  if (value instanceof Date) return value.toLocaleDateString("es-ES");
  if (typeof value === "object") {
    if (value.text) return String(value.text);
    if (value.result !== undefined) return excelCellToText(value.result);
    if (Array.isArray(value.richText)) return value.richText.map((part) => part.text || "").join("");
  }
  return String(value);
}

function normalizeHeader(value) {
  return safeText(value)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[º°]/g, "o")
    .replace(/[^a-z0-9]/g, "");
}

function buildHeaderMap(rowValues) {
  const map = {};
  for (let col = 1; col < rowValues.length; col += 1) {
    const key = normalizeHeader(excelCellToText(rowValues[col]));
    if (key) map[key] = col;
  }
  return map;
}

function importedValue(rowValues, headerMap, keys, fallbackCol) {
  const col = keys.map((key) => headerMap[key]).find(Boolean);
  if (!col) return "";
  return excelCellToText(rowValues[col]);
}

function rowToImportedRecord(rowValues, index, headerMap = {}) {
  return cleanRecord({
    id: `import-${Date.now()}-${index}-${Math.random().toString(16).slice(2)}`,
    cliente: importedValue(rowValues, headerMap, ["cliente"], 1),
    edificio: importedValue(rowValues, headerMap, ["edificio"], 2),
    cantidad: importedValue(rowValues, headerMap, ["defectos", "numerosyco", "numero", "num"], 3),
    ubicacion: importedValue(rowValues, headerMap, ["ubicacion"], 4),
    modelo: importedValue(rowValues, headerMap, ["modelo"], 5),
    numeroSerie: importedValue(rowValues, headerMap, ["recordar", "noserie", "numeroserie", "serie"], 6),
    fechaFabricacion: importedValue(rowValues, headerMap, ["informacion", "fechaanofabricacion", "fechafabricacion", "fabricacion"], 7),
    fechaProximoRetimbrado: importedValue(rowValues, headerMap, ["fecharetimbrado", "retimbrado"], 8),
    observaciones: importedValue(rowValues, headerMap, ["observaciones", "observacion"], 9),
    senal: importedValue(rowValues, headerMap, ["senal"], 10),
    visto: ["si", "true", "1"].includes(normalizeSpeechText(importedValue(rowValues, headerMap, ["visto"], 0))),
    ...Object.fromEntries(statusFields.map((key) => [key, statusValue(importedValue(rowValues, headerMap, [key], 0))])),
    origen: "importado",
  });
}

function openDatabase() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = () => request.result.createObjectStore(STORE_NAME);
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function readState(key = "records") {
  const db = await openDatabase();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, "readonly");
    const request = tx.objectStore(STORE_NAME).get(key);
    request.onsuccess = () => resolve(request.result || null);
    request.onerror = () => reject(request.error);
  });
}

async function writeState(value, key = "clients-v1") {
  const db = await openDatabase();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, "readwrite");
    tx.objectStore(STORE_NAME).put(value, key);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
    tx.onabort = () => reject(tx.error || new Error("No se han guardado los datos."));
  });
}

async function loadRecords() {
  const savedClients = await readState("clients-v1");
  if (Array.isArray(savedClients)) {
    clients = savedClients.map((client) => ({ ...client, records: (client.records || []).map(cleanRecord) }));
    return;
  }
  // Keep the legacy record key intact while creating the client-based storage.
  const saved = await readState();
  const previous = (Array.isArray(saved) ? saved : window.INITIAL_EXTINTORES_LISTADOS || []).map(cleanRecord);
  const groups = new Map();
  for (const record of previous) {
    const name = record.cliente.trim() || "Registros anteriores";
    if (!groups.has(name)) groups.set(name, { id: createId(), name, data1: "", data2: "", records: [] });
    groups.get(name).records.push(record);
  }
  clients = [...groups.values()];
  await writeState(clients);
}

async function saveRecords() {
  const client = clients.find((item) => item.id === activeClientId);
  if (!client) throw new Error("Selecciona un cliente.");
  records = records.map(cleanRecord);
  client.records = records;
  await writeState(clients);
  updateStats();
}

function openClient(id) {
  if (importing) return;
  const client = clients.find((item) => item.id === id);
  if (!client) return;
  activeClientId = id;
  records = client.records;
  $("clientHeading").textContent = client.name;
  for (const field of ["filterEdificio", "filterNumero", "filterSerie"]) $(field).value = "";
  $("sortOrder").value = "none";
  $("seenFilter").value = "all";
  $("importStatus").textContent = "";
  updateStats();
  showView("home");
}

function renderClients() {
  const list = $("clientsList");
  list.replaceChildren();
  for (const client of clients) {
    const form = document.createElement("form");
    form.className = "clientCard";
    const open = document.createElement("button");
    open.type = "button";
    open.className = "clientOpen";
    open.textContent = client.name;
    open.onclick = () => openClient(client.id);
    form.append(open);
    const count = document.createElement("p");
    count.textContent = `${client.records.length} registros`;
    form.append(count);
    const inputs = {};
    for (const [key, title] of [["name", "Nombre del cliente"], ["data1", "Datos 1"], ["data2", "Datos 2"]]) {
      const label = document.createElement("label");
      label.textContent = title;
      const input = document.createElement(key === "name" ? "input" : "textarea");
      input.value = client[key] || "";
      if (key === "name") { input.required = true; input.maxLength = 150; } else input.rows = 2;
      inputs[key] = input;
      label.append(input);
      form.append(label);
    }
    const save = document.createElement("button");
    save.type = "submit";
    save.className = "secondary";
    save.textContent = "Guardar datos del cliente";
    form.append(save);
    form.onsubmit = async (event) => {
      event.preventDefault();
      const name = inputs.name.value.trim();
      if (!name) return inputs.name.focus();
      const previous = { name: client.name, data1: client.data1, data2: client.data2 };
      save.disabled = true;
      Object.assign(client, { name, data1: inputs.data1.value.trim(), data2: inputs.data2.value.trim() });
      try { await writeState(clients); open.textContent = name; save.textContent = "Datos guardados"; }
      catch { Object.assign(client, previous); alert("No se han podido guardar los datos del cliente."); }
      finally { save.disabled = false; }
    };
    list.append(form);
  }
}

async function createClient(event) {
  event.preventDefault();
  const name = $("newClientName").value.trim();
  if (!name) return $("newClientName").focus();
  const button = event.target.querySelector('button[type="submit"]');
  button.disabled = true;
  const client = { id: createId(), name, data1: $("newClientData1").value.trim(), data2: $("newClientData2").value.trim(), records: [] };
  clients.push(client);
  try {
    await writeState(clients);
    event.target.reset();
    renderClients();
    openClient(client.id);
  } catch {
    clients = clients.filter((item) => item.id !== client.id);
    alert("No se ha podido crear el cliente.");
  } finally { button.disabled = false; }
}

function updateStats() {
  const total = records.length;
  const seen = records.filter((record) => record.visto).length;
  $("totalCount").textContent = total;
  $("seenCount").textContent = seen;
  $("pendingCount").textContent = total - seen;
}

function showView(name) {
  if (importing && name === "clients") return;
  $("clientsView").classList.toggle("hidden", name !== "clients");
  $("homeView").classList.toggle("hidden", name !== "home");
  $("listView").classList.toggle("hidden", name !== "list");
  $("formView").classList.toggle("hidden", name !== "form");
  if (name === "list") renderTable();
  if (name === "clients") {
    renderClients();
    activeClientId = null;
    records = [];
  }
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function compareText(a, b) {
  return safeText(a).localeCompare(safeText(b), "es", { numeric: true, sensitivity: "base" });
}

function filteredRecords() {
  const number = $("filterNumero").value.trim().toLowerCase();
  const serial = $("filterSerie").value.trim().toLowerCase();
  const building = $("filterEdificio").value.trim().toLowerCase();
  const state = $("seenFilter").value;
  const rows = records.filter((record) => safeText(record.cantidad).toLowerCase().includes(number) && safeText(record.numeroSerie).toLowerCase().includes(serial)
    && [record.edificio, record.ubicacion].join(" ").toLowerCase().includes(building)
    && (state === "all" || (state === "seen" ? record.visto : !record.visto)));
  if ($("sortOrder").value === "edificio") rows.sort((a, b) => compareText(a.edificio, b.edificio));
  if ($("sortOrder").value === "numero") rows.sort((a, b) => compareText(a.cantidad, b.cantidad));
  return rows;
}

function renderTable() {
  const body = $("recordsBody");
  body.replaceChildren();
  for (const record of filteredRecords()) {
    const row = document.createElement("tr");
    row.dataset.recordId = record.id;
    for (const field of fields) {
      const cell = document.createElement("td");
      cell.textContent = safeText(record[field]) || "-";
      if (["edificio", "ubicacion"].includes(field)) {
        cell.dataset.fieldEdit = field;
        cell.style.backgroundColor = record.cellColors?.[field] || "";
        bindEditableTableCell(cell);
      } else {
        cell.onclick = () => { openForm(record.id); $(field).focus(); };
      }
      row.append(cell);
      if (field === "edificio") {
        const seen = document.createElement("td");
        seen.className = `seenCell ${record.visto ? "seenYes" : "seenNo"}`;
        seen.textContent = record.visto ? "Sí" : "No";
        seen.tabIndex = 0;
        seen.setAttribute("role", "button");
        seen.setAttribute("aria-label", record.visto ? "Marcar como no visto" : "Marcar como visto");
        seen.onclick = () => toggleSeenFromTable(record.id);
        seen.onkeydown = (event) => { if (["Enter", " "].includes(event.key)) { event.preventDefault(); seen.click(); } };
        row.append(seen);
      }
    }
    for (const key of statusFields) {
      const cell = document.createElement("td");
      const toggle = document.createElement("button");
      toggle.type = "button";
      setStatusButton(toggle, record[key], key.toUpperCase());
      bindStatusButton(toggle, () => record[key], async (value) => {
        toggle.disabled = true;
        const previous = record[key];
        record[key] = value;
        try {
          await saveRecords();
          renderTable();
        } catch {
          record[key] = previous;
          toggle.disabled = false;
          alert("No se ha podido guardar el cambio.");
        }
      });
      cell.append(toggle);
      row.append(cell);
    }
    const client = clients.find((item) => item.id === activeClientId);
    for (const key of ["name", "data1", "data2"]) {
      const cell = document.createElement("td");
      cell.textContent = safeText(client?.[key]) || "-";
      row.append(cell);
    }
    for (let index = 0; index < 2; index += 1) {
      const cell = document.createElement("td");
      const photo = record.photos?.[index];
      if (photo) {
        const image = document.createElement("img");
        image.className = "tablePhoto";
        image.src = photo;
        image.alt = `Foto ${index + 1}`;
        cell.append(image);
      } else {
        cell.className = "noPhoto";
        cell.textContent = "-";
      }
      row.append(cell);
    }
    const action = document.createElement("td");
    const button = document.createElement("button");
    button.className = "editBtn";
    button.textContent = "Ver / corregir";
    button.onclick = () => openForm(record.id);
    action.append(button);
    row.append(action);
    body.append(row);
  }
  if (!body.children.length) body.innerHTML = '<tr><td colspan="19">No hay registros con ese filtro.</td></tr>';
}

async function toggleSeenFromTable(recordId) {
  const record = records.find((item) => item.id === recordId);
  if (!record) return;
  record.visto = !record.visto;
  await saveRecords();
  renderTable();
}

function cellColorStyle(record, field) {
  const color = safeText(record.cellColors?.[field]);
  return color ? `background-color:${color}` : "";
}

function bindEditableTableCell(cell) {
  let clickTimer = null;
  cell.addEventListener("click", () => {
    if (clickTimer) return;
    clickTimer = setTimeout(() => {
      const row = cell.closest("tr");
      openForm(row?.dataset.recordId);
      setTimeout(() => $(cell.dataset.fieldEdit)?.focus(), 120);
      clickTimer = null;
    }, 260);
  });
  cell.addEventListener("dblclick", (event) => {
    event.preventDefault();
    if (clickTimer) {
      clearTimeout(clickTimer);
      clickTimer = null;
    }
    openColorPicker(cell.closest("tr")?.dataset.recordId, cell.dataset.fieldEdit, event);
  });
}

function closeColorPicker() {
  document.querySelector(".colorPicker")?.remove();
}

function openColorPicker(recordId, field, event) {
  if (!recordId || !field) return;
  closeColorPicker();
  const picker = document.createElement("div");
  picker.className = "colorPicker";
  picker.style.left = `${Math.min(event.clientX, window.innerWidth - 260)}px`;
  picker.style.top = `${Math.min(event.clientY + 8, window.innerHeight - 90)}px`;
  picker.innerHTML = cellColorOptions
    .map((color) => {
      const label = color ? `Color ${color}` : "Sin color";
      const style = color ? `background:${color}` : "";
      return `<button type="button" class="colorSwatch ${color ? "" : "noColor"}" style="${style}" data-color="${color}" aria-label="${label}">${color ? "" : "×"}</button>`;
    })
    .join("");
  document.body.appendChild(picker);
  picker.querySelectorAll("[data-color]").forEach((button) => {
    button.addEventListener("click", async () => {
      const record = records.find((item) => item.id === recordId);
      if (!record) return;
      record.cellColors = { ...(record.cellColors || {}), [field]: button.dataset.color };
      await saveRecords();
      closeColorPicker();
      renderTable();
    });
  });
  setTimeout(() => document.addEventListener("click", closeColorPicker, { once: true }), 0);
}

function renderDefects(selected = []) {
  const box = $("defectsList");
  box.innerHTML = "";
  for (const option of defectOptions) {
    const label = document.createElement("label");
    label.className = "checkItem";
    label.innerHTML = `<input type="checkbox" value="${option}"><span>${option}</span>`;
    label.querySelector("input").checked = selected.includes(option);
    box.appendChild(label);
  }
}

function setPhotoPreview(index, dataUrl) {
  const photo = safeText(dataUrl);
  const img = $(`photoPreview${index + 1}`);
  const text = $(`photoBox${index + 1}`).querySelector("span");
  currentPhotos[index] = photo;
  img.src = photo;
  img.classList.toggle("hidden", !photo);
  text.classList.toggle("hidden", Boolean(photo));
  $(`deletePhoto${index + 1}`).disabled = !photo;
}

function updateLastNumberUsed(currentId = "") {
  const savedNumber = localStorage.getItem(LAST_NUMBER_KEY);
  const lastRecord = records.find((record) => record.id !== currentId && safeText(record.cantidad).trim());
  const value = safeText(savedNumber).trim() || safeText(lastRecord?.cantidad).trim() || "-";
  $("lastNumberUsed").textContent = value;
}

function normalizeSpeechText(text) {
  return safeText(text)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[.,;:!?¿¡]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function spokenDigit(token) {
  const digits = {
    cero: "0",
    uno: "1",
    un: "1",
    una: "1",
    dos: "2",
    tres: "3",
    cuatro: "4",
    cinco: "5",
    seis: "6",
    siete: "7",
    ocho: "8",
    nueve: "9",
  };
  return digits[token] || (/^\d+$/.test(token) ? token : "");
}

function speechToNumberValue(text) {
  const normalized = normalizeSpeechText(text);
  const numericParts = normalized.match(/\d+/g);
  if (numericParts?.length) return numericParts.join("");

  const tokens = normalized.split(" ").filter((token) => token && token !== "y");
  const compactDigits = tokens.map(spokenDigit).join("");
  if (compactDigits && tokens.every((token) => spokenDigit(token))) return compactDigits;

  const values = {
    cero: 0,
    uno: 1,
    un: 1,
    una: 1,
    dos: 2,
    tres: 3,
    cuatro: 4,
    cinco: 5,
    seis: 6,
    siete: 7,
    ocho: 8,
    nueve: 9,
    diez: 10,
    once: 11,
    doce: 12,
    trece: 13,
    catorce: 14,
    quince: 15,
    dieciseis: 16,
    diecisiete: 17,
    dieciocho: 18,
    diecinueve: 19,
    veinte: 20,
    veintiuno: 21,
    veintidos: 22,
    veintitres: 23,
    veinticuatro: 24,
    veinticinco: 25,
    veintiseis: 26,
    veintisiete: 27,
    veintiocho: 28,
    veintinueve: 29,
    treinta: 30,
    cuarenta: 40,
    cincuenta: 50,
    sesenta: 60,
    setenta: 70,
    ochenta: 80,
    noventa: 90,
    cien: 100,
    ciento: 100,
    doscientos: 200,
    trescientos: 300,
    cuatrocientos: 400,
    quinientos: 500,
    seiscientos: 600,
    setecientos: 700,
    ochocientos: 800,
    novecientos: 900,
  };

  let total = 0;
  let current = 0;
  let found = false;
  for (const token of tokens) {
    if (token === "mil") {
      total += (current || 1) * 1000;
      current = 0;
      found = true;
      continue;
    }
    if (values[token] === undefined) continue;
    current += values[token];
    found = true;
  }

  return found ? String(total + current) : speechToPlainValue(text).replace(/\s+/g, "");
}

function speechToSerial(text) {
  const tokens = normalizeSpeechText(text).split(" ").filter(Boolean);
  const parts = tokens.map((token) => spokenDigit(token) || token.toUpperCase());
  return parts.join("").replace(/[^A-Z0-9-]/g, "");
}

function speechToPlainValue(text) {
  const digitWords = {
    cero: "0",
    uno: "1",
    un: "1",
    una: "1",
    dos: "2",
    tres: "3",
    cuatro: "4",
    cinco: "5",
    seis: "6",
    siete: "7",
    ocho: "8",
    nueve: "9",
  };
  return normalizeSpeechText(text)
    .split(" ")
    .filter(Boolean)
    .map((token) => digitWords[token] || token.toUpperCase())
    .join(" ")
    .replace(/\bKG\b/g, "KG")
    .trim();
}

function speechToModel(text) {
  const normalized = normalizeSpeechText(text);
  const compactDigits = normalized
    .split(" ")
    .map(spokenDigit)
    .join("");
  const numeric = normalized.match(/\b\d+\b/)?.[0] || compactDigits || speechToNumberValue(text);
  const models = {
    1: "ABC 1 KG",
    2: "CO2 2 KG",
    3: "ABC 3 KG",
    5: "CO2 5 KG",
    6: "ABC 6 KG",
    9: "ABC 9 KG",
    10: "CO2 10 KG",
    25: "ABC 25 KG",
    50: "ABC 50 KG",
  };
  return models[numeric] || speechToPlainValue(text);
}

function speechToYear(text) {
  const normalized = normalizeSpeechText(text);
  const numeric = normalized.match(/\b(19|20)\d{2}\b/);
  if (numeric) return numeric[0];

  const parsedNumber = Number(speechToNumberValue(text));
  if (Number.isFinite(parsedNumber)) {
    if (parsedNumber >= 1900 && parsedNumber <= 2099) return String(parsedNumber);
    if (parsedNumber >= 0 && parsedNumber <= 99) return String(2000 + parsedNumber);
  }

  const compactDigits = normalized
    .split(" ")
    .map(spokenDigit)
    .join("");
  if (/^(19|20)\d{2}$/.test(compactDigits)) return compactDigits;

  const yearWords = {
    diez: 2010,
    once: 2011,
    doce: 2012,
    trece: 2013,
    catorce: 2014,
    quince: 2015,
    dieciseis: 2016,
    diecisiete: 2017,
    dieciocho: 2018,
    diecinueve: 2019,
    veinte: 2020,
    veintiuno: 2021,
    veintidos: 2022,
    veintitres: 2023,
    veinticuatro: 2024,
    veinticinco: 2025,
    veintiseis: 2026,
  };
  for (const [word, year] of Object.entries(yearWords)) {
    if (normalized.includes(word)) return String(year);
  }
  return speechToPlainValue(text);
}

function setSelectValue(id, value) {
  const select = $(id);
  const cleanValue = safeText(value).trim();
  const option = Array.from(select.options).find((item) => item.value === cleanValue);
  if (option) {
    select.value = cleanValue;
    return;
  }
  if (/^\d{4}$/.test(cleanValue)) {
    select.add(new Option(cleanValue, cleanValue));
    select.value = cleanValue;
    return;
  }
  select.value = "";
}

function resizePhoto(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const img = new Image();
      img.onload = () => {
        const maxSide = 1200;
        const scale = Math.min(1, maxSide / Math.max(img.width, img.height));
        const canvas = document.createElement("canvas");
        canvas.width = Math.round(img.width * scale);
        canvas.height = Math.round(img.height * scale);
        canvas.getContext("2d").drawImage(img, 0, 0, canvas.width, canvas.height);
        resolve(canvas.toDataURL("image/jpeg", 0.72));
      };
      img.onerror = reject;
      img.src = reader.result;
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function openForm(id = null) {
  const record = id ? records.find((item) => item.id === id) : null;
  $("recordId").value = record?.id || "";

  $("formTitle").textContent = record ? "Ver y corregir registro" : "Meter dato nuevo";
  $("formKicker").textContent = record ? "REGISTRO EXISTENTE" : "NUEVO REGISTRO";
  $("deleteBtn").classList.toggle("hidden", !record);
  for (const key of fields) $(key).value = safeText(record?.[key]);
  $("visto").checked = Boolean(record?.visto);
  for (const key of statusFields) setStatusButton($(key), record?.[key], key.toUpperCase());
  setPhotoPreview(0, record?.photos?.[0] || "");
  setPhotoPreview(1, record?.photos?.[1] || "");
  showView("form");
}

function collectForm() {
  const record = { ...records.find((item) => item.id === $("recordId").value), id: $("recordId").value || createId(), origen: $("recordId").value ? "editado" : "manual" };
  const existingRecord = records.find((item) => item.id === record.id);
  for (const key of fields) record[key] = $(key).value.trim();
  record.visto = $("visto").checked;
  for (const key of statusFields) record[key] = statusValue($(key).dataset.status);
  record.photos = [...currentPhotos];
  record.cellColors = existingRecord?.cellColors || {};
  return cleanRecord(record);
}

async function saveForm(event) {
  event.preventDefault();
  const record = collectForm();
  const isNewRecord = !$("recordId").value;
  const index = records.findIndex((item) => item.id === record.id);
  if (index >= 0) records[index] = record;
  else records.unshift(record);
  if (record.cantidad) localStorage.setItem(LAST_NUMBER_KEY, record.cantidad);
  await saveRecords();
  if (isNewRecord) openForm();
  else showView("list");
}

async function deleteCurrent() {
  const id = $("recordId").value;
  if (!id) return;
  if (!confirm("¿Seguro que quieres eliminar este registro?")) return;
  records = records.filter((record) => record.id !== id);
  await saveRecords();
  showView("list");
}

async function clearAllRecords() {
  if (!records.length) {
    alert("No hay registros para eliminar.");
    return;
  }
  if (!confirm("¿Seguro que quieres eliminar todos los registros de este cliente?")) return;
  records = [];
  localStorage.removeItem(LAST_NUMBER_KEY);
  await saveRecords();
  renderTable();
  showView("home");
  alert("Registros eliminados.");
}

async function importExcelFile(file) {
  if (!window.ExcelJS) return alert("No se ha cargado el lector de Excel.");
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await file.arrayBuffer());
  const sheet = workbook.worksheets[0];
  if (!sheet) return alert("No encuentro ninguna hoja en ese Excel.");
  const imported = [];
  const headerMap = buildHeaderMap(sheet.getRow(1).values);
  const clientDetails = {};
  const clientColumns = [["name", "cliente"], ["data1", "datos1"], ["data2", "datos2"]];
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const record = rowToImportedRecord(row.values, rowNumber, headerMap);
    const hasData = fields.map((field) => record[field]).some((value) => safeText(value).trim()) || statusFields.some((key) => record[key]);
    if (!hasData) return;
    for (const [key, header] of clientColumns) {
      if (!headerMap[header]) continue;
      const value = excelCellToText(row.getCell(headerMap[header]).value).trim();
      if (clientDetails[key] !== undefined && clientDetails[key] !== value) throw new Error("El Excel contiene datos de varios clientes. Importa un listado por cliente.");
      clientDetails[key] = value;
    }
    imported.push(record);
  });
  if (!imported.length) {
    $("importStatus").textContent = "No se encontraron registros para importar.";
    return alert("No se encontraron registros para importar.");
  }
  const client = clients.find((item) => item.id === activeClientId);
  if (!client) throw new Error("Selecciona un cliente.");
  const detailsSheet = workbook.getWorksheet("Cliente");
  if (detailsSheet) detailsSheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const entry = clientColumns.find(([, header]) => header === normalizeHeader(excelCellToText(row.getCell(1).value)));
    if (entry && clientDetails[entry[0]] === undefined) clientDetails[entry[0]] = excelCellToText(row.getCell(2).value).trim();
  });
  const previous = { name: client.name, data1: client.data1, data2: client.data2, records };
  if (clientDetails.name) client.name = clientDetails.name;
  for (const key of ["data1", "data2"]) if (clientDetails[key] !== undefined) client[key] = clientDetails[key];
  records = [...imported, ...records];
  try { await saveRecords(); }
  catch (error) { Object.assign(client, previous); records = previous.records; throw error; }
  $("clientHeading").textContent = client.name;
  $("importStatus").textContent = `Importados ${imported.length} registros. No se han descartado repetidos.`;
  alert(`Importación correcta.\nRegistros importados: ${imported.length}\nNo se han descartado repetidos.`);
}

function defectFlag(selected, defect) {
  return selected.includes(defect) ? "Sí" : "";
}

async function downloadExcel() {
  if (!window.ExcelJS) return alert("No se ha cargado el generador de Excel.");
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Edificios";
  const sheet = workbook.addWorksheet("Edificios");
  const client = clients.find((item) => item.id === activeClientId);
  const columns = [["cliente", "Cliente"], ["data1", "Datos 1"], ["data2", "Datos 2"], ["edificio", "Edificio"], ["visto", "Visto"], ["ubicacion", "Ubicación"], ["fechaFabricacion", "Información"], ["cantidad", "Defectos"], ["numeroSerie", "Recordar"], ["observaciones", "Observaciones"], ...statusFields.map((key) => [key, key.toUpperCase()]), ["foto1", "Foto 1"], ["foto2", "Foto 2"]];
  sheet.columns = columns.map(([key, header]) => ({ key, header, width: key === "visto" || statusFields.includes(key) ? 12 : 40 }));
  for (const record of records) {
    const row = sheet.addRow({ ...Object.fromEntries(fields.map((key) => [key, safeText(record[key])])), ...Object.fromEntries(statusFields.map((key) => [key, statusLabel(record[key])])), visto: record.visto ? "Sí" : "No", cliente: client?.name || "", data1: client?.data1 || "", data2: client?.data2 || "" });
    for (let index = 0; index < 2; index += 1) {
      const photo = record.photos?.[index];
      if (!photo) continue;
      const imageId = workbook.addImage({ base64: photo, extension: "jpeg" });
      sheet.addImage(imageId, { tl: { col: columns.findIndex(([key]) => key === `foto${index + 1}`), row: row.number - 1 }, ext: { width: 120, height: 85 }, editAs: "oneCell" });
      row.height = 92;
    }
  }
  sheet.getRow(1).font = { bold: true };
  sheet.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFCC99" } };
  sheet.eachRow((row) => { row.alignment = { vertical: "top", wrapText: true }; });
  sheet.views = [{ state: "frozen", ySplit: 1 }];
  sheet.autoFilter = "A1:R1";
  const blob = new Blob([await workbook.xlsx.writeBuffer()], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  const clientName = clients.find((item) => item.id === activeClientId)?.name || "Edificios";
  link.download = clientName.replace(/[<>:"/\\|?*]/g, "_") + "_" + new Date().toISOString().slice(0, 10) + ".xlsx";
  link.click();
  setTimeout(() => URL.revokeObjectURL(link.href), 1000);
}

function bindEvents() {
  $("newClientForm").addEventListener("submit", createClient);
  for (const key of statusFields) bindStatusButton($(key), () => statusValue($(key).dataset.status), (value) => setStatusButton($(key), value, key.toUpperCase()));
  $("openListBtn").addEventListener("click", () => showView("list"));
  $("newRecordBtn").addEventListener("click", () => openForm());
  $("newRecordFromListBtn").addEventListener("click", () => openForm());
  $("downloadExcelBtn").addEventListener("click", downloadExcel);
  $("downloadExcelFromTableBtn").addEventListener("click", downloadExcel);
  $("clearRecordsBtn").addEventListener("click", clearAllRecords);
  $("viewTableFromFormBtn").addEventListener("click", () => showView("list"));
  $("importExcelBtn").addEventListener("click", () => $("importExcelInput").click());
  $("importExcelInput").addEventListener("change", async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    importing = true;
    try {
      $("importStatus").textContent = "Importando Excel...";
      await importExcelFile(file);
      renderTable();
    } catch (error) {
      console.error(error);
      $("importStatus").textContent = "No se ha podido importar el Excel.";
      alert("No se ha podido importar el Excel. Revisa que tenga el mismo formato.");
    } finally {
      importing = false;
      event.target.value = "";
    }
  });
  ["filterEdificio", "filterNumero", "filterSerie", "sortOrder", "seenFilter"].forEach((id) => {
    $(id).addEventListener("input", renderTable);
    $(id).addEventListener("change", renderTable);
  });
  $("recordForm").addEventListener("submit", saveForm);
  $("deleteBtn").addEventListener("click", deleteCurrent);
  [0, 1].forEach((index) => {
    $(`photoInput${index + 1}`).addEventListener("change", async (event) => {
      const file = event.target.files?.[0];
      if (!file) return;
      try {
        setPhotoPreview(index, await resizePhoto(file));
      } catch {
        alert("No he podido cargar esa foto. Prueba con otra imagen.");
      } finally {
        event.target.value = "";
      }
    });
    $(`deletePhoto${index + 1}`).addEventListener("click", () => setPhotoPreview(index, ""));
  });
  document.querySelectorAll("[data-back]").forEach((button) => button.addEventListener("click", () => showView(button.dataset.back)));
}

if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => navigator.serviceWorker.register("./sw.js").catch(() => {}));
}

async function init() {
  await loadRecords();
  bindEvents();
  updateStats();
  showView("clients");
}

init().catch(() => alert("No se han podido cargar los clientes guardados. Recarga la aplicación."));
