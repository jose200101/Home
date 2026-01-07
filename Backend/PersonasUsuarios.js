/**
 * CRUD Personas y Usuarios
 * - PERSONAS: id_persona, nombre_persona, fecha_nacimiento, genero, telefono, estado
 * - USUARIOS: id, nombre_persona, nombre_usuario, contrasenia, rol, estado
 */

function _normHeader_(s) {
  return String(s || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, "_");
}

function _buildHeaderIndex_(headers) {
  const map = {};
  (headers || []).forEach((h, i) => {
    const k = _normHeader_(h);
    if (k) map[k] = i;
  });
  return map;
}

function _pickIdx_(idx, aliases) {
  for (const a of aliases) {
    const k = _normHeader_(a);
    if (idx.hasOwnProperty(k)) return idx[k];
  }
  return -1;
}

function _readSheetObjects_(sheetName) {
  const sh = obtenerSheet(sheetName);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) {
    return { sh, headers: [], idx: {}, values: [], rowStart: 2 };
  }
  const headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(h => String(h || "").trim());
  const idx = _buildHeaderIndex_(headers);
  const values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  return { sh, headers, idx, values, rowStart: 2 };
}

function _isRowEmpty_(row) {
  return !Array.isArray(row) || row.every(v => String(v ?? "").trim() === "");
}

function _toISODate_(v) {
  if (!v) return "";
  if (Object.prototype.toString.call(v) === "[object Date]") {
    if (isNaN(v.getTime())) return "";
    const y = v.getFullYear();
    const m = ("0" + (v.getMonth() + 1)).slice(-2);
    const d = ("0" + v.getDate()).slice(-2);
    return `${y}-${m}-${d}`;
  }
  const s = String(v).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  // dd-mm-yyyy o dd/mm/yyyy
  const m = s.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
  if (m) {
    const dd = ("0" + m[1]).slice(-2);
    const mm = ("0" + m[2]).slice(-2);
    const yy = m[3];
    return `${yy}-${mm}-${dd}`;
  }
  // Intento parse
  const dt = new Date(s);
  if (dt.toString() !== "Invalid Date") return _toISODate_(dt);
  return s;
}

function _findRowById_(values, idIdx, idValue) {
  const target = String(idValue || "").trim();
  if (!target || idIdx < 0) return -1;
  for (let r = 0; r < values.length; r++) {
    const cell = values[r][idIdx];
    if (String(cell ?? "").trim() === target) return r;
  }
  return -1;
}

function _nextNumericId_(values, idIdx) {
  let maxId = 0;
  if (idIdx < 0) return 1;
  values.forEach(row => {
    const n = parseInt(String(row[idIdx] ?? "").trim(), 10);
    if (!isNaN(n)) maxId = Math.max(maxId, n);
  });
  return maxId + 1;
}

// =====================
// PERSONAS
// =====================

function listarPersonas(params) {
  params = params || {};
  const env = (typeof env_ === "function") ? env_() : {};
  const sheetName = env.SH_PERSONAS || "PERSONAS";
  const { headers, idx, values } = _readSheetObjects_(sheetName);

  const iId = _pickIdx_(idx, ["id_persona", "id", "persona_id"]);
  const iNombre = _pickIdx_(idx, ["nombre_persona", "nombre", "persona"]);
  const iFecha = _pickIdx_(idx, ["fecha_nacimiento", "fecha_nacimient", "fecha_nac", "fecha"]);
  const iGenero = _pickIdx_(idx, ["genero", "género"]);
  const iTel = _pickIdx_(idx, ["telefono", "teléfono", "celular"]);
  const iEstado = _pickIdx_(idx, ["estado"]);

  const q = String(params.q || "").trim().toLowerCase();
  const estado = String(params.estado || "").trim().toLowerCase();

  const out = [];
  for (const row of values) {
    if (_isRowEmpty_(row)) continue;
    const item = {
      id_persona: iId >= 0 ? String(row[iId] ?? "").trim() : "",
      nombre_persona: iNombre >= 0 ? String(row[iNombre] ?? "").trim() : "",
      fecha_nacimiento: iFecha >= 0 ? _toISODate_(row[iFecha]) : "",
      genero: iGenero >= 0 ? String(row[iGenero] ?? "").trim() : "",
      telefono: iTel >= 0 ? String(row[iTel] ?? "").trim() : "",
      estado: iEstado >= 0 ? String(row[iEstado] ?? "").trim() : ""
    };

    if (estado && String(item.estado || "").toLowerCase() !== estado) continue;
    if (q) {
      const hay = `${item.id_persona} ${item.nombre_persona} ${item.telefono} ${item.genero}`.toLowerCase();
      if (!hay.includes(q)) continue;
    }
    out.push(item);
  }

  // Orden por id asc (numérico si aplica)
  out.sort((a, b) => {
    const na = parseInt(a.id_persona, 10);
    const nb = parseInt(b.id_persona, 10);
    if (!isNaN(na) && !isNaN(nb)) return na - nb;
    return String(a.id_persona).localeCompare(String(b.id_persona));
  });
  return out;
}

function listarPersonasActivas() {
  const all = listarPersonas({ estado: "Activo" });
  return (all || []).map(p => ({ id_persona: p.id_persona, nombre_persona: p.nombre_persona }));
}

function guardarPersona(data) {
  data = data || {};
  // Nota: este proyecto es un WebApp (script independiente). En scripts NO vinculados
  // a un Spreadsheet/Doc, getDocumentLock() retorna null y provoca el error:
  // "Cannot read properties of null (reading 'waitLock')".
  // Por eso usamos ScriptLock, que funciona en cualquier contexto.
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const env = (typeof env_ === "function") ? env_() : {};
    const sheetName = env.SH_PERSONAS || "PERSONAS";
    const { sh, idx, values } = _readSheetObjects_(sheetName);

    const iId = _pickIdx_(idx, ["id_persona", "id", "persona_id"]);
    const iNombre = _pickIdx_(idx, ["nombre_persona", "nombre", "persona"]);
    const iFecha = _pickIdx_(idx, ["fecha_nacimiento", "fecha_nacimient", "fecha_nac", "fecha"]);
    const iGenero = _pickIdx_(idx, ["genero", "género"]);
    const iTel = _pickIdx_(idx, ["telefono", "teléfono", "celular"]);
    const iEstado = _pickIdx_(idx, ["estado"]);

    const nombre = String(data.nombre_persona || "").trim();
    if (!nombre) throw new Error("El nombre de la persona es requerido.");

    let id = String(data.id_persona || "").trim();
    const isNew = !id;
    if (isNew) {
      // Genera consecutivo
      const next = _nextNumericId_(values, iId);
      id = String(next);
    }

    const rowIdx = _findRowById_(values, iId, id);

    const payload = {
      id_persona: id,
      nombre_persona: nombre,
      fecha_nacimiento: _toISODate_(data.fecha_nacimiento),
      genero: String(data.genero || "").trim(),
      telefono: String(data.telefono || "").trim(),
      estado: String(data.estado || "Activo").trim() || "Activo"
    };

    const lastCol = sh.getLastColumn();
    const headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
    const hIdx = _buildHeaderIndex_(headers);

    const colId = _pickIdx_(hIdx, ["id_persona", "id", "persona_id"]);
    const colNombre = _pickIdx_(hIdx, ["nombre_persona", "nombre", "persona"]);
    const colFecha = _pickIdx_(hIdx, ["fecha_nacimiento", "fecha_nacimient", "fecha_nac", "fecha"]);
    const colGenero = _pickIdx_(hIdx, ["genero", "género"]);
    const colTel = _pickIdx_(hIdx, ["telefono", "teléfono", "celular"]);
    const colEstado = _pickIdx_(hIdx, ["estado"]);

    const setCell = (arr, col, val) => { if (col >= 0) arr[col] = val; };

    if (rowIdx >= 0) {
      // Update
      const existing = sh.getRange(rowIdx + 2, 1, 1, lastCol).getValues()[0];
      const row = existing.slice();
      setCell(row, colId, payload.id_persona);
      setCell(row, colNombre, payload.nombre_persona);
      setCell(row, colFecha, payload.fecha_nacimiento);
      setCell(row, colGenero, payload.genero);
      setCell(row, colTel, payload.telefono);
      setCell(row, colEstado, payload.estado);
      sh.getRange(rowIdx + 2, 1, 1, lastCol).setValues([row]);
    } else {
      // Create
      const row = new Array(lastCol).fill("");
      setCell(row, colId, payload.id_persona);
      setCell(row, colNombre, payload.nombre_persona);
      setCell(row, colFecha, payload.fecha_nacimiento);
      setCell(row, colGenero, payload.genero);
      setCell(row, colTel, payload.telefono);
      setCell(row, colEstado, payload.estado);
      sh.appendRow(row);
    }

    return { ok: true, id_persona: payload.id_persona };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function eliminarPersona(id_persona) {
  const env = (typeof env_ === "function") ? env_() : {};
  const sheetName = env.SH_PERSONAS || "PERSONAS";
  const { sh, idx, values } = _readSheetObjects_(sheetName);

  const iId = _pickIdx_(idx, ["id_persona", "id", "persona_id"]);
  const rowIdx = _findRowById_(values, iId, id_persona);
  if (rowIdx < 0) return { ok: true };

  sh.deleteRow(rowIdx + 2);
  return { ok: true };
}

// =====================
// USUARIOS
// =====================

function listarUsuarios(params) {
  params = params || {};
  const env = (typeof env_ === "function") ? env_() : {};
  const sheetName = env.SH_REGISTRO_USUARIOS || "USUARIOS";
  const { idx, values } = _readSheetObjects_(sheetName);

  const iId = _pickIdx_(idx, ["id"]);
  const iPersona = _pickIdx_(idx, ["nombre_persona", "persona", "nombrecompleto", "nombre_completo"]);
  const iUser = _pickIdx_(idx, ["nombre_usuario", "usuario"]);
  const iPass = _pickIdx_(idx, ["contrasenia", "contraseña", "password"]);
  const iRol = _pickIdx_(idx, ["rol", "role"]);
  const iEstado = _pickIdx_(idx, ["estado"]);

  const q = String(params.q || "").trim().toLowerCase();
  const estado = String(params.estado || "").trim().toLowerCase();

  const out = [];
  for (const row of values) {
    if (_isRowEmpty_(row)) continue;
    const item = {
      id: iId >= 0 ? String(row[iId] ?? "").trim() : "",
      nombre_persona: iPersona >= 0 ? String(row[iPersona] ?? "").trim() : "",
      nombre_usuario: iUser >= 0 ? String(row[iUser] ?? "").trim() : "",
      contrasenia: iPass >= 0 ? String(row[iPass] ?? "").trim() : "",
      rol: iRol >= 0 ? String(row[iRol] ?? "").trim() : "",
      estado: iEstado >= 0 ? String(row[iEstado] ?? "").trim() : ""
    };

    if (estado && String(item.estado || "").toLowerCase() !== estado) continue;
    if (q) {
      const hay = `${item.id} ${item.nombre_persona} ${item.nombre_usuario} ${item.rol}`.toLowerCase();
      if (!hay.includes(q)) continue;
    }
    // Por seguridad, no enviamos password al frontend
    item.contrasenia = "";
    out.push(item);
  }

  out.sort((a, b) => String(a.nombre_usuario || "").localeCompare(String(b.nombre_usuario || "")));
  return out;
}

function guardarUsuario(data) {
  data = data || {};
  // WebApp (script independiente): usar ScriptLock para evitar null en DocumentLock
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const env = (typeof env_ === "function") ? env_() : {};
    const sheetName = env.SH_REGISTRO_USUARIOS || "USUARIOS";
    const { sh, idx, values } = _readSheetObjects_(sheetName);

    const iId = _pickIdx_(idx, ["id"]);
    const iPersona = _pickIdx_(idx, ["nombre_persona", "persona", "nombrecompleto", "nombre_completo"]);
    const iUser = _pickIdx_(idx, ["nombre_usuario", "usuario"]);
    const iPass = _pickIdx_(idx, ["contrasenia", "contraseña", "password"]);
    const iRol = _pickIdx_(idx, ["rol", "role"]);
    const iEstado = _pickIdx_(idx, ["estado"]);

    const nombre_usuario = String(data.nombre_usuario || "").trim();
    if (!nombre_usuario) throw new Error("El nombre de usuario es requerido.");

    let id = String(data.id || "").trim();
    const isNew = !id;
    if (isNew) id = Utilities.getUuid();

    // Username único (case-insensitive)
    const normUser = nombre_usuario.toLowerCase();
    for (const row of values) {
      if (_isRowEmpty_(row)) continue;
      const rid = iId >= 0 ? String(row[iId] ?? "").trim() : "";
      const u = iUser >= 0 ? String(row[iUser] ?? "").trim().toLowerCase() : "";
      if (u === normUser && rid !== id) {
        throw new Error("Ya existe un usuario con ese nombre. Usa otro.");
      }
    }

    const rowIdx = _findRowById_(values, iId, id);
    const lastCol = sh.getLastColumn();
    const headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
    const hIdx = _buildHeaderIndex_(headers);

    const colId = _pickIdx_(hIdx, ["id"]);
    const colPersona = _pickIdx_(hIdx, ["nombre_persona", "persona", "nombrecompleto", "nombre_completo"]);
    const colUser = _pickIdx_(hIdx, ["nombre_usuario", "usuario"]);
    const colPass = _pickIdx_(hIdx, ["contrasenia", "contraseña", "password"]);
    const colRol = _pickIdx_(hIdx, ["rol", "role"]);
    const colEstado = _pickIdx_(hIdx, ["estado"]);

    const payload = {
      id,
      nombre_persona: String(data.nombre_persona || "").trim(),
      nombre_usuario,
      // contrasenia: en edición, si viene vacía, se conserva
      contrasenia: String(data.contrasenia || ""),
      rol: String(data.rol || "").trim(),
      estado: String(data.estado || "Activo").trim() || "Activo"
    };

    if (!payload.nombre_persona) throw new Error("Selecciona una persona.");
    if (!payload.rol) throw new Error("El rol es requerido.");
    if (isNew && !String(payload.contrasenia || "").trim()) throw new Error("La contraseña es requerida para crear el usuario.");

    const setCell = (arr, col, val) => { if (col >= 0) arr[col] = val; };

    if (rowIdx >= 0) {
      const existing = sh.getRange(rowIdx + 2, 1, 1, lastCol).getValues()[0];
      const row = existing.slice();
      setCell(row, colId, payload.id);
      setCell(row, colPersona, payload.nombre_persona);
      setCell(row, colUser, payload.nombre_usuario);
      if (String(payload.contrasenia || "").trim()) setCell(row, colPass, payload.contrasenia);
      setCell(row, colRol, payload.rol);
      setCell(row, colEstado, payload.estado);
      sh.getRange(rowIdx + 2, 1, 1, lastCol).setValues([row]);
    } else {
      const row = new Array(lastCol).fill("");
      setCell(row, colId, payload.id);
      setCell(row, colPersona, payload.nombre_persona);
      setCell(row, colUser, payload.nombre_usuario);
      setCell(row, colPass, payload.contrasenia);
      setCell(row, colRol, payload.rol);
      setCell(row, colEstado, payload.estado);
      sh.appendRow(row);
    }

    return { ok: true, id: payload.id };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function eliminarUsuario(id) {
  const env = (typeof env_ === "function") ? env_() : {};
  const sheetName = env.SH_REGISTRO_USUARIOS || "USUARIOS";
  const { sh, idx, values } = _readSheetObjects_(sheetName);

  const iId = _pickIdx_(idx, ["id"]);
  const rowIdx = _findRowById_(values, iId, id);
  if (rowIdx < 0) return { ok: true };

  sh.deleteRow(rowIdx + 2);
  return { ok: true };
}
