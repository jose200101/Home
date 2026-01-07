/**
 * Web App - Control de acceso (Login) usando Google Sheets (pestaña: USUARIOS)
 * Tecnologías: Google Apps Script + HTML Service (Frontend con React + Bootstrap 5 + Babel)
 *
 * IMPORTANTE:
 *  - Configura SPREADSHEET_ID (o deja vacío si este script está "vinculado" al Spreadsheet).
 *  - La hoja debe llamarse exactamente: USUARIOS
 *  - Encabezados esperados (fila 1): id, nombreCompleto, nombre_usuario, contrasenia, rol, estado
 *    (si cambias nombres, se ajusta automáticamente por encabezado).
 */

/** ==== CONFIG ==== */
const CONFIG = {
  // Pega aquí el ID del Spreadsheet (lo que está entre /d/ y /edit en la URL).
  // Ejemplo: https://docs.google.com/spreadsheets/d/XXXXXXXXXXXX/edit
  SPREADSHEET_ID: "",
  SHEET_NAME: "USUARIOS",
  HEADER_ROW: 1,

  // Tiempo (segundos) para mantener token en cache (control de sesión simple)
  TOKEN_TTL_SECONDS: 60 * 60 * 6 // 6 horas
};

/** Web entry */
function doGet() {
  // Soporta proyectos con carpetas (clasp) y proyectos "planos" (editor).
  const tpl = _createTemplateFromAny_([
    "Frontend/Index",
    "Index",
    "Frontend/Index.html",
    "Index.html"
  ]);

  return tpl.evaluate()
    .setTitle("Acceso")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** ==== HTML HELPERS (soporta carpetas Frontend/) ==== */
function _buildHtmlCandidates_(name) {
  const clean = name.replace(/\s+/g, "");
  const base = clean.endsWith(".html") ? clean.slice(0, -5) : clean;

  // Orden de preferencia:
  // 1) tal cual (por si ya viene "Frontend/Index")
  // 2) versión sin .html
  // 3) con carpeta Frontend/
  // 4) con .html (para entornos que sí usan el nombre con extensión)
  const candidates = [];
  const pushUnique = (v) => { if (v && candidates.indexOf(v) === -1) candidates.push(v); };

  pushUnique(clean);
  pushUnique(base);

  if (!base.startsWith("Frontend/")) {
    pushUnique("Frontend/" + base);
    pushUnique("Frontend/" + clean);
  }

  pushUnique(base + ".html");
  pushUnique("Frontend/" + base + ".html");

  return candidates;
}

function _createTemplateFromAny_(candidates) {
  let lastErr = null;
  for (const name of candidates) {
    try {
      return HtmlService.createTemplateFromFile(name);
    } catch (e) {
      lastErr = e;
    }
  }
  throw new Error("No se encontró el archivo HTML principal. Probé: " + candidates.join(", ") + ". Error: " + lastErr);
}

function _createHtmlFromAny_(candidates) {
  let lastErr = null;
  for (const name of candidates) {
    try {
      return HtmlService.createHtmlOutputFromFile(name);
    } catch (e) {
      lastErr = e;
    }
  }
  throw new Error("No se encontró el parcial HTML. Probé: " + candidates.join(", ") + ". Error: " + lastErr);
}


/** Helper para incluir archivos HTML parciales */
function include(filename) {
  const name = String(filename || "").trim();
  const candidates = _buildHtmlCandidates_(name);
  return _createHtmlFromAny_(candidates).getContent();
}

/**
 * Login: valida usuario/contraseña contra la hoja USUARIOS.
 * @param {string} username
 * @param {string} password
 * @returns {{ok:boolean,message?:string,token?:string,user?:object}}
 */
function login(username, password) {
  username = String(username || "").trim();
  password = String(password || "").trim();

  if (!username || !password) {
    return { ok: false, message: "Completa usuario y contraseña." };
  }

  const users = _readUsers_();
  const uNorm = username.toLowerCase();
  const found = users.find(u => String(u.nombre_usuario || "").trim().toLowerCase() === uNorm);

  if (!found) return { ok: false, message: "Usuario o contraseña inválidos." };

  const estado = String(found.estado || "").trim().toLowerCase();
  if (estado && estado !== "activo") {
    return { ok: false, message: "Usuario inactivo. Contacta al administrador." };
  }

  // Comparación simple (tal como viene en tu hoja). Recomendado: guardar hashes.
  const stored = String(found.contrasenia ?? "").trim();
  if (stored !== password) return { ok: false, message: "Usuario o contraseña inválidos." };

  // Crear token y guardarlo en CacheService (sesión simple)
  const token = Utilities.getUuid();
  const cache = CacheService.getScriptCache();
  cache.put(_tokenKey_(token), JSON.stringify({
    id: found.id || "",
    nombreCompleto: found.nombreCompleto || "",
    nombre_usuario: found.nombre_usuario || "",
    rol: found.rol || ""
  }), CONFIG.TOKEN_TTL_SECONDS);

  return {
    ok: true,
    token,
    user: {
      id: found.id || "",
      nombreCompleto: found.nombreCompleto || "",
      nombre_usuario: found.nombre_usuario || "",
      rol: found.rol || ""
    }
  };
}

/**
 * Valida token (para mantener sesión en el frontend).
 * @param {string} token
 * @returns {{ok:boolean,user?:object}}
 */
function validateToken(token) {
  token = String(token || "").trim();
  if (!token) return { ok: false };

  const cache = CacheService.getScriptCache();
  const raw = cache.get(_tokenKey_(token));
  if (!raw) return { ok: false };

  try {
    const user = JSON.parse(raw);
    return { ok: true, user };
  } catch (e) {
    return { ok: false };
  }
}

/**
 * Cierra sesión: elimina el token del cache.
 * @param {string} token
 * @returns {{ok:boolean}}
 */
function logout(token) {
  token = String(token || "").trim();
  if (!token) return { ok: true };

  CacheService.getScriptCache().remove(_tokenKey_(token));
  return { ok: true };
}

/* ================== Internals ================== */

function _tokenKey_(token) {
  return `LOGIN_TOKEN_${token}`;
}

function _getSpreadsheet_() {
  // 0) Si existe conexion() úsala (lee env_() / Script Properties / bound)
  try {
    if (typeof conexion === "function") return conexion();
  } catch (e) {
    // seguimos con otros métodos
  }

  // 1) CONFIG (hardcode)
  const cfgId = String(CONFIG.SPREADSHEET_ID || "").trim();
  if (cfgId) return SpreadsheetApp.openById(cfgId);

  // 2) Script Properties
  const propId = String(PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID") || "").trim();
  if (propId) return SpreadsheetApp.openById(propId);

  // 3) Container-bound (si aplica)
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (e) {}

  throw new Error(
    "No se pudo obtener el Spreadsheet. " +
    "Solución: (a) define el ID en Backend/config/env.gs (ID_DATABASE), o (b) crea una Script Property SPREADSHEET_ID, " +
    "o (c) define CONFIG.SPREADSHEET_ID."
  );
}



function _readUsers_() {
  const ss = _getSpreadsheet_();

  // Preferimos el nombre desde env_() si existe
  let sheetName = CONFIG.SHEET_NAME;
  try {
    if (typeof env_ === "function") {
      const env = env_();
      if (env && env.SH_REGISTRO_USUARIOS) sheetName = String(env.SH_REGISTRO_USUARIOS);
    }
  } catch (e) {}

  const sh = ss.getSheetByName(sheetName) || _getSheetByNameInsensitive_(ss, sheetName);
  if (!sh) throw new Error(`No existe la hoja "${sheetName}".`);

  const values = sh.getDataRange().getDisplayValues();
  if (values.length <= CONFIG.HEADER_ROW) return [];

  const headers = values[CONFIG.HEADER_ROW - 1].map(h => String(h || "").trim());
  const rows = values.slice(CONFIG.HEADER_ROW);

  const idx = {};
  headers.forEach((h, i) => { if (h) idx[h] = i; });

  // Función para obtener valor por encabezado (si no existe, retorna "")
  const get = (row, name) => {
    const i = idx[name];
    return (i === undefined) ? "" : row[i];
  };

  return rows
    .filter(r => r.some(c => String(c || "").trim() !== ""))
    .map(r => ({
      id: get(r, "id"),
      nombreCompleto: get(r, "nombreCompleto"),
      nombre_usuario: get(r, "nombre_usuario"),
      contrasenia: get(r, "contrasenia"),
      rol: get(r, "rol"),
      estado: get(r, "estado"),
    }));
}

/**
 * Busca una hoja por nombre sin distinguir mayúsculas/minúsculas.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} name
 * @return {GoogleAppsScript.Spreadsheet.Sheet|null}
 */
function _getSheetByNameInsensitive_(ss, name) {
  try {
    const target = String(name || "").trim().toLowerCase();
    if (!target) return null;

    const sheets = ss.getSheets();
    return sheets.find(s => String(s.getName() || "").trim().toLowerCase() === target) || null;
  } catch (e) {
    return null;
  }
}



/**
 * Diagnóstico rápido de permisos / cuenta Google.
 * Útil para casos donde en móvil se abre con otra cuenta.
 */
function diagnosticoAcceso() {
  let eff = "";
  let act = "";
  try { eff = Session.getEffectiveUser().getEmail() || ""; } catch (_) {}
  try { act = Session.getActiveUser().getEmail() || ""; } catch (_) {}
  let id = "";
  try {
    if (typeof env_ === "function") {
      const env = env_();
      id = String(env?.ID_DATABASE || "");
    }
  } catch (_) {}
  return {
    spreadsheetId: id,
    effectiveUser: eff,
    activeUser: act,
    scriptTimeZone: Session.getScriptTimeZone(),
    now: new Date().toISOString(),
  };
}
