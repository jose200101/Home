/**
 * conexion
 * Retorna el Spreadsheet "base de datos".
 * @return {GoogleAppsScript.Spreadsheet.Spreadsheet}
 */
function conexion() {
  const id = String(env_().ID_DATABASE || "").trim();

  if (id) {
  try {
    return SpreadsheetApp.openById(id);
  } catch (e) {
    // Este error suele ocurrir cuando la WebApp está desplegada como
    // "Ejecutar como: Usuario que accede" y la cuenta Google del navegador
    // no tiene acceso al Spreadsheet (muy común en móviles con otra cuenta).
    let eff = "";
    let act = "";
    try { eff = Session.getEffectiveUser().getEmail() || ""; } catch (_) {}
    try { act = Session.getActiveUser().getEmail() || ""; } catch (_) {}

    throw new Error(
      "No cuentas con el permiso necesario para acceder al documento solicitado. " +
      "ID Spreadsheet: " + id + ". " +
      "Cuenta (effective): " + (eff || "(no disponible)") + ", " +
      "Cuenta (active): " + (act || "(no disponible)") + ". " +
      "Solución: (1) Re-despliega la WebApp con 'Ejecutar como: Yo (propietario)' o " +
      "(2) comparte el Spreadsheet con la cuenta Google que abre la WebApp."
    );
  }
}

// Fallback: si el proyecto está vinculado al Spreadsheet (container-bound)
  // y no configuraste ID_DATABASE/SPREADSHEET_ID.
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (e) {
    // ignoramos y lanzamos un error más claro abajo
  }

  throw new Error(
    "No hay ID_DATABASE en env_(). " +
    "Solución: (a) pega el ID en Backend/config/env.gs (ID_DATABASE), o " +
    "(b) crea una Script Property SPREADSHEET_ID."
  );
}

/**
 * Devuelve una hoja por nombre, con fallback case-insensitive.
 * @param {String} NAME nombre de la hoja
 * @return {GoogleAppsScript.Spreadsheet.Sheet}
 */
function obtenerSheet(NAME) {
  const ss = conexion();
  const name = String(NAME || "").trim();
  if (!name) throw new Error("obtenerSheet(NAME): NAME vacío.");

  // Primero, intento directo
  let sh = ss.getSheetByName(name);
  if (sh) return sh;

  // Fallback: buscar por nombre sin distinguir mayúsculas/minúsculas
  const target = name.toLowerCase();
  const all = ss.getSheets();
  sh = all.find(s => String(s.getName() || "").toLowerCase() === target) || null;
  if (sh) return sh;

  throw new Error(`No existe la hoja "${name}". Revisa el nombre en env_() o en el Spreadsheet.`);
}

/**
 * Retorna todos los datos como arreglo bidimensional (DisplayValues).
 * @param {String} NAME nombre de la hoja
 * @return {Array<Array<string>>}
 */
function obtenerDatos(NAME) {
  return obtenerSheet(NAME).getDataRange().getDisplayValues();
}
