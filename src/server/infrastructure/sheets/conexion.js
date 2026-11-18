/**
 * conexion_
 * Retorna el Spreadsheet "base de datos".
 * @return {GoogleAppsScript.Spreadsheet.Spreadsheet}
 */
function conexion_() {
  const id = String(env_().ID_DATABASE || "").trim();

  if (!id) {
    throw new Error("Configura SPREADSHEET_ID en Script Properties antes de conectar.");
  }
  try {
    return SpreadsheetApp.openById(id);
  } catch (e) {
    throw new Error(
      "No se pudo abrir el Spreadsheet configurado. Verifica SPREADSHEET_ID y los permisos de la cuenta de ejecución. No se intentó otra conexión."
    );
  }
}

/**
 * Devuelve una hoja por nombre, con fallback case-insensitive.
 * @param {String} NAME nombre de la hoja
 * @return {GoogleAppsScript.Spreadsheet.Sheet}
 */
function obtenerSheet_(NAME) {
  const ss = conexion_();
  const name = String(NAME || "").trim();
  if (!name) throw new Error("obtenerSheet_(NAME): NAME vacío.");

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
function obtenerDatos_(NAME) {
  return obtenerSheet_(NAME).getDataRange().getDisplayValues();
}
