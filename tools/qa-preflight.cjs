/** Requiere confirmar explícitamente el script QA antes de usar clasp. No despliega. */
const fs = require("node:fs");
const path = require("node:path");
const root = path.resolve(__dirname, "..");
const pathConfig = path.join(root, ".clasp.json");
// Identificador no secreto: se conserva únicamente para impedir publicaciones QA al original.
const originalScriptId = "1ZOHbILc2lIA9HTsi79gyB6EKlfuPLwhdGbasU0gJd6EruRLOlm-dnVwt";
function fail(message) {
  console.error("QA_PREFLIGHT_FAIL: " + message);
  process.exitCode = 1;
}
if (!fs.existsSync(pathConfig)) {
  fail("Falta .clasp.json local. Crear a partir de .clasp.example.json y usar un script QA independiente.");
} else {
  let config;
  try { config = JSON.parse(fs.readFileSync(pathConfig, "utf8")); }
  catch (_) { fail("Archivo .clasp.json no es JSON válido."); }
  if (!config || typeof config !== "object" || Array.isArray(config)) {
    fail("La configuración debe ser un objeto JSON válido.");
  } else {
    const expected = String(process.env.HOME_QA_SCRIPT_ID || "").trim();
    if (config.scriptId === originalScriptId || expected === originalScriptId) {
      fail("El Apps Script original está prohibido como destino QA. Usar un proyecto independiente.");
    } else if (!expected || expected.length < 20 || expected === "REEMPLAZAR_CON_SCRIPT_ID_QA") {
      fail("Configurar HOME_QA_SCRIPT_ID con el ID verificado del script QA.");
    } else if (config.scriptId !== expected) {
      fail("El scriptId de .clasp.json no coincide con HOME_QA_SCRIPT_ID.");
    } else if (config.rootDir !== "dist" || !fs.existsSync(path.join(root, "dist/appsscript.json"))) {
      fail("El destino debe ser dist y npm run build debe completarse antes.");
    } else {
      console.log("QA_PREFLIGHT_PASS: destino local coincide con el script QA indicado; comprobar SPREADSHEET_ID manualmente antes de publicar.");
    }
  }
}
