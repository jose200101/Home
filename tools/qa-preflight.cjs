/** Requiere confirmar explícitamente el script QA antes de usar clasp. No despliega. */
const fs = require("node:fs");
const path = require("node:path");
const root = path.resolve(__dirname, "..");
const pathConfig = path.join(root, ".clasp.json");
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
  if (config) {
    const expected = process.env.HOME_QA_SCRIPT_ID || "";
    if (!expected || expected.length < 20) {
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
