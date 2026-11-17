/** Genera un directorio desplegable por clasp sin alterar el código fuente. */
const fs = require("node:fs");
const path = require("node:path");
const manifest = require("./build-manifest.cjs");
const root = path.resolve(__dirname, "..");
const outputRoot = path.join(root, "dist");
const seen = new Set();

function safePath(relative) {
  if (!relative || path.isAbsolute(relative) || relative.split(/[\\/]/).includes("..")) {
    throw new Error(`Ruta no permitida en el manifiesto: ${relative}`);
  }
  return path.resolve(root, relative);
}

for (const [source, destination] of manifest) {
  safePath(source);
  safePath(destination);
  if (!source.startsWith("src/") || seen.has(destination)) {
    throw new Error(`Entrada inválida o duplicada: ${source} -> ${destination}`);
  }
  if (!fs.existsSync(safePath(source))) {
    throw new Error(`Falta archivo fuente: ${source}`);
  }
  seen.add(destination);
}

const config = JSON.parse(fs.readFileSync(path.join(root, "src/gas/appsscript.json"), "utf8"));
if (config.runtimeVersion !== "V8") throw new Error("Apps Script debe utilizar V8.");

fs.rmSync(outputRoot, { recursive: true, force: true });
for (const [source, destination] of manifest) {
  const to = path.join(outputRoot, destination);
  fs.mkdirSync(path.dirname(to), { recursive: true });
  fs.copyFileSync(safePath(source), to);
}
console.log(`HOME_BUILD_OK: ${manifest.length} archivos en dist/`);
