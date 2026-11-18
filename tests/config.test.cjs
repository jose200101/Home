const { test } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const vm = require("node:vm");
const { spawnSync } = require("node:child_process");

const root = path.resolve(__dirname, "..");
const originalScriptId = "1ZOHbILc2lIA9HTsi79gyB6EKlfuPLwhdGbasU0gJd6EruRLOlm-dnVwt";
const qaScriptId = "qa_script_ficticio_independiente_12345";

// Ejecuta la CLI real sobre archivos locales desechables; no usa clasp ni Google.
function preflight(t, config, expected = qaScriptId, built = true) {
  const fixture = fs.mkdtempSync(path.join(os.tmpdir(), "home-config-test-"));
  t.after(() => fs.rmSync(fixture, { recursive: true, force: true }));
  fs.mkdirSync(path.join(fixture, "tools"));
  fs.copyFileSync(path.join(root, "tools/qa-preflight.cjs"), path.join(fixture, "tools/qa-preflight.cjs"));
  if (config !== undefined) fs.writeFileSync(path.join(fixture, ".clasp.json"), JSON.stringify(config));
  if (built) {
    fs.mkdirSync(path.join(fixture, "dist"));
    fs.writeFileSync(path.join(fixture, "dist/appsscript.json"), "{}");
  }
  return spawnSync(process.execPath, [path.join(fixture, "tools/qa-preflight.cjs")], {
    encoding: "utf8",
    env: { ...process.env, HOME_QA_SCRIPT_ID: expected },
  });
}

test("preflight bloquea el original aunque coincidan configuración y variable QA", (t) => {
  const result = preflight(t, { scriptId: originalScriptId, rootDir: "dist" }, originalScriptId);
  assert.equal(result.status, 1);
  assert.match(result.stderr, /original está prohibido/);
  assert.doesNotMatch(result.stdout, /QA_PREFLIGHT_PASS/);
});

test("preflight rechaza el original en cualquiera de las dos entradas", (t) => {
  for (const [scriptId, expected] of [[originalScriptId, qaScriptId], [qaScriptId, originalScriptId]]) {
    const result = preflight(t, { scriptId, rootDir: "dist" }, expected);
    assert.equal(result.status, 1);
    assert.match(result.stderr, /original está prohibido/);
  }
});

test("preflight acepta únicamente un destino explícito coincidente con build", (t) => {
  const config = { scriptId: qaScriptId, rootDir: "dist" };
  assert.equal(preflight(t, config).status, 0);
  for (const result of [
    preflight(t, undefined),
    preflight(t, null),
    preflight(t, config, ""),
    preflight(t, { ...config, scriptId: "otro_script_qa_ficticio_12345" }),
    preflight(t, { ...config, rootDir: "src" }),
    preflight(t, config, qaScriptId, false),
    preflight(t, { scriptId: "REEMPLAZAR_CON_SCRIPT_ID_QA", rootDir: "dist" }, "REEMPLAZAR_CON_SCRIPT_ID_QA"),
  ]) assert.equal(result.status, 1);
});

function connection(property) {
  const opened = [];
  let activeReads = 0;
  const spreadsheet = { fixture: true };
  const context = vm.createContext({
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => property }) },
    SpreadsheetApp: {
      openById: (id) => { opened.push(id); return spreadsheet; },
      getActiveSpreadsheet: () => { activeReads++; return { wrong: true }; },
    },
  });
  for (const file of ["Backend/config/env.js", "Backend/database/conexion.js", "Backend/Code.js"]) {
    vm.runInContext(fs.readFileSync(path.join(root, "dist", file), "utf8"), context, { filename: file });
  }
  return { context, opened, spreadsheet, get activeReads() { return activeReads; } };
}

test("sin SPREADSHEET_ID explícito todas las rutas fallan sin abrir otro Spreadsheet", () => {
  for (const property of [null, "", "   ", "https://docs.google.com/spreadsheets/d/ficticio/edit"]) {
    const h = connection(property);
    for (const run of [() => h.context.env_(), () => h.context.conexion_(), () => h.context._getSpreadsheet_()]) {
      assert.throws(run, /SPREADSHEET_ID/);
    }
    assert.deepEqual(h.opened, []);
    assert.equal(h.activeReads, 0);
  }
});

test("conexión y lector de usuarios usan exclusivamente el ID de Script Properties", () => {
  const h = connection("  qa_sheet_ficticia  ");
  // Una configuración heredada en memoria tampoco debe habilitar un fallback.
  vm.runInContext("CONFIG.SPREADSHEET_ID = 'otro_destino_ficticio'", h.context);
  assert.equal(h.context.conexion_(), h.spreadsheet);
  assert.equal(h.context._getSpreadsheet_(), h.spreadsheet);
  assert.deepEqual(h.opened, ["qa_sheet_ficticia", "qa_sheet_ficticia"]);
  assert.equal(h.activeReads, 0);
});

test("un error de acceso detiene la conexión sin reintentos ni fallback activo", () => {
  const h = connection("qa_sheet_ficticia");
  let attempts = 0;
  h.context.SpreadsheetApp.openById = () => { attempts++; throw Error("detalle privado del proveedor"); };
  assert.throws(() => h.context._getSpreadsheet_(), (error) => {
    assert.match(error.message, /No se pudo abrir el Spreadsheet configurado/);
    assert.doesNotMatch(error.message, /detalle privado/);
    return true;
  });
  assert.equal(attempts, 1);
  assert.equal(h.activeReads, 0);
});

test("el lector propaga errores de conexión y no intenta una ruta alternativa", () => {
  const h = connection("qa_sheet_ficticia");
  const failure = new Error("conexión bloqueada");
  h.context.conexion_ = () => { throw failure; };
  assert.throws(() => h.context._getSpreadsheet_(), (error) => error === failure);
  assert.deepEqual(h.opened, []);
  assert.equal(h.activeReads, 0);
});
