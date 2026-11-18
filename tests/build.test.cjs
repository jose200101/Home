const { test } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const entries = require("../tools/build-manifest.cjs");
const root = path.resolve(__dirname, "..");
const read = (file) => fs.readFileSync(path.join(root, file));

test("build reproducible: cada fuente llega a dist sin modificaciones", () => {
  for (const [source, target] of entries) {
    assert.deepEqual(read(`dist/${target}`), read(source), `${source} -> ${target}`);
  }
  assert.equal(new Set(entries.map(([, target]) => target)).size, entries.length);
});

test("clasp solo publica el directorio generado", () => {
  const clasp = JSON.parse(read(".clasp.example.json").toString("utf8"));
  assert.equal(clasp.rootDir, "dist");
  assert.equal(clasp.scriptId, "REEMPLAZAR_CON_SCRIPT_ID_QA");
  assert.match(read(".gitignore").toString("utf8"), /^\.clasp\.json$/m);
  assert.equal(clasp.skipSubdirectories, false);
  assert.equal(JSON.parse(read("dist/appsscript.json").toString("utf8")).runtimeVersion, "V8");
});

test("las plantillas incluidas por Index existen en dist", () => {
  const html = read("dist/Frontend/Index.html").toString("utf8");
  const includes = [...html.matchAll(/include\('([^']+)'\)/g)].map((m) => m[1]);
  assert.ok(includes.length > 0, "Index debe tener plantillas incluidas");
  for (const name of includes) {
    assert.ok(fs.existsSync(path.join(root, "dist/Frontend", name + ".html")), name);
  }
  assert.match(read("dist/Backend/Code.js").toString("utf8"), /function doGet\(/);
});
