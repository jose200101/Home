const { test } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
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

test("el layout y los breakpoints de HOME tienen alcance independiente", () => {
  const css = read("dist/Frontend/_HomeStyles.html").toString("utf8")
    .replace(/\/\*[\s\S]*?\*\/|"(?:\\.|[^"\\])*"|'(?:\\.|[^'\\])*'/g, "");
  const blocks = [];
  const topLevel = new Set();
  let start = 0;
  for (const match of css.matchAll(/[{}]/g)) {
    if (match[0] === "{") {
      const rule = css.slice(start, match.index).replace(/^<style>/, "").trim();
      if (blocks.length === 0) topLevel.add(rule);
      blocks.push(rule);
    } else {
      assert.ok(blocks.length > 0, "Llave de cierre sin bloque CSS");
      blocks.pop();
    }
    start = match.index + 1;
  }
  assert.deepEqual(blocks, [], "Bloques CSS sin cerrar");
  for (const rule of [
    ".home-layout",
    "@media (max-width: 600px)",
    "@media (min-width: 900px)",
    "@media (max-width: 760px)",
    "@media (prefers-reduced-motion: reduce)",
  ]) {
    assert.ok(topLevel.has(rule), rule + " debe estar en el nivel raíz");
  }
});

test("doGet agrega el viewport a HtmlOutput y conserva el título", () => {
  const meta = [];
  const output = {
    addMetaTag(name, content) {
      meta.push([name, content]);
      return this;
    },
    setTitle(title) {
      this.title = title;
      return this;
    },
  };
  const context = {
    HtmlService: {
      createTemplateFromFile(name) {
        assert.equal(name, "Frontend/Index");
        return { evaluate: () => output };
      },
    },
  };
  vm.runInNewContext(read("dist/Backend/Code.js").toString("utf8"), context);
  assert.equal(context.doGet(), output);
  assert.deepEqual(meta, [["viewport", "width=device-width, initial-scale=1"]]);
  assert.equal(output.title, "Home · Cuentas del hogar");
});
