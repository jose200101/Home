const { test } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const { spawnSync } = require("node:child_process");
const { assertDeploymentPolicy, checkBaseManifest } = require("../tools/deployment-policy.cjs");
const policy = require("../tools/deployment-policy.json");
const root = path.resolve(__dirname, "..");

test("el manifiesto base y su política mantienen acceso exclusivo al propietario", () => {
  assert.equal(checkBaseManifest().webapp.access, "MYSELF");
  assert.deepEqual(policy.allowedBaseWebappAccess, ["MYSELF"]);
  assert.deepEqual(policy.allowedBaseExecutionApiAccess, ["MYSELF"]);
});

test("cambiar solo el manifiesto no permite acceso público, de dominio ni implícito", () => {
  for (const access of ["ANYONE", "ANYONE_ANONYMOUS", "DOMAIN", "", undefined, "myself"]) {
    assert.throws(() => assertDeploymentPolicy({ webapp: { access } }, policy), /DEPLOYMENT_POLICY_FAIL/);
  }
  for (const manifest of [{}, { webapp: null }, null]) {
    assert.throws(() => assertDeploymentPolicy(manifest, policy), /DEPLOYMENT_POLICY_FAIL/);
  }
});

test("una API ejecutable tampoco puede introducir acceso más amplio", () => {
  const manifest = { webapp: { access: "MYSELF" }, executionApi: { access: "MYSELF" } };
  assert.doesNotThrow(() => assertDeploymentPolicy(manifest, policy));
  for (const access of ["ANYONE", "ANYONE_ANONYMOUS", "DOMAIN", undefined]) {
    assert.throws(() => assertDeploymentPolicy({ ...manifest, executionApi: { access } }, policy), /executionApi/);
  }
});

test("una ampliación exige cambiar explícitamente la política, no solo el manifiesto", () => {
  const publicManifest = { webapp: { access: "ANYONE" } };
  assert.throws(() => assertDeploymentPolicy(publicManifest, policy));
  const reviewedPolicy = { ...policy, allowedBaseWebappAccess: ["MYSELF", "ANYONE"] };
  assert.doesNotThrow(() => assertDeploymentPolicy(publicManifest, reviewedPolicy));
  for (const invalidPolicy of [undefined, { ...policy, version: 2 }, { ...policy, allowedBaseWebappAccess: [] }]) {
    assert.throws(() => assertDeploymentPolicy(publicManifest, invalidPolicy), /política|lista/);
  }
});

test("un build con manifiesto público falla antes de reemplazar el artefacto existente", (t) => {
  const fixture = fs.mkdtempSync(path.join(os.tmpdir(), "home-deployment-test-"));
  t.after(() => fs.rmSync(fixture, { recursive: true, force: true }));
  for (const dir of ["tools", "src/gas", "dist"]) fs.mkdirSync(path.join(fixture, dir), { recursive: true });
  for (const file of ["build.cjs", "deployment-policy.cjs", "deployment-policy.json"]) {
    fs.copyFileSync(path.join(root, "tools", file), path.join(fixture, "tools", file));
  }
  fs.writeFileSync(path.join(fixture, "tools/build-manifest.cjs"), 'module.exports = [["src/gas/appsscript.json", "appsscript.json"]];');
  fs.writeFileSync(path.join(fixture, "src/gas/appsscript.json"), JSON.stringify({ runtimeVersion: "V8", webapp: { access: "ANYONE" } }));
  const sentinel = path.join(fixture, "dist/previous-build.txt");
  fs.writeFileSync(sentinel, "conservar");
  const result = spawnSync(process.execPath, [path.join(fixture, "tools/build.cjs")], { encoding: "utf8" });
  assert.equal(result.status, 1);
  assert.match(result.stderr, /DEPLOYMENT_POLICY_FAIL/);
  assert.equal(fs.readFileSync(sentinel, "utf8"), "conservar");
  assert.equal(fs.existsSync(path.join(fixture, "dist/appsscript.json")), false);
});
