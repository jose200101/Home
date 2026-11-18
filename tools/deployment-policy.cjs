/** Política del manifiesto base; no lee credenciales ni publica recursos. */
const fs = require("node:fs");
const path = require("node:path");

function assertDeploymentPolicy(manifest, policy) {
  const knownAccess = new Set(["MYSELF", "DOMAIN", "ANYONE", "ANYONE_ANONYMOUS"]);
  if (!policy || policy.version !== 1) throw new Error("DEPLOYMENT_POLICY_FAIL: política ausente o versión inválida.");
  for (const key of ["allowedBaseWebappAccess", "allowedBaseExecutionApiAccess"]) {
    if (!Array.isArray(policy[key]) || !policy[key].length || policy[key].some(value => !knownAccess.has(value))) {
      throw new Error("DEPLOYMENT_POLICY_FAIL: lista de acceso inválida en " + key);
    }
  }
  function check(resource, allowed) {
    const access = manifest && manifest[resource] && manifest[resource].access;
    if (!knownAccess.has(access) || !allowed.includes(access)) {
      throw new Error("DEPLOYMENT_POLICY_FAIL: " + resource + ".access=" + String(access) +
        " no está permitido. El manifiesto base debe respetar tools/deployment-policy.json; ampliar acceso requiere modificar y revisar explícitamente esa política.");
    }
  }
  check("webapp", policy.allowedBaseWebappAccess);
  if (manifest.executionApi !== undefined) check("executionApi", policy.allowedBaseExecutionApiAccess);
}

function checkBaseManifest() {
  const root = path.resolve(__dirname, "..");
  const manifest = JSON.parse(fs.readFileSync(path.join(root, "src/gas/appsscript.json"), "utf8"));
  const policy = JSON.parse(fs.readFileSync(path.join(__dirname, "deployment-policy.json"), "utf8"));
  assertDeploymentPolicy(manifest, policy);
  return manifest;
}

if (require.main === module) {
  try {
    checkBaseManifest();
    console.log("DEPLOYMENT_POLICY_PASS: manifiesto base conforme a la política versionada.");
  } catch (error) {
    console.error(error.message);
    process.exitCode = 1;
  }
}

module.exports = { assertDeploymentPolicy, checkBaseManifest };
