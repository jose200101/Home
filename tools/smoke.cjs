/** Prueba HTTP/RPC de la vista previa local. No usa Google ni datos reales. */
const assert = require("node:assert/strict");
const { spawn } = require("node:child_process");
const path = require("node:path");

const base = "http://127.0.0.1:4173";
const child = spawn(process.execPath, [path.join(__dirname, "preview.cjs")], {
  stdio: ["ignore", "pipe", "pipe"],
});
let output = "";
child.stdout.on("data", (chunk) => { output += chunk.toString("utf8"); });
child.stderr.on("data", (chunk) => { output += chunk.toString("utf8"); });

async function rpc(name, args) {
  const response = await fetch(base + "/rpc", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name, args }),
    signal: AbortSignal.timeout(3000),
  });
  assert.equal(response.status, 200);
  const body = await response.json();
  if (body.error) throw Error("RPC " + name + ": " + body.error);
  return body.result;
}

async function main() {
  let response;
  for (let n = 0; n < 60; n++) {
    if (child.exitCode !== null) throw Error("Preview terminó antes de iniciar:\n" + output);
    try {
      response = await fetch(base + "/", { signal: AbortSignal.timeout(800) });
      break;
    } catch (e) {
      await new Promise((resolve) => setTimeout(resolve, 100));
    }
  }
  assert.ok(response, "La vista previa no inició:\n" + output);
  assert.equal(response.status, 200);
  const html = await response.text();
  assert.match(html, /VISTA PREVIA LOCAL/);
  assert.match(html, /HomeCuentasView/);
  assert.doesNotMatch(html, /<\?!= include\(/);

  const admin = await rpc("homeGetState", ["token-admin"]);
  const member = await rpc("homeGetState", ["token-luis"]);
  assert.equal(admin.admin, true);
  assert.ok(admin.people.length >= 3);
  assert.ok(admin.templates.length >= 2);
  assert.ok(admin.charges.length >= 2);
  assert.equal(member.admin, false);
  assert.equal(member.templates.length, 0);
  assert.equal(member.audit.length, 0);
  assert.ok(member.me);
  console.log("HOME_SMOKE_PASS: HTML, RPC y privacidad en datos simulados");
}
main()
  .catch((error) => { console.error(error); process.exitCode = 1; })
  .finally(() => { child.kill("SIGTERM"); });
