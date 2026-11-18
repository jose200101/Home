const { test } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const c = vm.createContext({ Set, JSON, Number, String, Array, Object, Math });
const filename = path.join(__dirname, "../dist/Backend/domain/AgreementRules.js");
vm.runInContext(fs.readFileSync(filename, "utf8"), c, { filename });
const admin = { id: "admin", active: true, admin: true };
const ana = { id: "user_ana", personId: "ana", active: true };
const carlos = { id: "user_carlos", personId: "carlos", active: true };
const maria = { id: "user_maria", personId: "maria", active: true };
const apply = (s, action, actor) => c.homeAgreementApply_(s, action, actor);
const proposal = (participants = ["ana", "carlos"]) => ({
  kind: "recurring", resourceId: "energia", mode: "equal",
  participants: participants.map((personId) => ({ personId })),
});
const occasional = () => ({
  kind: "occasional", resourceId: "limpieza001", mode: "equal",
  amountCents: 60000, payerId: "ana",
  participants: ["ana", "carlos", "maria"].map((personId) => ({ personId })),
});
const start = (s, p, actor = admin) => apply(s, { type: "PROPOSE", proposal: p }, actor);
const answer = (s, actor, decision = "accept", version = s.pending.version) =>
  apply(s, { type: "RESPOND", version, decision }, actor);
const cancel = (s, actor = admin) => apply(s, { type: "CANCEL", version: s.pending.version }, actor);

test("primer acuerdo no factura hasta aceptar todos", () => {
  let s = start(null, proposal());
  assert.equal(s.pending.version, 1);
  assert.equal(c.homeAgreementCanInvoice_(s), false);
  s = answer(s, ana);
  assert.equal(c.homeAgreementCanInvoice_(s), false);
  s = answer(s, carlos);
  assert.equal(s.active.version, 1);
  assert.equal(c.homeAgreementCanInvoice_(s), true);
});
test("versión nueva suspende factura sin reemplazar acuerdo anterior", () => {
  let s = answer(answer(start(null, proposal()), ana), carlos);
  s = start(s, proposal(["ana", "maria"]));
  assert.equal(s.active.version, 1);
  assert.equal(s.pending.version, 2);
  assert.equal(c.homeAgreementCanInvoice_(s), false);
  s = answer(answer(s, ana), maria);
  assert.equal(s.active.version, 2);
  assert.equal(s.history[0].status, "superseded");
  assert.equal(c.homeAgreementCanInvoice_(s), true);
});
test("rechazo exige nueva versión y respuestas nuevas", () => {
  let s = answer(answer(start(null, proposal()), ana), carlos);
  s = start(s, proposal(["ana", "maria"]));
  s = answer(answer(s, ana), maria, "reject");
  assert.equal(s.pending.status, "rejected");
  assert.equal(c.homeAgreementCanInvoice_(s), false);
  assert.throws(() => answer(s, maria), /rechazada/);
  s = start(s, proposal(["ana", "maria"]));
  assert.equal(s.pending.version, 3);
  assert.ok(s.pending.participants.every((p) => p.decision === "pending"));
  s = answer(answer(s, ana), maria);
  assert.equal(c.homeAgreementCanInvoice_(s), true);
});
test("cancelación devuelve la versión anterior", () => {
  let s = answer(answer(start(null, proposal()), ana), carlos);
  s = cancel(start(s, proposal(["ana", "maria"])));
  assert.equal(s.active.version, 1);
  assert.equal(s.pending, null);
  assert.equal(s.history[0].status, "cancelled");
  assert.equal(c.homeAgreementCanInvoice_(s), true);
});
test("gasto ocasional exige pagador y aprobación individual", () => {
  assert.throws(() => start(null, occasional(), carlos), /Solo quien pagó/);
  let s = start(null, occasional(), ana);
  s = answer(answer(s, ana), carlos);
  assert.equal(c.homeAgreementCanActivate_(s), false);
  s = answer(s, maria);
  assert.equal(c.homeAgreementCanActivate_(s), true);
  assert.equal(c.homeAgreementCanInvoice_(s), false);
});
test("ni el administrador ni un extraño pueden aceptar por un integrante", () => {
  const s = start(null, proposal());
  assert.throws(() => answer(s, admin), /involucrado/);
  assert.throws(() => answer(s, maria), /involucrado/);
  assert.throws(() => answer(s, ana, "accept", 0), /versión/);
  assert.throws(() => answer(s, ana, "unknown"), /Decisión/);
  const approved = answer(s, ana);
  assert.throws(() => answer(approved, ana), /ya respondió/);
  assert.equal(s.pending.participants[0].decision, "pending");
});
test("porcentaje respeta base 100 y dos decimales", () => {
  const p = proposal(); p.mode = "percent";
  p.participants[0].value = "33.33"; p.participants[1].value = "66.67";
  const s = start(null, p);
  assert.equal(s.pending.participants[0].basis, 3333);
  p.participants[1].value = "66.66";
  assert.throws(() => start(null, p), /sumar 100/);
  p.participants[1].value = "66.666";
  assert.throws(() => start(null, p), /dos decimales/);
});
test("aportes fijos deben sumar centavos exactos", () => {
  const p = proposal(); p.mode = "amount"; p.amountCents = 10001;
  p.participants[0].value = "60.01"; p.participants[1].value = "40";
  assert.equal(start(null, p).pending.participants[0].cents, 6001);
  p.participants[1].value = "40.01";
  assert.throws(() => start(null, p), /sumar exactamente/);
});
test("valida participantes duplicados, montos y actor activo", () => {
  assert.throws(() => start(null, proposal(["ana", "ana"])), /repetido/);
  assert.throws(() => start(null, proposal([])), /entre 1 y 100/);
  assert.throws(() => start(null, proposal(), ana), /administrador/);
  assert.throws(() => start(null, proposal(), { ...admin, active: false }), /activa/);
  assert.throws(() => start(null, { ...occasional(), amountCents: 0 }, ana), /importe total/);
  assert.throws(() => start(null, { ...proposal(), resourceId: "../x" }), /Identificador/);
});
test("rechaza editar propuesta pendiente, respuestas tras retiro y cancelar sin permiso", () => {
  const s = start(null, proposal());
  assert.throws(() => start(s, proposal()), /Ya hay/);
  assert.throws(() => answer(s, { ...ana, active: false }), /activa/);
  assert.throws(() => cancel(s, carlos), /administrador/);
});
test("gasto rechazado requiere reenviar versión; gasto aprobado requiere corrección", () => {
  let s = start(null, occasional(), ana);
  s = answer(answer(s, ana), carlos, "reject");
  assert.equal(c.homeAgreementCanActivate_(s), false);
  const p = occasional(); p.participants = [{ personId: "ana" }, { personId: "maria" }];
  s = start(s, p, ana);
  assert.equal(s.pending.version, 2);
  s = answer(answer(s, ana), maria);
  assert.equal(c.homeAgreementCanActivate_(s), true);
  assert.throws(() => start(s, p, ana), /corrección/);
});

test("identidad de cuenta no sustituye el vínculo explícito con integrante", () => {
  const s = start(null, proposal());
  assert.throws(() => answer(s, { id: "ana", active: true }), /involucrado/);
  assert.throws(() => answer(s, { id: "user_ana", personId: "maria", active: true }), /involucrado/);
  const updated = answer(s, ana);
  assert.equal(updated.pending.participants[0].decision, "accepted");
  assert.equal(updated.pending.createdBy, "admin");
});
