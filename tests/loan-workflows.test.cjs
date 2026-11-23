const { test } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const { createHarness } = require("../tools/harness.cjs");

const html = fs.readFileSync(path.join(__dirname, "../dist/Frontend/views/HomeLoans.html"), "utf8");
const { homeLoanScope, homeLoanRequestPayload } = vm.runInNewContext(
  html.match(/<script>([\s\S]*?)<\/script>/)[1] + "\n({ homeLoanScope, homeLoanRequestPayload })",
);

test("el formulario envía solo los campos del contrato y admite el vencimiento opcional", () => {
  const h = createHarness();
  const [ana, luis] = h.seed();
  const state = h.context.homeGetState("token-ana");
  const { debtors } = homeLoanScope(state, "", "");
  const draft = { debtorId: luis.id, creditorId: "tercero", name: "  Compra de ejemplo  ", amount: "125,50", date: state.today, due: "" };
  const payload = homeLoanRequestPayload(draft, debtors, state.today);
  assert.deepEqual({ ...payload }, { debtorId: luis.id, name: "Compra de ejemplo", amount: "125.50", date: state.today });
  const result = h.execute("LOAN_REQUEST", payload, "ana", "ui-request-0001");
  assert.equal(result.state.loans[0].creditorId, ana.id);
  assert.equal(result.state.loans[0].cents, 12550);
  assert.equal(result.state.loans[0].due, "");
  assert.equal(result.state.obligations.length, 0);
  assert.equal(h.execute("LOAN_REQUEST", payload, "ana", "ui-request-0001").duplicate, true);
  const dated = homeLoanRequestPayload({ ...draft, due: "2026-12-01" }, debtors, state.today);
  assert.equal(dated.due, "2026-12-01");
  for (const bad of [
    { debtorId: ana.id }, { name: "   " }, { name: "x".repeat(501) }, { amount: "0" },
    { amount: "123.456" }, { amount: "100000000.01" }, { amount: "1,234.00" },
    { date: "2026-02-30" }, { date: "2026-11-17" }, { due: "2026-11-15" },
  ]) assert.throws(() => homeLoanRequestPayload({ ...draft, ...bad }, debtors, state.today));
});

test("la vista seleccionada cambia el alcance y las acciones siguen perteneciendo a la sesión", () => {
  const h = createHarness();
  const [ana, luis, maria] = h.seed();
  h.execute("PERSON", { name: "Administración", userId: "admin" });
  const admin = h.context.homeGetState("token-admin").me;
  const add = (user, debtor, id) => h.execute("LOAN_REQUEST", {
    debtorId: debtor.id, name: id, amount: "25", date: "2026-11-12",
  }, user, id).id;
  const ownId = add("admin", ana, "ui-request-admin");
  const otherId = add("luis", maria, "ui-request-other");
  const doneId = add("ana", luis, "ui-request-done");
  h.execute("LOAN_DECIDE", { loanId: doneId, decision: "accept" }, "luis");
  const state = h.context.homeGetState("token-admin");
  const all = homeLoanScope(state, "", "");
  assert.equal(all.me, admin);
  assert.deepEqual(Array.from(all.groups.sent, x => x.id), [ownId]);
  assert.deepEqual(Array.from(all.groups.other, x => x.id), [otherId]);
  assert.deepEqual(Array.from(all.groups.history, x => x.id), [doneId]);
  const selected = homeLoanScope(state, maria.id, "2026-11");
  assert.deepEqual(Array.from(selected.groups.received, x => x.id), [otherId]);
  assert.deepEqual(Array.from(selected.groups.received[0].allowedDecisions), []);
  assert.deepEqual(Array.from(all.groups.sent[0].allowedDecisions), ["cancel"]);
  assert.deepEqual(Array.from(all.groups.history[0].allowedDecisions), []);
  assert.equal(homeLoanScope(state, maria.id, "2026-10").groups.received.length, 0);
  const memberState = h.context.homeGetState("token-maria");
  const member = homeLoanScope(memberState, ana.id, "");
  assert.equal(member.groups.received[0].id, otherId);
  assert.deepEqual(Array.from(member.groups.received[0].allowedDecisions), ["accept", "reject"]);
  h.execute("PERSON", { id: luis.id, name: luis.name, userId: "luis", active: false });
  const archivedCreditor = h.context.homeGetState("token-maria").loans.find((loan) => loan.id === otherId);
  assert.deepEqual(Array.from(archivedCreditor.allowedDecisions), ["reject"]);
});

test("integrantes inactivos pueden cerrar sus solicitudes pendientes sin adquirir deuda", () => {
  const h = createHarness();
  const [ana, luis] = h.seed();
  const request = () => h.execute("LOAN_REQUEST", {
    debtorId: luis.id, name: "Préstamo pendiente", amount: "25", date: "2026-11-12",
  }, "ana").id;
  const rejectId = request();
  const cancelId = request();
  const archive = (person, userId) => h.execute("PERSON", {
    id: person.id, name: person.name, userId, active: false,
  });
  const state = (user) => h.context.homeGetState("token-" + user);
  const allowed = (current, id) => Array.from(current.loans.find((loan) => loan.id === id)?.allowedDecisions || []);

  archive(luis, "luis");
  const inactiveDebtor = state("luis");
  assert.equal(inactiveDebtor.me, "");
  assert.equal(homeLoanScope(inactiveDebtor, "", "").debtors.length, 0);
  const beforeReject = h.db.getSheetByName("HOME_EVENTOS").getLastRow();
  assert.throws(() => h.execute("LOAN_DECIDE", { loanId: rejectId, decision: "accept" }, "luis"), /no disponible/);
  assert.equal(h.db.getSheetByName("HOME_EVENTOS").getLastRow(), beforeReject);
  const rejected = h.execute("LOAN_DECIDE", { loanId: rejectId, decision: "reject" }, "luis").state;
  assert.deepEqual(allowed(inactiveDebtor, rejectId), ["reject"]);
  assert.deepEqual(allowed(rejected, rejectId), []);

  archive(ana, "ana");
  const inactiveCreditor = state("ana");
  assert.equal(inactiveCreditor.me, "");
  assert.equal(homeLoanScope(inactiveCreditor, "", "").debtors.length, 0);
  assert(inactiveCreditor.people.every((person) => !Object.hasOwn(person, "userId")));
  assert.deepEqual(allowed(inactiveCreditor, cancelId), ["cancel"]);
  assert.deepEqual(allowed(state("maria"), cancelId), []);
  const unlinked = state("admin");
  assert.equal(homeLoanScope(unlinked, "", "").debtors.length, 0);
  assert.deepEqual(Array.from(unlinked.loans.find((loan) => loan.id === cancelId).allowedDecisions), []);
  const cancelled = h.execute("LOAN_DECIDE", { loanId: cancelId, decision: "cancel" }, "ana").state;
  assert.deepEqual(allowed(cancelled, cancelId), []);
  assert.equal(state("admin").obligations.length, 0);
});
