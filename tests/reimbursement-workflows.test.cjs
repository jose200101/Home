const { test } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const { createHarness } = require("../tools/harness.cjs");

const html = fs.readFileSync(path.join(__dirname, "../dist/Frontend/views/HomeReimbursements.html"), "utf8");
const { homeReimbursementScope, homeReimbursementPayload } = vm.runInNewContext(
  html.match(/<script>([\s\S]*?)<\/script>/)[1] + "\n({ homeReimbursementScope, homeReimbursementPayload })",
);

function fixture() {
  const h = createHarness();
  const [ana, luis, maria] = h.seed();
  const participants = [ana, luis].map((p) => ({ personId: p.id }));
  h.execute("EXPENSE_SERVICE", {
    resourceId: "energy", name: "Energía nueva", mode: "equal", expectedVersion: 0, participants,
  });
  for (const user of ["ana", "luis"]) {
    h.execute("AGREEMENT_RESPOND", { resourceId: "energy", version: 1, decision: "accept" }, user);
  }
  const invoice = h.execute("EXPENSE_INVOICE", {
    resourceId: "energy", version: 1, amount: "1200", period: "2026-11", date: "2026-11-01", due: "2026-11-30",
  }).id;
  h.execute("EXPENSE_PAY", { chargeId: invoice, payerId: ana.id, amount: "1200", date: "2026-11-08" }, "ana");
  h.execute("EXPENSE_PURCHASE", {
    resourceId: "purchase", name: "Compra nueva", expectedVersion: 0, payerId: ana.id,
    participants, mode: "equal", amount: "300", date: "2026-11-09",
  }, "ana");
  for (const user of ["ana", "luis"]) {
    h.execute("EXPENSE_RESPOND", { resourceId: "purchase", version: 1, decision: "accept" }, user);
  }
  const loan = h.execute("LOAN_REQUEST", {
    debtorId: luis.id, name: "Préstamo nuevo", amount: "200", date: "2026-11-10",
  }, "ana").id;
  h.execute("LOAN_DECIDE", { loanId: loan, decision: "accept" }, "luis");
  const state = (user) => h.context.homeGetState("token-" + user);
  const source = (user, id) => state(user).obligations.find((obligation) => obligation.sourceId === id);
  const balance = (user, id) => {
    const { pending, reserved, available } = source(user, id);
    return [pending, reserved, available];
  };
  return { h, ana, luis, maria, invoice, loan, state, source, balance };
}

test("el formulario usa los límites proyectados para los tres orígenes y separa las rutas históricas", () => {
  const f = fixture();
  const { obligations, payments, today } = f.state("luis");
  const form = { amount: "125,50", date: "2026-11-11", note: " Nota ficticia " };
  for (const [kind, earliest] of [["recurring", "2026-11-08"], ["occasional", "2026-11-09"], ["personal-loan", "2026-11-10"]]) {
    const obligation = obligations.find((o) => (o.expenseKind || o.kind) === kind);
    assert.deepEqual({ ...obligation.reimbursement }, { earliestDate: earliest, canDeclare: true });
    assert.deepEqual({ ...homeReimbursementPayload({ obligation }, form, today) }, {
      obligationId: obligation.id, amount: "125.50", date: form.date, note: "Nota ficticia",
    });
    for (const draft of [
      { amount: "0" }, { amount: "10.001" }, { amount: (obligation.available / 100 + 0.01).toFixed(2) },
      { date: "2026-11-07" }, { date: "2026-11-17" }, { date: "2026-02-30" },
    ]) assert.throws(() => homeReimbursementPayload({ obligation }, { ...form, ...draft }, today));
  }
  assert.equal(homeReimbursementScope("charges", obligations, payments).balances.length, 2);
  assert.equal(homeReimbursementScope("debts", obligations, payments).balances.length, 3);
  assert.equal(homeReimbursementScope("history", obligations, payments).balances.length, 0);
  const historical = { status: "pending" };
  const modern = { status: "pending", canDecideReimbursement: false };
  assert.equal(homeReimbursementScope("history", obligations, [historical, modern]).entries[0], modern);
  assert.equal(homeReimbursementScope("history", obligations, [historical, modern]).entries.length, 1);
});

test("la proyección da acciones a cuentas vinculadas inactivas y refleja la reserva, el rechazo y la confirmación", () => {
  const f = fixture();
  const { h } = f;
  const archive = (person, userId) => h.execute("PERSON", {
    id: person.id, name: person.name, userId, active: false,
  });
  archive(f.luis, "luis");
  archive(f.ana, "ana");
  assert.equal(f.state("luis").me, "");
  assert.equal(f.state("ana").me, "");
  assert.equal(f.source("admin", f.invoice).reimbursement.canDeclare, false);
  assert.equal(f.source("ana", f.invoice).reimbursement.canDeclare, false);
  assert.equal(f.source("maria", f.invoice), undefined);
  const declare = (amount) => {
    const obligation = f.source("luis", f.invoice);
    return homeReimbursementPayload({ obligation }, { amount, date: "2026-11-11", note: "" }, "2026-11-16");
  };
  const payload = declare("200");
  const a = h.execute("REIMBURSEMENT_DECLARE", payload, "luis", "ui-reimbursement-01");
  assert(h.execute("REIMBURSEMENT_DECLARE", payload, "luis", "ui-reimbursement-01").duplicate);
  assert.deepEqual(f.balance("luis", f.invoice), [60000, 20000, 40000]);
  const payment = f.state("ana").payments.find((p) => p.id === a.id);
  assert.equal(payment.canDecideReimbursement, true);
  assert.equal(f.state("luis").payments[0].canDecideReimbursement, false);
  assert.equal(f.state("admin").payments[0].canDecideReimbursement, false);
  assert.equal(f.state("maria").payments.length, 0);
  assert(f.state("ana").people.every((person) => !Object.hasOwn(person, "userId")));
  assert.throws(() => homeReimbursementPayload({ payment, decision: "reject" }, { reason: "  " }), /motivo/);
  const reject = homeReimbursementPayload({ payment, decision: "reject" }, { reason: "  No recibido  " });
  const rejected = h.execute("REIMBURSEMENT_DECIDE", reject, "ana").state.payments.find((p) => p.id === a.id);
  assert.equal(rejected.reason, "No recibido");
  assert.equal(rejected.canDecideReimbursement, false);
  assert.deepEqual(f.balance("luis", f.invoice), [60000, 0, 60000]);
  const full = h.execute("REIMBURSEMENT_DECLARE", declare("600"), "luis").id;
  assert.equal(f.source("luis", f.invoice).reimbursement.canDeclare, false);
  const confirm = homeReimbursementPayload({ payment: { id: full }, decision: "confirm" }, { reason: "" });
  h.execute("REIMBURSEMENT_DECIDE", confirm, "ana");
  assert.deepEqual(f.balance("luis", f.invoice), [0, 0, 0]);
});

test("una corrección pendiente conserva los saldos visibles y oculta el abono", () => {
  const f = fixture();
  const obligationId = f.source("luis", f.loan).id;
  f.h.execute("CORRECTION_PROPOSE", { obligationId, expectedVersion: 0, amount: "100", reason: "Ejemplo" });
  const obligation = f.source("luis", f.loan);
  assert(obligation.pendingCorrectionId);
  assert.equal(obligation.reimbursement.canDeclare, false);
  assert.deepEqual(f.balance("luis", f.loan), [20000, 0, 20000]);
  assert.throws(() => homeReimbursementPayload({ obligation }, { amount: "10", date: "2026-11-11", note: "" }, "2026-11-16"), /no está disponible/);
});
