const { test } = require("node:test");
const assert = require("node:assert/strict");
const { createHarness } = require("../tools/harness.cjs");

function fixture() {
  const h = createHarness();
  const [ana, luis, maria] = h.seed();
  const info = (user = "admin") => h.context.homeGetState("token-" + user);
  const rows = h.db.getSheetByName("HOME_EVENTOS").rows;
  const issue = (overrides = {}, user = "ana", id) =>
    h.execute(
      "LOAN_REQUEST",
      {
        debtorId: luis.id,
        amount: "2000.00",
        date: "2026-10-01",
        name: "Préstamo ficticio",
        ...overrides,
      },
      user,
      id,
    );
  const decide = (loanId, decision = "accept", user = "luis", id) =>
    h.execute("LOAN_DECIDE", { loanId, decision, reason: "Prueba" }, user, id);
  function denied(action, expected) {
    const count = rows.length;
    assert.throws(action, expected);
    assert.equal(rows.length, count);
    assert.equal(h.locked, false);
  }
  return { h, ana, luis, maria, info, issue, decide, rows, denied };
}

test("solicitud pendiente, identidad del acreedor y vencimiento opcional", () => {
  const f = fixture();
  const first = f.issue({ creditorId: f.maria.id }, "ana", "loan-request-0001");
  assert.equal(first.state.loans[0].creditorId, f.ana.id);
  assert.equal(first.state.loans[0].status, "pending");
  assert.equal(first.state.loans[0].due, "");
  assert.equal(f.info().obligations.length, 0);
  assert.equal(f.info("luis").loans.length, 1);
  assert.equal(f.info("maria").loans.length, 0);
  assert(
    f.issue({ creditorId: f.maria.id }, "ana", "loan-request-0001").duplicate,
  );
  f.denied(
    () => f.issue({ amount: "5" }, "ana", "loan-request-0001"),
    /original/,
  );
  f.denied(() => f.issue({ debtorId: f.ana.id }), /otro integrante/);
  f.denied(() => f.issue({}, "admin"), /activo vinculado/);
  for (const amount of ["-1", "0", "1.001"])
    f.denied(() => f.issue({ amount }), /monto/);
  for (const invalid of [{ date: "2026-11-17" }, { due: "2026-09-30" }])
    f.denied(() => f.issue(invalid), /fechas/);
  const second = f.issue({ due: "2027-01-01" });
  assert.equal(
    f.info().loans.find((l) => l.id === second.id).due,
    "2027-01-01",
  );
  f.decide(first.id, "accept", "luis", "loan-accept-0001");
  assert(f.decide(first.id, "accept", "luis", "loan-accept-0001").duplicate);
  f.denied(() => f.decide(first.id), /atendido/);
  assert.equal(f.info().loans[0].obligationId, first.id);
  assert.equal(f.info().obligations[0].pending, 200000);
  assert.equal(f.info().plan.length, 0);
});

test("persona sin cuenta, vínculo posterior, decisiones auténticas y retiro", () => {
  const f = fixture();
  f.h.execute("PERSON", { name: "Administrador", userId: "admin" });
  const person = f.h
    .execute("PERSON", { name: "Persona sin cuenta" })
    .state.people.at(-1);
  const a = f.issue({ debtorId: person.id });
  for (const user of ["ana", "luis", "maria", "admin"])
    f.denied(() => f.decide(a.id, "accept", user), /no disponible/);
  assert.equal(f.info().obligations.length, 0);
  f.denied(
    () =>
      f.h.execute("PERSON", {
        id: person.id,
        name: person.name,
        userId: "luis",
      }),
    /ya está vinculada/,
  );
  f.h.users.push({
    id: "nueva",
    nombreCompleto: person.name,
    rol: "Integrante",
    estado: "Activo",
  });
  f.h.execute("PERSON", { id: person.id, name: person.name, userId: "nueva" });
  assert.equal(f.info("nueva").loans[0].id, a.id);
  f.decide(a.id, "accept", "nueva");
  const b = f.issue({ debtorId: person.id });
  f.h.execute("PERSON", {
    id: person.id,
    name: person.name,
    userId: "nueva",
    active: false,
  });
  f.denied(() => f.decide(b.id, "accept", "nueva"), /no disponible/);
  f.denied(() => f.issue({ debtorId: person.id }), /desactivado/);
  assert.equal(f.info("nueva").obligations[0].pending, 200000);
  f.decide(b.id, "reject", "nueva");
  assert.equal(f.info().loans[1].status, "rejected");
  assert.equal(f.info().obligations.length, 1);
});

test("rechazo y cancelación cierran la solicitud sin deuda ni edición posterior", () => {
  const f = fixture();
  const a = f.issue(),
    b = f.issue();
  f.denied(() => f.decide(a.id, "reject", "ana"), /no disponible/);
  f.denied(() => f.decide(b.id, "cancel", "luis"), /no disponible/);
  f.decide(a.id, "reject");
  f.decide(b.id, "cancel", "ana");
  assert.deepEqual(
    Array.from(f.info().loans, (l) => l.status),
    ["rejected", "cancelled"],
  );
  assert.equal(f.info().obligations.length, 0);
  f.denied(() => f.decide(a.id), /atendido/);
  f.denied(() => f.decide(b.id), /atendido/);
});

test("abonos de un préstamo vencido conservan principal y excluyen rutas históricas", () => {
  const f = fixture();
  const a = f.issue({ due: "2026-10-15" });
  f.decide(a.id);
  f.denied(
    () => f.h.execute("VOID", { targetId: a.id, reason: "Prueba" }),
    /anular/,
  );
  f.denied(
    () => f.h.execute("PAYMENT", { obligationId: a.id }),
    /comando de reembolsos/,
  );
  const declare = (amount, user = "luis", id, date = "2026-11-15") =>
    f.h.execute(
      "REIMBURSEMENT_DECLARE",
      { obligationId: a.id, date, amount },
      user,
      id,
    );
  const respond = (paymentId, decision, user = "ana") =>
    f.h.execute(
      "REIMBURSEMENT_DECIDE",
      { paymentId, decision, reason: "Prueba" },
      user,
    );
  const balance = () => {
    const { cents, paid, reserved, pending, available, status } =
      f.info().obligations[0];
    return [cents, paid, reserved, pending, available, status];
  };
  f.denied(() => declare("500", "ana"), /no disponible/);
  f.denied(
    () => declare("500", "luis", undefined, "2026-09-30"),
    /fecha del préstamo/,
  );
  const first = declare("500", "luis", "loan-payment-0001");
  assert(declare("500", "luis", "loan-payment-0001").duplicate);
  assert.deepEqual(balance(), [200000, 0, 50000, 200000, 150000, "open"]);
  f.denied(() => respond(first.id, "confirm", "admin"), /no disponible/);
  f.denied(() => f.h.execute("CONFIRM", { paymentId: first.id }), /no existe/);
  respond(first.id, "confirm");
  const second = declare("1000");
  f.denied(() => declare("500.01"), /saldo disponible/);
  respond(second.id, "reject");
  assert.deepEqual(balance(), [200000, 50000, 0, 150000, 150000, "open"]);
  respond(declare("1500").id, "confirm");
  assert.deepEqual(balance(), [200000, 200000, 0, 0, 0, "settled"]);
  assert.equal(f.info("maria").payments.length, 0);
  assert.deepEqual(
    Array.from(f.info("ana").payments, (p) => p.status),
    ["confirmed", "rejected", "confirmed"],
  );
  assert.equal(f.info().loans[0].cents, 200000);
});

test("replay mixto de esquemas 1–5 conserva gastos, abonos y saldo de préstamo por separado", () => {
  const f = fixture();
  f.h.execute("EXPENSE_SERVICE", {
    resourceId: "energia-loans-test",
    name: "Energía",
    mode: "equal",
    expectedVersion: 0,
    participants: [{ personId: f.luis.id }],
  });
  f.h.execute(
    "AGREEMENT_RESPOND",
    { resourceId: "energia-loans-test", version: 1, decision: "accept" },
    "luis",
  );
  const charge = f.h.execute("EXPENSE_INVOICE", {
    resourceId: "energia-loans-test",
    version: 1,
    period: "2026-11",
    amount: "123.45",
    date: "2026-11-01",
    due: "2026-11-30",
  });
  f.h.execute(
    "EXPENSE_PAY",
    {
      chargeId: charge.id,
      amount: "123.45",
      date: "2026-11-10",
      payerId: f.ana.id,
    },
    "ana",
  );
  const expense = f.info().obligations[0];
  const payExpense = f.h.execute(
    "REIMBURSEMENT_DECLARE",
    { obligationId: expense.id, amount: "23.45", date: "2026-11-15" },
    "luis",
  );
  f.h.execute(
    "REIMBURSEMENT_DECIDE",
    { paymentId: payExpense.id, decision: "confirm" },
    "ana",
  );
  const baseline = JSON.stringify(f.rows),
    length = f.rows.length;
  const obligation = f.issue({ amount: "10" });
  f.decide(obligation.id);
  f.h.execute(
    "REIMBURSEMENT_DECLARE",
    { obligationId: obligation.id, amount: "5", date: "2026-11-15" },
    "luis",
  );

  const recovered = createHarness();
  const sheet = recovered.db.insertSheet("HOME_EVENTOS");
  f.rows.forEach((row) => sheet.appendRow(row));
  const view = recovered.context.homeGetState("token-admin");
  assert.deepEqual(
    [...new Set(f.rows.slice(1).map((r) => JSON.parse(r[4]).schema))].sort(),
    [1, 2, 3, 4, 5],
  );
  assert.equal(JSON.stringify(f.rows.slice(0, length)), baseline);
  assert.equal(JSON.stringify(view), JSON.stringify(f.info()));
  assert.equal(view.obligations[0].pending, 10000);
  assert.equal(view.obligations[1].pending, 1000);
  assert.equal(view.obligations[1].reserved, 500);
  assert.equal(view.plan.length, 0);
});
