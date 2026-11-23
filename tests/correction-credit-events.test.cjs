const { test } = require("node:test");
const assert = require("node:assert/strict");
const { createHarness } = require("../tools/harness.cjs");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

function viewScript(name) {
  const html = fs.readFileSync(path.join(__dirname, "../dist/Frontend/views", name + ".html"), "utf8");
  return html.match(/<script>([\s\S]*?)<\/script>/)[1];
}
const { homeCreditPayload, homeDistinctLabels, homeCorrectionPayload } = vm.runInNewContext(
  viewScript("HomeCredits") + viewScript("HomeReimbursements") + viewScript("HomeCorrections") +
    "\n({ homeCreditPayload, homeDistinctLabels, homeCorrectionPayload })",
);

function fixture() {
  const h = createHarness();
  const [ana, luis, maria] = h.seed();
  const state = (user = "admin") => h.context.homeGetState("token-" + user);
  const rows = h.db.getSheetByName("HOME_EVENTOS").rows;
  function expense(
    resourceId,
    debtor = luis,
    payer = ana,
    amount = "600",
    period = "2026-11",
    paid = true,
  ) {
    h.execute("EXPENSE_SERVICE", {
      resourceId,
      name: resourceId,
      mode: "equal",
      expectedVersion: 0,
      participants: [{ personId: debtor.id }],
    });
    h.execute(
      "AGREEMENT_RESPOND",
      { resourceId, version: 1, decision: "accept" },
      debtor.userId,
    );
    const invoice = h.execute("EXPENSE_INVOICE", {
      resourceId,
      version: 1,
      amount,
      period,
      date: "2026-11-01",
      due: "2026-11-30",
    }).id;
    if (paid)
      h.execute(
        "EXPENSE_PAY",
        { chargeId: invoice, payerId: payer.id, date: "2026-11-10", amount },
        payer.userId,
      );
    return {
      invoice,
      obligation: state().obligations.find((o) => o.sourceId === invoice)?.id,
    };
  }
  const source = expense("energia");
  const propose = (
    amount,
    expectedVersion = 0,
    ob = source.obligation,
    user = "admin",
    id,
  ) =>
    h.execute(
      "CORRECTION_PROPOSE",
      {
        obligationId: ob,
        amount,
        expectedVersion,
        reason: "Importe conciliado",
      },
      user,
      id,
    );
  const vote = (id, user, version = 1, decision = "accept", operationId) =>
    h.execute(
      "CORRECTION_DECIDE",
      {
        correctionId: id,
        version,
        decision,
        ...(decision === "reject" ? { reason: "Diferencia" } : {}),
      },
      user,
      operationId,
    );
  const approve = (amount, version = 0, ob = source.obligation) => {
    const id = propose(amount, version, ob).id;
    vote(id, "ana", version + 1);
    vote(id, "luis", version + 1);
    return id;
  };
  const declare = (amount, ob = source.obligation) =>
    h.execute(
      "REIMBURSEMENT_DECLARE",
      { obligationId: ob, amount, date: "2026-11-15" },
      "luis",
    ).id;
  const receive = (paymentId, decision = "confirm") =>
    h.execute(
      "REIMBURSEMENT_DECIDE",
      { paymentId, decision, reason: "Comprobación" },
      "ana",
    );
  const credit = () => {
    receive(declare("500"));
    return approve("400");
  };
  const balance = (id = source.obligation) =>
    state().obligations.find((o) => o.id === id);
  const refund = (id, amount, user = "ana", op) =>
    h.execute("CREDIT_REFUND_DECLARE", { creditId: id, amount }, user, op).id;
  const apply = (id, obligationId, amount, user = "luis") =>
    h.execute(
      "CREDIT_APPLY_PROPOSE",
      { creditId: id, obligationId, amount },
      user,
    ).id;
  const decide = (operationId, user, decision = "accept", id) =>
    h.execute(
      "CREDIT_DECIDE",
      {
        operationId,
        decision,
        ...(decision === "reject" ? { reason: "No recibido" } : {}),
      },
      user,
      id,
    );
  function denied(fn, regex) {
    const n = rows.length;
    assert.throws(fn, regex);
    assert.equal(rows.length, n);
    assert.equal(h.locked, false);
  }
  return {
    h,
    ana,
    luis,
    maria,
    state,
    rows,
    expense,
    source,
    propose,
    vote,
    approve,
    declare,
    receive,
    credit,
    balance,
    refund,
    apply,
    decide,
    denied,
  };
}

test("la interfaz usa la versión proyectada y las acciones siguen a las partes desactivadas", () => {
  const f = fixture();
  const ob = () => f.balance();
  assert.equal(f.state("luis").obligations[0].correctionOffer, undefined);
  const original = homeCorrectionPayload({ obligation: ob() }, { amount: "0", reason: " Cuenta revisada " });
  assert.equal(original.expectedVersion, 0);
  assert.equal(original.amount, "0");
  const proposal = f.h.execute("CORRECTION_PROPOSE", original, "admin", "ui-correction-request");
  assert(f.h.execute("CORRECTION_PROPOSE", original, "admin", "ui-correction-request").duplicate);
  assert.match(ob().correctionOffer.blocked, /pendiente/);
  for (const p of [f.ana, f.luis]) {
    f.h.execute("PERSON", { id: p.id, name: p.name, userId: p.userId, active: false });
  }
  const correction = (user) => f.state(user).corrections[0];
  assert.equal(f.state("luis").me, "");
  assert.equal(correction("luis").canDecide, true);
  assert.equal(correction("admin").canDecide, false);
  assert.equal(correction("maria"), undefined);
  const approve = homeCorrectionPayload({ kind: "accept", correction: correction("luis") }, { reason: "" });
  assert.equal(approve.version, 1);
  f.h.execute("CORRECTION_DECIDE", approve, "luis");
  assert.equal(correction("luis").canDecide, false);
  assert.equal(correction("ana").canDecide, true);
  assert.equal(ob().effectiveCents ?? ob().cents, 60000);
  assert.throws(() => homeCorrectionPayload({ kind: "reject", correction: correction("ana") }, { reason: " " }), /motivo/);
  const reject = homeCorrectionPayload({ kind: "reject", correction: correction("ana") }, { reason: " Importe incorrecto " });
  f.h.execute("CORRECTION_DECIDE", reject, "ana");
  assert.equal(correction("ana").canDecide, false);
  assert.equal(correction("ana").decisions[1].reason, "Importe incorrecto");
  assert.equal(homeCorrectionPayload({ obligation: ob() }, { amount: "700", reason: "Importe revisado" }).expectedVersion, 1);
  f.denied(() => f.h.execute("CORRECTION_PROPOSE", original, "admin"), /desactualizada/);
});

test("las reservas y el crédito aplicado se reflejan sin recalcular la obligación", () => {
  const f = fixture();
  const creditId = f.credit();
  const target = f.expense("destino-ui").obligation;
  const allocation = f.apply(creditId, target, "40");
  const draft = { amount: "30", reason: "Conciliado" };
  assert.match(f.balance(target).correctionOffer.blocked, /reservas/);
  assert.throws(() => homeCorrectionPayload({ obligation: f.balance(target) }, draft), /reservas/);
  f.decide(allocation, "ana");
  assert.equal(f.balance(target).correctionOffer.minimum, 4000);
  assert.throws(() => homeCorrectionPayload({ obligation: f.balance(target) }, draft), /crédito ya aplicado/);
  f.declare("10", target);
  assert.match(f.balance(target).correctionOffer.blocked, /reservas/);
});

test("la proyección ofrece destinos compatibles y acciones de sesión aunque los integrantes estén archivados", () => {
  const f = fixture();
  const id = f.credit();
  const target = f.expense("destino", f.luis, f.ana, "40").obligation;
  f.expense("sin-pagar", f.luis, f.ana, "80", "2026-11", false);
  f.expense("otras-partes", f.luis, f.maria, "90");
  const blocked = f.expense("en-correccion").obligation;
  f.propose("100", 0, blocked);
  const loan = f.h.execute("LOAN_REQUEST", {
    debtorId: f.luis.id, name: "Préstamo", date: "2026-11-12", amount: "200",
  }, "ana").id;
  f.h.execute("LOAN_DECIDE", { loanId: loan, decision: "accept" }, "luis");
  for (const p of [f.ana, f.luis]) {
    f.h.execute("PERSON", { id: p.id, name: p.name, userId: p.userId, active: false });
  }
  const credit = (user) => f.state(user).credits.find((c) => c.id === id);
  assert.equal(f.state("luis").me, "");
  assert.equal(credit("ana").canRefund, true);
  assert.equal(credit("luis").canRefund, false);
  assert.equal(credit("admin").canRefund, false);
  assert.deepEqual(Array.from(credit("admin").targets), []);
  assert.equal(f.state("maria").credits.length, 0);
  for (const user of ["ana", "luis"]) {
    assert.deepEqual(Array.from(credit(user).targets, (x) => ({ ...x })), [{ obligationId: target, limit: 4000 }]);
  }
  const action = { kind: "application", credit: credit("luis") };
  const form = { obligationId: target, amount: "40,00", note: " Acuerdo " };
  assert.throws(() => homeCreditPayload(action, { ...form, amount: "40.01" }), /límite/);
  assert.throws(() => homeCreditPayload(action, { ...form, obligationId: loan }), /no|disponible/);
  const payload = homeCreditPayload(action, form);
  const proposal = f.h.execute("CREDIT_APPLY_PROPOSE", payload, "luis", "credit-ui-propose01");
  assert(f.h.execute("CREDIT_APPLY_PROPOSE", payload, "luis", "credit-ui-propose01").duplicate);
  assert.equal(credit("luis").reserved, 4000);
  assert.equal(f.balance(target).creditReserved, 4000);
  assert.equal(f.balance(target).pending, 4000);
  assert.deepEqual(Array.from(credit("luis").targets), []);
  const operation = f.state("ana").creditOperations.find((o) => o.id === proposal.id);
  assert.equal(operation.canDecide, true);
  assert.equal(f.state("admin").creditOperations[0].canDecide, false);
  assert.throws(() => homeCreditPayload({ kind: "reject", operation }, { reason: "  " }), /motivo/);
  const reject = homeCreditPayload({ kind: "reject", operation }, { reason: " Falta comprobante " });
  f.h.execute("CREDIT_DECIDE", reject, "ana");
  assert.equal(credit("luis").pending, 10000);
  assert.equal(credit("luis").available, 10000);
  assert.equal(f.balance(target).available, 4000);
  assert.equal(f.state("ana").creditOperations[0].canDecide, false);
});

test("los nombres accesibles distinguen registros cuando el concepto y los detalles coinciden", () => {
  const labels = homeDistinctLabels([{ id: "a" }, { id: "b" }, { id: "c" }], (x) => x.id === "c" ? "Otro" : "Mismo");
  assert.deepEqual(Array.from(labels.values()), ["Mismo, registro 1", "Mismo, registro 2", "Otro"]);
});

test("corrección versionada exige las dos partes, admite incremento y conserva factura y principal", () => {
  const f = fixture();
  const charge = JSON.stringify(f.state().charges[0]);
  f.denied(
    () => f.propose("700", 0, f.source.obligation, "ana"),
    /administrador/,
  );
  const p = f.propose("700").id;
  assert.equal(f.balance().pending, 60000);
  f.denied(() => f.propose("720", 1), /pendiente/);
  f.denied(() => f.declare("5"), /corrección pendiente/);
  f.denied(() => f.vote(p, "admin"), /no disponible/);
  f.denied(() => f.vote(p, "maria"), /no disponible/);
  f.denied(() => f.vote(p, "luis", 2), /versión/);
  f.vote(p, "ana");
  assert.equal(f.balance().pending, 60000);
  f.denied(() => f.vote(p, "ana"), /decisión/);
  f.vote(p, "luis", 1, "accept", "correction-vote-0001");
  assert(f.vote(p, "luis", 1, "accept", "correction-vote-0001").duplicate);
  assert.equal(f.balance().cents, 60000);
  assert.equal(f.balance().effectiveCents, 70000);
  assert.equal(f.balance().pending, 70000);
  assert.equal(f.state().credits.length, 0);
  assert.equal(JSON.stringify(f.state().charges[0]), charge);
  const increasedPayment = f.declare("700");
  assert.equal(f.balance().available, 0);
  f.receive(increasedPayment, "reject");
  f.denied(() => f.propose("450"), /desactualizada/);
  const rejected = f.propose("450", 1).id;
  f.vote(rejected, "ana", 2, "reject");
  assert.equal(f.balance().pending, 70000);
  assert.equal(f.state().credits.length, 0);
});

test("excedente de efectivo se reconoce una sola vez aunque suba y vuelva a bajar el importe", () => {
  const f = fixture();
  const pay = f.declare("500");
  f.denied(() => f.propose("400"), /reservas/);
  f.receive(pay);
  const baseline = JSON.stringify(f.state().payments);
  const cancelled = f.propose("390").id;
  f.vote(cancelled, "ana", 1, "reject");
  assert.equal(f.state().credits.length, 0);
  f.approve("400", 1);
  assert.deepEqual(
    [
      f.balance().cents,
      f.balance().paid,
      f.balance().effectiveCents,
      f.balance().pending,
    ],
    [60000, 50000, 40000, 0],
  );
  assert.equal(f.state().credits[0].cents, 10000);
  f.approve("500", 2);
  assert.equal(f.balance().pending, 10000);
  f.approve("400", 3);
  assert.equal(f.state().credits.length, 1);
  f.approve("350", 4);
  assert.equal(f.state().credits.length, 2);
  assert.equal(f.state().creditTotals.pending, 15000);
  assert.equal(JSON.stringify(f.state().payments), baseline);
  assert.equal(f.balance().creditedCents, 15000);
});

test("devolución parcial y completa exige receptor, reserva disponible y rechazo libera reserva", () => {
  const f = fixture(),
    credit = f.credit();
  f.denied(() => f.refund(credit, "60", "luis"), /no disponible/);
  const first = f.refund(credit, "60", "ana", "refund-unique-0001");
  assert.equal(f.refund(credit, "60", "ana", "refund-unique-0001"), first);
  assert.deepEqual(
    [
      f.state().creditTotals.pending,
      f.state().creditTotals.reserved,
      f.state().creditTotals.available,
    ],
    [10000, 6000, 4000],
  );
  f.denied(() => f.refund(credit, "40.01"), /crédito disponible/);
  f.denied(() => f.decide(first, "ana"), /no disponible/);
  f.denied(() => f.decide(first, "admin"), /no disponible/);
  f.decide(first, "luis", "reject");
  assert.equal(f.state().credits[0].available, 10000);
  f.decide(f.refund(credit, "60"), "luis");
  assert.equal(f.state().credits[0].pending, 4000);
  const last = f.refund(credit, "40");
  f.decide(last, "luis", "accept", "refund-decide-0001");
  assert(f.decide(last, "luis", "accept", "refund-decide-0001").duplicate);
  f.denied(() => f.decide(last, "luis"), /atendida/);
  assert.equal(f.state().credits[0].returned, 10000);
  assert.equal(f.state().creditTotals.pending, 0);
});

test("aplicación solo a gasto pagado entre las mismas partes y consume ambos saldos", () => {
  const f = fixture(),
    credit = f.credit();
  const other = f.expense("otra-persona", f.maria).obligation;
  const target = f.expense("agua", f.luis, f.ana, "200").obligation;
  const unpaid = f.expense("impago", f.luis, f.ana, "200", "2026-11", false);
  const loan = f.h.execute(
    "LOAN_REQUEST",
    {
      debtorId: f.luis.id,
      amount: "200",
      date: "2026-11-01",
      name: "Préstamo",
    },
    "ana",
  ).id;
  f.h.execute("LOAN_DECIDE", { loanId: loan, decision: "accept" }, "luis");
  f.denied(() => f.apply(credit, other, "100"), /mismas dos personas/);
  f.denied(() => f.apply(credit, loan, "100"), /gastos compartidos/);
  f.denied(() => f.apply(credit, unpaid.invoice, "100"), /gastos compartidos/);
  f.denied(() => f.apply(credit, target, "100", "maria"), /no disponible/);
  f.receive(f.declare("40", target));
  const cash = f.declare("80", target);
  f.denied(() => f.apply(credit, target, "100"), /saldo disponible/);
  const p = f.apply(credit, target, "80");
  assert.deepEqual(
    [
      f.balance(target).paid,
      f.balance(target).pending,
      f.balance(target).creditReserved,
      f.balance(target).available,
    ],
    [4000, 16000, 8000, 0],
  );
  f.denied(() => f.propose("250", 0, target), /reservas/);
  f.denied(() => f.refund(credit, "20.01"), /crédito disponible/);
  f.denied(() => f.declare("0.01", target), /saldo disponible/);
  f.decide(p, "ana", "reject");
  assert.equal(f.balance(target).available, 8000);
  f.receive(cash, "reject");
  const approved = f.apply(credit, target, "100", "ana");
  f.decide(approved, "luis", "accept", "apply-decision-0001");
  assert(f.decide(approved, "luis", "accept", "apply-decision-0001").duplicate);
  assert.deepEqual(
    [
      f.balance(target).cents,
      f.balance(target).paid,
      f.balance(target).appliedCreditCents,
      f.balance(target).pending,
    ],
    [20000, 4000, 10000, 6000],
  );
  assert.equal(f.state().credits[0].applied, 10000);
  assert.equal(f.state().creditTotals.pending, 0);
  f.denied(() => f.propose("99.99", 0, target), /crédito ya aplicado/);
  f.approve("100", 0, target);
  assert.equal(f.balance(target).pending, 0);
  assert.equal(f.state().credits[1].cents, 4000);
  assert.equal(f.state().creditTotals.pending, 4000);
});

test("solicitudes intercaladas no gastan dos veces el crédito", () => {
  const f = fixture(),
    credit = f.credit();
  const original = f.h.context.LockService.getScriptLock;
  f.h.context.LockService.getScriptLock = () => {
    const lock = original();
    return {
      ...lock,
      waitLock: () => {
        f.h.context.LockService.getScriptLock = original;
        f.refund(credit, "75");
        lock.waitLock();
      },
    };
  };
  assert.throws(() => f.refund(credit, "26"), /crédito disponible/);
  assert.equal(f.state().creditOperations.length, 1);
  assert.equal(f.state().creditTotals.available, 2500);
  assert.equal(f.h.locked, false);
});

test("replay de esquemas 1–6 conserva filas, privacidad y consolidado incluso tras retiro", () => {
  const f = fixture();
  const loan = f.h.execute(
    "LOAN_REQUEST",
    {
      debtorId: f.luis.id,
      amount: "200",
      date: "2026-11-01",
      name: "Préstamo",
    },
    "ana",
  ).id;
  f.h.execute("LOAN_DECIDE", { loanId: loan, decision: "accept" }, "luis");
  const baseline = JSON.stringify(f.rows),
    n = f.rows.length;
  const credit = f.credit();
  const refund = f.refund(credit, "10");
  assert.equal(f.state("maria").credits.length, 0);
  assert.equal(f.state("maria").corrections.length, 0);
  assert.equal(f.state("maria").creditOperations.length, 0);
  assert.equal(f.state("maria").creditTotals.pending, 0);
  f.h.execute("PERSON", {
    id: f.luis.id,
    userId: "luis",
    active: false,
    name: "Luis",
  });
  assert.equal(f.state("luis").credits[0].id, credit);
  f.decide(refund, "luis");
  const restored = createHarness(),
    sh = restored.db.insertSheet("HOME_EVENTOS");
  f.rows.forEach((row) => sh.appendRow(row));
  const view = restored.context.homeGetState("token-admin");
  assert.equal(JSON.stringify(view), JSON.stringify(f.state()));
  assert.equal(JSON.stringify(f.rows.slice(0, n)), baseline);
  assert.deepEqual(
    [...new Set(f.rows.slice(1).map((r) => JSON.parse(r[4]).schema))].sort(),
    [1, 2, 3, 4, 5, 6],
  );
  assert.equal(view.creditTotals.pending, 9000);
  assert.equal(view.obligations.find((o) => o.id === loan).pending, 20000);
  assert.equal(view.plan.length, 0);
});

test("una corrección también ajusta un préstamo sin reescribir el principal del contrato", () => {
  const f = fixture();
  const loan = f.h.execute(
    "LOAN_REQUEST",
    {
      debtorId: f.luis.id,
      amount: "200",
      date: "2026-11-01",
      name: "Préstamo",
    },
    "ana",
  ).id;
  f.h.execute("LOAN_DECIDE", { loanId: loan, decision: "accept" }, "luis");
  const original = JSON.stringify(f.state().loans);
  f.receive(f.declare("150", loan));
  const credit = f.approve("100", 0, loan);
  assert.equal(f.balance(loan).cents, 20000);
  assert.equal(f.balance(loan).pending, 0);
  assert.equal(f.state().credits.find((c) => c.id === credit).cents, 5000);
  assert.equal(JSON.stringify(f.state().loans), original);
});
