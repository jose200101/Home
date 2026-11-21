const { test } = require("node:test");
const assert = require("node:assert/strict");
const { createHarness } = require("../tools/harness.cjs");

function fixture() {
  const h = createHarness();
  const [ana, luis, maria] = h.seed();
  const state = (user = "admin") => h.context.homeGetState("token-" + user);
  const rows = h.db.getSheetByName("HOME_EVENTOS").rows;
  const people = (...members) => members.map(({ id }) => ({ personId: id }));
  function response(
    resourceId,
    user,
    version = 1,
    decision = "accept",
    type = "AGREEMENT_RESPOND",
  ) {
    return h.execute(type, { resourceId, version, decision }, user);
  }
  function service(resourceId, participants, overrides = {}) {
    return h.execute("EXPENSE_SERVICE", {
      resourceId,
      name: resourceId,
      provider: "Proveedor ficticio",
      expectedVersion: 0,
      mode: "equal",
      participants,
      ...overrides,
    });
  }
  function invoice(resourceId, overrides = {}) {
    return h.execute("EXPENSE_INVOICE", {
      resourceId,
      version: 1,
      amount: "1200.00",
      period: "2026-11",
      date: "2026-11-01",
      due: "2026-11-20",
      ...overrides,
    });
  }
  function purchase(overrides = {}, user = "maria") {
    return h.execute(
      "EXPENSE_PURCHASE",
      {
        resourceId: "limpieza",
        expectedVersion: 0,
        name: "Limpieza",
        date: "2026-11-15",
        amount: "100.01",
        payerId: maria.id,
        mode: "equal",
        participants: people(ana, luis),
        ...overrides,
      },
      user,
    );
  }
  function denied(run, reason) {
    const before = rows.length;
    assert.throws(run, reason);
    assert.equal(rows.length, before);
  }
  return {
    h,
    ana,
    luis,
    maria,
    people,
    service,
    response,
    invoice,
    purchase,
    denied,
    state,
    rows,
  };
}

test("servicios independientes: aprobación, suspensión, cancelación y snapshot exacto", () => {
  const f = fixture();
  f.service("energia", f.people(f.ana, f.luis, f.maria));
  f.service("claro", f.people(f.luis));
  f.denied(() => f.invoice("energia"), /acuerdo aprobado/);
  f.response("claro", "luis");
  f.invoice("claro");
  ["ana", "luis", "maria"].forEach((user) => f.response("energia", user));
  const first = f.invoice("energia");
  const original = JSON.stringify(
    first.state.charges.find((c) => c.id === first.id),
  );
  assert.equal(first.state.obligations.length, 0);
  assert.deepEqual(
    Array.from(first.state.charges[1].shares, (s) => s.cents),
    [40000, 40000, 40000],
  );
  f.denied(() => f.invoice("energia"), /Ya existe/);
  f.service("energia", f.people(f.ana), {
    expectedVersion: 1,
    name: "Energía revisada",
  });
  f.denied(
    () => f.invoice("energia", { period: "2026-12" }),
    /acuerdo aprobado/,
  );
  f.invoice("claro", { period: "2026-12" });
  f.h.execute("AGREEMENT_CANCEL", { resourceId: "energia", version: 2 });
  const second = f
    .invoice("energia", { period: "2026-12" })
    .state.charges.at(-1);
  assert.equal(second.name, "energia");
  assert.equal(second.agreementVersion, 1);
  assert.equal(
    JSON.stringify(f.state().charges.find((c) => c.id === first.id)),
    original,
  );
  f.service("energia", f.people(f.ana), {
    expectedVersion: 2,
    name: "Energía revisada",
  });
  f.response("energia", "ana", 3);
  f.denied(() => f.invoice("energia", { period: "2027-01" }), /versión/);
  const third = f
    .invoice("energia", { period: "2027-01", version: 3 })
    .state.charges.at(-1);
  assert.deepEqual(
    Array.from(third.shares, (s) => s.personId),
    [f.ana.id],
  );
});

test("factura exige pago íntegro propio y crea reembolsos una sola vez", () => {
  const f = fixture();
  f.service("claro", f.people(f.ana, f.luis));
  ["ana", "luis"].forEach((user) => f.response("claro", user));
  const chargeId = f.invoice("claro").id;
  const p = {
    chargeId,
    payerId: f.ana.id,
    amount: "1200.00",
    date: "2026-11-15",
  };
  assert.equal(f.state("maria").charges.length, 0);
  f.denied(() => f.h.execute("EXPENSE_PAY", p, "luis"), /realizaste tú/);
  f.denied(
    () => f.h.execute("EXPENSE_PAY", { ...p, amount: "1199.99" }, "ana"),
    /completo/,
  );
  f.denied(() => f.h.execute("FUND", p), /no existe/);
  f.denied(
    () => f.h.execute("CHARGE", { templateId: "claro", ...p }),
    /gasto activo/,
  );
  const op = "qa-pay-retry-0001";
  f.h.execute("EXPENSE_PAY", p, "ana", op);
  const before = f.rows.length;
  assert.equal(f.h.execute("EXPENSE_PAY", p, "ana", op).duplicate, true);
  assert.equal(f.rows.length, before);
  f.denied(() => f.h.execute("EXPENSE_PAY", p, "ana"), /pendiente de pago/);
  const [ob] = f.state().obligations;
  assert.equal(f.state().obligations.length, 1);
  assert.equal(ob.creditorId, f.ana.id);
  assert.equal(ob.debtorId, f.luis.id);
  assert.equal(ob.cents, 60000);
  assert.equal(ob.agreementVersion, 1);
  assert.equal(ob.sourceId, chargeId);
  assert.equal(f.state().charges[0].providerStatus, "paid");
  assert.equal(f.state().plan.length, 0);
  f.denied(() => f.h.execute("PAYMENT", { obligationId: ob.id }), /aún no/);

  const recovered = createHarness();
  const sheet = recovered.db.insertSheet("HOME_EVENTOS");
  f.rows.forEach((row) => sheet.appendRow(row));
  const replay = recovered.context.homeGetState("token-admin");
  assert.equal(
    JSON.stringify(replay.obligations),
    JSON.stringify(f.state().obligations),
  );
  assert.equal(
    JSON.stringify(replay.charges),
    JSON.stringify(f.state().charges),
  );
});

test("compra propia pagada por persona externa: rechazo, reenvío y reutilización sin consentimientos", () => {
  const f = fixture();
  f.denied(() => f.purchase({}, "ana"), /pagaste tú/);
  f.purchase();
  assert.equal(f.state().charges.length, 0);
  f.denied(() => f.response("limpieza", "ana"), /operación de gastos/);
  f.denied(
    () => f.response("limpieza", "admin", 1, "accept", "EXPENSE_RESPOND"),
    /vinculado/,
  );
  f.denied(
    () => f.response("limpieza", "maria", 1, "accept", "EXPENSE_RESPOND"),
    /participante/,
  );
  const respond = (user, version, decision) =>
    f.response("limpieza", user, version, decision, "EXPENSE_RESPOND");
  respond("ana", 1, "accept");
  respond("luis", 1, "reject");
  assert.equal(f.state().obligations.length, 0);
  f.purchase({ expectedVersion: 1 });
  f.denied(() => respond("ana", 1, "accept"), /vigente/);
  respond("luis", 2, "accept");
  assert.equal(f.state().obligations.length, 0);
  const lastResponse = {
    resourceId: "limpieza",
    version: 2,
    decision: "accept",
  };
  const operationId = "approve-purchase-0001";
  f.h.execute("EXPENSE_RESPOND", lastResponse, "ana", operationId);
  const beforeRetry = f.rows.length;
  assert.equal(
    f.h.execute("EXPENSE_RESPOND", lastResponse, "ana", operationId).duplicate,
    true,
  );
  assert.equal(f.rows.length, beforeRetry);
  const charge = f.state().charges[0];
  assert.equal(charge.providerStatus, "paid");
  assert.equal(charge.agreementVersion, 2);
  assert.deepEqual(
    Array.from(f.state().obligations, (o) => [o.creditorId, o.cents]),
    [
      [f.maria.id, 5001],
      [f.maria.id, 5000],
    ],
  );
  assert.equal(f.state("maria").charges.length, 1);
  const luis = f.state("luis");
  assert.equal(luis.obligations.length, 1);
  assert.equal(luis.charges[0].shares.length, 1);
  assert.equal(
    luis.agreements[0].active.participants.find((p) => p.personId === f.ana.id)
      .decision,
    undefined,
  );
  f.purchase({
    resourceId: "limpieza-2",
    reuseResourceId: "limpieza",
    mode: undefined,
    participants: undefined,
  });
  assert.equal(f.state().charges.length, 1);
  const reuse = f.state().agreements.find((a) => a.resourceId === "limpieza-2");
  assert.ok(reuse.pending.participants.every((p) => p.decision === "pending"));
  f.h.execute("PERSON", {
    id: f.ana.id,
    userId: "ana",
    name: "Ana",
    active: false,
  });
  assert.equal(f.state("ana").obligations[0].cents, 5001);
  f.denied(
    () => f.response("limpieza-2", "ana", 1, "accept", "EXPENSE_RESPOND"),
    /activo vinculado/,
  );
});

test("porcentajes y montos exactos conservan centavos y permiten un pagador externo a la factura", () => {
  const f = fixture();
  f.service(
    "porcentaje",
    [
      { personId: f.ana.id, value: "33.33" },
      { personId: f.luis.id, value: "66.67" },
    ],
    { mode: "percent" },
  );
  ["ana", "luis"].forEach((user) => f.response("porcentaje", user));
  const chargeId = f.invoice("porcentaje", { amount: "100.01" }).id;
  f.h.execute(
    "EXPENSE_PAY",
    { chargeId, payerId: f.maria.id, amount: "100.01", date: "2026-11-15" },
    "maria",
  );
  assert.deepEqual(
    Array.from(f.state().obligations, (o) => o.cents),
    [3334, 6667],
  );
  f.service(
    "fijo",
    [
      { personId: f.ana.id, value: "50.01" },
      { personId: f.luis.id, value: "50.00" },
    ],
    { mode: "amount", fixedAmount: "100.01" },
  );
  ["ana", "luis"].forEach((user) => f.response("fijo", user));
  f.denied(() => f.invoice("fijo", { amount: "100.00" }), /importe/);
  const invoice = f.invoice("fijo", { amount: "100.01" }).state.charges.at(-1);
  assert.deepEqual(
    Array.from(invoice.shares, (s) => s.cents),
    [5001, 5000],
  );
});

test("mezcla esquemas 1–3 y protecciones de recursos sin modificar eventos anteriores", () => {
  const f = fixture();
  const old = f.h.execute("TEMPLATE", {
    name: "Histórico",
    participants: f.people(f.ana),
  }).state.templates[0];
  f.h.execute("CHARGE", {
    templateId: old.id,
    amount: "100",
    period: "2026-11",
    date: "2026-11-01",
    due: "2026-11-20",
  });
  const historical = JSON.stringify(f.rows);
  const n = f.rows.length;
  f.service("agua", f.people(f.ana));
  f.denied(() => f.service(old.id, f.people(f.ana)), /otro gasto/);
  f.denied(
    () =>
      f.h.execute("AGREEMENT_PROPOSE", {
        expectedVersion: 1,
        proposal: { resourceId: "agua" },
      }),
    /operación de gastos/,
  );
  f.denied(
    () => f.h.execute("EXPENSE_SERVICE", { resourceId: "otro" }, "ana"),
    /administrador/,
  );
  f.response("agua", "ana");
  f.invoice("agua");
  assert.equal(JSON.stringify(f.rows.slice(0, n)), historical);
  assert.deepEqual(
    f.rows.slice(n).map((r) => JSON.parse(r[4]).schema),
    [3, 2, 3],
  );
  assert.equal(f.state("luis").expenseResources.length, 0);
  assert.equal(f.state("luis").charges.length, 0);
  assert.equal(f.state().obligations.length, 1);
  assert.equal(f.state().obligations[0].creditorId, "");
});
