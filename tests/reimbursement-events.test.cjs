const { test } = require("node:test");
const assert = require("node:assert/strict");
const { createHarness } = require("../tools/harness.cjs");

function fixture() {
  const h = createHarness();
  const [ana, luis, maria] = h.seed();
  const people = [ana, luis].map(({ id }) => ({ personId: id }));
  h.execute("EXPENSE_SERVICE", {
    resourceId: "energia",
    name: "Energía",
    mode: "equal",
    participants: people,
    expectedVersion: 0,
  });
  for (const user of ["ana", "luis"])
    h.execute(
      "AGREEMENT_RESPOND",
      { resourceId: "energia", version: 1, decision: "accept" },
      user,
    );
  const invoice = h.execute("EXPENSE_INVOICE", {
    resourceId: "energia",
    version: 1,
    amount: "1200.00",
    period: "2026-11",
    date: "2026-11-01",
    due: "2026-11-30",
  });
  h.execute(
    "EXPENSE_PAY",
    {
      chargeId: invoice.id,
      payerId: ana.id,
      amount: "1200.00",
      date: "2026-11-08",
    },
    "ana",
  );
  const state = (user = "admin") => h.context.homeGetState("token-" + user);
  const ob = state().obligations[0].id;
  const sheet = h.db.getSheetByName("HOME_EVENTOS");
  const declare = (amount, user = "luis", id, fields = {}) =>
    h.execute(
      "REIMBURSEMENT_DECLARE",
      { obligationId: ob, amount, date: "2026-11-10", ...fields },
      user,
      id,
    );
  const decide = (paymentId, decision, user = "ana", id, fields = {}) =>
    h.execute(
      "REIMBURSEMENT_DECIDE",
      { paymentId, decision, ...fields },
      user,
      id,
    );
  const balance = (items = state().obligations) => {
    const o = items.find((x) => x.id === ob);
    return [o.cents, o.paid, o.reserved, o.pending, o.available, o.status];
  };
  function denied(run, message) {
    const n = sheet.rows.length;
    assert.throws(run, message);
    assert.equal(sheet.rows.length, n);
    assert.equal(h.locked, false);
  }
  return {
    h,
    ana,
    luis,
    maria,
    state,
    ob,
    sheet,
    balance,
    declare,
    decide,
    denied,
  };
}

test("abonos parciales y confirmación liquidan exactamente el principal sin cambiar factura ni acuerdo", () => {
  const f = fixture();
  const snapshot = JSON.stringify([f.state().charges, f.state().agreements]);
  const before = JSON.stringify(f.sheet.rows);
  const n = f.sheet.rows.length;
  const a = f.declare("200.00");
  const b = f.declare("150.00");
  assert.deepEqual(f.balance(), [60000, 0, 35000, 60000, 25000, "open"]);
  f.denied(() => f.declare("250.01"), /saldo disponible/);
  f.decide(a.id, "confirm");
  assert.deepEqual(f.balance(), [60000, 20000, 15000, 40000, 25000, "open"]);
  const c = f.declare("250.00");
  assert.deepEqual(f.balance(), [60000, 20000, 40000, 40000, 0, "open"]);
  f.decide(b.id, "confirm");
  f.decide(c.id, "confirm");
  assert.deepEqual(f.balance(), [60000, 60000, 0, 0, 0, "settled"]);
  assert.equal(
    JSON.stringify([f.state().charges, f.state().agreements]),
    snapshot,
  );
  assert.equal(JSON.stringify(f.sheet.rows.slice(0, n)), before);
  assert.equal(f.state().plan.length, 0);
});

test("abono total, rechazo con motivo y reintentos conservan una sola reserva o confirmación", () => {
  const f = fixture();
  const a = f.declare("600", "luis", "abono-completo-0001", {
    note: "Transferencia ficticia",
  });
  assert.equal(
    f.declare("600", "luis", "abono-completo-0001", {
      note: "Transferencia ficticia",
    }).duplicate,
    true,
  );
  f.denied(() => f.declare("599", "luis", "abono-completo-0001"), /original/);
  f.denied(() => f.decide(a.id, "reject"), /Motivo/);
  const rejection = f.decide(a.id, "reject", "ana", "rechazo-abono-0001", {
    reason: "No se recibió",
  });
  assert.equal(
    f.decide(a.id, "reject", "ana", "rechazo-abono-0001", {
      reason: "No se recibió",
    }).duplicate,
    true,
  );
  assert.deepEqual(f.balance(), [60000, 0, 0, 60000, 60000, "open"]);
  const rejected = f.state("luis").payments[0];
  assert.equal(rejected.reason, "No se recibió");
  assert.ok(rejected.declaredAt && rejected.decidedAt);
  assert.equal(rejected.status, "rejected");
  assert.equal(f.state().audit.find((x) => x.id === rejection.id).actor, "ana");
  f.denied(() => f.decide(a.id, "confirm"), /atendido/);
  const b = f.declare("600");
  f.decide(b.id, "confirm", "ana", "confirmar-abono-0001");
  assert.equal(
    f.decide(b.id, "confirm", "ana", "confirmar-abono-0001").duplicate,
    true,
  );
  f.denied(() => f.decide(b.id, "confirm"), /atendido/);
  assert.deepEqual(f.balance(), [60000, 60000, 0, 0, 0, "settled"]);
});

test("identidad privada y vínculo real para ambas partes, sin delegación administrativa ni rutas antiguas", () => {
  const f = fixture();
  f.h.execute("PERSON", { name: "Administrador", userId: "admin" });
  for (const user of ["ana", "maria", "admin"])
    f.denied(
      () =>
        f.declare("10", user, undefined, {
          personId: f.luis.id,
          actor: "luis",
        }),
      /no disponible/,
    );
  for (const amount of ["0", "-1", "1.001", "600.01"])
    f.denied(() => f.declare(amount), /monto|saldo/);
  f.denied(
    () => f.declare("10", "luis", undefined, { date: "2026-11-07" }),
    /fecha/,
  );
  f.denied(
    () => f.declare("10", "luis", undefined, { date: "2026-11-17" }),
    /fecha/,
  );
  const a = f.declare("10");
  for (const user of ["luis", "maria", "admin"])
    f.denied(
      () => f.decide(a.id, "confirm", user, undefined, { personId: f.ana.id }),
      /no disponible/,
    );
  for (const type of ["CONFIRM", "REJECT"])
    f.denied(
      () => f.h.execute(type, { paymentId: a.id, reason: "intento" }),
      /no existe/,
    );
  f.denied(
    () => f.h.execute("VOID", { targetId: a.id, reason: "intento" }),
    /anular/,
  );
  assert.equal(f.state("maria").payments.length, 0);
  assert.equal(f.state("maria").obligations.length, 0);
  assert.equal(f.state("ana").payments.length, 1);
  assert.equal(f.state("luis").payments.length, 1);
  assert.equal(f.state().payments.length, 1);
  f.h.execute("PERSON", {
    id: f.ana.id,
    name: "Ana",
    userId: "ana",
    active: false,
  });
  assert.equal(f.state("ana").payments.length, 1);
  f.decide(a.id, "confirm");
  assert.equal(f.state("ana").payments[0].status, "confirmed");
});

test("una obligación de compra ocasional admite el mismo flujo sin cambiar su origen", () => {
  const f = fixture();
  f.h.execute(
    "EXPENSE_PURCHASE",
    {
      resourceId: "limpieza",
      name: "Limpieza",
      expectedVersion: 0,
      amount: "10.01",
      payerId: f.ana.id,
      date: "2026-11-09",
      mode: "equal",
      participants: [{ personId: f.luis.id }],
    },
    "ana",
  );
  f.h.execute(
    "EXPENSE_RESPOND",
    {
      resourceId: "limpieza",
      version: 1,
      decision: "accept",
    },
    "luis",
  );
  const original = f
    .state()
    .obligations.find((o) => o.expenseKind === "occasional");
  const abono = f.declare("10.01", "luis", undefined, {
    obligationId: original.id,
  });
  f.decide(abono.id, "confirm");
  const settled = f.state().obligations.find((o) => o.id === original.id);
  assert.equal(settled.cents, 1001);
  assert.equal(settled.paid, 1001);
  assert.equal(settled.pending, 0);
  assert.equal(settled.status, "settled");
  assert.equal(settled.sourceId, original.sourceId);
  assert.equal(settled.agreementVersion, original.agreementVersion);
  assert.equal(f.state().obligations.find((o) => o.id === f.ob).pending, 60000);
});

test("la cuenta y el saldo se vuelven a comprobar dentro del bloqueo para operaciones intercaladas", () => {
  const f = fixture();
  const native = f.h.context.LockService.getScriptLock;
  f.h.context.LockService.getScriptLock = () => {
    const lock = native();
    return {
      ...lock,
      waitLock: () => {
        f.h.context.LockService.getScriptLock = native;
        f.declare("500");
        lock.waitLock();
      },
    };
  };
  assert.throws(() => f.declare("101"), /saldo disponible/);
  assert.deepEqual(f.balance(), [60000, 0, 50000, 60000, 10000, "open"]);
  f.h.context.LockService.getScriptLock = () => {
    const lock = native();
    return {
      ...lock,
      waitLock: () => {
        f.h.users.find((u) => u.id === "luis").estado = "Inactivo";
        lock.waitLock();
      },
    };
  };
  f.denied(() => f.declare("1"), /no está activa/);
});

test("reconstrucción de esquemas 1 a 4 con estados históricos y sin doble contabilización", () => {
  const f = fixture();
  const old = f.h.execute("DEBT", {
    name: "Histórica",
    debtorId: f.luis.id,
    creditorId: f.maria.id,
    amount: "25",
    date: "2026-11-01",
    due: "2026-11-20",
  }).id;
  const legacy = f.h.execute(
    "PAYMENT",
    { obligationId: old, amount: "5", date: "2026-11-10" },
    "luis",
  );
  f.h.execute("CONFIRM", { paymentId: legacy.id }, "maria");
  const a = f.declare("200");
  f.decide(a.id, "confirm");
  f.declare("150");
  const recovered = createHarness();
  const sheet = recovered.db.insertSheet("HOME_EVENTOS");
  f.sheet.rows.forEach((row) => sheet.appendRow(row));
  const replay = recovered.context.homeGetState("token-admin");
  assert.deepEqual(f.balance(replay.obligations), [
    60000,
    20000,
    15000,
    40000,
    25000,
    "open",
  ]);
  assert.equal(replay.obligations.find((o) => o.id === old).pending, 2000);
  assert.equal(JSON.stringify(replay), JSON.stringify(f.state()));
  const schemas = new Set(
    f.sheet.rows.slice(1).map((row) => JSON.parse(row[4]).schema),
  );
  assert.deepEqual([...schemas].sort(), [1, 2, 3, 4]);
});
