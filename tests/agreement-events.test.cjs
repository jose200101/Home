const { test } = require("node:test");
const assert = require("node:assert/strict");
const { createHarness } = require("../tools/harness.cjs");
const plain = (value) => JSON.parse(JSON.stringify(value));

test("sin acuerdos la consulta histórica no introduce validaciones nuevas de vinculación", () => {
  const h = createHarness(); h.seed();
  h.db.getSheetByName("HOME_EVENTOS").appendRow(["historical-person", "2026-11-01", "admin", "PERSON",
    JSON.stringify({schema: 1, id: "historical-duplicate", name: "Persona histórica", userId: "ana", active: true})]);
  assert.equal(h.context.homeGetState("token-ana").agreements.length, 0);
  assert.equal(h.context.homeGetState("token-admin").agreements.length, 0);
});

function fixture() {
  const h = createHarness();
  const [a, b, c] = h.seed(); // Los IDs de personas son diferentes de ana/luis/maria.
  const template = h.execute("TEMPLATE", {
    name: "Energía",
    mode: "equal",
    participants: [a, b].map((person) => ({ personId: person.id })),
  }).state.templates[0];
  const resourceId = template.id;
  const proposal = (extra = {}) => ({
    resourceId,
    kind: "recurring",
    mode: "equal",
    participants: [a, b].map((person) => ({ personId: person.id })),
    ...extra,
  });
  const propose = (
    value = proposal(),
    expectedVersion = 0,
    user = "admin",
    id,
  ) =>
    h.execute(
      "AGREEMENT_PROPOSE",
      { proposal: value, expectedVersion },
      user,
      id,
    );
  const respond = (
    user,
    version = 1,
    decision = "accept",
    id,
    resource = resourceId,
  ) =>
    h.execute(
      "AGREEMENT_RESPOND",
      { resourceId: resource, version, decision },
      user,
      id,
    );
  const cancel = (version, id, user = "admin", resource = resourceId) =>
    h.execute("AGREEMENT_CANCEL", { resourceId: resource, version }, user, id);
  const raw = () => h.context.homeProject_(h.context.homeRead_());
  const approved = () => {
    propose();
    respond("ana");
    return respond("luis");
  };
  const charge = (extra) =>
    h.execute("CHARGE", {
      templateId: resourceId,
      period: "2026-11",
      amount: "100.01",
      date: "2026-11-01",
      due: "2026-11-20",
      ...extra,
    });
  return {
    h,
    a,
    b,
    c,
    template,
    resourceId,
    proposal,
    propose,
    respond,
    cancel,
    raw,
    approved,
    charge,
  };
}

test("eventos de acuerdos reconstruyen aprobaciones sin generar obligaciones", () => {
  const f = fixture();
  const first = f.propose();
  assert.equal(first.state.agreements[0].pending.version, 1);
  assert.equal(first.state.obligations.length, 0);
  assert.equal(f.respond("ana").state.agreements[0].active, null);
  const final = f.respond("luis").state;
  assert.equal(final.agreements[0].active.version, 1);
  assert.equal(final.agreements[0].pending, null);
  assert.equal(final.obligations.length, 0);
  const events = f.h.context
    .homeRead_()
    .filter((e) => e.type.startsWith("AGREEMENT_"));
  assert.equal(events.length, 3);
  for (const e of events) {
    assert.equal(e.data.schema, 2);
    assert.equal(e.data.eventVersion, 1);
    assert.equal(e.data.version, 1);
    assert.equal(e.actor, e.data.actor.id);
    assert.ok(!Number.isNaN(Date.parse(e.at)));
    assert.ok(e.id);
  }
  assert.notEqual(events[1].actor, events[1].data.actor.personId);
  const recovered = createHarness();
  recovered.db
    .insertSheet("HOME_EVENTOS")
    .rows.push(...structuredClone(f.h.db.getSheetByName("HOME_EVENTOS").rows));
  assert.deepEqual(
    plain(recovered.context.homeGetState("token-admin").agreements),
    plain(f.h.context.homeGetState("token-admin").agreements),
  );
});

test("rechazo, reenvío y respuestas requieren una versión vigente y aceptaciones nuevas", () => {
  const f = fixture();
  f.propose();
  f.respond("ana");
  f.respond("luis", 1, "reject");
  assert.equal(f.raw().agreements[0].pending.status, "rejected");
  assert.throws(() => f.respond("luis"), /rechazada/);
  f.propose(f.proposal(), 1);
  assert.equal(f.raw().agreements[0].pending.version, 2);
  assert.ok(
    f
      .raw()
      .agreements[0].pending.participants.every(
        (p) => p.decision === "pending",
      ),
  );
  const rows = f.h.db.getSheetByName("HOME_EVENTOS").getLastRow();
  assert.throws(() => f.respond("ana", 1), /versión/);
  assert.equal(f.h.db.getSheetByName("HOME_EVENTOS").getLastRow(), rows);
  f.respond("ana", 2);
  f.respond("luis", 2);
  assert.equal(f.raw().agreements[0].active.version, 2);
});

test("modificación pendiente o rechazada suspende CHARGE; cancelar restaura el acuerdo anterior", () => {
  const f = fixture();
  const oldCharge = f.charge({ period: "2026-10" }).state.charges[0];
  f.approved();
  f.propose(
    f.proposal({ participants: [f.a, f.c].map((p) => ({ personId: p.id })) }),
    1,
  );
  assert.equal(f.raw().agreements[0].active.version, 1);
  assert.throws(() => f.charge(), /acuerdo aprobado/);
  f.respond("maria", 2, "reject");
  assert.throws(() => f.charge(), /acuerdo aprobado/);
  f.cancel(2);
  const result = f.charge().state;
  assert.equal(result.agreements[0].active.version, 1);
  assert.equal(result.agreements[0].canInvoice, true);
  assert.equal(result.charges[1].agreementVersion, 1);
  assert.deepEqual(plain(result.charges[0]), plain(oldCharge));
  f.propose(f.proposal(), 2);
  assert.equal(f.raw().agreements[0].pending.version, 3);
  assert.throws(() => f.cancel(2), /versión/);
});

test("contrato aprobado fija versión y reparto; no admite sobreescrituras del cliente", () => {
  const f = fixture();
  f.propose(
    f.proposal({
      mode: "percent",
      participants: [
        { personId: f.a.id, value: "60" },
        { personId: f.b.id, value: "40" },
      ],
    }),
  );
  f.respond("ana");
  f.respond("luis");
  assert.throws(() => f.charge({ mode: "equal" }), /distribución/);
  assert.throws(
    () =>
      f.charge({
        participants: [
          { personId: f.a.id, value: "50" },
          { personId: f.b.id, value: "50" },
        ],
      }),
    /no coinciden/,
  );
  const contract = f.h.context.homeApprovedAgreement_(
    f.raw(),
    f.resourceId,
    "recurring",
  );
  assert.equal(contract.version, 1);
  assert.equal(contract.mode, "percent");
  contract.participants[0].value = "99";
  assert.equal(f.raw().agreements[0].active.participants[0].basis, 6000);
  assert.deepEqual(
    f.charge().state.charges[0].shares.map((p) => p.cents),
    [6001, 4000],
  );
  f.propose(
    f.proposal({
      mode: "amount",
      amountCents: 10001,
      participants: [
        { personId: f.a.id, value: "60.01" },
        { personId: f.b.id, value: "40" },
      ],
    }),
    1,
  );
  f.respond("ana", 2);
  f.respond("luis", 2);
  assert.throws(
    () => f.charge({ period: "2026-12", amount: "101" }),
    /importe/,
  );
  assert.equal(
    f.charge({ period: "2026-12" }).state.charges[1].agreementVersion,
    2,
  );
});

test("ocasionales: integrante activo propone su compra; aprobación no crea una deuda definitiva", () => {
  const f = fixture();
  const value = f.proposal({
    resourceId: "compra-1",
    kind: "occasional",
    amountCents: 60000,
    payerId: f.a.id,
    participants: [f.b, f.c].map((p) => ({ personId: p.id })),
  });
  assert.throws(() => f.propose(value, 0, "luis"), /Solo quien pagó/);
  assert.throws(
    () => f.propose({ ...value, resourceId: f.resourceId }, 0, "ana"),
    /identificador/,
  );
  f.propose(value, 0, "ana");
  assert.throws(
    () => f.respond("ana", 1, "accept", undefined, "compra-1"),
    /involucrado/,
  );
  f.respond("luis", 1, "accept", undefined, "compra-1");
  const state = f.respond("maria", 1, "accept", undefined, "compra-1").state;
  assert.equal(state.agreements[0].canActivate, true);
  assert.equal(state.agreements[0].canInvoice, false);
  assert.equal(state.obligations.length, 0);
  assert.equal(
    f.h.context.homeApprovedAgreement_(f.raw(), "compra-1", "occasional")
      .payerId,
    f.a.id,
  );
});

test("una compra cancelada conserva autoría; otro integrante o administrador no la reasigna", () => {
  const f = fixture();
  const value = f.proposal({
    resourceId: "compra-1",
    kind: "occasional",
    amountCents: 60000,
    payerId: f.a.id,
  });
  f.propose(value, 0, "ana");
  f.cancel(1, undefined, "ana", "compra-1");
  assert.throws(
    () => f.propose({ ...value, payerId: f.b.id }, 1, "luis"),
    /Solo quien registró/,
  );
  const adminPerson = f.h
    .execute("PERSON", { name: "Administrador integrante", userId: "admin" })
    .state.people.find((p) => p.userId === "admin");
  assert.throws(
    () => f.propose({ ...value, payerId: adminPerson.id }, 1, "admin"),
    /Solo quien registró/,
  );
  f.propose(value, 1, "ana");
  assert.equal(f.raw().agreements[0].pending.version, 2);
});

test("la sesión y el vínculo real impiden suplantación y aceptación administrativa ajena", () => {
  const f = fixture();
  assert.throws(() => f.propose(f.proposal(), 0, "ana"), /administrador/);
  f.propose();
  assert.throws(() => f.respond("admin"), /integrante activo/);
  assert.throws(
    () =>
      f.h.execute("AGREEMENT_RESPOND", {
        resourceId: f.resourceId,
        version: 1,
        decision: "accept",
        personId: f.a.id,
      }),
    /identidad/,
  );
  f.h.users.push({
    id: f.a.id,
    nombre_usuario: "colision",
    rol: "Integrante",
    estado: "Activo",
  });
  assert.throws(() => f.respond(f.a.id), /integrante activo/);
  assert.throws(() => f.respond("maria"), /involucrado/);
  assert.throws(() => f.cancel(1, undefined, "ana"), /administrador/);
  f.h.execute("PERSON", { ...f.a, userId: "" });
  assert.throws(() => f.respond("ana"), /integrante activo/);
  assert.throws(() => f.h.context.homeGetState("token-inventado"), /Sesión/);
  assert.equal(
    f.raw().agreements[0].pending.participants[0].decision,
    "pending",
  );
});

test("administrador vinculado acepta solo su propia participación, sin privilegio de delegación", () => {
  const f = fixture();
  const adminPerson = f.h
    .execute("PERSON", { name: "Administrador integrante", userId: "admin" })
    .state.people.find((p) => p.userId === "admin");
  f.propose(
    f.proposal({
      participants: [adminPerson, f.a].map((p) => ({ personId: p.id })),
    }),
  );
  f.respond("admin");
  assert.equal(f.raw().agreements[0].active, null);
  assert.throws(() => f.respond("admin"), /ya respondió/);
  f.respond("ana");
  assert.equal(f.raw().agreements[0].active.version, 1);
});

test("privacidad por versión y participante; no expone actores ni solicitudes ajenas", () => {
  const f = fixture();
  f.approved();
  f.propose(
    f.proposal({ participants: [f.a, f.c].map((p) => ({ personId: p.id })) }),
    1,
  );
  const luis = f.h.context.homeGetState("token-luis").agreements[0];
  assert.equal(luis.active.version, 1);
  assert.equal(luis.pending, null);
  assert.equal(luis.suspended, true);
  assert.equal(
    luis.active.participants.find((p) => p.personId === f.a.id).decision,
    undefined,
  );
  const maria = f.h.context.homeGetState("token-maria").agreements[0];
  assert.equal(maria.active, null);
  assert.equal(maria.pending.version, 2);
  assert.doesNotMatch(
    JSON.stringify(maria),
    /"createdBy":|requestJson|"actor":/,
  );
  f.h.users.push({
    id: "outsider",
    nombre_usuario: "outsider",
    estado: "Activo",
    rol: "Integrante",
  });
  assert.equal(f.h.context.homeGetState("token-outsider").agreements.length, 0);
  assert.equal(
    f.h.context.homeGetState("token-admin").agreements[0].pending.createdBy,
    "admin",
  );
});

test("retiro bloquea nuevas participaciones y aprobaciones, sin borrar historia ni obligaciones", () => {
  const f = fixture();
  f.approved();
  const priorCharge = f.charge().state.obligations;
  f.propose(f.proposal(), 1);
  f.respond("ana", 2);
  f.h.execute("PERSON", { ...f.a, active: false });
  assert.throws(() => f.respond("ana", 2), /integrante activo/);
  assert.throws(() => f.respond("luis", 2), /desactivado/);
  f.cancel(2);
  assert.throws(() => f.propose(f.proposal(), 2), /desactivado/);
  assert.throws(() => f.charge({ period: "2026-12" }), /integrantes activos/);
  assert.deepEqual(plain(f.raw().obligations), plain(priorCharge));
  const historical = f.h.context.homeGetState("token-ana").agreements[0];
  assert.equal(historical.active.version, 1);
  assert.equal(historical.canInvoice, false);
});

test("idempotencia de propuesta, respuesta y cancelación no duplica ni acepta una versión nueva", () => {
  const f = fixture();
  const proposalId = "op-propuesta-123456",
    responseId = "op-respuesta-123456",
    cancelId = "op-cancelar-123456";
  f.propose(f.proposal(), 0, "admin", proposalId);
  f.respond("ana", 1, "accept", responseId);
  f.respond("luis", 1, "reject");
  f.cancel(1, cancelId);
  f.propose(f.proposal(), 1);
  const before = f.h.db.getSheetByName("HOME_EVENTOS").getLastRow();
  assert.equal(f.propose(f.proposal(), 0, "admin", proposalId).duplicate, true);
  assert.equal(f.respond("ana", 1, "accept", responseId).duplicate, true);
  assert.equal(f.cancel(1, cancelId).duplicate, true);
  assert.equal(f.h.db.getSheetByName("HOME_EVENTOS").getLastRow(), before);
  assert.ok(
    f
      .raw()
      .agreements[0].pending.participants.every(
        (p) => p.decision === "pending",
      ),
  );
  assert.throws(
    () => f.respond("ana", 2, "accept", responseId),
    /datos diferentes/,
  );
  assert.throws(() => f.respond("luis", 1, "accept", responseId), /utilizado/);
  assert.throws(
    () => f.propose(f.proposal(), 2, "admin", 123456789012),
    /Identificador/,
  );
  assert.equal(f.h.locked, false);
});

test("una escritura ganadora durante la espera del lock invalida la propuesta obsoleta", () => {
  const f = fixture();
  f.propose();
  f.respond("ana", 1, "reject");
  const baseGetLock = f.h.context.LockService.getScriptLock;
  let interleave = true;
  f.h.context.LockService.getScriptLock = () => {
    const lock = baseGetLock();
    return {
      waitLock(timeout) {
        if (interleave) {
          interleave = false;
          f.propose(f.proposal(), 1, "admin", "winner-operation-123456");
        }
        lock.waitLock(timeout);
      },
      releaseLock() {
        lock.releaseLock();
      },
    };
  };
  const sheet = f.h.db.getSheetByName("HOME_EVENTOS"),
    append = sheet.appendRow;
  sheet.appendRow = (row) => {
    assert.equal(f.h.locked, true);
    return append(row);
  };
  const before = sheet.getLastRow();
  assert.throws(
    () => f.propose(f.proposal(), 1, "admin", "loser-operation-123456"),
    /versión esperada/,
  );
  assert.equal(sheet.getLastRow(), before + 1);
  assert.equal(f.raw().agreements[0].pending.version, 2);
  assert.equal(f.h.locked, false);
});

test("se revalida el rol después de adquirir el lock y se libera ante un rechazo", () => {
  const f = fixture(),
    baseGetLock = f.h.context.LockService.getScriptLock;
  f.h.context.LockService.getScriptLock = () => {
    const lock = baseGetLock();
    return {
      waitLock(timeout) {
        lock.waitLock(timeout);
        f.h.users.find((u) => u.id === "admin").rol = "Integrante";
      },
      releaseLock() {
        lock.releaseLock();
      },
    };
  };
  const before = f.h.db.getSheetByName("HOME_EVENTOS").getLastRow();
  assert.throws(() => f.propose(), /administrador/);
  assert.equal(f.h.db.getSheetByName("HOME_EVENTOS").getLastRow(), before);
  assert.equal(f.h.locked, false);
});

test("lector distingue pares tipo/esquema y rechaza versiones futuras o actores inconsistentes", () => {
  const f = fixture();
  f.propose();
  const rows = structuredClone(f.h.db.getSheetByName("HOME_EVENTOS").rows);
  const row = rows[rows.length - 1];
  for (const change of [
    (data) => (data.schema = 1),
    (data) => (data.eventVersion = 2),
    (data) => (data.schema = 3),
  ]) {
    const copy = structuredClone(rows),
      data = JSON.parse(row[4]);
    change(data);
    copy[copy.length - 1][4] = JSON.stringify(data);
    const h = createHarness();
    h.db.insertSheet("HOME_EVENTOS").rows.push(...copy);
    assert.throws(
      () => h.context.homeGetState("token-admin"),
      /Historial inválido/,
    );
  }
  const corrupt = structuredClone(rows),
    data = JSON.parse(row[4]);
  data.actor.id = "otro";
  corrupt[corrupt.length - 1][4] = JSON.stringify(data);
  const h = createHarness();
  h.db.insertSheet("HOME_EVENTOS").rows.push(...corrupt);
  assert.throws(
    () => h.context.homeGetState("token-admin"),
    /Historial de acuerdos inválido/,
  );
  const old = createHarness(),
    sh = old.db.insertSheet("HOME_EVENTOS");
  sh.rows.push(rows[0], [
    "old-init",
    "2026-11-01",
    "admin",
    "INIT",
    JSON.stringify({ schema: 2, name: "Casa" }),
  ]);
  assert.throws(() => old.context.homeRead_(), /Historial inválido/);
});

test("los diez tipos históricos conservan saldos y registros al mezclar eventos nuevos", () => {
  const f = fixture();
  const charge = f.charge();
  f.h.execute("FUND", {
    chargeId: charge.id,
    payerId: f.a.id,
    date: "2026-11-10",
  });
  const debt = f.h.execute("DEBT", {
    name: "Deuda anterior",
    debtorId: f.a.id,
    creditorId: f.b.id,
    amount: "50",
    date: "2026-11-01",
    due: "2026-11-20",
  });
  const p1 = f.h.execute(
    "PAYMENT",
    { obligationId: debt.id, amount: "10", date: "2026-11-10" },
    "ana",
  );
  f.h.execute("REJECT", { paymentId: p1.id, reason: "No recibido" }, "luis");
  const p2 = f.h.execute(
    "PAYMENT",
    { obligationId: debt.id, amount: "10", date: "2026-11-10" },
    "ana",
  );
  f.h.execute("CONFIRM", { paymentId: p2.id }, "luis");
  f.h.execute("VOID", { targetId: p2.id, reason: "Corrección" });
  const before = plain(f.raw()),
    rows = structuredClone(f.h.db.getSheetByName("HOME_EVENTOS").rows);
  const types = new Set(f.h.context.homeRead_().map((e) => e.type));
  for (const type of [
    "INIT",
    "PERSON",
    "TEMPLATE",
    "CHARGE",
    "DEBT",
    "PAYMENT",
    "CONFIRM",
    "REJECT",
    "FUND",
    "VOID",
  ])
    assert.ok(types.has(type), type);
  f.propose(f.proposal({ resourceId: "otro-servicio" }));
  const after = plain(f.raw());
  for (const key of [
    "initialized",
    "household",
    "people",
    "templates",
    "obligations",
    "charges",
    "payments",
  ])
    assert.deepEqual(after[key], before[key], key);
  assert.deepEqual(
    f.h.db.getSheetByName("HOME_EVENTOS").rows.slice(0, rows.length),
    rows,
  );
  assert.ok(
    f.h.context
      .homeRead_()
      .filter((e) => !e.type.startsWith("AGREEMENT_"))
      .every((e) => e.data.schema === 1),
  );
});
