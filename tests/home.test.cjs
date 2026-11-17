const { test } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const { createHarness } = require("../tools/harness.cjs");
function fixture() {
  const h = createHarness();
  const [a, b, c] = h.seed();
  const t = h.execute("TEMPLATE", {
    name: "Claro",
    provider: "Claro",
    mode: "equal",
    participants: [a, b, c].map((p) => ({ personId: p.id })),
  }).state.templates[0];
  const charge = (payload) =>
    h.execute("CHARGE", {
      templateId: t.id,
      period: "2026-11",
      amount: "900",
      date: "2026-11-01",
      due: "2026-11-20",
      ...payload,
    });
  return { h, a, b, c, t, charge };
}
test("reparto exacto de centavos, porcentajes y montos; validación de entradas", () => {
  const { context: c } = createHarness();
  const p = ["a", "b", "c"].map((personId) => ({ personId }));
  assert.deepEqual(
    c.homeSplit_(100, p, "equal").map((p) => p.cents),
    [34, 33, 33],
  );
  assert.deepEqual(
    c
      .homeSplit_(
        999,
        [
          { personId: "a", value: "25" },
          { personId: "b", value: "75" },
        ],
        "percent",
      )
      .map((p) => p.cents),
    [250, 749],
  );
  assert.deepEqual(
    c
      .homeSplit_(
        300,
        [
          { personId: "a", value: "1" },
          { personId: "b", value: "2" },
        ],
        "amount",
      )
      .map((p) => p.cents),
    [100, 200],
  );
  for (const amount of ["-1", "1.001", "NaN", "1,000", "Infinity", ""])
    assert.throws(() => c.homeCents_(amount));
  assert.throws(() => c.homeSplit_(100, [], "equal"));
  assert.throws(() => c.homeSplit_(100, [p[0], p[0]], "equal"));
  assert.throws(() =>
    c.homeSplit_(100, [{ personId: "a", value: "99" }], "percent"),
  );
  assert.throws(() =>
    c.homeSplit_(100, [{ personId: "a", value: "2" }], "amount"),
  );
  assert.throws(() => c.homeDate_("2026-02-30"));
});
test("energía entre dos y Claro entre tres mantienen participantes independientes", () => {
  const { h, a, b, c, t, charge } = fixture();
  charge();
  const energy = h
    .execute("TEMPLATE", {
      name: "Energía",
      mode: "equal",
      participants: [a, b].map((p) => ({ personId: p.id })),
    })
    .state.templates.find((v) => v.name === "Energía");
  const result = h.execute("CHARGE", {
    templateId: energy.id,
    period: "2026-11",
    amount: "1000",
    date: "2026-11-01",
    due: "2026-11-20",
  });
  assert.equal(result.state.obligations.length, 5);
  h.execute("TEMPLATE", {
    ...t,
    participants: [a, b].map((p) => ({ personId: p.id })),
  });
  const next = charge({
    period: "2026-12",
    date: "2026-11-16",
    amount: "900",
  }).state;
  assert.equal(next.charges[0].shares.length, 3);
  assert.equal(next.charges[2].shares.length, 2);
  assert.equal(
    next.charges[0].shares.find((p) => p.personId === c.id).cents,
    30000,
  );
});
test("adelanto al proveedor genera reembolsos sin cobrar al integrante su propio aporte", () => {
  const { h, a, charge } = fixture();
  const state = charge({ payerId: a.id }).state;
  assert.equal(state.obligations.length, 2);
  assert.equal(
    state.obligations.reduce((sum, o) => sum + o.pending, 0),
    60000,
  );
  assert.ok(state.obligations.every((o) => o.creditorId === a.id));
  const other = fixture();
  const s = other.charge().state;
  const funded = other.h.execute("FUND", {
    chargeId: s.charges[0].id,
    payerId: other.a.id,
    date: "2026-11-10",
  }).state;
  assert.equal(funded.charges[0].payerId, other.a.id);
  assert.equal(funded.obligations.filter((o) => o.pending > 0).length, 2);
  assert.throws(() =>
    other.h.execute("FUND", {
      chargeId: s.charges[0].id,
      payerId: other.b.id,
      date: "2026-11-10",
    }),
  );
});
test("pagos reservan saldo, requieren receptor y resisten sobrepagos", () => {
  const { h, a, charge } = fixture();
  const ob = charge({ payerId: a.id }).state.obligations[0];
  const p = h.execute(
    "PAYMENT",
    { obligationId: ob.id, amount: "100", date: "2026-11-10" },
    "luis",
  );
  const after = p.state.obligations.find((o) => o.id === ob.id);
  assert.equal(after.pending, 30000);
  assert.equal(after.reserved, 10000);
  assert.throws(() =>
    h.execute(
      "PAYMENT",
      { obligationId: ob.id, amount: "201", date: "2026-11-10" },
      "luis",
    ),
  );
  assert.throws(() => h.execute("CONFIRM", { paymentId: p.id }, "luis"));
  const confirmed = h
    .execute("CONFIRM", { paymentId: p.id }, "ana")
    .state.obligations.find((o) => o.id === ob.id);
  assert.equal(confirmed.pending, 20000);
  assert.equal(confirmed.reserved, 0);
  assert.throws(() => h.execute("CONFIRM", { paymentId: p.id }, "ana"));
  assert.throws(() =>
    h.execute(
      "PAYMENT",
      { obligationId: ob.id, amount: "1", date: "2026-11-17" },
      "luis",
    ),
  );
});
test("rechazo libera reserva; reversión conserva historia y restaura saldo", () => {
  const { h, a, charge } = fixture();
  const ob = charge({ payerId: a.id }).state.obligations[0];
  const p = h.execute(
    "PAYMENT",
    { obligationId: ob.id, amount: "300", date: "2026-11-10" },
    "luis",
  );
  assert.throws(() =>
    h.execute("REJECT", { paymentId: p.id, reason: "" }, "ana"),
  );
  const rejected = h.execute(
    "REJECT",
    { paymentId: p.id, reason: "No recibí el dinero" },
    "ana",
  ).state;
  assert.equal(rejected.obligations[0].reserved, 0);
  const p2 = h.execute(
    "PAYMENT",
    { obligationId: ob.id, amount: "300", date: "2026-11-10" },
    "luis",
  );
  h.execute("CONFIRM", { paymentId: p2.id }, "ana");
  const reverted = h.execute("VOID", {
    targetId: p2.id,
    reason: "Registro duplicado",
  }).state;
  assert.equal(reverted.obligations[0].pending, 30000);
  assert.equal(reverted.payments.length, 1);
  assert.ok(reverted.audit.find((e) => e.id === p2.id).voided);
});
test("idempotencia, conflicto de contenido y bloqueo liberado tras fallos", () => {
  const { h, a, charge } = fixture();
  const id = "retry-operation-123";
  const payload = {
    name: "Préstamo",
    debtorId: a.id,
    creditorId: h.context.homeGetState("token-admin").people[1].id,
    amount: "50",
    date: "2026-11-01",
    due: "2026-11-20",
  };
  const first = h.execute("DEBT", payload, "admin", id),
    rows = h.db.getSheetByName("HOME_EVENTOS").getLastRow();
  assert.equal(h.execute("DEBT", payload, "admin", id).duplicate, true);
  assert.equal(h.db.getSheetByName("HOME_EVENTOS").getLastRow(), rows);
  assert.throws(() =>
    h.execute("DEBT", { ...payload, amount: "60" }, "admin", id),
  );
  assert.equal(h.locked, false);
  assert.equal(first.state.obligations[0].pending, 5000);
  charge();
  assert.throws(() => charge());
  assert.equal(h.locked, false);
});
test("privacidad, revocación y autorización de cada operación", () => {
  const { h, a, b, c, charge } = fixture();
  charge({ payerId: a.id });
  h.execute("DEBT", {
    name: "Privado",
    debtorId: b.id,
    creditorId: c.id,
    amount: "50",
    date: "2026-11-01",
    due: "2026-11-20",
  });
  const mine = h.context.homeGetState("token-ana");
  assert.equal(mine.obligations.length, 2);
  assert.equal(mine.audit.length, 0);
  assert.equal(mine.templates.length, 0);
  assert.equal(mine.plan.length, 0);
  assert.equal(mine.charges[0].shares.length, 1);
  assert.equal(mine.charges[0].requestJson, undefined);
  assert.ok(mine.people.every((p) => p.userId === undefined));
  assert.throws(() => h.execute("DEBT", {}, "ana"));
  assert.throws(() => h.context.homeMigrationPreview("token-ana"));
  assert.throws(() =>
    h.context.homeLegacyAuthorize_([{ homeToken: "token-ana" }]),
  );
  assert.ok(
    h.context.homeLegacyAuthorize_([{ homeToken: "token-admin" }]).admin,
  );
  h.users.find((u) => u.id === "ana").estado = "Inactivo";
  assert.throws(() => h.context.homeGetState("token-ana"));
  assert.throws(() => h.context.homeGetState("forged"));
});
test("anulación rechaza movimientos relacionados; pagos al proveedor y facturas inexistentes", () => {
  const { h, a, charge } = fixture();
  const result = charge();
  const ob = result.state.obligations[0];
  assert.throws(() =>
    h.execute(
      "PAYMENT",
      { obligationId: ob.id, amount: "1", date: "2026-11-10" },
      "ana",
    ),
  );
  h.execute("PAYMENT", {
    obligationId: ob.id,
    amount: "10",
    date: "2026-11-10",
  });
  assert.throws(() =>
    h.execute("VOID", { targetId: result.id, reason: "Error" }),
  );
  assert.throws(() =>
    h.execute("FUND", {
      chargeId: result.id,
      payerId: a.id,
      date: "2026-11-10",
    }),
  );
  assert.throws(() =>
    h.execute("PAYMENT", {
      obligationId: "missing",
      amount: "1",
      date: "2026-11-10",
    }),
  );
});
test("deudas cruzadas: el plan conserva saldo y no modifica deudas", () => {
  const { h, a, b, c } = fixture();
  for (const [debtorId, creditorId, amount] of [
    [a.id, b.id, "100"],
    [b.id, c.id, "100"],
  ])
    h.execute("DEBT", {
      name: "Préstamo",
      debtorId,
      creditorId,
      amount,
      date: "2026-11-01",
      due: "2026-11-20",
    });
  const state = h.context.homeGetState("token-admin");
  assert.equal(state.obligations.length, 2);
  assert.deepEqual(JSON.parse(JSON.stringify(state.plan)), [
    { debtorId: a.id, creditorId: c.id, cents: 10000 },
  ]);
});
test("esquema inválido falla sin escribir; lecturas y previsualización no crean hojas", () => {
  const h = createHarness();
  assert.equal(h.context.homeGetState("token-admin").initialized, false);
  assert.equal(h.sheets.size, 0);
  h.context.homeMigrationPreview("token-admin");
  assert.equal(h.sheets.size, 0);
  const sh = h.db.insertSheet("HOME_EVENTOS");
  sh.appendRow(["incorrecto"]);
  assert.throws(() => h.execute("INIT", { name: "Casa" }));
  assert.equal(sh.getLastRow(), 1);
});
test("backend completo compila y todos los endpoints históricos están protegidos", () => {
  const root = path.join(__dirname, "../Backend");
  const files = [];
  const walk = (dir) => {
    for (const f of fs.readdirSync(dir, { withFileTypes: true })) {
      const p = path.join(dir, f.name);
      if (f.isDirectory()) walk(p);
      else if (p.endsWith(".js")) files.push(p);
    }
  };
  walk(root);
  const combined = files.map((f) => fs.readFileSync(f, "utf8")).join("\n");
  new vm.Script(combined);
  const allowed = new Set([
    "login",
    "logout",
    "validateToken",
    "doGet",
    "include",
    "homeGetState",
    "homeExecute",
    "homeMigrationPreview",
  ]);
  for (const m of combined.matchAll(
    /function\s+(\w+)\s*\([^)]*\)\s*\{([^]*?)(?=\n\})/g,
  )) {
    if (!m[1].endsWith("_") && !allowed.has(m[1]))
      assert.match(m[2], /^\s*homeLegacyAuthorize_\(arguments\)/, m[1]);
  }
});

test("abonos históricos rechazan montos negativos, gasto ausente, exceso y fechas inválidas antes de crear hojas", () => {
  const h = createHarness(),
    c = h.context;
  vm.runInContext(
    fs.readFileSync(
      path.join(__dirname, "../Backend/GastosVariables.js"),
      "utf8",
    ),
    c,
  );
  let writes = 0;
  c._gvEnsureSheet_ = () => {
    writes++;
    throw Error("No debe llegar a escritura");
  };
  c.listarGastosVariablesLegacy_ = () => [
    { id: "g1", pendiente: 100, fecha: "2026-11-01" },
  ];
  for (const monto of ["-1", "0", "101", "1.001", "abc"])
    assert.throws(() =>
      c.registrarPagoGastoVariableLegacy_({
        gastoId: "g1",
        monto,
        fechaPago: "2026-11-10",
      }),
    );
  assert.throws(() =>
    c.registrarPagoGastoVariableLegacy_({
      gastoId: "g1",
      monto: "10",
      fechaPago: "2026-11-17",
    }),
  );
  c.listarGastosVariablesLegacy_ = () => [];
  assert.throws(() =>
    c.registrarPagoGastoVariableLegacy_({
      gastoId: "missing",
      monto: "10",
      fechaPago: "2026-11-10",
    }),
  );
  assert.equal(writes, 0);
  assert.equal(h.locked, false);
});

test("archivar integrantes y tipos conserva historial; obliga a revisar participantes futuros", () => {
  const { h, a, b, t, charge } = fixture();
  charge();
  h.execute("PERSON", { ...a, active: false });
  assert.throws(() => charge({ period: "2026-12" }));
  const archived = h.execute("TEMPLATE", { ...t, active: false }).state;
  assert.equal(archived.charges[0].shares.length, 3);
  assert.equal(archived.templates[0].active, false);
  assert.throws(() =>
    h.execute("PERSON", { name: "Duplicado", userId: b.userId }),
  );
});

test("wrappers históricos verifican sesión antes de llamar implementación privada y preservan argumentos", () => {
  const h = createHarness(),
    c = h.context;
  vm.runInContext(
    fs.readFileSync(
      path.join(__dirname, "../Backend/LegacyEndpoints.js"),
      "utf8",
    ),
    c,
  );
  let received;
  c.guardarPersonaLegacy_ = (...args) => {
    received = args;
    return { ok: true };
  };
  assert.throws(() => c.guardarPersona({ nombre_persona: "Ana" }));
  assert.equal(received, undefined);
  assert.throws(() =>
    c.guardarPersona({ nombre_persona: "Ana" }, { homeToken: "token-ana" }),
  );
  assert.equal(
    c.guardarPersona({ nombre_persona: "Ana" }, { homeToken: "token-admin" })
      .ok,
    true,
  );
  assert.equal(received.length, 1);
  assert.equal(received[0].nombre_persona, "Ana");
});

test("login reconoce el encabezado real nombre_persona y rechaza estado vacío", () => {
  const h = createHarness(),
    c = h.context;
  vm.runInContext(
    fs.readFileSync(path.join(__dirname, "../Backend/Code.js"), "utf8"),
    c,
  );
  let status = "Activo";
  const cache = new Map();
  c.CacheService = {
    getScriptCache: () => ({
      put: (k, v) => cache.set(k, v),
      get: (k) => cache.get(k),
      remove: (k) => cache.delete(k),
    }),
  };
  c._getSpreadsheet_ = () => ({
    getSheetByName: () => ({
      getDataRange: () => ({
        getDisplayValues: () => [
          [
            "id",
            "nombre_persona",
            "nombre_usuario",
            "contrasenia",
            "rol",
            "estado",
          ],
          [
            "a",
            "Nombre correcto",
            "demo",
            "solo-ficticia",
            "Integrante",
            status,
          ],
        ],
      }),
    }),
  });
  c.env_ = () => ({ SH_REGISTRO_USUARIOS: "USUARIOS" });
  assert.equal(
    c.login("demo", "solo-ficticia").user.nombreCompleto,
    "Nombre correcto",
  );
  status = "";
  assert.equal(c.login("demo", "solo-ficticia").ok, false);
  status = "Inactivo";
  assert.equal(c.login("demo", "solo-ficticia").ok, false);
});
