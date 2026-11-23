const { test } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const { createHarness } = require("../tools/harness.cjs");

const html = fs.readFileSync(path.join(__dirname, "../dist/Frontend/views/HomeDashboard.html"), "utf8");
const model = vm.runInNewContext(html.match(/<script>([\s\S]*?)<\/script>/)[1] + "\nhomeDashboardModel");
const plain = (value) => JSON.parse(JSON.stringify(value));
const people = ["ana", "luis", "maria"].map((id) => ({ id, active: true, userId: id }));

function obligation(id, debtorId, creditorId, pending, due = "", period = "2026-11") {
  return { id, debtorId, creditorId, pending, due, period };
}

const base = {
  admin: true, me: "", people, today: "2026-11-23",
  obligations: [
    obligation("shared", "ana", "luis", 8000, "2026-11-20"),
    obligation("provider", "ana", "", 7000, "2026-11-23"),
    obligation("loan", "luis", "ana", 10000),
    obligation("other", "luis", "maria", 20000, "2026-11-28"),
    obligation("settled", "luis", "ana", 0, "2026-11-01"),
    obligation("old", "ana", "luis", 2500, "", "2026-10"),
  ],
};

test("totales por integrante distinguen cobros, deudas y proveedor sin mezclar créditos", () => {
  const result = model(base, "ana", "2026-11");
  assert.equal(result.memberDue, 8000);
  assert.equal(result.providerDue, 7000);
  assert.equal(result.receivable, 10000);
  assert.deepEqual(plain(result.upcoming.map((item) => [item.id, item.overdue])), [
    ["shared", true], ["provider", false], ["loan", false],
  ]);
});

test("la vista general registra cada obligación una vez y conserva su alcance", () => {
  const result = model(base, "", "2026-11");
  assert.equal(result.household, true);
  assert.equal(result.memberDue, 38000);
  assert.equal(result.providerDue, 7000);
  assert.equal(result.receivable, 0);
  assert.deepEqual(plain(result.credits), { pending: 0, reserved: 0, available: 0 });
  assert.equal(result.upcoming[3].due, "");
  assert.equal(result.upcoming[3].overdue, false);
});

test("los créditos sin gastar se atribuyen al beneficiario y se separan de la reserva", () => {
  const state = {
    ...base,
    credits: [
      { beneficiaryId: "ana", holderId: "luis", pending: 10000, reserved: 6000, available: 4000 },
      { beneficiaryId: "ana", holderId: "maria", pending: 2000, reserved: 2000, available: 0 },
      { beneficiaryId: "luis", holderId: "ana", pending: 5000, reserved: 0, available: 5000 },
    ],
  };
  const result = model(state, "ana", "2026-10");
  assert.deepEqual(plain(result.credits), { pending: 12000, reserved: 8000, available: 4000 });
  assert.equal(result.memberDue, 2500);
});

test("atención cuenta operaciones por resolver sin confundir pagos enviados ni votos emitidos", () => {
  const state = {
    ...base,
    payments: [
      { id: "receive", creditorId: "ana", debtorId: "luis", status: "pending", cents: 3500, date: "2026-11-20" },
      { id: "send", debtorId: "ana", creditorId: "luis", status: "pending" },
      { id: "done", debtorId: "luis", creditorId: "ana", status: "confirmed" },
    ],
    loans: [{ debtorId: "ana", status: "pending" }, { creditorId: "ana", debtorId: "luis", status: "pending" }],
    agreements: [{ pending: { status: "pending", participants: [
      { personId: "ana", decision: "pending" }, { personId: "maria", decision: "pending" },
    ] } }],
    corrections: [{ debtorId: "ana", creditorId: "luis", status: "pending", decisions: [{ personId: "ana" }] }],
    creditOperations: [{ beneficiaryId: "ana", deciderId: "maria", status: "pending" }],
  };
  const result = model(state, "ana", "2026-11");
  assert.deepEqual(plain(result.counts), { payments: 1, loans: 1, agreements: 1, corrections: 0, credits: 0 });
  assert.equal(result.attention, 3);
  assert.equal(result.receivable, 10000);
  assert.equal(model(state, "", "").counts.agreements, 1);
});

test("sin vínculo no se exponen datos y una cuenta vacía conserva sus estados cero", () => {
  const state = { ...base, admin: false, me: "" };
  const unlinked = model(state, "ana", "");
  assert.equal(unlinked.scoped, false);
  assert.equal(unlinked.memberDue, 0);
  assert.equal(unlinked.activity.length + unlinked.upcoming.length + unlinked.attention, 0);
  const empty = model({ ...base, obligations: [], audit: [] }, "", "");
  assert.equal(empty.scoped, true);
  assert.equal(empty.memberDue + empty.providerDue + empty.attention, 0);
  assert.equal(empty.upcoming.length + empty.activity.length, 0);
});

test("la actividad por integrante respeta la proyección real y excluye auditoría global", () => {
  const h = createHarness();
  const [ana, luis, maria] = h.seed();
  const addDebt = (debtorId, creditorId, name) => h.execute("DEBT", {
    debtorId, creditorId, name, amount: "100", date: "2026-11-02", due: "2026-11-14",
  }).state.obligations.find((item) => item.name === name);
  const mine = addDebt(ana.id, luis.id, "Propia");
  const other = addDebt(luis.id, maria.id, "Ajena");
  h.execute("PAYMENT", { obligationId: mine.id, amount: "10", date: "2026-11-04" }, "ana");
  h.execute("PAYMENT", { obligationId: other.id, amount: "20", date: "2026-11-05" }, "luis");
  const member = model(h.context.homeGetState("token-ana"), maria.id, "");
  const admin = model(h.context.homeGetState("token-admin"), ana.id, "");
  for (const result of [member, admin]) {
    assert.equal(result.personId, ana.id);
    assert.deepEqual(plain(result.upcoming.map((item) => item.name)), ["Propia"]);
    assert.deepEqual(plain(result.activity.map((item) => [item.label, item.amount])), [["Pago", 1000]]);
    assert.equal(result.attention, 0);
  }
});
