/* Servicios de Apps Script en memoria: no llama a Google ni contiene datos reales. */
const vm = require("node:vm");
const fs = require("node:fs");
const path = require("node:path");
function createHarness() {
  let sequence = 0,
    locked = false;
  const sheets = new Map();
  const users = [
    {
      id: "admin",
      nombreCompleto: "Administración",
      nombre_usuario: "admin",
      rol: "Administrador",
      estado: "Activo",
    },
    {
      id: "ana",
      nombreCompleto: "Ana",
      nombre_usuario: "ana",
      rol: "Integrante",
      estado: "Activo",
    },
    {
      id: "luis",
      nombreCompleto: "Luis",
      nombre_usuario: "luis",
      rol: "Integrante",
      estado: "Activo",
    },
    {
      id: "maria",
      nombreCompleto: "María",
      nombre_usuario: "maria",
      rol: "Integrante",
      estado: "Activo",
    },
  ];
  function sheet(name) {
    const rows = [];
    return {
      name,
      rows,
      getName: () => name,
      getLastRow: () => rows.length,
      getLastColumn: () => rows[0]?.length || 0,
      getDataRange: () => ({ getValues: () => structuredClone(rows) }),
      getRange: (r, c, n = 1, m = 1) => ({
        getDisplayValues: () =>
          rows
            .slice(r - 1, r - 1 + n)
            .map((row) => row.slice(c - 1, c - 1 + m).map(String)),
      }),
      appendRow: (row) => rows.push(structuredClone(row)),
      setFrozenRows: () => {},
    };
  }
  const db = {
    getSheetByName: (name) => sheets.get(name),
    getSheets: () => [...sheets.values()],
    insertSheet: (name) => {
      if (sheets.has(name)) throw Error("Hoja duplicada");
      const sh = sheet(name);
      sheets.set(name, sh);
      return sh;
    },
  };
  const context = vm.createContext({
    console,
    Set,
    Date,
    JSON,
    Math,
    Number,
    String,
    Object,
    Array,
    validateToken: (token) => {
      const user = users.find((u) => "token-" + u.id === token);
      return user ? { ok: true, user: { ...user } } : { ok: false };
    },
    _readUsers_: () => users,
    conexion_: () => db,
    Utilities: {
      getUuid: () => `generated-${++sequence}`,
      formatDate: () => "2026-11-16",
    },
    LockService: {
      getScriptLock: () => ({
        waitLock: () => {
          if (locked) throw Error("Bloqueo concurrente");
          locked = true;
        },
        releaseLock: () => {
          locked = false;
        },
      }),
    },
    SpreadsheetApp: { flush: () => {} },
  });
  for (const name of ["domain/AgreementRules.js", "AgreementEvents.js", "ExpenseEvents.js", "ReimbursementEvents.js", "LoanEvents.js", "CorrectionEvents.js", "CreditEvents.js", "HomeLedger.js", "HomeAccess.js"])
    vm.runInContext(
      fs.readFileSync(path.join(__dirname, "../dist/Backend", name), "utf8"),
      context,
      { filename: name },
    );
  function execute(type, payload = {}, user = "admin", id) {
    return context.homeExecute("token-" + user, {
      id: id || "request-" + String(++sequence).padStart(12, "0"),
      type,
      payload,
    });
  }
  function seed() {
    execute("INIT", { name: "Casa compartida" });
    ["ana", "luis", "maria"].forEach((id, i) =>
      execute("PERSON", { name: ["Ana", "Luis", "María"][i], userId: id }),
    );
    return context.homeGetState("token-admin").people;
  }
  return {
    context,
    execute,
    seed,
    users,
    db,
    sheets,
    get locked() {
      return locked;
    },
  };
}
module.exports = { createHarness };
