/** HOME v1: acuerdos de participación. Dominio puro, sin operaciones de Sheets ni RPC. */
function homeAgreementId_(value) {
  const id = typeof value === "string" ? value.trim() : "";
  if (!/^[a-zA-Z0-9_-]{1,100}$/.test(id)) throw new Error("Identificador inválido.");
  return id;
}

function homeAgreementCents_(value) {
  const text = String(value == null ? "" : value).trim();
  if (!/^\d+(\.\d{1,2})?$/.test(text)) throw new Error("Importe inválido: máximo dos decimales.");
  const p = text.split(".");
  const cents = Number(p[0]) * 100 + Number((p[1] || "").padEnd(2, "0"));
  if (!Number.isSafeInteger(cents) || cents < 1 || cents > 10000000000)
    throw new Error("Importe fuera de rango.");
  return cents;
}

function homeAgreementBasis_(value) {
  const text = String(value == null ? "" : value).trim();
  if (!/^\d{1,3}(\.\d{1,2})?$/.test(text))
    throw new Error("Porcentaje inválido: máximo dos decimales.");
  const p = text.split(".");
  const basis = Number(p[0]) * 100 + Number((p[1] || "").padEnd(2, "0"));
  if (basis < 1 || basis > 10000) throw new Error("Porcentaje debe estar entre 0.01 y 100.");
  return basis;
}

function homeAgreementActor_(actor) {
  if (!actor || actor.active !== true) throw new Error("Se requiere una cuenta activa.");
  homeAgreementId_(actor.id);
  return actor;
}

function homeAgreementProposal_(proposal, actor) {
  homeAgreementActor_(actor);
  if (!proposal || typeof proposal !== "object" || Array.isArray(proposal))
    throw new Error("Falta la propuesta.");
  const resourceId = homeAgreementId_(proposal.resourceId);
  const kind = proposal.kind;
  if (kind !== "recurring" && kind !== "occasional") throw new Error("Tipo de gasto inválido.");
  if (kind === "recurring" && actor.admin !== true)
    throw new Error("Un acuerdo recurrente requiere administrador.");
  const mode = proposal.mode;
  if (!["equal", "percent", "amount"].includes(mode)) throw new Error("Distribución inválida.");
  const entries = proposal.participants;
  if (!Array.isArray(entries) || !entries.length || entries.length > 100)
    throw new Error("Se requieren entre 1 y 100 participantes.");
  const seen = new Set();
  const participants = entries.map((entry) => {
    const id = homeAgreementId_(entry && entry.personId);
    if (seen.has(id)) throw new Error("Participante repetido.");
    seen.add(id);
    const p = { personId: id, decision: "pending" };
    if (mode === "percent") p.basis = homeAgreementBasis_(entry.value);
    else if (mode === "amount") p.cents = homeAgreementCents_(entry.value);
    else if (entry.value !== undefined && entry.value !== null && String(entry.value).trim() !== "")
      throw new Error("Partes iguales no usan importes por integrante.");
    return p;
  });
  const out = { resourceId, kind, mode, participants, createdBy: actor.id, status: "pending" };
  if (kind === "occasional") {
    if (!actor.personId || proposal.payerId !== actor.personId)
      throw new Error("Solo quien pagó puede registrar su compra.");
    homeAgreementId_(actor.personId);
    out.payerId = actor.personId;
  } else if (proposal.payerId) throw new Error("El pagador recurrente se elige al pagar la factura.");
  if (kind === "occasional" || mode === "amount") {
    if (!Number.isSafeInteger(proposal.amountCents) ||
        proposal.amountCents < 1 || proposal.amountCents > 10000000000)
      throw new Error("Se requiere importe total en centavos enteros.");
    out.amountCents = proposal.amountCents;
  } else if (proposal.amountCents != null)
    throw new Error("El acuerdo recurrente variable no fija importe mensual.");
  if (mode === "percent" && participants.reduce((s, p) => s + p.basis, 0) !== 10000)
    throw new Error("Los porcentajes deben sumar 100 %.");
  if (mode === "amount" && participants.reduce((s, p) => s + p.cents, 0) !== out.amountCents)
    throw new Error("Los aportes deben sumar exactamente el importe.");
  return out;
}

/** Comandos: PROPOSE {proposal}; RESPOND {version,decision}; CANCEL {version}. */
function homeAgreementApply_(state, command, actor) {
  homeAgreementActor_(actor);
  if (!command || typeof command !== "object") throw new Error("Comando inválido.");
  const next = state ? JSON.parse(JSON.stringify(state)) : null;
  if (command.type === "PROPOSE") {
    const proposal = homeAgreementProposal_(command.proposal, actor);
    if (next && (next.resourceId !== proposal.resourceId || next.kind !== proposal.kind))
      throw new Error("La propuesta pertenece a otro gasto.");
    if (next && next.pending && next.pending.status === "pending")
      throw new Error("Ya hay una propuesta pendiente.");
    if (next && next.kind === "occasional" && next.active)
      throw new Error("Un gasto aprobado requiere una corrección formal.");
    if (next && next.kind === "occasional" && next.pending &&
        next.pending.createdBy !== actor.id && actor.admin !== true)
      throw new Error("Solo quien registró la compra puede reenviar.");
    const result = next || { resourceId: proposal.resourceId, kind: proposal.kind,
      active: null, pending: null, history: [] };
    let lastVersion = 0;
    [result.active, result.pending].concat(result.history).forEach((v) => {
      if (v) lastVersion = Math.max(lastVersion, v.version);
    });
    if (result.pending) result.history.push(Object.assign({}, result.pending, { status: "replaced" }));
    proposal.version = lastVersion + 1;
    result.pending = proposal;
    return result;
  }
  if (!next || !next.pending) throw new Error("No hay propuesta pendiente.");
  if (!Number.isSafeInteger(command.version) || command.version !== next.pending.version)
    throw new Error("La versión no está vigente.");
  if (command.type === "RESPOND") {
    if (next.pending.status !== "pending") throw new Error("La propuesta fue rechazada.");
    if (!["accept", "reject"].includes(command.decision)) throw new Error("Decisión inválida.");
    const person = next.pending.participants.find((p) => p.personId === actor.personId);
    if (!person) throw new Error("Solo el participante involucrado puede responder.");
    if (person.decision !== "pending") throw new Error("El integrante ya respondió.");
    person.decision = command.decision === "accept" ? "accepted" : "rejected";
    if (person.decision === "rejected") next.pending.status = "rejected";
    else if (next.pending.participants.every((p) => p.decision === "accepted")) {
      if (next.active) next.history.push(Object.assign({}, next.active, { status: "superseded" }));
      next.active = Object.assign({}, next.pending, { status: "approved" });
      next.pending = null;
    }
    return next;
  }
  if (command.type === "CANCEL") {
    if (next.kind === "recurring" && actor.admin !== true)
      throw new Error("Solo el administrador puede cancelar acuerdos recurrentes.");
    if (next.kind === "occasional" && actor.id !== next.pending.createdBy && actor.admin !== true)
      throw new Error("No puedes cancelar el gasto de otra persona.");
    next.history.push(Object.assign({}, next.pending, { status: "cancelled" }));
    next.pending = null;
    return next;
  }
  throw new Error("Comando desconocido.");
}

function homeAgreementCanInvoice_(state) {
  return !!(state && state.kind === "recurring" && state.active && !state.pending);
}

function homeAgreementCanActivate_(state) {
  return !!(state && state.active && !state.pending);
}
