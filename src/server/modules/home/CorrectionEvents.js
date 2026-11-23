/** Correcciones bilaterales en HOME_EVENTOS, sin alterar el principal histórico. */
const HOME_ADJUSTMENT_SCHEMA_ = 6;

function homeIsCorrectionType_(type) {
  return type === "CORRECTION_PROPOSE" || type === "CORRECTION_DECIDE";
}

function homeObligationBalance_(ob) {
  return Math.max(
    0,
    (ob.effectiveCents ?? ob.cents) -
      ob.paid +
      (ob.creditedCents || 0) -
      (ob.appliedCreditCents || 0),
  );
}

function homeObligationAvailable_(ob) {
  return homeObligationBalance_(ob) - ob.reserved - (ob.creditReserved || 0);
}

function homeAdjustmentParty_(state, session, allowed) {
  const member = homeAgreementMember_(state, session, false);
  if (!member || !allowed.includes(member.id))
    throw new Error("Operación no disponible para esta cuenta.");
  return member.id;
}

function homeAdjustmentDecision_(payload) {
  if (!["accept", "reject"].includes(payload.decision))
    throw new Error("Decisión inválida.");
  return {
    decision: payload.decision,
    reason:
      payload.decision === "reject" || payload.reason
        ? homeText_(payload.reason, "Motivo", 500)
        : "",
  };
}

function homeCheckCorrectionCapacity_(ob, amount) {
  if (ob.reserved || ob.creditReserved)
    throw new Error("Resuelve las reservas pendientes antes de corregir.");
  if (amount < (ob.appliedCreditCents || 0))
    throw new Error(
      "El importe corregido no puede ser inferior al crédito ya aplicado.",
    );
}

function homeCorrectionObligation_(ob) {
  return homeUsesParticipantPayments_(ob) && !!ob.creditorId && ob.debtorId !== ob.creditorId;
}

function homeCorrectionsVisible_(state, result, session) {
  if (session.admin) {
    result.obligations = result.obligations.map((ob) => {
      if (!homeCorrectionObligation_(ob)) return ob;
      const minimum = ob.appliedCreditCents || 0;
      const blocked = ob.pendingCorrectionId ? "Ya existe una corrección pendiente."
        : ob.reserved || ob.creditReserved ? "Resuelve las reservas pendientes antes de corregir."
        : "";
      return { ...ob, correctionOffer: { minimum, blocked } };
    });
  }
  if (!result.corrections.length) return;
  const member = homeAgreementMember_(state, session, false);
  result.corrections = result.corrections.map((correction) => ({
    ...correction,
    canDecide: !!member && correction.status === "pending" &&
      [correction.debtorId, correction.creditorId].includes(member.id) &&
      !correction.decisions.some((entry) => entry.personId === member.id),
  }));
}

function homeCorrectionCommand_(request, state, session) {
  const p = request.payload || {};
  if (request.type === "CORRECTION_PROPOSE") {
    if (!session.admin)
      throw new Error("Esta operación requiere administrador.");
    const ob = state.obligations.find(
      (o) => o.id === p.obligationId && homeCorrectionObligation_(o),
    );
    if (!ob) throw new Error("Obligación no disponible para corrección.");
    if (ob.pendingCorrectionId)
      throw new Error("Ya existe una corrección pendiente.");
    if (p.expectedVersion !== (ob.correctionVersion || 0))
      throw new Error("Versión de corrección desactualizada.");
    const proposedCents = homeCents_(p.amount, true);
    const previousCents = ob.effectiveCents ?? ob.cents;
    if (proposedCents === previousCents)
      throw new Error("El importe propuesto debe cambiar.");
    homeCheckCorrectionCapacity_(ob, proposedCents);
    return {
      obligationId: ob.id,
      debtorId: ob.debtorId,
      creditorId: ob.creditorId,
      version: p.expectedVersion + 1,
      originalCents: ob.cents,
      previousCents,
      proposedCents,
      reason: homeText_(p.reason, "Motivo", 500),
    };
  }
  const correction = state.corrections.find((c) => c.id === p.correctionId);
  const personId = homeAdjustmentParty_(
    state,
    session,
    correction ? [correction.debtorId, correction.creditorId] : [],
  );
  if (correction.status !== "pending" || p.version !== correction.version)
    throw new Error("La corrección o su versión ya no está pendiente.");
  if (correction.decisions.some((d) => d.personId === personId))
    throw new Error("Ya registraste tu decisión.");
  return {
    correctionId: correction.id,
    version: p.version,
    personId,
    ...homeAdjustmentDecision_(p),
  };
}

function homeCorrectionReplay_(ledger, event, data) {
  const { obligations, corrections, credits } = ledger;
  if (event.type === "CORRECTION_PROPOSE") {
    const ob = obligations[data.obligationId];
    if (
      !homeUsesParticipantPayments_(ob) ||
      ob.pendingCorrectionId ||
      corrections[event.id] ||
      !ob.creditorId ||
      ob.creditorId === ob.debtorId ||
      data.creditorId !== ob.creditorId ||
      data.debtorId !== ob.debtorId ||
      data.originalCents !== ob.cents ||
      data.previousCents !== (ob.effectiveCents ?? ob.cents) ||
      data.version !== (ob.correctionVersion || 0) + 1 ||
      !Number.isSafeInteger(data.proposedCents) ||
      data.proposedCents < 0 ||
      data.proposedCents > 10000000000 ||
      data.proposedCents === data.previousCents
    )
      throw new Error("Historial de correcciones inválido.");
    homeCheckCorrectionCapacity_(ob, data.proposedCents);
    corrections[event.id] = {
      ...data,
      id: event.id,
      date: event.at,
      status: "pending",
      decisions: [],
    };
    ob.correctionVersion = data.version;
    ob.pendingCorrectionId = event.id;
    return;
  }
  const correction = corrections[data.correctionId];
  if (
    !correction ||
    correction.status !== "pending" ||
    data.version !== correction.version ||
    ![correction.debtorId, correction.creditorId].includes(data.personId) ||
    correction.decisions.some((d) => d.personId === data.personId) ||
    !["accept", "reject"].includes(data.decision) ||
    (data.decision === "reject" && !data.reason)
  )
    throw new Error("Historial de decisiones de corrección inválido.");
  correction.decisions.push({
    personId: data.personId,
    decision: data.decision,
    reason: data.reason,
    at: event.at,
  });
  const ob = obligations[correction.obligationId];
  if (data.decision === "accept" && correction.decisions.length === 1) return;
  delete ob.pendingCorrectionId;
  correction.decidedAt = event.at;
  correction.status = data.decision === "reject" ? "rejected" : "accepted";
  if (correction.status === "rejected") return;
  homeCheckCorrectionCapacity_(ob, correction.proposedCents);
  ob.effectiveCents = correction.proposedCents;
  // Solo pagos en efectivo aún no reconocidos como crédito pueden originar un nuevo excedente.
  const excess = Math.max(
    0,
    ob.paid -
      (ob.creditedCents || 0) +
      (ob.appliedCreditCents || 0) -
      ob.effectiveCents,
  );
  if (!excess) return;
  ob.creditedCents = (ob.creditedCents || 0) + excess;
  correction.creditId = correction.id;
  credits[correction.id] = {
    id: correction.id,
    correctionId: correction.id,
    obligationId: ob.id,
    beneficiaryId: ob.debtorId,
    holderId: ob.creditorId,
    date: event.at,
    cents: excess,
    returned: 0,
    applied: 0,
    reserved: 0,
  };
}
