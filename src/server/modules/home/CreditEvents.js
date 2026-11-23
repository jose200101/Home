/** Reservas y disposición bilateral de saldos a favor. */
function homeIsCreditType_(type) {
  return [
    "CREDIT_REFUND_DECLARE",
    "CREDIT_APPLY_PROPOSE",
    "CREDIT_DECIDE",
  ].includes(type);
}

function homeCreditBalance_(credit) {
  return credit.cents - credit.returned - credit.applied;
}

function homeCreditMatchesObligation_(credit, ob) {
  return !!ob.expenseKind && !ob.loanId &&
    ob.creditorId === credit.holderId && ob.debtorId === credit.beneficiaryId;
}

function homeCreditsVisible_(state, result, session) {
  if (!result.credits.length && !result.creditOperations.length) return;
  const member = homeAgreementMember_(state, session, false);
  const paid = new Set(state.charges.filter((charge) => charge.providerStatus === "paid").map((charge) => charge.id));
  result.credits = result.credits.map((credit) => {
    const party = member && [credit.holderId, credit.beneficiaryId].includes(member.id);
    const targets = party && credit.available > 0 ? result.obligations
      .filter((ob) => homeCreditMatchesObligation_(credit, ob) &&
        paid.has(ob.sourceId) && !ob.pendingCorrectionId && ob.available > 0)
      .map((ob) => ({ obligationId: ob.id, limit: Math.min(ob.available, credit.available) })) : [];
    return { ...credit, canRefund: !!party && member.id === credit.holderId && credit.available > 0, targets };
  });
  result.creditOperations = result.creditOperations.map((operation) => ({
    ...operation,
    canDecide: !!member && operation.deciderId === member.id && operation.status === "pending",
  }));
}

function homeCreditCapacity_(credit, cents, ob) {
  if (
    !Number.isSafeInteger(cents) ||
    cents <= 0 ||
    cents > homeCreditBalance_(credit) - credit.reserved
  )
    throw new Error("El importe supera el crédito disponible.");
  if (!ob) return;
  if (!homeCreditMatchesObligation_(credit, ob))
    throw new Error(
      "Solo procede para gastos compartidos pagados entre las mismas dos personas.",
    );
  if (ob.pendingCorrectionId)
    throw new Error("La obligación tiene una corrección pendiente.");
  if (cents > homeObligationAvailable_(ob))
    throw new Error("El importe supera el saldo disponible de la obligación.");
}

function homeCreditCommand_(request, state, session) {
  const p = request.payload || {};
  if (request.type === "CREDIT_DECIDE") {
    const operation = state.creditOperations.find(
      (o) => o.id === p.operationId,
    );
    const personId = homeAdjustmentParty_(
      state,
      session,
      operation ? [operation.deciderId] : [],
    );
    if (operation.status !== "pending")
      throw new Error("La operación ya fue atendida.");
    return {
      operationId: operation.id,
      personId,
      ...homeAdjustmentDecision_(p),
    };
  }

  const credit = state.credits.find((c) => c.id === p.creditId);
  const refund = request.type === "CREDIT_REFUND_DECLARE";
  const parties = credit
    ? refund
      ? [credit.holderId]
      : [credit.holderId, credit.beneficiaryId]
    : [];
  const personId = homeAdjustmentParty_(state, session, parties);
  const cents = homeCents_(p.amount);
  let ob;
  if (!refund) {
    ob = state.obligations.find((o) => o.id === p.obligationId);
    const charge = ob && state.charges.find((c) => c.id === ob.sourceId);
    if (!charge || charge.providerStatus !== "paid")
      throw new Error(
        "Solo procede para gastos compartidos pagados entre las mismas dos personas.",
      );
  }
  homeCreditCapacity_(credit, cents, ob);
  return {
    creditId: credit.id,
    holderId: credit.holderId,
    beneficiaryId: credit.beneficiaryId,
    ...(ob ? { obligationId: ob.id } : {}),
    cents,
    personId,
    deciderId:
      personId === credit.holderId ? credit.beneficiaryId : credit.holderId,
    note: p.note ? homeText_(p.note, "Nota", 500) : "",
  };
}

function homeCreditReplay_(ledger, event, data) {
  const { obligations, credits, creditOperations } = ledger;
  if (event.type !== "CREDIT_DECIDE") {
    const credit = credits[data.creditId];
    const refund = event.type === "CREDIT_REFUND_DECLARE";
    const ob = refund ? null : obligations[data.obligationId];
    if (
      !credit ||
      creditOperations[event.id] ||
      (!refund && !ob) ||
      data.holderId !== credit.holderId ||
      data.beneficiaryId !== credit.beneficiaryId ||
      ![credit.holderId, credit.beneficiaryId].includes(data.personId) ||
      (refund && data.personId !== credit.holderId) ||
      data.deciderId !==
        (data.personId === credit.holderId
          ? credit.beneficiaryId
          : credit.holderId)
    )
      throw new Error("Historial de operaciones de crédito inválido.");
    homeCreditCapacity_(credit, data.cents, ob);
    credit.reserved += data.cents;
    if (ob) ob.creditReserved = (ob.creditReserved || 0) + data.cents;
    creditOperations[event.id] = {
      ...data,
      id: event.id,
      kind: refund ? "refund" : "application",
      status: "pending",
      declaredAt: event.at,
    };
    return;
  }
  const op = creditOperations[data.operationId];
  if (
    !op ||
    op.status !== "pending" ||
    op.deciderId !== data.personId ||
    !["accept", "reject"].includes(data.decision) ||
    (data.decision === "reject" && !data.reason)
  )
    throw new Error("Historial de decisiones de crédito inválido.");
  const credit = credits[op.creditId];
  const ob = op.kind === "application" ? obligations[op.obligationId] : null;
  credit.reserved -= op.cents;
  if (ob) ob.creditReserved -= op.cents;
  if (data.decision === "accept") {
    if (ob) {
      credit.applied += op.cents;
      ob.appliedCreditCents = (ob.appliedCreditCents || 0) + op.cents;
    } else credit.returned += op.cents;
  }
  op.status = data.decision === "accept" ? "confirmed" : "rejected";
  op.decidedAt = event.at;
  op.reason = data.reason;
}

function homeCreditTotals_(credits) {
  return credits.reduce(
    (total, credit) => ({
      pending: total.pending + credit.pending,
      reserved: total.reserved + credit.reserved,
      available: total.available + credit.available,
    }),
    { pending: 0, reserved: 0, available: 0 },
  );
}
