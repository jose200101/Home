/** Abonos entre integrantes para gastos compartidos y préstamos personales. */
const HOME_REIMBURSEMENT_SCHEMA_ = 4;

function homeUsesParticipantPayments_(item) {
  return !!(item && (item.expenseKind || item.loanId));
}

function homeIsReimbursementType_(type) {
  return type === "REIMBURSEMENT_DECLARE" || type === "REIMBURSEMENT_DECIDE";
}

function homeReimbursementAccess_(state, session, personId) {
  // El retiro conserva la liquidación de obligaciones existentes.
  const member = homeAgreementMember_(state, session, false);
  if (!member || member.id !== personId)
    throw new Error("Operación de reembolso no disponible para esta cuenta.");
}

function homeReimbursementCommand_(request, state, session) {
  const p = request.payload || {};
  if (request.type === "REIMBURSEMENT_DECLARE") {
    const ob = state.obligations.find(
      (item) =>
        item.id === p.obligationId &&
        homeUsesParticipantPayments_(item) &&
        item.creditorId,
    );
    homeReimbursementAccess_(state, session, ob && ob.debtorId);
    const cents = homeCents_(p.amount);
    if (cents > ob.available)
      throw new Error("El abono supera el saldo disponible.");
    const date = homeDate_(p.date);
    const earliest = ob.loanId
      ? ob.date
      : state.charges.find((item) => item.id === ob.sourceId).paidDate;
    if (date < earliest || date > homeToday_())
      throw new Error(
        ob.loanId
          ? "La fecha debe estar entre la fecha del préstamo y hoy."
          : "La fecha debe estar entre el pago al proveedor y hoy.",
      );
    const note =
      p.note == null || String(p.note).trim() === ""
        ? ""
        : homeText_(p.note, "Nota", 500);
    return {
      obligationId: ob.id,
      debtorId: ob.debtorId,
      creditorId: ob.creditorId,
      ...(ob.loanId ? { loanId: ob.loanId } : { expenseKind: ob.expenseKind }),
      cents,
      date,
      note,
    };
  }

  const payment = state.payments.find(
    (item) =>
      item.id === p.paymentId && item.schema === HOME_REIMBURSEMENT_SCHEMA_,
  );
  homeReimbursementAccess_(state, session, payment && payment.creditorId);
  if (payment.status !== "pending")
    throw new Error("El abono ya fue atendido.");
  if (p.decision !== "confirm" && p.decision !== "reject")
    throw new Error("Decisión inválida.");
  const reason =
    p.decision === "reject" || p.reason
      ? homeText_(p.reason, "Motivo", 500)
      : "";
  return {
    paymentId: payment.id,
    creditorId: payment.creditorId,
    decision: p.decision,
    reason,
  };
}

function homeReimbursementReplay_(obligations, payments, event, data) {
  if (event.type === "REIMBURSEMENT_DECLARE") {
    const ob = obligations[data.obligationId];
    if (
      !ob ||
      !homeUsesParticipantPayments_(ob) ||
      !ob.creditorId ||
      ob.debtorId === ob.creditorId ||
      ob.debtorId !== data.debtorId ||
      ob.creditorId !== data.creditorId ||
      ob.expenseKind !== data.expenseKind ||
      ob.loanId !== data.loanId ||
      payments[event.id] ||
      !Number.isSafeInteger(data.cents) ||
      data.cents <= 0 ||
      data.cents > ob.cents - ob.paid - ob.reserved
    )
      throw new Error("Historial de abonos inválido.");
    ob.reserved += data.cents;
    payments[event.id] = {
      ...data,
      id: event.id,
      declaredAt: event.at,
      status: "pending",
    };
    return;
  }

  const payment = payments[data.paymentId];
  if (
    !payment ||
    payment.schema !== HOME_REIMBURSEMENT_SCHEMA_ ||
    payment.status !== "pending" ||
    payment.creditorId !== data.creditorId ||
    (data.decision !== "confirm" && data.decision !== "reject") ||
    (data.decision === "reject" && !data.reason)
  )
    throw new Error("Historial de decisiones de abonos inválido.");
  const ob = obligations[payment.obligationId];
  ob.reserved -= payment.cents;
  if (data.decision === "confirm") ob.paid += payment.cents;
  payment.status = data.decision === "confirm" ? "confirmed" : "rejected";
  payment.decidedAt = event.at;
  payment.reason = data.reason;
}
