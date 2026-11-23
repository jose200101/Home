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

function homeReimbursementEarliest_(obligation, charges) {
  return obligation.loanId
    ? obligation.date
    : charges.find((charge) => charge.id === obligation.sourceId)?.paidDate;
}

function homeReimbursementsVisible_(state, result, session) {
  if (!result.obligations.some(homeUsesParticipantPayments_) &&
      !result.payments.some((payment) => payment.schema === HOME_REIMBURSEMENT_SCHEMA_)) return;
  const member = homeAgreementMember_(state, session, false);
  result.obligations = result.obligations.map((obligation) => {
    if (!homeUsesParticipantPayments_(obligation) || !obligation.creditorId) return obligation;
    const earliestDate = homeReimbursementEarliest_(obligation, state.charges) || "";
    return {
      ...obligation,
      reimbursement: {
        earliestDate,
        canDeclare: !!member && member.id === obligation.debtorId &&
          !obligation.pendingCorrectionId && obligation.available > 0 && !!earliestDate &&
          earliestDate <= result.today,
      },
    };
  });
  result.payments = result.payments.map((payment) => payment.schema === HOME_REIMBURSEMENT_SCHEMA_
    ? { ...payment, canDecideReimbursement: !!member && member.id === payment.creditorId && payment.status === "pending" }
    : payment);
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
    if (ob.pendingCorrectionId) throw new Error("La obligación tiene una corrección pendiente.");
    const cents = homeCents_(p.amount);
    if (cents > ob.available)
      throw new Error("El abono supera el saldo disponible.");
    const date = homeDate_(p.date);
    const earliest = homeReimbursementEarliest_(ob, state.charges);
    if (!earliest || date < earliest || date > homeToday_())
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
      ob.pendingCorrectionId ||
      data.cents > homeObligationAvailable_(ob)
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
