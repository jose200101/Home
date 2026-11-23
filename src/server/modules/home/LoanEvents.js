/** Solicitudes de préstamos personales y aceptación explícita del deudor. */
const HOME_LOAN_SCHEMA_ = 5;

function homeIsLoanType_(type) {
  return type === "LOAN_REQUEST" || type === "LOAN_DECIDE";
}

function homeLoanAccess_(state, session, personId, activeOnly) {
  const person = homeAgreementMember_(state, session, activeOnly);
  if (!person || person.id !== personId)
    throw new Error("Préstamo no disponible para esta cuenta.");
}

function homeLoanCommand_(request, state, session) {
  if (!state.initialized) throw new Error("Primero configura el hogar.");
  const p = request.payload || {};
  if (request.type === "LOAN_REQUEST") {
    const creditor = homeAgreementMember_(state, session, true);
    if (!creditor)
      throw new Error("Se requiere un integrante activo vinculado.");
    const debtorId = homePerson_(state, p.debtorId, true).id;
    if (creditor.id === debtorId)
      throw new Error("Selecciona otro integrante como deudor.");
    const date = homeDate_(p.date);
    const due = p.due == null || p.due === "" ? "" : homeDate_(p.due);
    if (date > homeToday_() || (due && due < date))
      throw new Error("Revisa las fechas del préstamo.");
    return {
      creditorId: creditor.id,
      debtorId,
      date,
      due,
      cents: homeCents_(p.amount),
      name: homeText_(p.name, "Descripción", 500),
    };
  }

  const loan = state.loans.find((item) => item.id === p.loanId);
  const cancel = p.decision === "cancel";
  const accept = p.decision === "accept";
  const personId = loan && (cancel ? loan.creditorId : loan.debtorId);
  homeLoanAccess_(state, session, personId, accept);
  if (loan.status !== "pending")
    throw new Error("El préstamo ya fue atendido.");
  if (!accept && !cancel && p.decision !== "reject")
    throw new Error("Decisión de préstamo inválida.");
  if (accept) homePerson_(state, loan.creditorId, true);
  const reason = p.reason ? homeText_(p.reason, "Motivo", 500) : "";
  return { loanId: loan.id, personId, decision: p.decision, reason };
}

function homeLoanReplay_(loans, obligations, event, data) {
  if (event.type === "LOAN_REQUEST") {
    if (
      loans[event.id] ||
      !data.debtorId ||
      !data.creditorId ||
      data.debtorId === data.creditorId ||
      !Number.isSafeInteger(data.cents) ||
      data.cents <= 0 ||
      data.cents > 10000000000
    )
      throw new Error("Historial de préstamos inválido.");
    loans[event.id] = {
      ...data,
      id: event.id,
      requestedAt: event.at,
      status: "pending",
    };
    return;
  }
  const loan = loans[data.loanId];
  const statuses = {
    accept: "accepted",
    reject: "rejected",
    cancel: "cancelled",
  };
  if (
    !loan ||
    loan.status !== "pending" ||
    !Object.prototype.hasOwnProperty.call(statuses, data.decision) ||
    data.personId !==
      (data.decision === "cancel" ? loan.creditorId : loan.debtorId)
  )
    throw new Error("Historial de decisiones de préstamos inválido.");
  loan.status = statuses[data.decision];
  loan.decidedAt = event.at;
  loan.reason = data.reason;
  if (data.decision !== "accept") return;

  loan.obligationId = loan.id;
  obligations[loan.id] = {
    id: loan.id,
    sourceId: loan.id,
    loanId: loan.id,
    kind: "personal-loan",
    name: loan.name,
    date: loan.date,
    due: loan.due,
    debtorId: loan.debtorId,
    creditorId: loan.creditorId,
    cents: loan.cents,
    paid: 0,
    reserved: 0,
    pending: loan.cents,
  };
}
