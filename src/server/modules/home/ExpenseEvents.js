/** Operaciones de gastos guardadas en HOME_EVENTOS. */
const HOME_EXPENSE_SCHEMA_ = 3;
const HOME_EXPENSE_TRANSITIONS_ = {
  EXPENSE_SERVICE: "AGREEMENT_PROPOSE",
  EXPENSE_PURCHASE: "AGREEMENT_PROPOSE",
  EXPENSE_RESPOND: "AGREEMENT_RESPOND",
};

function homeIsExpenseType_(type) {
  return [
    "EXPENSE_SERVICE",
    "EXPENSE_PURCHASE",
    "EXPENSE_RESPOND",
    "EXPENSE_INVOICE",
    "EXPENSE_PAY",
  ].includes(type);
}

function homeExpenseResource_(state, id) {
  return state.expenseResources.find((resource) => resource.resourceId === id);
}

function homeExpenseProposal_(state, p, session, kind) {
  const resourceId = homeAgreementId_(p.resourceId);
  const previous = homeExpenseResource_(state, resourceId);
  if (
    previous
      ? previous.kind !== kind
      : state.templates.some((item) => item.id === resourceId) ||
        state.agreements.some((item) => item.resourceId === resourceId)
  )
    throw new Error("El identificador pertenece a otro gasto.");

  const details =
    previous && kind === "occasional"
      ? previous.details
      : {
          name: homeText_(p.name, "Descripción"),
          provider: p.provider ? homeText_(p.provider, "Proveedor") : "",
        };
  if (
    previous &&
    kind === "occasional" &&
    ((p.name !== undefined && String(p.name).trim() !== details.name) ||
      (p.provider !== undefined &&
        String(p.provider).trim() !== details.provider))
  )
    throw new Error("El reenvío solo permite revisar el reparto.");
  let payment = {};
  if (kind === "occasional") {
    const member = homeAgreementMember_(state, session, true);
    if (!member || p.payerId !== member.id)
      throw new Error("Solo puedes registrar una compra que pagaste tú.");
    const date = homeDate_(p.date);
    if (date > homeToday_())
      throw new Error("No se puede registrar una compra futura.");
    payment = { date, payerId: member.id, cents: homeCents_(p.amount) };
    if (
      previous &&
      JSON.stringify(payment) !== JSON.stringify(previous.payment)
    )
      throw new Error("El reenvío solo permite revisar el reparto.");
  }

  let mode = p.mode;
  let participants = p.participants;
  if (p.reuseResourceId !== undefined) {
    if (
      kind !== "occasional" ||
      previous ||
      mode !== undefined ||
      participants !== undefined
    )
      throw new Error(
        "La reutilización requiere una compra nueva sin reparto explícito.",
      );
    const source = homeAgreementsVisible_(state, session).find(
      (item) => item.resourceId === p.reuseResourceId && item.active,
    );
    if (
      !source ||
      !homeExpenseResource_(state, source.resourceId) ||
      source.kind !== "occasional"
    )
      throw new Error("El reparto seleccionado no está disponible.");
    const terms = homeApprovedAgreement_(
      state,
      source.resourceId,
      "occasional",
    );
    mode = terms.mode;
    participants = terms.participants;
  }

  const proposal = { resourceId, kind, mode, participants };
  if (kind === "occasional") {
    proposal.payerId = payment.payerId;
    proposal.amountCents = payment.cents;
  } else if (p.fixedAmount !== undefined) {
    proposal.amountCents = homeCents_(p.fixedAmount);
  }
  const agreement = homeAgreementCommand_(
    {
      type: "AGREEMENT_PROPOSE",
      payload: { expectedVersion: p.expectedVersion, proposal },
    },
    state,
    session,
  );
  return { resourceId, kind, details, payment, agreement };
}

function homeExpenseSnapshot_(terms, cents) {
  if (terms.amountCents !== undefined && terms.amountCents !== cents)
    throw new Error("El importe no coincide con el acuerdo aprobado.");
  return {
    cents,
    mode: terms.mode,
    agreementVersion: terms.version,
    shares: homeSplit_(cents, terms.participants, terms.mode),
    currency: "HNL",
    agreementStatus: "approved",
  };
}

function homeExpenseCommand_(request, state, events, session) {
  if (!state.initialized) throw new Error("Primero configura el hogar.");
  const type = request.type;
  const p = request.payload || {};
  if (type === "EXPENSE_SERVICE") {
    if (!session.admin)
      throw new Error("Esta operación requiere administrador.");
    return homeExpenseProposal_(state, p, session, "recurring");
  }
  if (type === "EXPENSE_PURCHASE")
    return homeExpenseProposal_(state, p, session, "occasional");

  if (type === "EXPENSE_PAY") {
    const charge = state.charges.find(
      (item) => item.id === p.chargeId && item.expenseKind === "recurring",
    );
    if (!charge || charge.payerId)
      throw new Error("La factura no está pendiente de pago.");
    const member = homeAgreementMember_(state, session, true);
    if (!member || member.id !== p.payerId)
      throw new Error("Solo puedes registrar un pago que realizaste tú.");
    const date = homeDate_(p.date);
    if (date < charge.date || date > homeToday_())
      throw new Error("Revisa la fecha del pago.");
    if (homeCents_(p.amount) !== charge.cents)
      throw new Error("El pago al proveedor debe cubrir el importe completo.");
    return {
      chargeId: charge.id,
      payerId: member.id,
      date,
      cents: charge.cents,
    };
  }

  const resource = homeExpenseResource_(state, homeAgreementId_(p.resourceId));
  if (type === "EXPENSE_RESPOND") {
    if (!resource || resource.kind !== "occasional")
      throw new Error("La compra no está disponible.");
    const agreement = homeAgreementCommand_(
      { type: "AGREEMENT_RESPOND", payload: p },
      state,
      session,
    );
    const prior = state.agreements.find(
      (item) => item.resourceId === resource.resourceId,
    );
    const next = homeAgreementApply_(prior, agreement.command, agreement.actor);
    const data = { resourceId: resource.resourceId, agreement };
    if (next.active) {
      const terms = homeApprovedAgreement_(
        { ...state, agreements: [next] },
        resource.resourceId,
        "occasional",
      );
      data.charge = {
        ...resource.details,
        ...resource.payment,
        ...homeExpenseSnapshot_(terms, resource.payment.cents),
        resourceId: resource.resourceId,
        expenseKind: "occasional",
        period: resource.payment.date.slice(0, 7),
        due: resource.payment.date,
        paidDate: resource.payment.date,
      };
    }
    return data;
  }

  if (!session.admin) throw new Error("Esta operación requiere administrador.");
  if (!resource || resource.kind !== "recurring")
    throw new Error("El servicio no está disponible.");
  if (
    p.payerId !== undefined ||
    p.participants !== undefined ||
    p.mode !== undefined
  )
    throw new Error("Usa el reparto aprobado y registra el pago por separado.");
  const period = String(p.period || "");
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(period))
    throw new Error("Selecciona un período válido.");
  if (
    state.charges.some(
      (item) =>
        item.resourceId === resource.resourceId && item.period === period,
    )
  )
    throw new Error("Ya existe una factura de este gasto en ese período.");
  const date = homeDate_(p.date);
  const due = homeDate_(p.due);
  if (date > homeToday_() || due < date)
    throw new Error("Revisa las fechas de la factura.");
  const terms = homeApprovedAgreement_(state, resource.resourceId, "recurring");
  if (p.version !== terms.version)
    throw new Error("La versión aprobada del reparto cambió.");
  const details = resource.versions.find(
    (version) => version.version === terms.version,
  );
  return {
    name: details.name,
    provider: details.provider,
    ...homeExpenseSnapshot_(terms, homeCents_(p.amount)),
    resourceId: resource.resourceId,
    expenseKind: "recurring",
    period,
    date,
    due,
    payerId: "",
    paidDate: "",
  };
}

function homeExpenseReplay_(state, obligations, event, data) {
  const agreementType = HOME_EXPENSE_TRANSITIONS_[event.type];
  if (agreementType) {
    if (!data.agreement || data.agreement.resourceId !== data.resourceId)
      throw new Error("Historial de gastos inválido.");
    homeAgreementReplay_(state.agreements, {
      actor: event.actor,
      type: agreementType,
      data: { ...data.agreement, schema: HOME_AGREEMENT_SCHEMA_ },
    });
  }
  if (event.type === "EXPENSE_SERVICE" || event.type === "EXPENSE_PURCHASE") {
    const { resourceId, kind, details, payment } = data;
    let resource = homeExpenseResource_(state, resourceId);
    if (!resource) {
      resource = { resourceId, kind, versions: [] };
      state.expenseResources.push(resource);
    }
    resource.details = details;
    resource.payment = payment;
    resource.versions.push({
      version: data.agreement.version,
      ...details,
      ...payment,
    });
  }
  if (event.type === "EXPENSE_INVOICE") {
    state.charges.push({ id: event.id, ...data, providerStatus: "unpaid" });
  }
  if (event.type === "EXPENSE_PAY") {
    const charge = state.charges.find(
      (item) => item.id === data.chargeId && item.expenseKind === "recurring",
    );
    if (
      !charge ||
      !data.payerId ||
      charge.payerId ||
      charge.cents !== data.cents
    )
      throw new Error("Historial de pagos al proveedor inválido.");
    Object.assign(charge, {
      payerId: data.payerId,
      paidDate: data.date,
      providerStatus: "paid",
    });
    homeAddChargeObligations_(obligations, charge);
  }
  if (event.type === "EXPENSE_RESPOND" && data.charge) {
    const charge = { id: event.id, ...data.charge, providerStatus: "paid" };
    state.charges.push(charge);
    homeAddChargeObligations_(obligations, charge);
  }
}

function homeExpenseResourcesVisible_(state, agreements) {
  return state.expenseResources.flatMap(({ resourceId, kind, versions }) => {
    const agreement = agreements.find((item) => item.resourceId === resourceId);
    if (!agreement) return [];
    const visible = new Set(
      [agreement.active, agreement.pending, ...agreement.history]
        .filter(Boolean)
        .map((version) => version.version),
    );
    return [
      {
        resourceId,
        kind,
        versions: versions.filter((version) => visible.has(version.version)),
      },
    ];
  });
}
