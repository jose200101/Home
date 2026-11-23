/** Integración de acuerdos: eventos nuevos, sin migración ni generación de obligaciones. */
const HOME_AGREEMENT_SCHEMA_ = 2;
const HOME_AGREEMENT_EVENT_VERSION_ = 1;

function homeIsAgreementType_(type) {
  return [
    "AGREEMENT_PROPOSE",
    "AGREEMENT_RESPOND",
    "AGREEMENT_CANCEL",
  ].includes(type);
}

function homeEventSupported_(type, data) {
  if (!data || typeof data !== "object") return false;
  if (homeIsCorrectionType_(type) || homeIsCreditType_(type))
    return data.schema === HOME_ADJUSTMENT_SCHEMA_ && data.eventVersion === 1;
  if (homeIsLoanType_(type))
    return data.schema === HOME_LOAN_SCHEMA_ && data.eventVersion === 1;
  if (homeIsReimbursementType_(type))
    return data.schema === HOME_REIMBURSEMENT_SCHEMA_ && data.eventVersion === 1;
  if (homeIsExpenseType_(type)) {
    return data.schema === HOME_EXPENSE_SCHEMA_ && data.eventVersion === 1;
  }
  if (homeIsAgreementType_(type)) {
    return (
      data.schema === HOME_AGREEMENT_SCHEMA_ &&
      data.eventVersion === HOME_AGREEMENT_EVENT_VERSION_
    );
  }
  return (
    data.schema === HOME_SCHEMA_ &&
    [
      "INIT",
      "PERSON",
      "TEMPLATE",
      "CHARGE",
      "DEBT",
      "PAYMENT",
      "CONFIRM",
      "REJECT",
      "FUND",
      "VOID",
    ].includes(type)
  );
}

function homeAgreementLatestVersion_(agreement) {
  if (!agreement) return 0;
  return [agreement.active, agreement.pending]
    .concat(agreement.history)
    .reduce(
      (max, version) => (version ? Math.max(max, version.version) : max),
      0,
    );
}

/** La cuenta y la persona son identidades diferentes; no aceptar un vínculo del cliente. */
function homeAgreementMember_(state, session, activeOnly) {
  const matches = state.people.filter((person) => person.userId === session.id);
  if (matches.length > 1)
    throw new Error("La cuenta tiene una vinculación ambigua con integrantes.");
  const member = matches[0];
  return member && (!activeOnly || member.active === true) ? member : null;
}

function homeAgreementCommand_(request, state, session) {
  if (!state.initialized) throw new Error("Primero configura el hogar.");
  const p = request.payload || {};
  if (
    p.actor !== undefined ||
    p.personId !== undefined ||
    p.userId !== undefined
  )
    throw new Error(
      "La identidad del actor se obtiene de la sesión, no del cliente.",
    );
  const member = homeAgreementMember_(state, session, true);
  const actor = { id: session.id, active: true, admin: session.admin };
  if (member) actor.personId = member.id;
  const resourceId = homeAgreementId_(
    request.type === "AGREEMENT_PROPOSE"
      ? p.proposal && p.proposal.resourceId
      : p.resourceId,
  );
  const previous =
    state.agreements.find((agreement) => agreement.resourceId === resourceId) ||
    null;
  let command;
  if (request.type === "AGREEMENT_PROPOSE") {
    if (
      !Number.isSafeInteger(p.expectedVersion) ||
      p.expectedVersion !== homeAgreementLatestVersion_(previous)
    )
      throw new Error(
        "La versión esperada cambió. Actualiza el acuerdo antes de proponer.",
      );
    const raw = p.proposal;
    if (!raw || typeof raw !== "object" || Array.isArray(raw))
      throw new Error("Falta la propuesta.");
    if (
      !Array.isArray(raw.participants) ||
      !raw.participants.length ||
      raw.participants.length > 100
    )
      throw new Error("Se requieren entre 1 y 100 participantes.");
    if (raw.kind === "occasional" && !member)
      throw new Error("Se requiere un integrante activo vinculado.");
    if (previous && previous.kind === "occasional") {
      const owner =
        previous.pending ||
        previous.active ||
        previous.history[previous.history.length - 1];
      if (owner.createdBy !== session.id)
        throw new Error("Solo quien registró la compra puede reenviar.");
    }
    if (
      raw.kind === "occasional" &&
      state.templates.some((template) => template.id === resourceId)
    )
      throw new Error(
        "El identificador ya pertenece a un servicio recurrente.",
      );
    const proposal = {
      resourceId,
      kind: raw.kind,
      mode: raw.mode,
      participants: Array.isArray(raw.participants)
        ? raw.participants.map((entry) => ({
            personId: entry && entry.personId,
            ...(entry && entry.value !== undefined
              ? { value: entry.value }
              : {}),
          }))
        : raw.participants,
    };
    if (raw.amountCents !== undefined) proposal.amountCents = raw.amountCents;
    if (raw.payerId !== undefined) proposal.payerId = raw.payerId;
    // El dominio valida importes y estructura antes de comprobar referencias.
    homeAgreementProposal_(proposal, actor).participants.forEach(
      (participant) => homePerson_(state, participant.personId, true),
    );
    command = { type: "PROPOSE", proposal };
  } else {
    if (!previous) throw new Error("El acuerdo no existe.");
    if (request.type === "AGREEMENT_RESPOND") {
      if (!member)
        throw new Error(
          "Se requiere un integrante activo vinculado para responder.",
        );
      if (p.decision === "accept" && previous.pending)
        previous.pending.participants.forEach((participant) =>
          homePerson_(state, participant.personId, true),
        );
      command = { type: "RESPOND", version: p.version, decision: p.decision };
    } else if (request.type === "AGREEMENT_CANCEL")
      command = { type: "CANCEL", version: p.version };
    else throw new Error("Comando de acuerdo desconocido.");
  }
  const next = homeAgreementApply_(previous, command, actor);
  return {
    eventVersion: HOME_AGREEMENT_EVENT_VERSION_,
    resourceId,
    kind: next.kind,
    version: homeAgreementLatestVersion_(next),
    command,
    actor,
  };
}

/** Replay usa la identidad y el rol históricos, nunca los permisos o vínculos actuales. */
function homeAgreementReplay_(agreements, event) {
  const data = event.data;
  const commands = {
    AGREEMENT_PROPOSE: "PROPOSE",
    AGREEMENT_RESPOND: "RESPOND",
    AGREEMENT_CANCEL: "CANCEL",
  };
  try {
    if (
      !homeEventSupported_(event.type, data) ||
      !data.actor ||
      data.actor.id !== event.actor ||
      !data.command ||
      data.command.type !== commands[event.type]
    )
      throw new Error("Contrato de evento inválido.");
    homeAgreementId_(data.resourceId);
    const index = agreements.findIndex(
      (agreement) => agreement.resourceId === data.resourceId,
    );
    const next = homeAgreementApply_(
      index >= 0 ? agreements[index] : null,
      data.command,
      data.actor,
    );
    if (
      next.resourceId !== data.resourceId ||
      next.kind !== data.kind ||
      homeAgreementLatestVersion_(next) !== data.version
    )
      throw new Error("Versión de evento inválida.");
    if (index >= 0) agreements[index] = next;
    else agreements.push(next);
  } catch (_) {
    throw new Error(
      "Historial de acuerdos inválido. No se aplicaron nuevas operaciones.",
    );
  }
}

function homeAgreementEligible_(state, agreement) {
  return (
    homeAgreementCanActivate_(agreement) &&
    agreement.active.participants.every((participant) =>
      state.people.some(
        (person) =>
          person.id === participant.personId && person.active === true,
      ),
    ) &&
    (agreement.kind !== "occasional" ||
      state.people.some(
        (person) =>
          person.id === agreement.active.payerId && person.active === true,
      ))
  );
}

/** Vista por versión: no revelar una modificación en la que el integrante no participa. */
function homeAgreementsVisible_(state, session) {
  if (!state.agreements.length) return [];
  const member = session.admin ? null : homeAgreementMember_(state, session, false);
  const allowed = (version) =>
    version &&
    member &&
    (version.participants.some(
      (participant) => participant.personId === member.id,
    ) ||
      version.payerId === member.id);
  function visibleVersion(version) {
    if (!version || (!session.admin && !allowed(version))) return null;
    const result = JSON.parse(JSON.stringify(version));
    if (!session.admin) {
      result.createdByMe = result.createdBy === session.id;
      delete result.createdBy;
      result.participants.forEach((participant) => {
        if (participant.personId !== member.id) delete participant.decision;
      });
    }
    return result;
  }
  return state.agreements
    .map((agreement) => {
      const active = visibleVersion(agreement.active),
        pending = visibleVersion(agreement.pending);
      const history = agreement.history.map(visibleVersion).filter(Boolean);
      if (!session.admin && !active && !pending && !history.length) return null;
      return {
        resourceId: agreement.resourceId,
        kind: agreement.kind,
        active,
        pending,
        history,
        latestVersion: homeAgreementLatestVersion_(agreement),
        suspended: !!agreement.pending,
        canActivate: homeAgreementEligible_(state, agreement),
        canInvoice:
          agreement.kind === "recurring" &&
          homeAgreementEligible_(state, agreement),
      };
    })
    .filter(Boolean);
}

/** Contrato interno para facturación futura; no escribe eventos ni crea obligaciones. */
function homeApprovedAgreement_(state, resourceId, kind) {
  const agreement = state.agreements.find(
    (item) => item.resourceId === resourceId,
  );
  if (
    !agreement ||
    agreement.kind !== kind ||
    !homeAgreementEligible_(state, agreement)
  )
    throw new Error(
      "El gasto requiere un acuerdo aprobado, sin modificaciones pendientes y con integrantes activos.",
    );
  const active = agreement.active;
  return {
    resourceId,
    kind,
    version: active.version,
    mode: active.mode,
    ...(active.amountCents !== undefined
      ? { amountCents: active.amountCents }
      : {}),
    ...(active.payerId ? { payerId: active.payerId } : {}),
    participants: active.participants.map((participant) => ({
      personId: participant.personId,
      ...(active.mode === "percent"
        ? { value: (participant.basis / 100).toFixed(2) }
        : {}),
      ...(active.mode === "amount"
        ? { value: (participant.cents / 100).toFixed(2) }
        : {}),
    })),
  };
}

/** Puente mínimo al CHARGE existente, solo para recursos que optaron por acuerdos. */
function homeAgreementChargeTerms_(state, payload, cents) {
  if (
    !state.agreements.some(
      (agreement) => agreement.resourceId === payload.templateId,
    )
  )
    return null;
  const contract = homeApprovedAgreement_(
    state,
    payload.templateId,
    "recurring",
  );
  if (contract.amountCents !== undefined && contract.amountCents !== cents)
    throw new Error("El importe no coincide con el acuerdo aprobado.");
  if (payload.mode !== undefined && payload.mode !== contract.mode)
    throw new Error("La distribución no coincide con el acuerdo aprobado.");
  if (payload.participants !== undefined) {
    const expected = homeSplit_(cents, contract.participants, contract.mode);
    const requested = homeSplit_(cents, payload.participants, contract.mode);
    if (JSON.stringify(expected) !== JSON.stringify(requested))
      throw new Error(
        "Los participantes o aportes no coinciden con el acuerdo aprobado.",
      );
  }
  return contract;
}
