/**
 * Home v2: un hogar por Spreadsheet. Importes enteros en centavos.
 * Cada operación se confirma con UNA fila de eventos, bajo ScriptLock.
 * Las proyecciones son derivadas: ninguna escritura financiera cruza hojas.
 * Los endpoints públicos están limitados a homeGetState y homeExecute.
 */
const HOME_SCHEMA_ = 1;
const HOME_SHEET_ = "HOME_EVENTOS";
const HOME_HEADERS_ = ["id", "fecha", "actor", "tipo", "datos"];

function homeSession_(token) {
  const result = validateToken(token);
  if (!result.ok || !result.user || !result.user.id)
    throw new Error("Sesión vencida. Vuelve a ingresar.");
  // Releer usuario: desactivar o cambiar su rol surte efecto inmediatamente.
  const user = _readUsers_().find(
    (u) => String(u.id) === String(result.user.id),
  );
  if (!user || String(user.estado).trim().toLowerCase() !== "activo")
    throw new Error("La cuenta no está activa.");
  const role = String(user.rol || "")
    .trim()
    .toLowerCase();
  return {
    id: String(user.id),
    nombre: user.nombreCompleto || user.nombre_usuario,
    admin: ["admin", "administrador", "propietario"].indexOf(role) >= 0,
  };
}

function homeRead_() {
  const sh = conexion_().getSheetByName(HOME_SHEET_);
  if (!sh) return [];
  if (sh.getLastRow() < 1)
    throw new Error("HOME_EVENTOS no tiene encabezados válidos.");
  const rows = sh.getDataRange().getValues();
  if (HOME_HEADERS_.some((h, i) => rows[0][i] !== h))
    throw new Error(
      "Esquema HOME_EVENTOS no reconocido. No se escribió ningún movimiento.",
    );
  return rows
    .slice(1)
    .filter((r) => r[0])
    .map((r) => {
      try {
        const data = JSON.parse(r[4]);
        if (!homeEventSupported_(String(r[3]), data))
          throw new Error("Versión desconocida.");
        return {
          id: String(r[0]),
          at: String(r[1]),
          actor: String(r[2]),
          type: String(r[3]),
          data,
        };
      } catch (_) {
        throw new Error(
          "Historial inválido. Solicita revisar el respaldo antes de continuar.",
        );
      }
    });
}

function homeText_(value, name, max) {
  const text = String(value == null ? "" : value).trim();
  if (!text || text.length > (max || 160))
    throw new Error(
      name +
        " es requerido y debe tener como máximo " +
        (max || 160) +
        " caracteres.",
    );
  return text;
}

function homeCents_(value, allowZero) {
  const text = String(value == null ? "" : value).trim();
  if (!/^\d+(\.\d{1,2})?$/.test(text))
    throw new Error(
      "Ingresa un monto válido con hasta dos decimales, sin separadores de miles.",
    );
  const parts = text.split(".");
  const cents =
    Number(parts[0]) * 100 + Number((parts[1] || "").padEnd(2, "0"));
  if (
    !Number.isSafeInteger(cents) ||
    cents > 10000000000 ||
    cents < (allowZero ? 0 : 1)
  )
    throw new Error("El monto está fuera del rango permitido.");
  return cents;
}

function homeDate_(value) {
  const str = String(value || "");
  if (!/^\d{4}-\d{2}-\d{2}$/.test(str))
    throw new Error("Selecciona una fecha válida.");
  const date = new Date(str + "T12:00:00Z");
  if (isNaN(date.getTime()) || date.toISOString().slice(0, 10) !== str)
    throw new Error("La fecha no existe.");
  return str;
}

function homeToday_() {
  return Utilities.formatDate(new Date(), "America/Tegucigalpa", "yyyy-MM-dd");
}

/** Reparto determinista: el residuo de centavos sigue el orden de los participantes. */
function homeSplit_(total, participants, mode) {
  if (
    !Array.isArray(participants) ||
    !participants.length ||
    participants.length > 100
  )
    throw new Error("Selecciona entre 1 y 100 participantes.");
  if (new Set(participants.map((p) => p.personId)).size !== participants.length)
    throw new Error("Un participante no puede repetirse.");
  if (["equal", "percent", "amount"].indexOf(mode) < 0)
    throw new Error("Selecciona una distribución válida.");
  let shares;
  if (mode === "amount") {
    shares = participants.map((p) => homeCents_(p.value, true));
    if (shares.reduce((a, b) => a + b, 0) !== total)
      throw new Error("Los aportes deben sumar exactamente el total.");
  } else {
    const weights =
      mode === "equal"
        ? participants.map(() => 1)
        : participants.map((p) => homeCents_(p.value, true));
    const sum = weights.reduce((a, b) => a + b, 0);
    if (mode === "percent" && sum !== 10000)
      throw new Error("Los porcentajes deben sumar 100 %.");
    shares = weights.map((w) => Math.floor((total * w) / sum));
    let remainder = total - shares.reduce((a, b) => a + b, 0);
    for (let i = 0; remainder > 0; i = (i + 1) % shares.length) {
      if (weights[i] > 0) {
        shares[i]++;
        remainder--;
      }
    }
  }
  return participants.map((p, i) => ({
    personId: p.personId,
    cents: shares[i],
  }));
}

function homeProject_(events) {
  const voided = new Set(
    events.filter((e) => e.type === "VOID").map((e) => e.data.targetId),
  );
  const state = {
    initialized: false,
    household: "Home",
    people: [],
    templates: [],
    obligations: [],
    charges: [],
    payments: [],
    audit: [],
    agreements: [],
  };
  const people = {},
    templates = {},
    obligations = {},
    payments = {};
  events.forEach((e) => {
    const d = Object.assign({}, e.data);
    delete d.requestJson;
    state.audit.push({
      id: e.id,
      at: e.at,
      actor: e.actor,
      type: e.type,
      note: d.reason || d.name || d.note || "",
      voided: voided.has(e.id),
    });
    if (voided.has(e.id)) return;
    if (homeIsAgreementType_(e.type)) {
      homeAgreementReplay_(state.agreements, e);
      return;
    }
    if (e.type === "INIT") {
      state.initialized = true;
      state.household = d.name;
    }
    if (e.type === "PERSON") people[d.id] = Object.assign({}, d);
    if (e.type === "TEMPLATE") templates[d.id] = Object.assign({}, d);
    if (e.type === "CHARGE") {
      state.charges.push(Object.assign({ id: e.id }, d));
      d.shares.forEach((s, i) => {
        if (d.payerId === s.personId || !s.cents) return;
        const id = e.id + ":" + i;
        obligations[id] = {
          id,
          sourceId: e.id,
          kind: "charge",
          name: d.name,
          period: d.period,
          date: d.date,
          due: d.due,
          debtorId: s.personId,
          creditorId: d.payerId || "",
          cents: s.cents,
          paid: 0,
          reserved: 0,
          pending: s.cents,
        };
      });
    }
    if (e.type === "DEBT")
      obligations[e.id] = {
        id: e.id,
        sourceId: e.id,
        kind: "debt",
        name: d.name,
        date: d.date,
        due: d.due,
        debtorId: d.debtorId,
        creditorId: d.creditorId,
        cents: d.cents,
        paid: 0,
        reserved: 0,
        pending: d.cents,
      };
    if (e.type === "PAYMENT")
      payments[e.id] = Object.assign(
        { id: e.id, status: d.creditorId ? "pending" : "confirmed" },
        d,
      );
    if (e.type === "CONFIRM" && payments[d.paymentId])
      payments[d.paymentId].status = "confirmed";
    if (e.type === "REJECT" && payments[d.paymentId])
      payments[d.paymentId].status = "rejected";
    if (e.type === "FUND") {
      const charge = state.charges.find((c) => c.id === d.chargeId);
      if (charge) {
        charge.payerId = d.payerId;
        charge.paidDate = d.date;
      }
      Object.keys(obligations).forEach((id) => {
        const ob = obligations[id];
        if (ob.sourceId === d.chargeId && !ob.creditorId) {
          if (ob.debtorId === d.payerId) ob.paid += ob.cents - ob.paid;
          else ob.creditorId = d.payerId;
        }
      });
    }
  });
  Object.keys(payments).forEach((id) => {
    const p = payments[id],
      ob = obligations[p.obligationId];
    if (ob && p.status === "confirmed") ob.paid += p.cents;
    if (ob && p.status === "pending") ob.reserved += p.cents;
  });
  state.people = Object.keys(people).map((id) => people[id]);
  state.templates = Object.keys(templates).map((id) => templates[id]);
  state.obligations = Object.keys(obligations).map((id) => {
    const ob = obligations[id];
    ob.pending = ob.cents - ob.paid;
    return ob;
  });
  state.payments = Object.keys(payments).map((id) => payments[id]);
  return state;
}

function homePerson_(state, id, activeOnly) {
  const person = state.people.find(
    (p) => p.id === id && (!activeOnly || p.active),
  );
  if (!person) throw new Error("El integrante no existe o está desactivado.");
  return person;
}

function homePlan_(obligations) {
  const net = {};
  obligations
    .filter((o) => o.creditorId && o.pending > 0)
    .forEach((o) => {
      // Los pagos por confirmar no se liquidan de nuevo.
      const amount = o.pending - o.reserved;
      net[o.debtorId] = (net[o.debtorId] || 0) - amount;
      net[o.creditorId] = (net[o.creditorId] || 0) + amount;
    });
  const debtors = Object.keys(net)
    .filter((id) => net[id] < 0)
    .map((id) => ({ id, cents: -net[id] }));
  const creditors = Object.keys(net)
    .filter((id) => net[id] > 0)
    .map((id) => ({ id, cents: net[id] }));
  const plan = [];
  let i = 0,
    j = 0;
  while (i < debtors.length && j < creditors.length) {
    const cents = Math.min(debtors[i].cents, creditors[j].cents);
    plan.push({ debtorId: debtors[i].id, creditorId: creditors[j].id, cents });
    debtors[i].cents -= cents;
    creditors[j].cents -= cents;
    if (!debtors[i].cents) i++;
    if (!creditors[j].cents) j++;
  }
  return plan;
}

function homeVisible_(state, session) {
  const me = state.people.find((p) => p.userId === session.id && p.active);
  const result = Object.assign({}, state, {
    admin: session.admin,
    me: me ? me.id : "",
    today: homeToday_(),
    agreements: homeAgreementsVisible_(state, session),
  });
  if (!session.admin) {
    const id = me ? me.id : "__unlinked__";
    result.obligations = state.obligations.filter(
      (o) => o.debtorId === id || o.creditorId === id,
    );
    result.payments = state.payments.filter(
      (p) => p.debtorId === id || p.creditorId === id,
    );
    result.charges = state.charges
      .filter(
        (c) => c.shares.some((s) => s.personId === id) || c.payerId === id,
      )
      .map((c) =>
        Object.assign({}, c, {
          shares: c.shares.filter((s) => s.personId === id),
        }),
      );
    result.templates = [];
    result.audit = [];
    result.people = state.people.map((p) => ({
      id: p.id,
      name: p.name,
      active: p.active,
    }));
    result.plan = []; // Un neteo global podría revelar deudas de terceros.
  } else result.plan = homePlan_(state.obligations);
  return result;
}

function homeGetState(token) {
  const session = homeSession_(token);
  return homeVisible_(homeProject_(homeRead_()), session);
}

function homeCommand_(request, state, events, session) {
  const type = request.type;
  const p = request.payload || {};
  if (homeIsAgreementType_(type)) return homeAgreementCommand_(request, state, session);
  if (!session.admin && ["PAYMENT", "CONFIRM", "REJECT"].indexOf(type) < 0)
    throw new Error("Esta operación requiere administrador.");
  if (!state.initialized && type !== "INIT")
    throw new Error("Primero configura el hogar.");
  if (type === "INIT") {
    if (state.initialized) throw new Error("El hogar ya está configurado.");
    return { name: homeText_(p.name, "Nombre del hogar") };
  }
  if (type === "PERSON") {
    const id = p.id ? homePerson_(state, p.id, false).id : Utilities.getUuid();
    const userId = String(p.userId || "").trim();
    if (userId && !_readUsers_().some((u) => String(u.id) === userId))
      throw new Error("La cuenta seleccionada no existe.");
    if (
      userId &&
      state.people.some(
        (person) => person.id !== id && person.userId === userId,
      )
    )
      throw new Error("Esa cuenta ya está vinculada a otro integrante.");
    return {
      id,
      name: homeText_(p.name, "Nombre"),
      userId,
      active: p.active !== false,
    };
  }
  if (type === "TEMPLATE") {
    if (p.id && !state.templates.some((t) => t.id === p.id))
      throw new Error("El gasto no existe.");
    const mode = p.mode || "equal";
    const participants = (p.participants || []).map((v) => ({
      personId: homePerson_(state, v.personId, p.active !== false).id,
      value: String(v.value || "0"),
    }));
    // Para montos, el total se valida al crear la factura.
    const testTotal =
      mode === "amount"
        ? participants.reduce((sum, v) => sum + homeCents_(v.value, true), 0)
        : 10000;
    homeSplit_(testTotal, participants, mode);
    return {
      id: p.id || Utilities.getUuid(),
      name: homeText_(p.name, "Nombre del gasto"),
      provider: String(p.provider || "")
        .trim()
        .slice(0, 160),
      mode,
      participants,
      active: p.active !== false,
    };
  }
  if (type === "CHARGE") {
    const template = state.templates.find(
      (t) => t.id === p.templateId && t.active,
    );
    if (!template) throw new Error("Selecciona un gasto activo.");
    const period = String(p.period || "");
    if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(period))
      throw new Error("Selecciona un período válido.");
    if (
      state.charges.some(
        (c) => c.templateId === template.id && c.period === period,
      )
    )
      throw new Error("Ya existe una factura de este gasto en ese período.");
    const cents = homeCents_(p.amount);
    const agreement = homeAgreementChargeTerms_(state, p, cents);
    const participants = agreement ? agreement.participants : p.participants || template.participants;
    participants.forEach((v) => homePerson_(state, v.personId, true));
    const payerId = p.payerId ? homePerson_(state, p.payerId, true).id : "";
    const date = homeDate_(p.date),
      due = homeDate_(p.due);
    if (due < date)
      throw new Error(
        "El vencimiento no puede ser anterior a la fecha de emisión.",
      );
    if (payerId && date > homeToday_())
      throw new Error("No se puede registrar un pago futuro.");
    return {
      templateId: template.id,
      name: template.name,
      provider: template.provider,
      period,
      date,
      due,
      cents,
      payerId,
      mode: agreement ? agreement.mode : p.mode || template.mode,
      shares: homeSplit_(cents, participants, agreement ? agreement.mode : p.mode || template.mode),
      ...(agreement ? {agreementVersion: agreement.version} : {}),
    };
  }
  if (type === "DEBT") {
    const debtorId = homePerson_(state, p.debtorId, true).id;
    const creditorId = homePerson_(state, p.creditorId, true).id;
    if (debtorId === creditorId)
      throw new Error("Selecciona dos integrantes diferentes.");
    const date = homeDate_(p.date),
      due = homeDate_(p.due);
    if (date > homeToday_() || due < date)
      throw new Error("Revisa las fechas de la deuda.");
    return {
      name: homeText_(p.name, "Concepto"),
      debtorId,
      creditorId,
      cents: homeCents_(p.amount),
      date,
      due,
    };
  }
  if (type === "PAYMENT") {
    const ob = state.obligations.find((o) => o.id === p.obligationId);
    if (!ob) throw new Error("La obligación no existe.");
    const me = state.people.find(
      (person) => person.userId === session.id && person.active,
    );
    if (!session.admin && (!me || ob.debtorId !== me.id))
      throw new Error("Solo puedes registrar pagos de tus obligaciones.");
    if (!session.admin && !ob.creditorId)
      throw new Error("Un administrador debe registrar el pago al proveedor.");
    const cents = homeCents_(p.amount);
    if (cents > ob.pending - ob.reserved)
      throw new Error(
        "El monto supera el saldo disponible. Hay pagos pendientes de confirmar o la deuda ya está cubierta.",
      );
    const date = homeDate_(p.date);
    if (date > homeToday_() || date < ob.date)
      throw new Error(
        "El pago debe estar entre la fecha de la obligación y hoy.",
      );
    return {
      obligationId: ob.id,
      debtorId: ob.debtorId,
      creditorId: ob.creditorId,
      cents,
      date,
      note: homeText_(p.note || "Pago registrado", "Nota", 500),
    };
  }
  if (type === "CONFIRM" || type === "REJECT") {
    const payment = state.payments.find(
      (v) => v.id === p.paymentId && v.status === "pending",
    );
    if (!payment) throw new Error("El pago ya fue atendido o no existe.");
    const me = state.people.find(
      (person) => person.userId === session.id && person.active,
    );
    if (!session.admin && (!me || payment.creditorId !== me.id))
      throw new Error(
        "Solo quien recibe el dinero puede confirmar o rechazar este pago.",
      );
    return {
      paymentId: payment.id,
      reason: homeText_(
        p.reason || (type === "CONFIRM" ? "Recepción confirmada" : ""),
        "Motivo",
        500,
      ),
    };
  }
  if (type === "FUND") {
    const charge = state.charges.find((c) => c.id === p.chargeId);
    if (
      !charge ||
      charge.payerId ||
      events.some((e) => e.type === "FUND" && e.data.chargeId === charge.id)
    )
      throw new Error(
        "La factura ya fue pagada por un integrante o no existe.",
      );
    const obs = state.obligations.filter((o) => o.sourceId === charge.id);
    if (
      state.payments.some(
        (pay) =>
          obs.some((o) => o.id === pay.obligationId) &&
          pay.status !== "rejected",
      )
    )
      throw new Error(
        "Esta factura tiene abonos. Registra sus aportes restantes por separado.",
      );
    const date = homeDate_(p.date);
    if (date > homeToday_() || date < charge.date)
      throw new Error("Revisa la fecha del pago.");
    return {
      chargeId: charge.id,
      payerId: homePerson_(state, p.payerId, true).id,
      date,
      cents: charge.cents,
    };
  }
  if (type === "VOID") {
    const target = events.find((e) => e.id === p.targetId);
    if (!target || ["CHARGE", "DEBT", "PAYMENT"].indexOf(target.type) < 0)
      throw new Error("Solo se pueden anular facturas, deudas o pagos.");
    if (events.some((e) => e.type === "VOID" && e.data.targetId === target.id))
      throw new Error("El movimiento ya está anulado.");
    if (target.type !== "PAYMENT") {
      const obs = state.obligations.filter((o) => o.sourceId === target.id);
      if (
        state.payments.some(
          (pay) =>
            obs.some((o) => o.id === pay.obligationId) &&
            pay.status !== "rejected",
        ) ||
        events.some((e) => e.type === "FUND" && e.data.chargeId === target.id)
      )
        throw new Error(
          "Tiene movimientos relacionados. Revierte primero sus pagos; una factura financiada no se puede anular.",
        );
    }
    return {
      targetId: target.id,
      reason: homeText_(p.reason, "Motivo de anulación", 500),
    };
  }
  throw new Error("Operación desconocida.");
}

function homeExecute(token, request) {
  let session = homeSession_(token);
  if (request && homeIsAgreementType_(request.type) && typeof request.id !== "string")
    throw new Error("Identificador de operación inválido.");
  if (!request || !/^[a-zA-Z0-9_-]{12,100}$/.test(String(request.id || "")))
    throw new Error("Identificador de operación inválido.");
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (homeIsAgreementType_(request.type)) session = homeSession_(token);
    const events = homeRead_();
    const previous = events.find((e) => e.id === request.id);
    if (previous) {
      if (previous.actor !== session.id || previous.type !== request.type)
        throw new Error("Identificador de operación ya utilizado.");
      if (previous.data.requestJson !== JSON.stringify(request.payload || {}))
        throw new Error(
          "La operación original ya está registrada. Actualiza las cuentas antes de registrar datos diferentes.",
        );
      return {
        ok: true,
        duplicate: true,
        id: previous.id,
        state: homeVisible_(homeProject_(events), session),
      };
    }
    const state = homeProject_(events);
    const data = homeCommand_(request, state, events, session);
    data.schema = homeIsAgreementType_(request.type) ? HOME_AGREEMENT_SCHEMA_ : HOME_SCHEMA_;
    data.requestJson = JSON.stringify(request.payload || {});
    const event = {
      id: request.id,
      at: new Date().toISOString(),
      actor: session.id,
      type: request.type,
      data,
    };
    const json = JSON.stringify(data);
    if (json.length > 40000)
      throw new Error(
        "La operación es demasiado grande. Reduce el número de participantes.",
      );
    const ss = conexion_();
    let sh = ss.getSheetByName(HOME_SHEET_);
    if (!sh) {
      sh = ss.insertSheet(HOME_SHEET_);
      sh.appendRow(HOME_HEADERS_);
      sh.setFrozenRows(1);
    }
    // Una sola escritura es la fuente de verdad de la operación.
    sh.appendRow([event.id, event.at, event.actor, event.type, json]);
    SpreadsheetApp.flush();
    return {
      ok: true,
      id: event.id,
      state: homeVisible_(homeProject_(events.concat([event])), session),
    };
  } finally {
    lock.releaseLock();
  }
}

function homeMigrationPreview(token) {
  const session = homeSession_(token);
  if (!session.admin) throw new Error("Requiere administrador.");
  // Solo lectura: nunca invocar listados antiguos que crean hojas automáticamente.
  return conexion_()
    .getSheets()
    .filter((sh) => sh.getName() !== HOME_SHEET_)
    .map((sh) => ({
      sheet: sh.getName(),
      rows: Math.max(0, sh.getLastRow() - 1),
      headers: sh.getLastColumn()
        ? sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0]
        : [],
    }));
}
