/**
 * =========================
 *  PRÉSTAMOS (Sin tabla de cuotas)
 * =========================
 *
 * NUEVO ENFOQUE:
 *  - NO se escriben filas por cuota en Google Sheets (PRESTAMOS_CUOTAS).
 *  - Se guarda SOLO:
 *      - Cabecera del préstamo en PRESTAMOS
 *      - Ledger de pagos en Prestamos_Pagos (1 fila por pago)
 *  - La amortización y el estado de cuotas (Pagada/Pendiente/Vencida) se calculan "al vuelo"
 *    usando:
 *      - Datos del préstamo (monto, plazo, tasa, fechas)
 *      - Pagos registrados
 *
 * IMPORTANTE:
 *  - Si ya tienes datos viejos en PRESTAMOS_CUOTAS, este backend ya NO los usa.
 *  - Puedes borrar/limpiar PRESTAMOS_CUOTAS para aligerar tu Sheet.
 */

// =========================
// Helpers básicos
// =========================

function _prSheetName_(key, fallback) {
  try {
    const cfg = (typeof env_ === "function") ? env_() : {};
    return String(cfg[key] || fallback || "").trim() || fallback;
  } catch (e) {
    return fallback;
  }
}

function _prNowIso_() {
  const tz = Session.getScriptTimeZone() || "GMT";
  return Utilities.formatDate(new Date(), tz, "yyyy-MM-dd'T'HH:mm:ss");
}

function _prRound2_(n) {
  const x = Number(n || 0);
  return Math.round((isNaN(x) ? 0 : x) * 100) / 100;
}

function _prNum_(v) {
  if (typeof v === "number") return isNaN(v) ? 0 : v;
  const s = String(v ?? "").replace(/[^0-9.\-]/g, "");
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function _prIso_(d) {
  if (!d) return "";
  if (Object.prototype.toString.call(d) === "[object Date]") {
    if (isNaN(d.getTime())) return "";
    const tz = Session.getScriptTimeZone() || "GMT";
    return Utilities.formatDate(d, tz, "yyyy-MM-dd");
  }
  // ya viene string
  const s = String(d || "").trim();
  // intenta normalizar si trae "yyyy-mm-dd ..."
  const m = s.match(/^(\d{4}-\d{2}-\d{2})/);
  return m ? m[1] : s;
}

function _prDateFromIso_(iso) {
  const s = String(iso || "").trim();
  if (!s) return null;
  // yyyy-mm-dd
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return null;
  const y = parseInt(m[1], 10);
  const mo = parseInt(m[2], 10) - 1;
  const d = parseInt(m[3], 10);
  const dt = new Date(y, mo, d);
  return isNaN(dt.getTime()) ? null : dt;
}

function _prDateOnly_(d) {
  if (!d) return null;
  const dt = (Object.prototype.toString.call(d) === "[object Date]") ? d : _prDateFromIso_(d);
  if (!dt || isNaN(dt.getTime())) return null;
  return new Date(dt.getFullYear(), dt.getMonth(), dt.getDate());
}

function _prDiffDays_(a, b) {
  // a - b (en días), con fechas "dateOnly"
  if (!a || !b) return 0;
  const ms = _prDateOnly_(a).getTime() - _prDateOnly_(b).getTime();
  return Math.floor(ms / 86400000);
}

function _prDaysInMonth_(year, monthIndex0) {
  return new Date(year, monthIndex0 + 1, 0).getDate();
}

function _prDateWithDay_(year, monthIndex0, day1to31) {
  const dmax = _prDaysInMonth_(year, monthIndex0);
  const dd = Math.max(1, Math.min(dmax, parseInt(day1to31, 10) || 1));
  return new Date(year, monthIndex0, dd);
}

function _prCuotaEstadoAuto_(totalPendiente, dueDate, today) {
  const tp = _prRound2_(totalPendiente);
  if (tp <= 0.000001) return "Pagada";
  if (dueDate && today && today.getTime() > dueDate.getTime()) return "Vencida";
  return "Pendiente";
}

function _prLoanEstadoAuto_(totalPendiente) {
  const tp = _prRound2_(totalPendiente);
  return (tp <= 0.000001) ? "Finalizado" : "Activo";
}

// =========================
// Sheets (schema)
// =========================

function _prEnsureSheet_(name, headers) {
  const ss = conexion_(); // viene de database/conexion_.js
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  const lastCol = sh.getLastColumn();
  const lastRow = sh.getLastRow();

  if (lastRow === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }

  // asegurar headers
  const current = sh.getRange(1, 1, 1, Math.max(lastCol, headers.length)).getDisplayValues()[0] || [];
  const norm = (s) => String(s || "").trim();
  const existing = current.map(norm);
  let changed = false;

  headers.forEach((h, i) => {
    const need = norm(h);
    if (!need) return;
    if (existing[i] !== need) {
      // si está vacío, lo coloca
      if (!existing[i]) {
        existing[i] = need;
        changed = true;
      }
    }
  });

  if (changed) {
    sh.getRange(1, 1, 1, existing.length).setValues([existing]);
  }
}

function _prEnsureSchema_() {
  const shPrestamos = _prSheetName_("SH_PRESTAMOS", "PRESTAMOS");
  const shPagos = _prSheetName_("SH_PRESTAMOS_PAGOS", "Prestamos_Pagos");
  const shParams = _prSheetName_("SH_PARAMETROS", "PARAMETROS");

  _prEnsureSheet_(shPrestamos, [
    "idPrestamo",
    "origen",
    "tipoCliente",
    "idPersona",
    "personaNombre",
    "montoPrincipal",
    "plazoMeses",
    "tasaMensual",
    "moratorioModo",
    "tasaMoratoriaMensual",
    "adminTipo",
    "adminValor",
    "adminMonto",
    "fechaDesembolso",
    "fechaHoraDesembolso",
    "metodoDesembolso",
    "refDesembolso",
    "notaDesembolso",
    "fechaPrimerPago",
    "diaCorte",
    "diaPago",
    "cuotaMensual",
    "totalInteresEstimado",
    "totalPagarEstimado",
    "estado",
    "creadoPor",
    "fechaCreacion",
    "actualizadoPor",
    "fechaActualizacion"
  ]);

  _prEnsureSheet_(shPagos, [
    "idPago",
    "idPrestamo",
    "fechaHora",
    "monto",
    "metodo",
    "referencia",
    "nota",
    "aplicadoA",
    "moratorioCobrado",
    "interesCobrado",
    "capitalCobrado",
    "saldoFavor"
  ]);

  _prEnsureSheet_(shParams, ["clave", "valor"]);
}

function _prRead_(sh) {
  const values = sh.getDataRange().getDisplayValues();
  const headers = (values[0] || []).map(h => String(h || "").trim());
  const idx = {};
  headers.forEach((h, i) => { if (h) idx[h] = i; });
  const rows = values.length > 1
    ? values.slice(1).filter(r => r.some(c => String(c || "").trim() !== ""))
    : [];
  return { headers, rows, idx };
}

function _prGet_(row, idx, name) {
  const i = idx[name];
  return (i === undefined) ? "" : row[i];
}

function _prFindRowIndexById_(rows, idx, idPrestamo) {
  const i = idx["idPrestamo"];
  if (i === undefined) return -1;
  const needle = String(idPrestamo || "").trim();
  if (!needle) return -1;
  for (let r = 0; r < rows.length; r++) {
    if (String(rows[r][i] || "").trim() === needle) return r;
  }
  return -1;
}

function _prLookupPersonaNombre_(idPersona) {
  try {
    const sh = obtenerSheet_(_prSheetName_("SH_PERSONAS", "PERSONAS"));
    const data = sh.getDataRange().getDisplayValues();
    const headers = (data[0] || []).map(h => String(h || "").trim().toLowerCase());
    const hIdx = {};
    headers.forEach((h, i) => { if (h) hIdx[h] = i; });

    const norm = (s) => String(s || "").trim().toLowerCase();
    const iId = hIdx[norm("id_persona")] ?? hIdx[norm("id")] ?? hIdx[norm("persona_id")] ?? -1;
    const iNombre = hIdx[norm("nombre_persona")] ?? hIdx[norm("nombre")] ?? hIdx[norm("persona")] ?? -1;
    if (iId < 0 || iNombre < 0) return "";

    const pid = String(idPersona || "").trim();
    for (let r = 1; r < data.length; r++) {
      const row = data[r];
      if (!row || row.every(c => String(c || "").trim() === "")) continue;
      if (String(row[iId] || "").trim() === pid) return String(row[iNombre] || "").trim();
    }
    return "";
  } catch (e) {
    return "";
  }
}

// =========================
// Amortización base (NO se guarda en Sheet)
// =========================

function _prComputeCuota_(P, r, N) {
  const principal = _prNum_(P);
  const rate = Number(r || 0);
  const n = Math.max(1, parseInt(N, 10) || 1);
  if (principal <= 0) return 0;
  if (!rate || Math.abs(rate) < 1e-12) return principal / n;
  const pow = Math.pow(1 + rate, n);
  return principal * rate * pow / (pow - 1);
}

/**
 * Construye cuotas base desde el préstamo (sin pagos).
 * Usa:
 *  - fechaPrimerPago si existe
 *  - si no existe: próximo mes en diaPago (como estaba)
 */
function _prBuildCuotasBaseFromPrestamo_(prestamo) {
  const P = _prNum_(prestamo.montoPrincipal);
  const N = Math.max(1, parseInt(prestamo.plazoMeses, 10) || 1);
  const r = Number(prestamo.tasaMensual || 0);
  const diaPago = Math.max(1, Math.min(28, parseInt(prestamo.diaPago, 10) || 1));
  const fechaDesembolso = _prIso_(prestamo.fechaDesembolso);

  const cuotaRaw = _prComputeCuota_(P, r, N);
  const cuotaMensual = _prRound2_(cuotaRaw);

  // fecha primer pago
  let dtPago = _prDateFromIso_(prestamo.fechaPrimerPago);
  if (!dtPago) {
    const disb = _prDateFromIso_(fechaDesembolso) || new Date();
    let y = disb.getFullYear();
    let m = disb.getMonth() + 1;
    if (m > 11) { m = 0; y += 1; }
    dtPago = _prDateWithDay_(y, m, diaPago);
  }
  const fechaPrimerPago = _prIso_(dtPago);

  const cuotas = [];
  let saldo = P;
  let totalInteres = 0;

  for (let i = 1; i <= N; i++) {
    const interes = _prRound2_(saldo * r);
    let capital = _prRound2_(cuotaMensual - interes);
    if (i === N) {
      capital = _prRound2_(saldo); // ajuste final
    }
    saldo = _prRound2_(Math.max(0, saldo - capital));
    totalInteres = _prRound2_(totalInteres + interes);

    const fechaPago = _prIso_(dtPago);
    cuotas.push({
      idCuota: `${String(prestamo.idPrestamo || "").trim()}-${i}`,
      idPrestamo: String(prestamo.idPrestamo || "").trim(),
      nroCuota: i,
      fechaPago,
      cuota: (i === N) ? _prRound2_(capital + interes) : cuotaMensual,
      interesCorriente: interes,
      capital,
      saldoDespues: saldo,
      // estado/pagos (se llenan al vuelo)
      interesPagado: 0,
      capitalPagado: 0,
      moraAcumulada: 0,
      moraPagada: 0,
      moraCalculadaHasta: fechaPago
    });

    // siguiente mes
    let yy = dtPago.getFullYear();
    let mm = dtPago.getMonth() + 1;
    if (mm > 11) { mm = 0; yy += 1; }
    dtPago = _prDateWithDay_(yy, mm, diaPago);
  }

  return { cuotaMensual, fechaPrimerPago, totalInteres, cuotas };
}

/**
 * Lee pagos del préstamo (ledger) y retorna arreglo
 */
function _prReadPagosByPrestamo_(idPrestamo) {
  const shPagos = obtenerSheet_(_prSheetName_("SH_PRESTAMOS_PAGOS", "Prestamos_Pagos"));
  const read = _prRead_(shPagos);
  const out = [];
  const id = String(idPrestamo || "").trim();
  read.rows.forEach(r => {
    const pid = String(_prGet_(r, read.idx, "idPrestamo") || "").trim();
    if (pid !== id) return;
    out.push({
      idPago: String(_prGet_(r, read.idx, "idPago") || "").trim(),
      idPrestamo: pid,
      fechaHora: String(_prGet_(r, read.idx, "fechaHora") || "").trim(),
      monto: _prNum_(_prGet_(r, read.idx, "monto")),
      metodo: String(_prGet_(r, read.idx, "metodo") || "").trim(),
      referencia: String(_prGet_(r, read.idx, "referencia") || "").trim(),
      nota: String(_prGet_(r, read.idx, "nota") || "").trim(),
      aplicadoA: String(_prGet_(r, read.idx, "aplicadoA") || "").trim(),
      moratorioCobrado: _prNum_(_prGet_(r, read.idx, "moratorioCobrado")),
      interesCobrado: _prNum_(_prGet_(r, read.idx, "interesCobrado")),
      capitalCobrado: _prNum_(_prGet_(r, read.idx, "capitalCobrado")),
      saldoFavor: _prNum_(_prGet_(r, read.idx, "saldoFavor"))
    });
  });

  // ordenar asc para replay (y desc para UI si se requiere)
  out.sort((a, b) => String(a.fechaHora || "").localeCompare(String(b.fechaHora || "")));
  return out;
}

/**
 * Acumula mora hasta una fecha (dateOnly), SIN duplicar: usa moraCalculadaHasta.
 */
function _prAccrueMoraHasta_(c, asOfDateOnly, dailyRate) {
  const due = _prDateOnly_(_prDateFromIso_(c.fechaPago));
  if (!due || !asOfDateOnly) return;

  const basePend = _prRound2_(Math.max(0, (c.interesCorriente - c.interesPagado)) + Math.max(0, (c.capital - c.capitalPagado)));
  if (dailyRate <= 0 || basePend <= 0) return;

  // solo si está vencida al asOf
  if (asOfDateOnly.getTime() <= due.getTime()) return;

  const calcHasta = _prDateOnly_(_prDateFromIso_(c.moraCalculadaHasta)) || due;
  const last = (calcHasta.getTime() < due.getTime()) ? due : calcHasta;

  const desde = (asOfDateOnly.getTime() > last.getTime()) ? last : due;
  const dias = _prDiffDays_(asOfDateOnly, desde);
  if (dias <= 0) return;

  const moraExtra = _prRound2_(basePend * dailyRate * dias);
  c.moraAcumulada = _prRound2_(_prNum_(c.moraAcumulada) + moraExtra);
  c.moraCalculadaHasta = _prIso_(asOfDateOnly);
}

/**
 * Aplica un pago a cuotas (mutando el arreglo), siguiendo regla:
 *  mora -> interés -> capital, en orden de cuotas.
 */
function _prApplyPayment_(cuotasState, prestamo, montoPago, fechaPagoDateOnly) {
  let rem = _prRound2_(Math.max(0, _prNum_(montoPago)));
  const dailyRate = (_prNum_(prestamo.tasaMoratoriaMensual) || 0) / 30;

  let sumMora = 0;
  let sumInt = 0;
  let sumCap = 0;
  const aplicado = [];

  const pushDet = (c, mora, interes, capital) => {
    if (!(mora || interes || capital)) return;
    aplicado.push({
      idCuota: c.idCuota,
      nroCuota: c.nroCuota,
      mora: _prRound2_(mora),
      interes: _prRound2_(interes),
      capital: _prRound2_(capital)
    });
  };

  for (let i = 0; i < cuotasState.length; i++) {
    if (rem <= 0) break;
    const c = cuotasState[i];

    // 1) actualizar mora acumulada hasta fecha de pago
    _prAccrueMoraHasta_(c, fechaPagoDateOnly, dailyRate);

    // pendientes actuales a esa fecha
    const interesPend = _prRound2_(Math.max(0, c.interesCorriente - c.interesPagado));
    const capitalPend = _prRound2_(Math.max(0, c.capital - c.capitalPagado));
    const moraPend = _prRound2_(Math.max(0, _prNum_(c.moraAcumulada) - _prNum_(c.moraPagada)));
    const totalPend = _prRound2_(moraPend + interesPend + capitalPend);

    if (totalPend <= 0) continue;

    let pagaMora = 0, pagaInt = 0, pagaCap = 0;

    // mora
    if (rem > 0 && moraPend > 0) {
      pagaMora = Math.min(rem, moraPend);
      c.moraPagada = _prRound2_(_prNum_(c.moraPagada) + pagaMora);
      rem = _prRound2_(rem - pagaMora);
    }

    // interés
    if (rem > 0 && interesPend > 0) {
      pagaInt = Math.min(rem, interesPend);
      c.interesPagado = _prRound2_(_prNum_(c.interesPagado) + pagaInt);
      rem = _prRound2_(rem - pagaInt);
    }

    // capital
    if (rem > 0 && capitalPend > 0) {
      pagaCap = Math.min(rem, capitalPend);
      c.capitalPagado = _prRound2_(_prNum_(c.capitalPagado) + pagaCap);
      rem = _prRound2_(rem - pagaCap);
    }

    sumMora = _prRound2_(sumMora + pagaMora);
    sumInt = _prRound2_(sumInt + pagaInt);
    sumCap = _prRound2_(sumCap + pagaCap);

    pushDet(c, pagaMora, pagaInt, pagaCap);
  }

  return {
    moratorioCobrado: sumMora,
    interesCobrado: sumInt,
    capitalCobrado: sumCap,
    saldoFavor: _prRound2_(Math.max(0, rem)),
    aplicado
  };
}

/**
 * Replay de pagos históricos para reconstruir estado de cuotas (sin guardar en Sheets)
 */
function _prReplayPagos_(cuotasBase, prestamo, pagosAsc) {
  const cuotasState = JSON.parse(JSON.stringify(cuotasBase));
  pagosAsc.forEach(p => {
    const dt = _prDateOnly_(new Date(String(p.fechaHora || "").replace("T", " "))) || _prDateOnly_(new Date());
    _prApplyPayment_(cuotasState, prestamo, p.monto, dt);
  });
  return cuotasState;
}

/**
 * Calcula "al día" (mora extra hasta hoy, pendientes, estado, resumen)
 */
function _prComputeAlDia_(cuotasState, prestamo, todayDateOnly) {
  const today = todayDateOnly || _prDateOnly_(new Date());
  const dailyRate = (_prNum_(prestamo.tasaMoratoriaMensual) || 0) / 30;

  let totalMoraPendiente = 0;
  let totalBasePendiente = 0;
  let saldoInsoluto = 0;
  let cuotasVencidas = 0;
  let montoVencido = 0;
  const upcoming = [];

  const cuotasAlDia = cuotasState.map(c => {
    const interesPend = _prRound2_(Math.max(0, (c.interesCorriente || 0) - (c.interesPagado || 0)));
    const capitalPend = _prRound2_(Math.max(0, (c.capital || 0) - (c.capitalPagado || 0)));
    const basePend = _prRound2_(interesPend + capitalPend);

    const moraAcc = _prNum_(c.moraAcumulada);
    const moraPag = _prNum_(c.moraPagada);

    const due = _prDateOnly_(_prDateFromIso_(c.fechaPago));
    const calcHasta = _prDateOnly_(_prDateFromIso_(c.moraCalculadaHasta)) || due;
    const last = (due && calcHasta && calcHasta.getTime() < due.getTime()) ? due : calcHasta;

    let moraExtra = 0;
    if (dailyRate > 0 && basePend > 0 && due && today && today.getTime() > due.getTime()) {
      const desde = (last && today.getTime() > last.getTime()) ? last : due;
      const dias = _prDiffDays_(today, desde);
      if (dias > 0) moraExtra = _prRound2_(basePend * dailyRate * dias);
    }

    const moraAlDia = _prRound2_(moraAcc + moraExtra);
    const moraPend = _prRound2_(Math.max(0, moraAlDia - moraPag));
    const totalPend = _prRound2_(basePend + moraPend);

    const estadoAlDia = _prCuotaEstadoAuto_(totalPend, due, today);

    // agregados
    saldoInsoluto = _prRound2_(saldoInsoluto + capitalPend);
    if (estadoAlDia === "Vencida") {
      cuotasVencidas += 1;
      montoVencido = _prRound2_(montoVencido + totalPend);
    }

    // próximas (>= hoy) y no pagadas
    if (estadoAlDia !== "Pagada" && due && today && due.getTime() >= today.getTime()) {
      upcoming.push({
        nroCuota: c.nroCuota,
        fechaPago: c.fechaPago,
        pendiente: totalPend,
        moraPendiente: moraPend,
        interesPendiente: interesPend,
        capitalPendiente: capitalPend
      });
    }

    totalMoraPendiente = _prRound2_(totalMoraPendiente + moraPend);
    totalBasePendiente = _prRound2_(totalBasePendiente + basePend);

    return {
      ...c,
      interesPendiente: interesPend,
      capitalPendiente: capitalPend,
      moraPendiente: moraPend,
      totalPendiente: totalPend,
      estado: estadoAlDia
    };
  });

  upcoming.sort((a, b) =>
    String(a.fechaPago || "").localeCompare(String(b.fechaPago || "")) || (a.nroCuota - b.nroCuota)
  );

  const proximasCuotas = upcoming.slice(0, 3);
  const proxima = proximasCuotas[0] || null;

  const resumen = {
    basePendiente: totalBasePendiente,
    moraPendiente: totalMoraPendiente,
    totalPendiente: _prRound2_(totalBasePendiente + totalMoraPendiente),
    saldoInsoluto: _prRound2_(saldoInsoluto),
    cuotasVencidas,
    montoVencido: _prRound2_(montoVencido),
    proximaCuotaFecha: proxima ? proxima.fechaPago : "",
    proximaCuotaMonto: proxima ? proxima.pendiente : 0,
    proximasCuotas
  };

  return { cuotasAlDia, resumen };
}

// =========================
// Endpoints (Frontend)
// =========================

function listarPrestamosLegacy_(params) {
  params = params || {};
  _prEnsureSchema_();

  const sh = obtenerSheet_(_prSheetName_("SH_PRESTAMOS", "PRESTAMOS"));
  const main = _prRead_(sh);

  const personaId = String(params.personaId || "").trim();
  const origenParamRaw = String(params.origen || "").trim();
  const origenParam = (origenParamRaw ? origenParamRaw.toUpperCase() : "SOLICITADO");
  const estado = String(params.estado || "").trim().toLowerCase();
  const estadoOperativo = (estado === "activo" || estado === "finalizado");
  const q = String(params.q || "").trim().toLowerCase();
  const from = _prIso_(params.fechaFrom);
  const to = _prIso_(params.fechaTo);

  // Leer todos los pagos una sola vez y agrupar por préstamo (para performance)
  const shPagos = obtenerSheet_(_prSheetName_("SH_PRESTAMOS_PAGOS", "Prestamos_Pagos"));
  const pagosRead = _prRead_(shPagos);
  const pagosPorPrestamo = {};
  pagosRead.rows.forEach(r => {
    const pid = String(_prGet_(r, pagosRead.idx, "idPrestamo") || "").trim();
    if (!pid) return;
    if (!pagosPorPrestamo[pid]) pagosPorPrestamo[pid] = [];
    pagosPorPrestamo[pid].push({
      fechaHora: String(_prGet_(r, pagosRead.idx, "fechaHora") || "").trim(),
      monto: _prNum_(_prGet_(r, pagosRead.idx, "monto"))
    });
  });
  Object.keys(pagosPorPrestamo).forEach(pid => {
    pagosPorPrestamo[pid].sort((a, b) => String(a.fechaHora || "").localeCompare(String(b.fechaHora || "")));
  });

  const today = _prDateOnly_(new Date());
  const out = [];

  main.rows.forEach(r => {
    const idPrestamo = String(_prGet_(r, main.idx, "idPrestamo") || "").trim();
    if (!idPrestamo) return;

    const origenRow = String(_prGet_(r, main.idx, "origen") || "").trim().toUpperCase() || "SOLICITADO";
    if (origenParam && origenRow !== origenParam) return;

    const idPersona = String(_prGet_(r, main.idx, "idPersona") || "").trim();
    if (personaId && idPersona !== personaId) return;

    const personaNombre = String(_prGet_(r, main.idx, "personaNombre") || "").trim();
    if (q) {
      const hay = `${idPrestamo} ${personaNombre} ${idPersona}`.toLowerCase();
      if (hay.indexOf(q) === -1) return;
    }

    const fechaDesembolso = _prIso_(_prGet_(r, main.idx, "fechaDesembolso"));
    if (from && fechaDesembolso && fechaDesembolso < from) return;
    if (to && fechaDesembolso && fechaDesembolso > to) return;

    const prestamo = {
      idPrestamo,
      origen: origenRow,
      tipoCliente: String(_prGet_(r, main.idx, "tipoCliente") || "").trim() || "PERSONA",
      idPersona,
      personaNombre,
      montoPrincipal: _prNum_(_prGet_(r, main.idx, "montoPrincipal")),
      plazoMeses: parseInt(String(_prGet_(r, main.idx, "plazoMeses") || "0"), 10) || 0,
      tasaMensual: _prNum_(_prGet_(r, main.idx, "tasaMensual")),
      moratorioModo: String(_prGet_(r, main.idx, "moratorioModo") || "").trim(),
      tasaMoratoriaMensual: _prNum_(_prGet_(r, main.idx, "tasaMoratoriaMensual")),
      adminTipo: String(_prGet_(r, main.idx, "adminTipo") || "").trim(),
      adminValor: _prNum_(_prGet_(r, main.idx, "adminValor")),
      adminMonto: _prNum_(_prGet_(r, main.idx, "adminMonto")),
      fechaDesembolso,
      fechaHoraDesembolso: String(_prGet_(r, main.idx, "fechaHoraDesembolso") || "").trim(),
      metodoDesembolso: String(_prGet_(r, main.idx, "metodoDesembolso") || "").trim(),
      refDesembolso: String(_prGet_(r, main.idx, "refDesembolso") || "").trim(),
      notaDesembolso: String(_prGet_(r, main.idx, "notaDesembolso") || "").trim(),
      fechaPrimerPago: _prIso_(_prGet_(r, main.idx, "fechaPrimerPago")),
      diaCorte: parseInt(String(_prGet_(r, main.idx, "diaCorte") || ""), 10) || "",
      diaPago: parseInt(String(_prGet_(r, main.idx, "diaPago") || ""), 10) || "",
      cuotaMensual: _prNum_(_prGet_(r, main.idx, "cuotaMensual")),
      totalInteresEstimado: _prNum_(_prGet_(r, main.idx, "totalInteresEstimado")),
      totalPagarEstimado: _prNum_(_prGet_(r, main.idx, "totalPagarEstimado")),
      estado: String(_prGet_(r, main.idx, "estado") || "").trim() || "Borrador",
      fechaCreacion: String(_prGet_(r, main.idx, "fechaCreacion") || "").trim(),
      fechaActualizacion: String(_prGet_(r, main.idx, "fechaActualizacion") || "").trim()
    };

    // cálculo al vuelo (sin cuotas sheet)
    const base = _prBuildCuotasBaseFromPrestamo_(prestamo);
    const pagos = pagosPorPrestamo[idPrestamo] || [];
    const cuotasState = _prReplayPagos_(base.cuotas, prestamo, pagos);
    const calc = _prComputeAlDia_(cuotasState, prestamo, today);

    prestamo.totalPendiente = calc.resumen.totalPendiente;
    prestamo.saldoInsoluto = calc.resumen.saldoInsoluto;
    prestamo.cuotasVencidas = calc.resumen.cuotasVencidas;
    prestamo.montoVencido = calc.resumen.montoVencido;
    prestamo.proximaCuotaFecha = calc.resumen.proximaCuotaFecha;
    prestamo.proximaCuotaMonto = calc.resumen.proximaCuotaMonto;

    const estadoSistema = _prLoanEstadoAuto_(prestamo.totalPendiente);
    prestamo.estadoSistema = estadoSistema;

    // filtro operativo (Activo/Finalizado)
    if (estadoOperativo && estado) {
      const st = String(estadoSistema || prestamo.estado || "").trim().toLowerCase();
      if (st !== estado) return;
    }

    out.push(prestamo);
  });

  // Orden: fecha desc, luego id
  out.sort((a, b) => {
    const fa = String(a.fechaDesembolso || "");
    const fb = String(b.fechaDesembolso || "");
    if (fa !== fb) return fb.localeCompare(fa);
    return String(b.idPrestamo || "").localeCompare(String(a.idPrestamo || ""));
  });

  return out;
}

function guardarPrestamoLegacy_(payload) {
  payload = payload || {};
  _prEnsureSchema_();

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const sh = obtenerSheet_(_prSheetName_("SH_PRESTAMOS", "PRESTAMOS"));
    const main = _prRead_(sh);

    const idPersona = String(payload.idPersona || "").trim();
    const montoPrincipal = _prNum_(payload.montoPrincipal);
    const plazoMeses = Math.max(1, parseInt(payload.plazoMeses, 10) || 0);
    const tasaMensual = Number(payload.tasaMensual || 0);

    const origen = String(payload.origen || "SOLICITADO").trim().toUpperCase() || "SOLICITADO";
    const tipoCliente = String(payload.tipoCliente || "PERSONA").trim() || "PERSONA";

    const moratorioModo = String(payload.moratorioModo || "25").trim(); // compat
    let tasaMoratoriaMensual = _prNum_(payload.tasaMoratoriaMensual);
    if (!tasaMoratoriaMensual) {
      if (moratorioModo === "manual") {
        tasaMoratoriaMensual = _prNum_(payload.tasaMoratoriaMensual);
      } else if (moratorioModo === "50") {
        tasaMoratoriaMensual = 0.50;
      } else {
        // default 25%
        tasaMoratoriaMensual = 0.25;
      }
    }

    const adminTipo = String(payload.adminTipo || "").trim();
    const adminValor = _prNum_(payload.adminValor);
    const adminMonto = _prNum_(payload.adminMonto);

    const diaCorte = parseInt(payload.diaCorte, 10) || "";
    const diaPago = parseInt(payload.diaPago, 10) || "";

    const estado = String(payload.estado || (origen === "OTORGADO" ? "Activo" : "Solicitado")).trim() || "Borrador";

    const personaNombre = String(payload.personaNombre || _prLookupPersonaNombre_(idPersona) || "").trim();

    const now = _prNowIso_();
    const idPrestamo = String(payload.idPrestamo || "").trim() || Utilities.getUuid();

    // construir amortización base (solo para guardar datos resumen, NO para guardar filas)
    const tmpPrest = {
      idPrestamo,
      montoPrincipal,
      plazoMeses,
      tasaMensual,
      diaPago: diaPago || 1,
      fechaDesembolso: _prIso_(payload.fechaDesembolso || now),
      fechaPrimerPago: _prIso_(payload.fechaPrimerPago || "")
    };
    const amort = _prBuildCuotasBaseFromPrestamo_(tmpPrest);
    const cuotaMensual = amort.cuotaMensual;
    const totalInteresEstimado = amort.totalInteres;
    const totalPagarEstimado = _prRound2_(montoPrincipal + totalInteresEstimado);

    const colCount = main.headers.length;
    const setCell = (row, name, val) => {
      const i = main.idx[name];
      if (i === undefined) return;
      row[i] = val;
    };

    const rowIndex = _prFindRowIndexById_(main.rows, main.idx, idPrestamo);

    if (rowIndex >= 0) {
      const row = main.rows[rowIndex].slice(0);
      setCell(row, "origen", origen);
      setCell(row, "tipoCliente", tipoCliente);
      setCell(row, "idPersona", idPersona);
      setCell(row, "personaNombre", personaNombre);
      setCell(row, "montoPrincipal", montoPrincipal);
      setCell(row, "plazoMeses", plazoMeses);
      setCell(row, "tasaMensual", tasaMensual);
      setCell(row, "moratorioModo", moratorioModo);
      setCell(row, "tasaMoratoriaMensual", tasaMoratoriaMensual);
      setCell(row, "adminTipo", adminTipo);
      setCell(row, "adminValor", adminValor);
      setCell(row, "adminMonto", adminMonto);

      // mantener desembolso si ya existía; si viene en payload se actualiza
      if (payload.fechaDesembolso) setCell(row, "fechaDesembolso", _prIso_(payload.fechaDesembolso));
      if (payload.fechaHoraDesembolso) setCell(row, "fechaHoraDesembolso", String(payload.fechaHoraDesembolso || ""));
      if (payload.metodoDesembolso) setCell(row, "metodoDesembolso", String(payload.metodoDesembolso || ""));
      if (payload.refDesembolso) setCell(row, "refDesembolso", String(payload.refDesembolso || ""));
      if (payload.notaDesembolso) setCell(row, "notaDesembolso", String(payload.notaDesembolso || ""));

      setCell(row, "fechaPrimerPago", amort.fechaPrimerPago);
      setCell(row, "diaCorte", diaCorte);
      setCell(row, "diaPago", diaPago);
      setCell(row, "cuotaMensual", cuotaMensual);
      setCell(row, "totalInteresEstimado", totalInteresEstimado);
      setCell(row, "totalPagarEstimado", totalPagarEstimado);
      setCell(row, "estado", estado);
      setCell(row, "actualizadoPor", String(payload.actualizadoPor || payload.user || ""));
      setCell(row, "fechaActualizacion", now);

      sh.getRange(rowIndex + 2, 1, 1, colCount).setValues([row]);
    } else {
      const row = new Array(colCount).fill("");
      setCell(row, "idPrestamo", idPrestamo);
      setCell(row, "origen", origen);
      setCell(row, "tipoCliente", tipoCliente);
      setCell(row, "idPersona", idPersona);
      setCell(row, "personaNombre", personaNombre);
      setCell(row, "montoPrincipal", montoPrincipal);
      setCell(row, "plazoMeses", plazoMeses);
      setCell(row, "tasaMensual", tasaMensual);
      setCell(row, "moratorioModo", moratorioModo);
      setCell(row, "tasaMoratoriaMensual", tasaMoratoriaMensual);
      setCell(row, "adminTipo", adminTipo);
      setCell(row, "adminValor", adminValor);
      setCell(row, "adminMonto", adminMonto);

      setCell(row, "fechaDesembolso", _prIso_(payload.fechaDesembolso || ""));
      setCell(row, "fechaHoraDesembolso", String(payload.fechaHoraDesembolso || ""));
      setCell(row, "metodoDesembolso", String(payload.metodoDesembolso || ""));
      setCell(row, "refDesembolso", String(payload.refDesembolso || ""));
      setCell(row, "notaDesembolso", String(payload.notaDesembolso || ""));

      setCell(row, "fechaPrimerPago", amort.fechaPrimerPago);
      setCell(row, "diaCorte", diaCorte);
      setCell(row, "diaPago", diaPago);
      setCell(row, "cuotaMensual", cuotaMensual);
      setCell(row, "totalInteresEstimado", totalInteresEstimado);
      setCell(row, "totalPagarEstimado", totalPagarEstimado);
      setCell(row, "estado", estado);

      setCell(row, "creadoPor", String(payload.creadoPor || payload.user || ""));
      setCell(row, "fechaCreacion", now);

      sh.appendRow(row);
    }

    // YA NO se crean cuotas en PRESTAMOS_CUOTAS ✅
    return { ok: true, idPrestamo, cuotaMensual, fechaPrimerPago: amort.fechaPrimerPago };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function eliminarPrestamoLegacy_(idPrestamo) {
  _prEnsureSchema_();
  const id = String(idPrestamo || "").trim();
  if (!id) return { ok: false, error: "idPrestamo requerido" };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    // borrar cabecera
    const sh = obtenerSheet_(_prSheetName_("SH_PRESTAMOS", "PRESTAMOS"));
    const main = _prRead_(sh);
    const idx = _prFindRowIndexById_(main.rows, main.idx, id);
    if (idx >= 0) sh.deleteRow(idx + 2);

    // borrar pagos
    const shPagos = obtenerSheet_(_prSheetName_("SH_PRESTAMOS_PAGOS", "Prestamos_Pagos"));
    const pagos = shPagos.getDataRange().getDisplayValues();
    if (pagos && pagos.length > 1) {
      const headers = (pagos[0] || []).map(h => String(h || "").trim());
      const map = {};
      headers.forEach((h, i) => { if (h) map[h] = i; });
      const colPrest = map["idPrestamo"];
      if (colPrest !== undefined) {
        for (let r = pagos.length - 1; r >= 1; r--) {
          if (String(pagos[r][colPrest] || "").trim() === id) shPagos.deleteRow(r + 1);
        }
      }
    }

    // opcional: limpiar cuotas legacy si existe la hoja (pero NO se usa)
    try {
      const shCuotas = obtenerSheet_(_prSheetName_("SH_PRESTAMOS_CUOTAS", "PRESTAMOS_CUOTAS"));
      const all = shCuotas.getDataRange().getDisplayValues();
      if (all && all.length > 1) {
        const headers = (all[0] || []).map(h => String(h || "").trim());
        const map = {};
        headers.forEach((h, i) => { if (h) map[h] = i; });
        const colPrest = map["idPrestamo"];
        if (colPrest !== undefined) {
          for (let r = all.length - 1; r >= 1; r--) {
            if (String(all[r][colPrest] || "").trim() === id) shCuotas.deleteRow(r + 1);
          }
        }
      }
    } catch (e) {}

    return { ok: true };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function obtenerDetallePrestamoLegacy_(idPrestamo) {
  _prEnsureSchema_();
  const id = String(idPrestamo || "").trim();
  if (!id) return { prestamo: null, cuotas: [], pagos: [], resumen: null };

  const sh = obtenerSheet_(_prSheetName_("SH_PRESTAMOS", "PRESTAMOS"));
  const main = _prRead_(sh);
  const rowIndex = _prFindRowIndexById_(main.rows, main.idx, id);
  if (rowIndex < 0) return { prestamo: null, cuotas: [], pagos: [], resumen: null };

  const r = main.rows[rowIndex];
  const prestamo = {
    idPrestamo: String(_prGet_(r, main.idx, "idPrestamo") || "").trim(),
    origen: String(_prGet_(r, main.idx, "origen") || "").trim().toUpperCase() || "SOLICITADO",
    tipoCliente: String(_prGet_(r, main.idx, "tipoCliente") || "").trim() || "PERSONA",
    idPersona: String(_prGet_(r, main.idx, "idPersona") || "").trim(),
    personaNombre: String(_prGet_(r, main.idx, "personaNombre") || "").trim(),
    montoPrincipal: _prNum_(_prGet_(r, main.idx, "montoPrincipal")),
    plazoMeses: parseInt(String(_prGet_(r, main.idx, "plazoMeses") || "0"), 10) || 0,
    tasaMensual: _prNum_(_prGet_(r, main.idx, "tasaMensual")),
    moratorioModo: String(_prGet_(r, main.idx, "moratorioModo") || "").trim(),
    tasaMoratoriaMensual: _prNum_(_prGet_(r, main.idx, "tasaMoratoriaMensual")),
    adminTipo: String(_prGet_(r, main.idx, "adminTipo") || "").trim(),
    adminValor: _prNum_(_prGet_(r, main.idx, "adminValor")),
    adminMonto: _prNum_(_prGet_(r, main.idx, "adminMonto")),
    fechaDesembolso: _prIso_(_prGet_(r, main.idx, "fechaDesembolso")),
    fechaHoraDesembolso: String(_prGet_(r, main.idx, "fechaHoraDesembolso") || "").trim(),
    metodoDesembolso: String(_prGet_(r, main.idx, "metodoDesembolso") || "").trim(),
    refDesembolso: String(_prGet_(r, main.idx, "refDesembolso") || "").trim(),
    notaDesembolso: String(_prGet_(r, main.idx, "notaDesembolso") || "").trim(),
    fechaPrimerPago: _prIso_(_prGet_(r, main.idx, "fechaPrimerPago")),
    diaCorte: parseInt(String(_prGet_(r, main.idx, "diaCorte") || ""), 10) || "",
    diaPago: parseInt(String(_prGet_(r, main.idx, "diaPago") || ""), 10) || "",
    cuotaMensual: _prNum_(_prGet_(r, main.idx, "cuotaMensual")),
    totalInteresEstimado: _prNum_(_prGet_(r, main.idx, "totalInteresEstimado")),
    totalPagarEstimado: _prNum_(_prGet_(r, main.idx, "totalPagarEstimado")),
    estado: String(_prGet_(r, main.idx, "estado") || "").trim() || "Borrador",
    fechaCreacion: String(_prGet_(r, main.idx, "fechaCreacion") || "").trim(),
    fechaActualizacion: String(_prGet_(r, main.idx, "fechaActualizacion") || "").trim()
  };

  // pagos (UI los quiere desc)
  const pagosAsc = _prReadPagosByPrestamo_(id);
  const pagosDesc = pagosAsc.slice().sort((a, b) => String(b.fechaHora || "").localeCompare(String(a.fechaHora || "")));

  // cuotas base + replay pagos + calcular al día
  const base = _prBuildCuotasBaseFromPrestamo_(prestamo);
  const cuotasState = _prReplayPagos_(base.cuotas, prestamo, pagosAsc);
  const calc = _prComputeAlDia_(cuotasState, prestamo, _prDateOnly_(new Date()));

  const resumen = calc.resumen;
  const estadoSistema = _prLoanEstadoAuto_(resumen.totalPendiente);
  prestamo.estadoSistema = estadoSistema;

  // auto actualiza estado (solo si estaba en operativo)
  try {
    const estActual = String(prestamo.estado || "").trim();
    const operativos = { "Activo": 1, "Finalizado": 1, "Aprobado": 1, "": 1 };
    if (operativos[estActual] && estActual !== estadoSistema) {
      const colEstado = main.idx["estado"];
      const colFechaAct = main.idx["fechaActualizacion"];
      if (colEstado !== undefined) {
        const rowToWrite = rowIndex + 2;
        sh.getRange(rowToWrite, colEstado + 1).setValue(estadoSistema);
        if (colFechaAct !== undefined) sh.getRange(rowToWrite, colFechaAct + 1).setValue(_prNowIso_());
      }
      prestamo.estado = estadoSistema;
    }
  } catch (e) {}

  return { prestamo, cuotas: calc.cuotasAlDia, pagos: pagosDesc, resumen };
}

function listarPagosPrestamoLegacy_(idPrestamo) {
  _prEnsureSchema_();
  const id = String(idPrestamo || "").trim();
  if (!id) return [];
  const pagosAsc = _prReadPagosByPrestamo_(id);
  return pagosAsc.slice().sort((a, b) => String(b.fechaHora || "").localeCompare(String(a.fechaHora || "")));
}

/**
 * Simula pago (preview) o registra pago real según el endpoint.
 * Frontend usa:
 *  - getPagoSugeridoLegacy_ (preview)
 *  - registrarPagoPrestamoLegacy_ (guardar)
 */
function getPagoSugeridoLegacy_(payload) {
  return getPagoSugeridoPrestamo_(payload);
}

function getPagoSugeridoPrestamo_(payload) {
  payload = payload || {};
  _prEnsureSchema_();

  const idPrestamo = String(payload.idPrestamo || "").trim();
  if (!idPrestamo) throw new Error("idPrestamo requerido");

  // modo
  const modoRaw = String(payload.modo || payload.mode || "PROXIMA_CUOTA").trim().toUpperCase();
  const modoMap = {
    "PROXIMA": "PROXIMA_CUOTA",
    "PROXIMA_CUOTA": "PROXIMA_CUOTA",
    "LIQUIDAR": "LIQUIDAR_HOY",
    "LIQUIDAR_HOY": "LIQUIDAR_HOY",
    "MONTO": "MONTO_LIBRE",
    "MONTO_LIBRE": "MONTO_LIBRE"
  };
  const modo = modoMap[modoRaw] || "PROXIMA_CUOTA";

  // fecha asOf (si viene)
  let asOf = _prDateOnly_(new Date());
  const fp = String(payload.fechaPago || payload.fechaHora || "").trim();
  if (fp) {
    const dt = new Date(fp);
    if (!isNaN(dt.getTime())) asOf = _prDateOnly_(dt);
  }

  // cargar préstamo
  const sh = obtenerSheet_(_prSheetName_("SH_PRESTAMOS", "PRESTAMOS"));
  const main = _prRead_(sh);
  const rowIndex = _prFindRowIndexById_(main.rows, main.idx, idPrestamo);
  if (rowIndex < 0) throw new Error("Préstamo no encontrado");

  const rr = main.rows[rowIndex];
  const prestamo = {
    idPrestamo,
    montoPrincipal: _prNum_(_prGet_(rr, main.idx, "montoPrincipal")),
    plazoMeses: parseInt(String(_prGet_(rr, main.idx, "plazoMeses") || "0"), 10) || 0,
    tasaMensual: _prNum_(_prGet_(rr, main.idx, "tasaMensual")),
    tasaMoratoriaMensual: _prNum_(_prGet_(rr, main.idx, "tasaMoratoriaMensual")),
    fechaDesembolso: _prIso_(_prGet_(rr, main.idx, "fechaDesembolso")),
    fechaPrimerPago: _prIso_(_prGet_(rr, main.idx, "fechaPrimerPago")),
    diaPago: parseInt(String(_prGet_(rr, main.idx, "diaPago") || ""), 10) || 1
  };

  const pagosAsc = _prReadPagosByPrestamo_(idPrestamo);
  const base = _prBuildCuotasBaseFromPrestamo_(prestamo);
  const cuotasState = _prReplayPagos_(base.cuotas, prestamo, pagosAsc);
  const calc = _prComputeAlDia_(cuotasState, prestamo, asOf);

  // monto sugerido
  let montoSugerido = 0;
  if (modo === "LIQUIDAR_HOY") {
    montoSugerido = calc.resumen.totalPendiente;
  } else if (modo === "PROXIMA_CUOTA") {
    const proxima = (calc.resumen.proximasCuotas && calc.resumen.proximasCuotas.length)
      ? calc.resumen.proximasCuotas[0]
      : null;

    if (proxima) montoSugerido = proxima.pendiente;
    else {
      // si no hay próxima futura, pagar la primera no pagada (aunque esté vencida)
      const first = calc.cuotasAlDia.find(c => String(c.estado) !== "Pagada");
      montoSugerido = first ? first.totalPendiente : 0;
    }
  } else {
    // monto libre
    montoSugerido = _prNum_(payload.monto || payload.amount || 0);
  }

  montoSugerido = _prRound2_(Math.max(0, montoSugerido));

  // preview aplicación (sin guardar)
  const cuotasSim = JSON.parse(JSON.stringify(cuotasState));
  const previewRes = _prApplyPayment_(cuotasSim, prestamo, montoSugerido, asOf);

  return {
    ok: true,
    idPrestamo,
    modo,
    montoSugerido,
    suggestedAmount: montoSugerido, // alias por si el front usa otro nombre
    preview: {
      moratorioCobrado: previewRes.moratorioCobrado,
      interesCobrado: previewRes.interesCobrado,
      capitalCobrado: previewRes.capitalCobrado,
      saldoFavor: previewRes.saldoFavor
    },
    aplicado: previewRes.aplicado,
    resumen: calc.resumen
  };
}

function registrarPagoPrestamoLegacy_(payload) {
  payload = payload || {};
  _prEnsureSchema_();

  const idPrestamo = String(payload.idPrestamo || "").trim();
  if (!idPrestamo) throw new Error("idPrestamo requerido");

  const monto = _prRound2_(Math.max(0, _prNum_(payload.monto || payload.amount)));
  if (monto <= 0) throw new Error("Monto inválido");

  const metodo = String(payload.metodo || "EFECTIVO").trim();
  const referencia = String(payload.referencia || "").trim();
  const nota = String(payload.nota || "").trim();

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    // cargar préstamo
    const sh = obtenerSheet_(_prSheetName_("SH_PRESTAMOS", "PRESTAMOS"));
    const main = _prRead_(sh);
    const rowIndex = _prFindRowIndexById_(main.rows, main.idx, idPrestamo);
    if (rowIndex < 0) throw new Error("Préstamo no encontrado");

    const rr = main.rows[rowIndex];
    const prestamo = {
      idPrestamo,
      montoPrincipal: _prNum_(_prGet_(rr, main.idx, "montoPrincipal")),
      plazoMeses: parseInt(String(_prGet_(rr, main.idx, "plazoMeses") || "0"), 10) || 0,
      tasaMensual: _prNum_(_prGet_(rr, main.idx, "tasaMensual")),
      tasaMoratoriaMensual: _prNum_(_prGet_(rr, main.idx, "tasaMoratoriaMensual")),
      fechaDesembolso: _prIso_(_prGet_(rr, main.idx, "fechaDesembolso")),
      fechaPrimerPago: _prIso_(_prGet_(rr, main.idx, "fechaPrimerPago")),
      diaPago: parseInt(String(_prGet_(rr, main.idx, "diaPago") || ""), 10) || 1,
      estado: String(_prGet_(rr, main.idx, "estado") || "").trim() || "Activo"
    };

    // estado actual por replay (antes del nuevo pago)
    const pagosAsc = _prReadPagosByPrestamo_(idPrestamo);
    const base = _prBuildCuotasBaseFromPrestamo_(prestamo);
    const cuotasState = _prReplayPagos_(base.cuotas, prestamo, pagosAsc);

    const fechaHora = _prNowIso_();
    const asOf = _prDateOnly_(new Date());

    // aplicar pago real
    const res = _prApplyPayment_(cuotasState, prestamo, monto, asOf);

    // construir aplicadoA
    let aplicadoA = "";
    if (res.aplicado && res.aplicado.length) {
      const nums = res.aplicado.map(x => x.nroCuota).sort((a, b) => a - b);
      const min = nums[0], max = nums[nums.length - 1];
      aplicadoA = (min === max) ? `Cuota ${min}` : `Cuotas ${min}-${max}`;
    } else {
      aplicadoA = "Saldo a favor";
    }

    // guardar pago en ledger
    const shPagos = obtenerSheet_(_prSheetName_("SH_PRESTAMOS_PAGOS", "Prestamos_Pagos"));
    const readPagos = _prRead_(shPagos);
    const headers = readPagos.headers;
    const idx = readPagos.idx;

    const row = new Array(headers.length).fill("");
    const set = (name, val) => {
      const i = idx[name];
      if (i === undefined) return;
      row[i] = val;
    };

    const idPago = Utilities.getUuid();
    set("idPago", idPago);
    set("idPrestamo", idPrestamo);
    set("fechaHora", fechaHora);
    set("monto", monto);
    set("metodo", metodo);
    set("referencia", referencia);
    set("nota", nota);
    set("aplicadoA", aplicadoA);
    set("moratorioCobrado", res.moratorioCobrado);
    set("interesCobrado", res.interesCobrado);
    set("capitalCobrado", res.capitalCobrado);
    set("saldoFavor", res.saldoFavor);

    shPagos.appendRow(row);

    // calcular estado del préstamo y actualizar si procede
    const calcHoy = _prComputeAlDia_(cuotasState, prestamo, asOf);
    const estadoSistema = _prLoanEstadoAuto_(calcHoy.resumen.totalPendiente);

    try {
      const estActual = String(prestamo.estado || "").trim();
      const operativos = { "Activo": 1, "Finalizado": 1, "Aprobado": 1, "": 1 };
      if (operativos[estActual] && estActual !== estadoSistema) {
        const colEstado = main.idx["estado"];
        const colFechaAct = main.idx["fechaActualizacion"];
        if (colEstado !== undefined) {
          const rowToWrite = rowIndex + 2;
          sh.getRange(rowToWrite, colEstado + 1).setValue(estadoSistema);
          if (colFechaAct !== undefined) sh.getRange(rowToWrite, colFechaAct + 1).setValue(_prNowIso_());
        }
      }
    } catch (e) {}

    return {
      ok: true,
      idPago,
      idPrestamo,
      aplicadoA,
      preview: {
        moratorioCobrado: res.moratorioCobrado,
        interesCobrado: res.interesCobrado,
        capitalCobrado: res.capitalCobrado,
        saldoFavor: res.saldoFavor
      },
      aplicado: res.aplicado
    };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function registrarDesembolsoLegacy_(payload) {
  return registrarDesembolsoPrestamo_(payload);
}

function registrarDesembolsoPrestamo_(payload) {
  payload = payload || {};
  _prEnsureSchema_();

  const idPrestamo = String(payload.idPrestamo || "").trim();
  if (!idPrestamo) throw new Error("idPrestamo requerido");

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const sh = obtenerSheet_(_prSheetName_("SH_PRESTAMOS", "PRESTAMOS"));
    const main = _prRead_(sh);
    const rowIndex = _prFindRowIndexById_(main.rows, main.idx, idPrestamo);
    if (rowIndex < 0) throw new Error("Préstamo no encontrado");

    const nowIso = _prNowIso_();
    const fechaIso = _prIso_(payload.fechaDesembolso || new Date());
    const fechaHoraIso = String(payload.fechaHoraDesembolso || nowIso).trim();

    const setInRow = (rowArr, name, val) => {
      const i = main.idx[name];
      if (i === undefined) return;
      rowArr[i] = val;
    };

    const row = main.rows[rowIndex].slice(0);

    // set desembolso
    setInRow(row, "origen", "OTORGADO");
    setInRow(row, "fechaDesembolso", fechaIso);
    setInRow(row, "fechaHoraDesembolso", fechaHoraIso);
    setInRow(row, "metodoDesembolso", String(payload.metodoDesembolso || "").trim());
    setInRow(row, "refDesembolso", String(payload.refDesembolso || "").trim());
    setInRow(row, "notaDesembolso", String(payload.notaDesembolso || "").trim());

    // si no hay fechaPrimerPago, se recalcula
    const diaPago = parseInt(String(_prGet_(row, main.idx, "diaPago") || ""), 10) || 1;
    const fpExist = _prIso_(_prGet_(row, main.idx, "fechaPrimerPago"));
    if (!fpExist) {
      const disb = _prDateFromIso_(fechaIso) || new Date();
      let y = disb.getFullYear();
      let m = disb.getMonth() + 1;
      if (m > 11) { m = 0; y += 1; }
      const dtPrimer = _prDateWithDay_(y, m, diaPago);
      setInRow(row, "fechaPrimerPago", _prIso_(dtPrimer));
    }

    // estado
    setInRow(row, "estado", "Activo");
    setInRow(row, "actualizadoPor", String(payload.actualizadoPor || payload.user || ""));
    setInRow(row, "fechaActualizacion", nowIso);

    sh.getRange(rowIndex + 2, 1, 1, main.headers.length).setValues([row]);

    return {
      ok: true,
      idPrestamo,
      estado: "Activo",
      fechaDesembolso: fechaIso,
      fechaHoraDesembolso: fechaHoraIso
    };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}
