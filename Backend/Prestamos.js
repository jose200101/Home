/**
 * =========================
 *  PRÉSTAMOS (Parte 1)
 * =========================
 *
 * Enfoque:
 *  - Registrar solicitudes / préstamos (cabecera)
 *  - Generar automáticamente tabla de amortización (saldo insoluto)
 *  - Vista de detalle (Resumen + Amortización)
 *
 * Hojas (configurables en env_()):
 *  - SH_PRESTAMOS           (default: PRESTAMOS)
 *  - SH_PRESTAMOS_CUOTAS    (default: PRESTAMOS_CUOTAS)
 *  - SH_PRESTAMOS_PAGOS     (default: Prestamos_Pagos)  // se usa en Parte 2
 *  - SH_PARAMETROS          (default: PARAMETROS)
 */

function _prSheetName_(key, fallback) {
  try {
    if (typeof env_ === "function") {
      const env = env_();
      const val = env && env[key];
      if (val) return String(val);
    }
  } catch (e) {}
  return String(fallback);
}

function _prNowIso_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "America/Tegucigalpa", "yyyy-MM-dd'T'HH:mm:ss");
}

function _prRound2_(n) {
  const x = Number(n || 0);
  return Math.round((x + Number.EPSILON) * 100) / 100;
}

// =========================
// Estados automáticos (Parte 3)
// =========================
// Cuota: Pendiente | Vencida | Pagada
// Préstamo: Activo | Finalizado

function _prCuotaEstadoAuto_(totalPendiente, dueDate, asOfDate) {
  const tp = _prRound2_(Number(totalPendiente || 0));
  if (tp <= 0.000001) return "Pagada";
  if (dueDate && asOfDate && asOfDate.getTime() > dueDate.getTime()) return "Vencida";
  return "Pendiente";
}

function _prLoanEstadoAuto_(totalPendiente) {
  const tp = _prRound2_(Number(totalPendiente || 0));
  return (tp <= 0.000001) ? "Finalizado" : "Activo";
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
    return Utilities.formatDate(d, Session.getScriptTimeZone() || "America/Tegucigalpa", "yyyy-MM-dd");
  }
  const s = String(d || "").trim();
  // si ya viene en ISO
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  // intento parse
  const dt = new Date(s);
  if (dt.toString() !== "Invalid Date") return _prIso_(dt);
  return s;
}

function _prDateFromIso_(iso) {
  const s = String(iso || "").trim();
  if (!s) return null;
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) {
    const dt = new Date(s);
    return (dt.toString() === "Invalid Date") ? null : dt;
  }
  const y = parseInt(m[1], 10);
  const mo = parseInt(m[2], 10) - 1;
  const d = parseInt(m[3], 10);
  return new Date(y, mo, d);
}

function _prDateOnly_(d) {
  if (!d) return null;
  const dt = (Object.prototype.toString.call(d) === "[object Date]") ? d : new Date(d);
  if (dt.toString() === "Invalid Date") return null;
  return new Date(dt.getFullYear(), dt.getMonth(), dt.getDate());
}

function _prDiffDays_(later, earlier) {
  const a = _prDateOnly_(later);
  const b = _prDateOnly_(earlier);
  if (!a || !b) return 0;
  const ms = a.getTime() - b.getTime();
  return Math.max(0, Math.floor(ms / 86400000));
}

function _prDaysInMonth_(y, monthIndex0) {
  return new Date(y, monthIndex0 + 1, 0).getDate();
}

function _prDateWithDay_(y, monthIndex0, day) {
  const dd = Math.max(1, Math.min(31, parseInt(day, 10) || 1));
  const max = _prDaysInMonth_(y, monthIndex0);
  return new Date(y, monthIndex0, Math.min(dd, max));
}

function _prEnsureSheet_(name, headers) {
  const ss = conexion();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);

  const lastCol = sh.getLastColumn();
  const headerRange = sh.getRange(1, 1, 1, Math.max(lastCol, headers.length));
  const existing = headerRange.getValues()[0].map(v => String(v || "").trim());

  if (existing.every(v => !v)) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    return sh;
  }

  const lower = existing.map(h => h.toLowerCase());
  const toAdd = headers.filter(h => lower.indexOf(String(h).toLowerCase()) === -1);
  if (toAdd.length) {
    sh.getRange(1, existing.length + 1, 1, toAdd.length).setValues([toAdd]);
  }
  return sh;
}

function _prEnsureSchema_() {
  const shPrestamos = _prSheetName_("SH_PRESTAMOS", "PRESTAMOS");
  const shCuotas = _prSheetName_("SH_PRESTAMOS_CUOTAS", "PRESTAMOS_CUOTAS");
  const shPagos = _prSheetName_("SH_PRESTAMOS_PAGOS", "Prestamos_Pagos");
  const shParams = _prSheetName_("SH_PARAMETROS", "PARAMETROS");

  _prEnsureSheet_(shPrestamos, [
    // Parte 1 (Otorgar): se agrega soporte para "origen" (SOLICITADO/OTORGADO) y desembolso
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

  _prEnsureSheet_(shCuotas, [
    "idCuota",
    "idPrestamo",
    "nroCuota",
    "fechaPago",
    "cuota",
    "interesCorriente",
    "capital",
    "saldoDespues",
    "estado",
    "createdAt",
    // Parte 2: control real (pagos + moratorio)
    "interesPagado",
    "capitalPagado",
    "moraAcumulada",
    "moraPagada",
    "moraCalculadaHasta",
    "updatedAt"
  ]);

  // Parte 2 (se crea desde ya)
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
  if (!values || values.length === 0) return { headers: [], rows: [], idx: {} };

  const headers = (values[0] || []).map(h => String(h || "").trim());
  const idx = {};
  headers.forEach((h, i) => { if (h) idx[h] = i; });
  const rows = (values.length > 1)
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
  const pid = String(idPersona || "").trim();
  if (!pid) return "";
  const env = (typeof env_ === "function") ? env_() : {};
  const shName = env.SH_PERSONAS || "PERSONAS";
  let sh;
  try { sh = obtenerSheet(shName); } catch (e) { return ""; }
  const data = sh.getDataRange().getDisplayValues();
  if (!data || data.length < 2) return "";
  const headers = data[0] || [];
  // usa helpers existentes si están
  const norm = (typeof _normHeader_ === "function") ? _normHeader_ : (s) => String(s||"").trim().toLowerCase();
  const hIdx = {};
  headers.forEach((h, i) => { hIdx[norm(h)] = i; });
  const iId = hIdx[norm("id_persona")] ?? hIdx[norm("id")] ?? hIdx[norm("persona_id")] ?? -1;
  const iNombre = hIdx[norm("nombre_persona")] ?? hIdx[norm("nombre")] ?? hIdx[norm("persona")] ?? -1;
  if (iId < 0 || iNombre < 0) return "";
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (!row || row.every(c => String(c || "").trim() === "")) continue;
    if (String(row[iId] || "").trim() === pid) return String(row[iNombre] || "").trim();
  }
  return "";
}

function _prComputeCuota_(P, r, N) {
  const principal = _prNum_(P);
  const rate = Number(r || 0);
  const n = Math.max(1, parseInt(N, 10) || 1);
  if (principal <= 0) return 0;
  if (!rate || Math.abs(rate) < 1e-12) return principal / n;
  const pow = Math.pow(1 + rate, n);
  return principal * rate * pow / (pow - 1);
}

function _prBuildAmortization_(opts) {
  const P = _prNum_(opts.montoPrincipal);
  const N = Math.max(1, parseInt(opts.plazoMeses, 10) || 1);
  const r = Number(opts.tasaMensual || 0);
  const fechaDesembolso = _prIso_(opts.fechaDesembolso);
  const diaPago = Math.max(1, Math.min(28, parseInt(opts.diaPago, 10) || 1));

  const cuotaRaw = _prComputeCuota_(P, r, N);
  const cuota = _prRound2_(cuotaRaw);

  const disb = _prDateFromIso_(fechaDesembolso) || new Date();
  // Por simplicidad (Parte 1): primer pago = próximo mes en diaPago
  let y = disb.getFullYear();
  let m = disb.getMonth() + 1; // siguiente mes
  if (m > 11) { m = 0; y += 1; }
  let dtPago = _prDateWithDay_(y, m, diaPago);
  const fechaPrimerPago = _prIso_(dtPago);

  const cuotas = [];
  let saldo = P;
  let totalInteres = 0;
  for (let i = 1; i <= N; i++) {
    const interes = _prRound2_(saldo * r);
    let capital = _prRound2_(cuota - interes);
    if (i === N) {
      // Ajuste final por redondeos
      capital = _prRound2_(saldo);
    }
    saldo = _prRound2_(Math.max(0, saldo - capital));
    totalInteres = _prRound2_(totalInteres + interes);

    const fechaPago = _prIso_(dtPago);
    cuotas.push({
      nroCuota: i,
      fechaPago,
      cuota: (i === N) ? _prRound2_(capital + interes) : cuota,
      interesCorriente: interes,
      capital,
      saldoDespues: saldo,
      estado: "Pendiente",
    });

    // siguiente mes
    let yy = dtPago.getFullYear();
    let mm = dtPago.getMonth() + 1;
    if (mm > 11) { mm = 0; yy += 1; }
    dtPago = _prDateWithDay_(yy, mm, diaPago);
  }

  return {
    cuotaMensual: cuota,
    fechaPrimerPago,
    totalInteres,
    cuotas,
  };
}

function _prReplaceCuotas_(idPrestamo, cuotas) {
  const shName = _prSheetName_("SH_PRESTAMOS_CUOTAS", "PRESTAMOS_CUOTAS");
  const sh = obtenerSheet(shName);
  const data = sh.getDataRange().getDisplayValues();
  const headers = (data[0] || []).map(h => String(h || "").trim());
  const idx = {};
  headers.forEach((h, i) => { if (h) idx[h] = i; });
  const colPrestamo = idx["idPrestamo"];
  if (colPrestamo === undefined) throw new Error("La hoja PRESTAMOS_CUOTAS no tiene columna idPrestamo.");

  // Borra filas existentes del préstamo (de abajo hacia arriba)
  if (data.length > 1) {
    for (let r = data.length - 1; r >= 1; r--) {
      const row = data[r];
      if (String(row[colPrestamo] || "").trim() === String(idPrestamo || "").trim()) {
        sh.deleteRow(r + 1);
      }
    }
  }

  if (!cuotas || !cuotas.length) return;

  const now = _prNowIso_();
  const mkRow = (c) => {
    const row = new Array(headers.length).fill("");
    const set = (name, val) => {
      const i = idx[name];
      if (i === undefined) return;
      row[i] = val;
    };
    set("idCuota", Utilities.getUuid());
    set("idPrestamo", idPrestamo);
    set("nroCuota", c.nroCuota);
    set("fechaPago", c.fechaPago);
    set("cuota", c.cuota);
    set("interesCorriente", c.interesCorriente);
    set("capital", c.capital);
    set("saldoDespues", c.saldoDespues);
    set("estado", c.estado || "Pendiente");
    set("createdAt", now);
    // Parte 2 (defaults)
    set("interesPagado", 0);
    set("capitalPagado", 0);
    set("moraAcumulada", 0);
    set("moraPagada", 0);
    set("moraCalculadaHasta", c.fechaPago);
    set("updatedAt", now);
    return row;
  };

  const rows = cuotas.map(mkRow);
  const start = sh.getLastRow() + 1;
  sh.getRange(start, 1, rows.length, headers.length).setValues(rows);
}

// =========================
// Endpoints (Frontend)
// =========================

function listarPrestamos(params) {
  params = params || {};
  _prEnsureSchema_();

  const shName = _prSheetName_("SH_PRESTAMOS", "PRESTAMOS");
  let sh;
  try { sh = obtenerSheet(shName); } catch (e) { return []; }

  const main = _prRead_(sh);
  const personaId = String(params.personaId || "").trim();
  const origenParamRaw = String(params.origen || "").trim();
  const origenParam = (origenParamRaw ? origenParamRaw.toUpperCase() : "SOLICITADO");
  const estado = String(params.estado || "").trim().toLowerCase();
  const estadoOperativo = (estado === "activo" || estado === "finalizado");
  const q = String(params.q || "").trim().toLowerCase();
  const from = _prIso_(params.fechaFrom);
  const to = _prIso_(params.fechaTo);

  const out = [];
  main.rows.forEach(r => {
    const idPrestamo = String(_prGet_(r, main.idx, "idPrestamo") || "").trim();
    if (!idPrestamo) return;

    const origenRow = String(_prGet_(r, main.idx, "origen") || "").trim().toUpperCase() || "SOLICITADO";
    const tipoCliente = String(_prGet_(r, main.idx, "tipoCliente") || "").trim() || "PERSONA";
    const idPersona = String(_prGet_(r, main.idx, "idPersona") || "").trim();
    const personaNombre = String(_prGet_(r, main.idx, "personaNombre") || "").trim();
    const montoPrincipal = _prNum_(_prGet_(r, main.idx, "montoPrincipal"));
    const plazoMeses = parseInt(String(_prGet_(r, main.idx, "plazoMeses") || "0"), 10) || 0;
    const tasaMensual = _prNum_(_prGet_(r, main.idx, "tasaMensual"));
    const moratorioModo = String(_prGet_(r, main.idx, "moratorioModo") || "").trim();
    const tasaMoratoriaMensual = _prNum_(_prGet_(r, main.idx, "tasaMoratoriaMensual"));
    const adminTipo = String(_prGet_(r, main.idx, "adminTipo") || "").trim();
    const adminValor = _prNum_(_prGet_(r, main.idx, "adminValor"));
    const adminMonto = _prNum_(_prGet_(r, main.idx, "adminMonto"));
    const fechaDesembolso = _prIso_(_prGet_(r, main.idx, "fechaDesembolso"));
    const fechaHoraDesembolso = String(_prGet_(r, main.idx, "fechaHoraDesembolso") || "").trim();
    const metodoDesembolso = String(_prGet_(r, main.idx, "metodoDesembolso") || "").trim();
    const refDesembolso = String(_prGet_(r, main.idx, "refDesembolso") || "").trim();
    const notaDesembolso = String(_prGet_(r, main.idx, "notaDesembolso") || "").trim();
    const fechaPrimerPago = _prIso_(_prGet_(r, main.idx, "fechaPrimerPago"));
    const diaCorte = parseInt(String(_prGet_(r, main.idx, "diaCorte") || ""), 10) || "";
    const diaPago = parseInt(String(_prGet_(r, main.idx, "diaPago") || ""), 10) || "";
    const cuotaMensual = _prNum_(_prGet_(r, main.idx, "cuotaMensual"));
    const totalInteresEstimado = _prNum_(_prGet_(r, main.idx, "totalInteresEstimado"));
    const totalPagarEstimado = _prNum_(_prGet_(r, main.idx, "totalPagarEstimado"));
    const est = String(_prGet_(r, main.idx, "estado") || "").trim() || "Borrador";
    const fechaCreacion = String(_prGet_(r, main.idx, "fechaCreacion") || "").trim();
    const actualizadoPor = String(_prGet_(r, main.idx, "actualizadoPor") || "").trim();
    const fechaActualizacion = String(_prGet_(r, main.idx, "fechaActualizacion") || "").trim();

    // origen
    if (origenParam && origenParam !== "TODOS" && origenParam !== "*" && origenRow !== origenParam) return;
    if (personaId && idPersona !== personaId) return;
    // Si filtran por Activo/Finalizado, se resuelve después con estadoSistema
    if (estado && !estadoOperativo && String(est || "").toLowerCase() !== estado) return;
    if (from && fechaDesembolso && fechaDesembolso < from) return;
    if (to && fechaDesembolso && fechaDesembolso > to) return;
    if (q) {
      const hay = `${idPrestamo} ${personaNombre} ${idPersona}`.toLowerCase();
      if (hay.indexOf(q) === -1) return;
    }

    out.push({
      idPrestamo,
      origen: origenRow,
      tipoCliente,
      idPersona,
      personaNombre,
      montoPrincipal,
      plazoMeses,
      tasaMensual,
      moratorioModo,
      tasaMoratoriaMensual,
      adminTipo,
      adminValor,
      adminMonto,
      fechaDesembolso,
      fechaHoraDesembolso,
      metodoDesembolso,
      refDesembolso,
      notaDesembolso,
      fechaPrimerPago,
      diaCorte,
      diaPago,
      cuotaMensual,
      totalInteresEstimado,
      totalPagarEstimado,
      estado: est,
      fechaCreacion,
      actualizadoPor,
      fechaActualizacion,
    });
  });

  // Enriquecer listado con métricas operativas (saldo insoluto, vencidos, próximas cuotas, estado automático)
  try {
    const today = _prDateOnly_(new Date());
    const byId = {};
    out.forEach(it => {
      byId[String(it.idPrestamo || "").trim()] = {
        item: it,
        saldoInsoluto: 0,
        basePendiente: 0,
        moraPendiente: 0,
        totalPendiente: 0,
        cuotasVencidas: 0,
        montoVencido: 0,
        proximaCuotaFecha: "",
        proximaCuotaMonto: 0,
      };
    });

    const shCuotas = obtenerSheet(_prSheetName_("SH_PRESTAMOS_CUOTAS", "PRESTAMOS_CUOTAS"));
    const all = shCuotas.getDataRange().getDisplayValues();
    if (all && all.length > 1) {
      const headers = (all[0] || []).map(h => String(h || "").trim());
      const idx = {};
      headers.forEach((h, i) => { if (h) idx[h] = i; });
      const colPrestamo = idx["idPrestamo"];
      if (colPrestamo !== undefined) {
        const get = (row, name) => (idx[name] === undefined ? "" : row[idx[name]]);
        for (let r = 1; r < all.length; r++) {
          const row = all[r];
          if (!row || row.every(c => String(c || "").trim() === "")) continue;
          const pid = String(row[colPrestamo] || "").trim();
          const bucket = byId[pid];
          if (!bucket) continue;

          const it = bucket.item;
          const dailyRate = (_prNum_(it.tasaMoratoriaMensual) || 0) / 30;

          const interesCorriente = _prNum_(get(row, "interesCorriente"));
          const capital = _prNum_(get(row, "capital"));
          const interesPagado = _prNum_(get(row, "interesPagado"));
          const capitalPagado = _prNum_(get(row, "capitalPagado"));
          const moraAcumulada = _prNum_(get(row, "moraAcumulada"));
          const moraPagada = _prNum_(get(row, "moraPagada"));
          const fechaPago = _prIso_(get(row, "fechaPago"));
          const moraCalculadaHasta = _prIso_(get(row, "moraCalculadaHasta")) || fechaPago;

          const interesPend = _prRound2_(Math.max(0, interesCorriente - interesPagado));
          const capitalPend = _prRound2_(Math.max(0, capital - capitalPagado));
          const basePend = _prRound2_(interesPend + capitalPend);

          const due = _prDateOnly_(_prDateFromIso_(fechaPago));
          const calcHasta = _prDateOnly_(_prDateFromIso_(moraCalculadaHasta)) || due;
          const last = (due && calcHasta && calcHasta.getTime() < due.getTime()) ? due : calcHasta;

          let moraExtra = 0;
          if (dailyRate > 0 && basePend > 0 && due && today && today.getTime() > due.getTime()) {
            const desde = (last && today.getTime() > last.getTime()) ? last : due;
            const dias = _prDiffDays_(today, desde);
            if (dias > 0) moraExtra = _prRound2_(basePend * dailyRate * dias);
          }
          const moraAlDia = _prRound2_(moraAcumulada + moraExtra);
          const moraPend = _prRound2_(Math.max(0, moraAlDia - moraPagada));
          const totalPend = _prRound2_(basePend + moraPend);
          const estadoCuota = _prCuotaEstadoAuto_(totalPend, due, today);

          bucket.saldoInsoluto = _prRound2_(bucket.saldoInsoluto + capitalPend);
          bucket.basePendiente = _prRound2_(bucket.basePendiente + basePend);
          bucket.moraPendiente = _prRound2_(bucket.moraPendiente + moraPend);
          bucket.totalPendiente = _prRound2_(bucket.totalPendiente + totalPend);

          if (estadoCuota === "Vencida") {
            bucket.cuotasVencidas += 1;
            bucket.montoVencido = _prRound2_(bucket.montoVencido + totalPend);
          }
          // próxima cuota (la primera no pagada y no vencida en fecha >= hoy)
          if (estadoCuota !== "Pagada" && due && today && due.getTime() >= today.getTime()) {
            const cur = bucket.proximaCuotaFecha;
            if (!cur || String(fechaPago).localeCompare(String(cur)) < 0) {
              bucket.proximaCuotaFecha = fechaPago;
              bucket.proximaCuotaMonto = totalPend;
            }
          }
        }
      }
    }

    // escribir campos calculados en cada item
    Object.keys(byId).forEach(idp => {
      const b = byId[idp];
      const it = b.item;
      it.saldoInsoluto = _prRound2_(b.saldoInsoluto);
      it.basePendiente = _prRound2_(b.basePendiente);
      it.moraPendiente = _prRound2_(b.moraPendiente);
      it.totalPendiente = _prRound2_(b.totalPendiente);
      it.cuotasVencidas = b.cuotasVencidas;
      it.montoVencido = _prRound2_(b.montoVencido);
      it.proximaCuotaFecha = b.proximaCuotaFecha;
      it.proximaCuotaMonto = _prRound2_(b.proximaCuotaMonto);

      const estadoSistema = _prLoanEstadoAuto_(it.totalPendiente);
      it.estadoSistema = estadoSistema;
      // si ya estaba en estado operativo, reflejar el estado sistema en el campo estado de respuesta
      const estActual = String(it.estado || "").trim();
      const operativos = {"Activo":1,"Finalizado":1,"Aprobado":1,"":1};
      if (operativos[estActual] && estActual !== estadoSistema) {
        it.estado = estadoSistema;
      }
    });
  } catch (e) {
    // si falla, no rompe el listado
  }

  // Filtro operativo (Activo/Finalizado) después del cálculo
  const finalOut = (estadoOperativo && estado) 
    ? out.filter(it => String(it.estadoSistema || it.estado || "").trim().toLowerCase() === estado)
    : out;

  // Orden: fecha desc, luego id
  finalOut.sort((a, b) => {
    const fa = String(a.fechaDesembolso || "");
    const fb = String(b.fechaDesembolso || "");
    if (fa !== fb) return fb.localeCompare(fa);
    return String(b.idPrestamo || "").localeCompare(String(a.idPrestamo || ""));
  });

  return finalOut;
}

function guardarPrestamo(payload) {
  payload = payload || {};
  _prEnsureSchema_();

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const shName = _prSheetName_("SH_PRESTAMOS", "PRESTAMOS");
    const sh = obtenerSheet(shName);
    const main = _prRead_(sh);

    const idPersona = String(payload.idPersona || "").trim();
    const montoPrincipal = _prNum_(payload.montoPrincipal);
    const plazoMeses = Math.max(1, parseInt(payload.plazoMeses, 10) || 0);
    const tasaMensual = Number(payload.tasaMensual || 0);

    if (!idPersona) throw new Error("Selecciona una persona.");
    if (!(montoPrincipal > 0)) throw new Error("El monto debe ser mayor a 0.");
    if (!(plazoMeses > 0)) throw new Error("El plazo (meses) debe ser mayor a 0.");
    if (tasaMensual < 0) throw new Error("La tasa mensual no puede ser negativa.");

    const origen = String(payload.origen || "").trim().toUpperCase() || "SOLICITADO";
    const tipoCliente = String(payload.tipoCliente || "").trim() || "PERSONA";

    const moratorioModo = String(payload.moratorioModo || "25").trim();
    const adminTipo = String(payload.adminTipo || "fijo").trim();
    const adminValor = _prNum_(payload.adminValor);
    const fechaDesembolso = _prIso_(payload.fechaDesembolso) || _prIso_(new Date());
    // Campos de desembolso (principalmente para OTORGADO)
    const fechaHoraDesembolso = (payload.fechaHoraDesembolso !== undefined) ? String(payload.fechaHoraDesembolso || "").trim() : null;
    const metodoDesembolso = (payload.metodoDesembolso !== undefined) ? String(payload.metodoDesembolso || "").trim() : null;
    const refDesembolso = (payload.refDesembolso !== undefined) ? String(payload.refDesembolso || "").trim() : null;
    const notaDesembolso = (payload.notaDesembolso !== undefined) ? String(payload.notaDesembolso || "").trim() : null;
    const diaCorte = Math.max(1, Math.min(28, parseInt(payload.diaCorte, 10) || 1));
    const diaPago = Math.max(1, Math.min(28, parseInt(payload.diaPago, 10) || 1));
    const estado = String(payload.estado || "Borrador").trim() || "Borrador";

    // personaNombre (por conveniencia)
    const personaNombre = String(payload.personaNombre || "").trim() || _prLookupPersonaNombre_(idPersona);

    // Admin monto calculado
    let adminMonto = 0;
    if (adminTipo === "porcentaje") adminMonto = _prRound2_(montoPrincipal * (adminValor / 100));
    else adminMonto = _prRound2_(adminValor);

    // Moratoria mensual (guardamos la tasa total de mora)
    let tasaMoratoriaMensual = 0;
    if (moratorioModo === "manual") {
      tasaMoratoriaMensual = _prRound2_(_prNum_(payload.tasaMoratoriaMensual));
    } else if (moratorioModo === "50") {
      tasaMoratoriaMensual = _prRound2_(tasaMensual * 1.5);
    } else {
      // default 25
      tasaMoratoriaMensual = _prRound2_(tasaMensual * 1.25);
    }

    // Amortización
    const amort = _prBuildAmortization_({ montoPrincipal, plazoMeses, tasaMensual, fechaDesembolso, diaPago });
    const cuotaMensual = _prRound2_(amort.cuotaMensual);
    const totalInteresEstimado = _prRound2_(amort.totalInteres);
    const totalPagarEstimado = _prRound2_(montoPrincipal + totalInteresEstimado + adminMonto);

    let idPrestamo = String(payload.idPrestamo || "").trim();
    const isNew = !idPrestamo;
    if (isNew) idPrestamo = Utilities.getUuid();

    const now = _prNowIso_();

    // Upsert row
    const rowIndex = _prFindRowIndexById_(main.rows, main.idx, idPrestamo);
    const lastCol = sh.getLastColumn();
    const headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(h => String(h || "").trim());
    const idx = {};
    headers.forEach((h, i) => { if (h) idx[h] = i; });

    const setCell = (arr, colName, val) => {
      const i = idx[colName];
      if (i === undefined) return;
      arr[i] = val;
    };
    const setIf = (arr, colName, valOrNull) => {
      if (valOrNull === null) return;
      setCell(arr, colName, valOrNull);
    };

    if (rowIndex >= 0) {
      const existing = sh.getRange(rowIndex + 2, 1, 1, lastCol).getValues()[0];
      const row = existing.slice();
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
      setCell(row, "fechaDesembolso", fechaDesembolso);
      setIf(row, "fechaHoraDesembolso", fechaHoraDesembolso);
      setIf(row, "metodoDesembolso", metodoDesembolso);
      setIf(row, "refDesembolso", refDesembolso);
      setIf(row, "notaDesembolso", notaDesembolso);
      setCell(row, "fechaPrimerPago", amort.fechaPrimerPago);
      setCell(row, "diaCorte", diaCorte);
      setCell(row, "diaPago", diaPago);
      setCell(row, "cuotaMensual", cuotaMensual);
      setCell(row, "totalInteresEstimado", totalInteresEstimado);
      setCell(row, "totalPagarEstimado", totalPagarEstimado);
      setCell(row, "estado", estado);
      setCell(row, "actualizadoPor", String(payload.actualizadoPor || payload.user || ""));
      setCell(row, "fechaActualizacion", now);
      sh.getRange(rowIndex + 2, 1, 1, lastCol).setValues([row]);
    } else {
      const row = new Array(lastCol).fill("");
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
      setCell(row, "fechaDesembolso", fechaDesembolso);
      setCell(row, "fechaHoraDesembolso", fechaHoraDesembolso || "");
      setCell(row, "metodoDesembolso", metodoDesembolso || "");
      setCell(row, "refDesembolso", refDesembolso || "");
      setCell(row, "notaDesembolso", notaDesembolso || "");
      setCell(row, "fechaPrimerPago", amort.fechaPrimerPago);
      setCell(row, "diaCorte", diaCorte);
      setCell(row, "diaPago", diaPago);
      setCell(row, "cuotaMensual", cuotaMensual);
      setCell(row, "totalInteresEstimado", totalInteresEstimado);
      setCell(row, "totalPagarEstimado", totalPagarEstimado);
      setCell(row, "estado", estado);
      setCell(row, "creadoPor", String(payload.creadoPor || payload.user || ""));
      setCell(row, "fechaCreacion", now);
      setCell(row, "actualizadoPor", String(payload.actualizadoPor || payload.user || ""));
      setCell(row, "fechaActualizacion", now);
      sh.appendRow(row);
    }

    // Reemplazar cuotas del préstamo
    _prReplaceCuotas_(idPrestamo, amort.cuotas);

    return { ok: true, idPrestamo };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function eliminarPrestamo(idPrestamo) {
  _prEnsureSchema_();
  const id = String(idPrestamo || "").trim();
  if (!id) return { ok: true };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const shName = _prSheetName_("SH_PRESTAMOS", "PRESTAMOS");
    const sh = obtenerSheet(shName);
    const main = _prRead_(sh);
    const rowIndex = _prFindRowIndexById_(main.rows, main.idx, id);
    if (rowIndex >= 0) sh.deleteRow(rowIndex + 2);

    // borrar cuotas
    const shCuotas = obtenerSheet(_prSheetName_("SH_PRESTAMOS_CUOTAS", "PRESTAMOS_CUOTAS"));
    const data = shCuotas.getDataRange().getDisplayValues();
    if (data.length > 1) {
      const headers = data[0] || [];
      let colPrestamo = -1;
      headers.forEach((h, i) => { if (String(h || "").trim() === "idPrestamo") colPrestamo = i; });
      if (colPrestamo >= 0) {
        for (let r = data.length - 1; r >= 1; r--) {
          if (String(data[r][colPrestamo] || "").trim() === id) shCuotas.deleteRow(r + 1);
        }
      }
    }
    return { ok: true };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function obtenerDetallePrestamo(idPrestamo) {
  _prEnsureSchema_();
  const id = String(idPrestamo || "").trim();
  if (!id) return { prestamo: null, cuotas: [], pagos: [], resumen: null };

  const shName = _prSheetName_("SH_PRESTAMOS", "PRESTAMOS");
  let sh;
  try { sh = obtenerSheet(shName); } catch (e) { return { prestamo: null, cuotas: [], pagos: [], resumen: null }; }
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
    fechaActualizacion: String(_prGet_(r, main.idx, "fechaActualizacion") || "").trim(),
  };

  const shCuotas = obtenerSheet(_prSheetName_("SH_PRESTAMOS_CUOTAS", "PRESTAMOS_CUOTAS"));
  const det = _prRead_(shCuotas);
  const cuotas = [];
  det.rows.forEach(rr => {
    const pid = String(_prGet_(rr, det.idx, "idPrestamo") || "").trim();
    if (pid !== id) return;
    cuotas.push({
      idCuota: String(_prGet_(rr, det.idx, "idCuota") || "").trim(),
      idPrestamo: pid,
      nroCuota: parseInt(String(_prGet_(rr, det.idx, "nroCuota") || "0"), 10) || 0,
      fechaPago: _prIso_(_prGet_(rr, det.idx, "fechaPago")),
      cuota: _prNum_(_prGet_(rr, det.idx, "cuota")),
      interesCorriente: _prNum_(_prGet_(rr, det.idx, "interesCorriente")),
      capital: _prNum_(_prGet_(rr, det.idx, "capital")),
      saldoDespues: _prNum_(_prGet_(rr, det.idx, "saldoDespues")),
      estado: String(_prGet_(rr, det.idx, "estado") || "Pendiente").trim() || "Pendiente",
      interesPagado: _prNum_(_prGet_(rr, det.idx, "interesPagado")),
      capitalPagado: _prNum_(_prGet_(rr, det.idx, "capitalPagado")),
      moraAcumulada: _prNum_(_prGet_(rr, det.idx, "moraAcumulada")),
      moraPagada: _prNum_(_prGet_(rr, det.idx, "moraPagada")),
      moraCalculadaHasta: _prIso_(_prGet_(rr, det.idx, "moraCalculadaHasta")),
    });
  });
  cuotas.sort((a, b) => a.nroCuota - b.nroCuota);

  // Pagos del préstamo
  const shPagos = obtenerSheet(_prSheetName_("SH_PRESTAMOS_PAGOS", "Prestamos_Pagos"));
  const pagosRead = _prRead_(shPagos);
  const pagos = [];
  pagosRead.rows.forEach(rr => {
    const pid = String(_prGet_(rr, pagosRead.idx, "idPrestamo") || "").trim();
    if (pid !== id) return;
    pagos.push({
      idPago: String(_prGet_(rr, pagosRead.idx, "idPago") || "").trim(),
      idPrestamo: pid,
      fechaHora: String(_prGet_(rr, pagosRead.idx, "fechaHora") || "").trim(),
      monto: _prNum_(_prGet_(rr, pagosRead.idx, "monto")),
      metodo: String(_prGet_(rr, pagosRead.idx, "metodo") || "").trim(),
      referencia: String(_prGet_(rr, pagosRead.idx, "referencia") || "").trim(),
      nota: String(_prGet_(rr, pagosRead.idx, "nota") || "").trim(),
      moratorioCobrado: _prNum_(_prGet_(rr, pagosRead.idx, "moratorioCobrado")),
      interesCobrado: _prNum_(_prGet_(rr, pagosRead.idx, "interesCobrado")),
      capitalCobrado: _prNum_(_prGet_(rr, pagosRead.idx, "capitalCobrado")),
      saldoFavor: _prNum_(_prGet_(rr, pagosRead.idx, "saldoFavor")),
      aplicadoA: String(_prGet_(rr, pagosRead.idx, "aplicadoA") || "").trim(),
    });
  });
  pagos.sort((a, b) => String(b.fechaHora || "").localeCompare(String(a.fechaHora || "")));

  // Cálculo al día: vencidos + mora (sin escribir en Sheets)
  const today = _prDateOnly_(new Date());
  const dailyRate = (_prNum_(prestamo.tasaMoratoriaMensual) || 0) / 30;
  let totalMoraPendiente = 0;
  let totalBasePendiente = 0;
  let saldoInsoluto = 0;
  let cuotasVencidas = 0;
  let montoVencido = 0;

  // próximas cuotas (top 3)
  const upcoming = [];

  const cuotasAlDia = cuotas.map(c => {
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
    if (estadoAlDia !== "Pagada" && due && today && due.getTime() >= today.getTime()) {
      upcoming.push({
        nroCuota: c.nroCuota,
        fechaPago: c.fechaPago,
        pendiente: totalPend,
        moraPendiente: moraPend,
        interesPendiente: interesPend,
        capitalPendiente: capitalPend,
      });
    }

    totalMoraPendiente = _prRound2_(totalMoraPendiente + moraPend);
    totalBasePendiente = _prRound2_(totalBasePendiente + basePend);

    return {
      ...c,
      interesPendiente: interesPend,
      capitalPendiente: capitalPend,
      moraAlDia,
      moraPendiente: moraPend,
      totalPendiente: totalPend,
      estadoAlDia,
    };
  });

  upcoming.sort((a, b) => String(a.fechaPago || "").localeCompare(String(b.fechaPago || "")) || (a.nroCuota - b.nroCuota));
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
    proximasCuotas,
  };

  const estadoSistema = _prLoanEstadoAuto_(resumen.totalPendiente);
  prestamo.estadoSistema = estadoSistema;

  // Auto-actualiza estado del préstamo en Sheets SOLO si ya estaba en modo operativo
  // (Evita pisar Borrador/Solicitado/Rechazado)
  try {
    const estActual = String(prestamo.estado || "").trim();
    const operativos = {"Activo":1,"Finalizado":1,"Aprobado":1};
    if (operativos[estActual] || !estActual) {
      if (String(estActual || "") !== estadoSistema) {
        const colEstado = main.idx["estado"];
        const colFechaAct = main.idx["fechaActualizacion"];
        if (colEstado !== undefined) {
          const rowToWrite = rowIndex + 2;
          sh.getRange(rowToWrite, colEstado + 1).setValue(estadoSistema);
          if (colFechaAct !== undefined) sh.getRange(rowToWrite, colFechaAct + 1).setValue(_prNowIso_());
        }
        prestamo.estado = estadoSistema;
      }
    }
  } catch (e) {}

  return { prestamo, cuotas: cuotasAlDia, pagos, resumen };
}

function listarPagosPrestamo(idPrestamo) {
  _prEnsureSchema_();
  const id = String(idPrestamo || "").trim();
  if (!id) return [];
  const shPagos = obtenerSheet(_prSheetName_("SH_PRESTAMOS_PAGOS", "Prestamos_Pagos"));
  const read = _prRead_(shPagos);
  const pagos = [];
  read.rows.forEach(rr => {
    const pid = String(_prGet_(rr, read.idx, "idPrestamo") || "").trim();
    if (pid !== id) return;
    pagos.push({
      idPago: String(_prGet_(rr, read.idx, "idPago") || "").trim(),
      idPrestamo: pid,
      fechaHora: String(_prGet_(rr, read.idx, "fechaHora") || "").trim(),
      monto: _prNum_(_prGet_(rr, read.idx, "monto")),
      metodo: String(_prGet_(rr, read.idx, "metodo") || "").trim(),
      referencia: String(_prGet_(rr, read.idx, "referencia") || "").trim(),
      nota: String(_prGet_(rr, read.idx, "nota") || "").trim(),
      moratorioCobrado: _prNum_(_prGet_(rr, read.idx, "moratorioCobrado")),
      interesCobrado: _prNum_(_prGet_(rr, read.idx, "interesCobrado")),
      capitalCobrado: _prNum_(_prGet_(rr, read.idx, "capitalCobrado")),
      saldoFavor: _prNum_(_prGet_(rr, read.idx, "saldoFavor")),
      aplicadoA: String(_prGet_(rr, read.idx, "aplicadoA") || "").trim(),
    });
  });
  pagos.sort((a, b) => String(b.fechaHora || "").localeCompare(String(a.fechaHora || "")));
  return pagos;
}

function registrarPagoPrestamo(payload) {
  payload = payload || {};
  _prEnsureSchema_();

  const idPrestamo = String(payload.idPrestamo || "").trim();
  if (!idPrestamo) throw new Error("Falta idPrestamo.");

  const montoPago = _prRound2_(_prNum_(payload.monto));
  if (!(montoPago > 0)) throw new Error("El monto debe ser mayor a 0.");

  const metodo = String(payload.metodo || "").trim();
  const referencia = String(payload.referencia || "").trim();
  const nota = String(payload.nota || "").trim();

  // fecha/hora
  let dtPago = null;
  const fh = String(payload.fechaHora || "").trim();
  if (fh) {
    // soporta "YYYY-MM-DDTHH:mm" y ISO
    const tryDt = new Date(fh);
    if (tryDt.toString() !== "Invalid Date") dtPago = tryDt;
  }
  if (!dtPago) dtPago = new Date();
  const dtPagoDate = _prDateOnly_(dtPago);
  const fechaHoraIso = Utilities.formatDate(dtPago, Session.getScriptTimeZone() || "America/Tegucigalpa", "yyyy-MM-dd'T'HH:mm:ss");
  const fechaPagoIso = Utilities.formatDate(dtPagoDate, Session.getScriptTimeZone() || "America/Tegucigalpa", "yyyy-MM-dd");

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    // Cargar préstamo
    const shPrestamos = obtenerSheet(_prSheetName_("SH_PRESTAMOS", "PRESTAMOS"));
    const main = _prRead_(shPrestamos);
    const pIdx = _prFindRowIndexById_(main.rows, main.idx, idPrestamo);
    if (pIdx < 0) throw new Error("Préstamo no encontrado.");
    const prRow = main.rows[pIdx];
    const origen = String(_prGet_(prRow, main.idx, "origen") || "").trim().toUpperCase() || "SOLICITADO";
    const estadoPrestamo = String(_prGet_(prRow, main.idx, "estado") || "").trim() || "";
    const estadoLower = estadoPrestamo.toLowerCase();
    const fechaHoraDesembolso = String(_prGet_(prRow, main.idx, "fechaHoraDesembolso") || "").trim();

    if (estadoLower === "cancelado") throw new Error("Este préstamo está cancelado.");
    if (estadoLower === "finalizado") throw new Error("Este préstamo ya está finalizado.");
    // Para OTORGADO: no se permiten pagos si no se ha registrado el desembolso.
    if (origen === "OTORGADO" && !fechaHoraDesembolso) {
      throw new Error("Primero registra el desembolso para poder registrar pagos.");
    }

    const tasaMoratoriaMensual = _prNum_(_prGet_(prRow, main.idx, "tasaMoratoriaMensual"));
    const dailyRate = (tasaMoratoriaMensual || 0) / 30;

    // Cargar cuotas (con row numbers)
    const shCuotas = obtenerSheet(_prSheetName_("SH_PRESTAMOS_CUOTAS", "PRESTAMOS_CUOTAS"));
    const all = shCuotas.getDataRange().getDisplayValues();
    const headers = (all[0] || []).map(h => String(h || "").trim());
    const idx = {};
    headers.forEach((h, i) => { if (h) idx[h] = i; });
    const colPrestamo = idx["idPrestamo"];
    if (colPrestamo === undefined) throw new Error("PRESTAMOS_CUOTAS sin idPrestamo.");

    const cuotas = [];
    for (let r = 1; r < all.length; r++) {
      const row = all[r];
      if (!row || row.every(c => String(c || "").trim() === "")) continue;
      if (String(row[colPrestamo] || "").trim() !== idPrestamo) continue;
      const get = (name) => (idx[name] === undefined ? "" : row[idx[name]]);
      cuotas.push({
        sheetRow: r + 1,
        raw: row.slice(),
        idCuota: String(get("idCuota") || "").trim(),
        nroCuota: parseInt(String(get("nroCuota") || "0"), 10) || 0,
        fechaPago: _prIso_(get("fechaPago")),
        interesCorriente: _prNum_(get("interesCorriente")),
        capital: _prNum_(get("capital")),
        interesPagado: _prNum_(get("interesPagado")),
        capitalPagado: _prNum_(get("capitalPagado")),
        moraAcumulada: _prNum_(get("moraAcumulada")),
        moraPagada: _prNum_(get("moraPagada")),
        moraCalculadaHasta: _prIso_(get("moraCalculadaHasta")) || _prIso_(get("fechaPago")),
        estado: String(get("estado") || "Pendiente").trim() || "Pendiente",
      });
    }
    cuotas.sort((a, b) => a.nroCuota - b.nroCuota);
    if (!cuotas.length) throw new Error("Este préstamo no tiene cuotas.");

    // 1) Calcular mora acumulada hasta la fecha del pago (solo sobre saldo vencido)
    if (dailyRate > 0) {
      cuotas.forEach(c => {
        const interesPend = _prRound2_(Math.max(0, c.interesCorriente - c.interesPagado));
        const capitalPend = _prRound2_(Math.max(0, c.capital - c.capitalPagado));
        const basePend = _prRound2_(interesPend + capitalPend);
        if (basePend <= 0) return;

        const due = _prDateOnly_(_prDateFromIso_(c.fechaPago));
        if (!due) return;
        if (dtPagoDate.getTime() <= due.getTime()) return; // no vencida a la fecha del pago

        const calcHasta = _prDateOnly_(_prDateFromIso_(c.moraCalculadaHasta)) || due;
        const last = (calcHasta.getTime() < due.getTime()) ? due : calcHasta;
        const dias = _prDiffDays_(dtPagoDate, last);
        if (dias <= 0) return;

        const add = _prRound2_(basePend * dailyRate * dias);
        c.moraAcumulada = _prRound2_(c.moraAcumulada + add);
        c.moraCalculadaHasta = fechaPagoIso;
      });
    }

    // 2) Aplicar pago a cuotas: mora -> interes -> capital
    let rem = montoPago;
    const aplicado = [];
    let sumMora = 0;
    let sumInt = 0;
    let sumCap = 0;

    const payPart = (c, field, available, amount) => {
      const pay = Math.min(amount, Math.max(0, available));
      if (pay <= 0) return 0;
      c[field] = _prRound2_(c[field] + pay);
      return pay;
    };

    cuotas.forEach(c => {
      if (rem <= 0) return;
      const interesPend = _prRound2_(Math.max(0, c.interesCorriente - c.interesPagado));
      const capitalPend = _prRound2_(Math.max(0, c.capital - c.capitalPagado));
      const moraPend = _prRound2_(Math.max(0, c.moraAcumulada - c.moraPagada));
      const totalPend = _prRound2_(interesPend + capitalPend + moraPend);
      if (totalPend <= 0) return;

      const det = { idCuota: c.idCuota, nroCuota: c.nroCuota, mora: 0, interes: 0, capital: 0 };

      let p = payPart(c, "moraPagada", moraPend, rem);
      if (p) { det.mora = _prRound2_(det.mora + p); rem = _prRound2_(rem - p); sumMora = _prRound2_(sumMora + p); }

      p = payPart(c, "interesPagado", interesPend, rem);
      if (p) { det.interes = _prRound2_(det.interes + p); rem = _prRound2_(rem - p); sumInt = _prRound2_(sumInt + p); }

      p = payPart(c, "capitalPagado", capitalPend, rem);
      if (p) { det.capital = _prRound2_(det.capital + p); rem = _prRound2_(rem - p); sumCap = _prRound2_(sumCap + p); }

      if (det.mora || det.interes || det.capital) aplicado.push(det);

      // estado en Sheets (a la fecha del pago)
      const interesPend2 = _prRound2_(Math.max(0, c.interesCorriente - c.interesPagado));
      const capitalPend2 = _prRound2_(Math.max(0, c.capital - c.capitalPagado));
      const moraPend2 = _prRound2_(Math.max(0, c.moraAcumulada - c.moraPagada));
      const totalPend2 = _prRound2_(interesPend2 + capitalPend2 + moraPend2);
      const due = _prDateOnly_(_prDateFromIso_(c.fechaPago));
      c.estado = _prCuotaEstadoAuto_(totalPend2, due, dtPagoDate);
    });

    const saldoFavor = _prRound2_(Math.max(0, rem));

    // 3) Persistir actualización de cuotas
    const nowIso = _prNowIso_();
    const setCell = (row, name, val) => {
      const i = idx[name];
      if (i === undefined) return;
      row[i] = val;
    };

    const updates = [];
    cuotas.forEach(c => {
      // solo actualizamos cuotas tocadas por mora o pago
      const row = c.raw.slice();
      setCell(row, "interesPagado", c.interesPagado);
      setCell(row, "capitalPagado", c.capitalPagado);
      setCell(row, "moraAcumulada", c.moraAcumulada);
      setCell(row, "moraPagada", c.moraPagada);
      setCell(row, "moraCalculadaHasta", c.moraCalculadaHasta);
      setCell(row, "estado", c.estado);
      setCell(row, "updatedAt", nowIso);
      updates.push({ sheetRow: c.sheetRow, row });
    });
    updates.forEach(u => {
      shCuotas.getRange(u.sheetRow, 1, 1, headers.length).setValues([u.row]);
    });

    // 3.1) Estado del préstamo (a la fecha del pago)
    try {
      const totalPendPrestamo = _prRound2_(cuotas.reduce((acc, c) => {
        const iPend = _prRound2_(Math.max(0, c.interesCorriente - c.interesPagado));
        const cPend = _prRound2_(Math.max(0, c.capital - c.capitalPagado));
        const mPend = _prRound2_(Math.max(0, c.moraAcumulada - c.moraPagada));
        return acc + iPend + cPend + mPend;
      }, 0));
      const estadoSistema = _prLoanEstadoAuto_(totalPendPrestamo);
      const colEstado = main.idx["estado"];
      const colFechaAct = main.idx["fechaActualizacion"];
      const actual = String(_prGet_(prRow, main.idx, "estado") || "").trim();
      const operativos = {"Activo":1,"Finalizado":1,"Aprobado":1,"":1};
      if (colEstado !== undefined && operativos[actual] && actual !== estadoSistema) {
        shPrestamos.getRange(pIdx + 2, colEstado + 1).setValue(estadoSistema);
        if (colFechaAct !== undefined) shPrestamos.getRange(pIdx + 2, colFechaAct + 1).setValue(_prNowIso_());
      }
    } catch (e) {}

    // 4) Insertar pago en PRESTAMOS_PAGOS
    const shPagos = obtenerSheet(_prSheetName_("SH_PRESTAMOS_PAGOS", "Prestamos_Pagos"));
    const pagosData = shPagos.getDataRange().getDisplayValues();
    const pagosHeaders = (pagosData[0] || []).map(h => String(h || "").trim());
    const pMap = {};
    pagosHeaders.forEach((h, i) => { if (h) pMap[h] = i; });
    const pr = new Array(pagosHeaders.length).fill("");
    const pSet = (name, val) => {
      const i = pMap[name];
      if (i === undefined) return;
      pr[i] = val;
    };

    const idPago = Utilities.getUuid();
    pSet("idPago", idPago);
    pSet("idPrestamo", idPrestamo);
    pSet("fechaHora", fechaHoraIso);
    pSet("monto", montoPago);
    pSet("metodo", metodo);
    pSet("referencia", referencia);
    pSet("nota", nota);
    pSet("aplicadoA", JSON.stringify({ aplicado, saldoFavor }));
    pSet("moratorioCobrado", sumMora);
    pSet("interesCobrado", sumInt);
    pSet("capitalCobrado", sumCap);
    pSet("saldoFavor", saldoFavor);
    shPagos.appendRow(pr);

    return {
      ok: true,
      idPago,
      idPrestamo,
      fechaHora: fechaHoraIso,
      monto: montoPago,
      metodo,
      referencia,
      nota,
      moratorioCobrado: sumMora,
      interesCobrado: sumInt,
      capitalCobrado: sumCap,
      saldoFavor,
      aplicado,
    };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// =========================================================
//  OTORGAR V1 (Parte 1): desembolso + pago sugerido
// =========================================================

/**
 * Registrar desembolso del préstamo (especialmente útil para origen=OTORGADO).
 * payload: { idPrestamo, fechaHora, metodo, referencia, nota, user }
 */
function registrarDesembolso(payload) {
  return registrarDesembolsoPrestamo_(payload);
}

function registrarDesembolsoPrestamo_(payload) {
  payload = payload || {};
  _prEnsureSchema_();

  const idPrestamo = String(payload.idPrestamo || "").trim();
  if (!idPrestamo) throw new Error("Falta idPrestamo.");

  // fecha/hora
  let dt = null;
  const fh = String(payload.fechaHora || payload.fechaHoraDesembolso || "").trim();
  if (fh) {
    const tryDt = new Date(fh);
    if (tryDt.toString() !== "Invalid Date") dt = tryDt;
  }
  if (!dt) dt = new Date();

  const tz = Session.getScriptTimeZone() || "America/Tegucigalpa";
  const dtDate = _prDateOnly_(dt);
  const fechaHoraIso = Utilities.formatDate(dt, tz, "yyyy-MM-dd'T'HH:mm:ss");
  const fechaIso = Utilities.formatDate(dtDate, tz, "yyyy-MM-dd");

  const metodo = String(payload.metodo || "").trim();
  const referencia = String(payload.referencia || payload.refDesembolso || "").trim();
  const nota = String(payload.nota || payload.notaDesembolso || "").trim();

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    // 1) Cargar préstamo
    const shPrestamos = obtenerSheet(_prSheetName_("SH_PRESTAMOS", "PRESTAMOS"));
    const main = _prRead_(shPrestamos);
    const pIdx = _prFindRowIndexById_(main.rows, main.idx, idPrestamo);
    if (pIdx < 0) throw new Error("Préstamo no encontrado.");

    const prRow = main.rows[pIdx];
    const estado = String(_prGet_(prRow, main.idx, "estado") || "").trim();
    const estadoLower = estado.toLowerCase();
    if (estadoLower === "cancelado") throw new Error("Este préstamo está cancelado.");
    if (estadoLower === "finalizado") throw new Error("Este préstamo ya está finalizado.");

    // 2) Si ya hay pagos o cuotas con abonos, no permitimos re-generar el plan.
    const shPagos = obtenerSheet(_prSheetName_("SH_PRESTAMOS_PAGOS", "Prestamos_Pagos"));
    const pagos = _prRead_(shPagos);
    const tienePagos = pagos.rows.some(rr => String(_prGet_(rr, pagos.idx, "idPrestamo") || "").trim() === idPrestamo);
    if (tienePagos) throw new Error("Este préstamo ya tiene pagos registrados. No se puede registrar desembolso (o reprogramar) desde aquí.");

    const shCuotas = obtenerSheet(_prSheetName_("SH_PRESTAMOS_CUOTAS", "PRESTAMOS_CUOTAS"));
    const all = shCuotas.getDataRange().getDisplayValues();
    const headers = (all[0] || []).map(h => String(h || "").trim());
    const idx = {};
    headers.forEach((h, i) => { if (h) idx[h] = i; });
    const colPrestamo = idx["idPrestamo"];
    if (colPrestamo === undefined) throw new Error("PRESTAMOS_CUOTAS sin idPrestamo.");

    const colIntPag = idx["interesPagado"]; 
    const colCapPag = idx["capitalPagado"]; 
    const colMoraPag = idx["moraPagada"]; 
    if (all && all.length > 1) {
      for (let r = 1; r < all.length; r++) {
        const row = all[r];
        if (!row || row.every(c => String(c || "").trim() === "")) continue;
        if (String(row[colPrestamo] || "").trim() !== idPrestamo) continue;
        const ip = (colIntPag === undefined) ? 0 : _prNum_(row[colIntPag]);
        const cp = (colCapPag === undefined) ? 0 : _prNum_(row[colCapPag]);
        const mp = (colMoraPag === undefined) ? 0 : _prNum_(row[colMoraPag]);
        if (ip > 0 || cp > 0 || mp > 0) {
          throw new Error("Este préstamo ya tiene abonos en cuotas. No se puede registrar desembolso (o reprogramar) desde aquí.");
        }
      }
    }

    // 3) Recalcular amortización usando la fecha de desembolso real
    const montoPrincipal = _prNum_(_prGet_(prRow, main.idx, "montoPrincipal"));
    const plazoMeses = Math.max(1, parseInt(String(_prGet_(prRow, main.idx, "plazoMeses") || "1"), 10) || 1);
    const tasaMensual = Number(_prNum_(_prGet_(prRow, main.idx, "tasaMensual")) || 0);
    const diaPago = Math.max(1, Math.min(28, parseInt(String(_prGet_(prRow, main.idx, "diaPago") || "1"), 10) || 1));

    const amort = _prBuildAmortization_({ montoPrincipal, plazoMeses, tasaMensual, fechaDesembolso: fechaIso, diaPago });
    const cuotaMensual = _prRound2_(amort.cuotaMensual);
    const totalInteresEstimado = _prRound2_(amort.totalInteres);
    const adminMonto = _prNum_(_prGet_(prRow, main.idx, "adminMonto"));
    const totalPagarEstimado = _prRound2_(montoPrincipal + totalInteresEstimado + adminMonto);

    // 4) Actualizar cabecera
    const lastCol = shPrestamos.getLastColumn();
    const h = shPrestamos.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(v => String(v || "").trim());
    const map = {}; h.forEach((hh, i) => { if (hh) map[hh] = i; });
    const row = shPrestamos.getRange(pIdx + 2, 1, 1, lastCol).getValues()[0].slice();
    const set = (name, val) => {
      const i = map[name];
      if (i === undefined) return;
      row[i] = val;
    };

    set("fechaDesembolso", fechaIso);
    set("fechaHoraDesembolso", fechaHoraIso);
    set("metodoDesembolso", metodo);
    set("refDesembolso", referencia);
    set("notaDesembolso", nota);
    set("fechaPrimerPago", amort.fechaPrimerPago);
    set("cuotaMensual", cuotaMensual);
    set("totalInteresEstimado", totalInteresEstimado);
    set("totalPagarEstimado", totalPagarEstimado);
    set("estado", "Activo");
    set("actualizadoPor", String(payload.actualizadoPor || payload.user || ""));
    set("fechaActualizacion", _prNowIso_());

    shPrestamos.getRange(pIdx + 2, 1, 1, lastCol).setValues([row]);

    // 5) Reemplazar cuotas
    _prReplaceCuotas_(idPrestamo, amort.cuotas);

    return {
      ok: true,
      idPrestamo,
      estado: "Activo",
      fechaDesembolso: fechaIso,
      fechaHoraDesembolso: fechaHoraIso,
      fechaPrimerPago: amort.fechaPrimerPago,
      cuotaMensual,
      totalPagarEstimado,
    };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/**
 * Sugiere monto de pago y muestra una pre-aplicación (sin escribir en Sheets).
 * payload: { idPrestamo, modo, cuotaId?, nroCuota?, fechaPago?, monto? }
 * modos: PROXIMA_CUOTA | CUOTA_ESPECIFICA | LIQUIDAR_HOY | MONTO_LIBRE
 */
function getPagoSugerido(payload) {
  return getPagoSugeridoPrestamo_(payload);
}

function getPagoSugeridoPrestamo_(payload) {
  payload = payload || {};
  _prEnsureSchema_();

  const idPrestamo = String(payload.idPrestamo || "").trim();
  if (!idPrestamo) throw new Error("Falta idPrestamo.");

  const modoRaw = String(payload.modo || payload.mode || "PROXIMA_CUOTA").trim().toUpperCase();
  const modoMap = {
    "PROXIMA": "PROXIMA_CUOTA",
    "PROXIMA_CUOTA": "PROXIMA_CUOTA",
    "CUOTA": "CUOTA_ESPECIFICA",
    "CUOTA_ESPECIFICA": "CUOTA_ESPECIFICA",
    "LIQUIDAR": "LIQUIDAR_HOY",
    "LIQUIDAR_HOY": "LIQUIDAR_HOY",
    "MONTO": "MONTO_LIBRE",
    "MONTO_LIBRE": "MONTO_LIBRE",
  };
  const modo = modoMap[modoRaw] || "PROXIMA_CUOTA";

  // fecha (as of)
  let dt = null;
  const fp = String(payload.fechaPago || payload.fechaHora || "").trim();
  if (fp) {
    const tryDt = new Date(fp);
    if (tryDt.toString() !== "Invalid Date") dt = tryDt;
  }
  if (!dt) dt = new Date();
  const asOf = _prDateOnly_(dt);
  const tz = Session.getScriptTimeZone() || "America/Tegucigalpa";
  const asOfIso = Utilities.formatDate(asOf, tz, "yyyy-MM-dd");

  const cuotaId = String(payload.cuotaId || payload.idCuota || "").trim();
  const nroCuota = parseInt(String(payload.nroCuota || ""), 10) || 0;
  const montoInput = _prRound2_(_prNum_(payload.monto));

  // 1) Cargar préstamo
  const shPrestamos = obtenerSheet(_prSheetName_("SH_PRESTAMOS", "PRESTAMOS"));
  const main = _prRead_(shPrestamos);
  const pIdx = _prFindRowIndexById_(main.rows, main.idx, idPrestamo);
  if (pIdx < 0) throw new Error("Préstamo no encontrado.");
  const prRow = main.rows[pIdx];
  const prestamo = {
    idPrestamo,
    origen: String(_prGet_(prRow, main.idx, "origen") || "").trim().toUpperCase() || "SOLICITADO",
    estado: String(_prGet_(prRow, main.idx, "estado") || "").trim() || "",
    tasaMoratoriaMensual: _prNum_(_prGet_(prRow, main.idx, "tasaMoratoriaMensual")),
  };
  const dailyRate = (prestamo.tasaMoratoriaMensual || 0) / 30;

  // 2) Cargar cuotas
  const shCuotas = obtenerSheet(_prSheetName_("SH_PRESTAMOS_CUOTAS", "PRESTAMOS_CUOTAS"));
  const all = shCuotas.getDataRange().getDisplayValues();
  const headers = (all[0] || []).map(h => String(h || "").trim());
  const idx = {}; headers.forEach((h, i) => { if (h) idx[h] = i; });
  const colPrestamo = idx["idPrestamo"];
  if (colPrestamo === undefined) throw new Error("PRESTAMOS_CUOTAS sin idPrestamo.");

  const cuotas = [];
  for (let r = 1; r < all.length; r++) {
    const row = all[r];
    if (!row || row.every(c => String(c || "").trim() === "")) continue;
    if (String(row[colPrestamo] || "").trim() !== idPrestamo) continue;
    const get = (name) => (idx[name] === undefined ? "" : row[idx[name]]);
    cuotas.push({
      idCuota: String(get("idCuota") || "").trim(),
      nroCuota: parseInt(String(get("nroCuota") || "0"), 10) || 0,
      fechaPago: _prIso_(get("fechaPago")),
      interesCorriente: _prNum_(get("interesCorriente")),
      capital: _prNum_(get("capital")),
      interesPagado: _prNum_(get("interesPagado")),
      capitalPagado: _prNum_(get("capitalPagado")),
      moraAcumulada: _prNum_(get("moraAcumulada")),
      moraPagada: _prNum_(get("moraPagada")),
      moraCalculadaHasta: _prIso_(get("moraCalculadaHasta")) || _prIso_(get("fechaPago")),
    });
  }
  cuotas.sort((a, b) => a.nroCuota - b.nroCuota);
  if (!cuotas.length) throw new Error("Este préstamo no tiene cuotas.");

  // 3) Calcular cuotas al día (asOf)
  let saldoInsoluto = 0;
  let basePendiente = 0;
  let moraPendiente = 0;
  let cuotasVencidas = 0;
  let montoVencido = 0;

  const cuotasAlDia = cuotas.map(c => {
    const interesPend = _prRound2_(Math.max(0, (c.interesCorriente || 0) - (c.interesPagado || 0)));
    const capitalPend = _prRound2_(Math.max(0, (c.capital || 0) - (c.capitalPagado || 0)));
    const basePend = _prRound2_(interesPend + capitalPend);

    const due = _prDateOnly_(_prDateFromIso_(c.fechaPago));
    const calcHasta = _prDateOnly_(_prDateFromIso_(c.moraCalculadaHasta)) || due;
    const last = (due && calcHasta && calcHasta.getTime() < due.getTime()) ? due : calcHasta;
    let moraExtra = 0;
    if (dailyRate > 0 && basePend > 0 && due && asOf && asOf.getTime() > due.getTime()) {
      const desde = (last && asOf.getTime() > last.getTime()) ? last : due;
      const dias = _prDiffDays_(asOf, desde);
      if (dias > 0) moraExtra = _prRound2_(basePend * dailyRate * dias);
    }
    const moraAlDia = _prRound2_((c.moraAcumulada || 0) + moraExtra);
    const moraPend = _prRound2_(Math.max(0, moraAlDia - (c.moraPagada || 0)));
    const totalPend = _prRound2_(basePend + moraPend);

    const estadoAlDia = _prCuotaEstadoAuto_(totalPend, due, asOf);

    saldoInsoluto = _prRound2_(saldoInsoluto + capitalPend);
    basePendiente = _prRound2_(basePendiente + basePend);
    moraPendiente = _prRound2_(moraPendiente + moraPend);
    if (estadoAlDia === "Vencida") {
      cuotasVencidas += 1;
      montoVencido = _prRound2_(montoVencido + totalPend);
    }

    return {
      ...c,
      interesPendiente: interesPend,
      capitalPendiente: capitalPend,
      moraAlDia,
      moraPendiente: moraPend,
      totalPendiente: totalPend,
      estadoAlDia,
    };
  });

  const totalPendiente = _prRound2_(basePendiente + moraPendiente);
  const firstPending = cuotasAlDia.find(c => (c.totalPendiente || 0) > 0.000001) || null;

  // 4) Monto sugerido
  const warnings = [];
  let montoSugerido = 0;

  if (modo === "LIQUIDAR_HOY") {
    montoSugerido = totalPendiente;
  } else if (modo === "CUOTA_ESPECIFICA") {
    const target = cuotaId
      ? cuotasAlDia.find(c => String(c.idCuota || "").trim() === cuotaId)
      : (nroCuota ? cuotasAlDia.find(c => c.nroCuota === nroCuota) : null);
    if (!target) throw new Error("No se encontró la cuota indicada.");

    const prevPend = cuotasAlDia.filter(c => c.nroCuota < target.nroCuota && (c.totalPendiente || 0) > 0.000001);
    if (prevPend.length) warnings.push("Hay cuotas anteriores pendientes. El pago siempre se aplicará primero a las cuotas más antiguas.");

    montoSugerido = _prRound2_(cuotasAlDia
      .filter(c => c.nroCuota <= target.nroCuota)
      .reduce((acc, c) => acc + (c.totalPendiente || 0), 0));
  } else if (modo === "MONTO_LIBRE") {
    montoSugerido = (montoInput > 0) ? montoInput : 0;
  } else {
    // PROXIMA_CUOTA
    montoSugerido = firstPending ? _prRound2_(firstPending.totalPendiente || 0) : 0;
  }

  montoSugerido = _prRound2_(montoSugerido);

  // 5) Previsualizar aplicación del pago
  const previewAmount = (montoSugerido > 0) ? montoSugerido : 0;
  const aplicado = [];
  let rem = previewAmount;
  let sumMora = 0;
  let sumInt = 0;
  let sumCap = 0;

  const pushDet = (c, mora, interes, capital) => {
    if (!(mora || interes || capital)) return;
    aplicado.push({ idCuota: c.idCuota, nroCuota: c.nroCuota, mora, interes, capital });
  };

  cuotasAlDia.forEach(c => {
    if (rem <= 0) return;
    const moraPend = _prRound2_(Math.max(0, c.moraPendiente || 0));
    const interesPend = _prRound2_(Math.max(0, c.interesPendiente || 0));
    const capitalPend = _prRound2_(Math.max(0, c.capitalPendiente || 0));
    const totalPend = _prRound2_(moraPend + interesPend + capitalPend);
    if (totalPend <= 0) return;

    let mora = 0, interes = 0, capital = 0;
    let p = Math.min(rem, moraPend);
    if (p > 0) { mora = _prRound2_(p); rem = _prRound2_(rem - p); sumMora = _prRound2_(sumMora + p); }
    p = Math.min(rem, interesPend);
    if (p > 0) { interes = _prRound2_(p); rem = _prRound2_(rem - p); sumInt = _prRound2_(sumInt + p); }
    p = Math.min(rem, capitalPend);
    if (p > 0) { capital = _prRound2_(p); rem = _prRound2_(rem - p); sumCap = _prRound2_(sumCap + p); }
    pushDet(c, mora, interes, capital);
  });

  const saldoFavor = _prRound2_(Math.max(0, rem));
  if (previewAmount > 0 && saldoFavor > 0.000001) warnings.push("El monto excede lo pendiente; quedará saldo a favor.");

  return {
    ok: true,
    idPrestamo,
    modo,
    fechaPago: asOfIso,
    montoSugerido,
    resumen: {
      saldoInsoluto: _prRound2_(saldoInsoluto),
      basePendiente: _prRound2_(basePendiente),
      moraPendiente: _prRound2_(moraPendiente),
      totalPendiente,
      cuotasVencidas,
      montoVencido: _prRound2_(montoVencido),
      proximaCuota: firstPending ? { idCuota: firstPending.idCuota, nroCuota: firstPending.nroCuota, fechaPago: firstPending.fechaPago, pendiente: _prRound2_(firstPending.totalPendiente || 0) } : null,
    },
    preview: (previewAmount > 0) ? {
      monto: previewAmount,
      moratorioCobrado: sumMora,
      interesCobrado: sumInt,
      capitalCobrado: sumCap,
      saldoFavor,
      aplicado,
    } : null,
    warnings,
  };
}
