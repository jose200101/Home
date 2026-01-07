/**
 * =========================
 *  GASTOS NETOS (Dashboard)
 * =========================
 * Fase 1:
 *  - Resumen por periodo (cards) + filtros
 *  - Total gastos (Fijos + Variables)
 *  - Total pagado / pendiente
 *  - Conteo/monto de saldos a favor / en contra (por ahora basado en VARIABLES pendientes)
 *
 * Nota: El netting completo y el balance detallado por persona se implementan en fases siguientes.
 */

/**
 * =========================
 *  FASE 2 (Balance por persona)
 * =========================
 * - Tabla resumida por persona (fijos + variables)
 * - Botón “Ver detalle” (UI) – el detalle completo se desarrolla en fases siguientes
 *
 * Importante (modelo actual):
 * - GASTOS FIJOS usa aportes en DETALLE_GASTO_FIJO con (personaId, personaNombre, pagado).
 * - GASTOS VARIABLES representa deudas DEUDOR->ACREEDOR y soporta abonos en DETALLE_GASTO_VARIABLE.
 * - En este proyecto, los IDs de "personas" pueden venir de PERSONAS y/o USUARIOS.
 *   Para evitar que el balance quede vacío, construimos un directorio unificado (PERSONAS + USUARIOS)
 *   y además agregamos IDs que aparezcan en los datos aunque no existan en el directorio.
 */

function _gnNowIso_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "America/Tegucigalpa", "yyyy-MM-dd'T'HH:mm:ss");
}

function _gnNum_(v) {
  const n = (typeof v === "number") ? v : parseFloat(String(v || "").replace(/[^0-9.\-]/g, ""));
  return isNaN(n) ? 0 : n;
}

function _gnBool_(v, defVal) {
  if (typeof v === "boolean") return v;
  const s = String(v ?? "").trim().toLowerCase();
  if (!s) return !!defVal;
  return s === "true" || s === "1" || s === "si" || s === "sí" || s === "yes";
}

function _gnNormalizeISODate_(raw) {
  const s = String(raw || "").trim();
  if (!s) return "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;

  const m1 = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m1) {
    const a = parseInt(m1[1], 10);
    const b = parseInt(m1[2], 10);
    const y = parseInt(m1[3], 10);
    // Prefer dd/mm en es-HN
    let d = a, m = b;
    if (a <= 12 && b > 12) { m = a; d = b; } // mm/dd
    return `${y}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
  }

  const dt = new Date(s);
  if (dt.toString() !== "Invalid Date") {
    const y = dt.getFullYear();
    const m = String(dt.getMonth() + 1).padStart(2, "0");
    const d = String(dt.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }
  return s;
}

function _gnNormalizeYYYYMM_(raw) {
  const s = String(raw || "").trim();
  if (!s) return "";
  if (/^\d{4}-\d{2}$/.test(s)) return s;

  const m1 = s.match(/^(\d{1,2})\/(\d{4})$/); // mm/yyyy
  if (m1) {
    const mm = String(parseInt(m1[1], 10)).padStart(2, "0");
    const yy = String(parseInt(m1[2], 10));
    return `${yy}-${mm}`;
  }

  const iso = _gnNormalizeISODate_(s);
  if (/^\d{4}-\d{2}-\d{2}$/.test(iso)) return iso.slice(0, 7);

  const dt = new Date(s);
  if (dt.toString() !== "Invalid Date") {
    const y = dt.getFullYear();
    const m = String(dt.getMonth() + 1).padStart(2, "0");
    return `${y}-${m}`;
  }

  return s;
}

function _gnMonthToRange_(yyyyMm) {
  const p = _gnNormalizeYYYYMM_(yyyyMm);
  if (!/^\d{4}-\d{2}$/.test(p)) return { desde: "", hasta: "" };
  const [y, m] = p.split("-").map(Number);
  const desde = `${y}-${String(m).padStart(2, "0")}-01`;
  const last = new Date(y, m, 0); // day 0 of next month => last of current
  const hasta = `${y}-${String(m).padStart(2, "0")}-${String(last.getDate()).padStart(2, "0")}`;
  return { desde, hasta };
}

function _gnKeyName_(s) {
  return String(s || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ");
}


/**
 * =========================
 *  Normalización de personas (CRÍTICO)
 * =========================
 * Regla: siempre preferir personaId (estable).
 * Fallback robusto cuando falta idPersona en alguna hoja:
 *  1) intenta resolver por nombre con un mapa nombre_normalizado -> idPersona
 *  2) si el nombre no existe o es ambiguo (duplicado), crea un id sintético determinístico: "name:<clave>"
 *     Esto evita perder movimientos en el balance y ayuda a detectar datos históricos sin id.
 */

function _gnHashKey_(s) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s || ''), Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function _gnCacheGet_(cacheKey) {
  try {
    const cache = CacheService.getScriptCache();
    const raw = cache.get(cacheKey);
    if (!raw) return null;
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
}

function _gnCachePut_(cacheKey, obj, ttlSec) {
  try {
    const cache = CacheService.getScriptCache();
    cache.put(cacheKey, JSON.stringify(obj), Math.max(5, Math.min(600, Number(ttlSec || 60))));
  } catch (e) {}
}

function _gnBuildNameToIdMap_(dir) {
  const map = {};
  const ambiguous = {};
  Object.keys(dir || {}).forEach(id => {
    const nombre = String(dir[id]?.nombre || '').trim();
    if (!nombre) return;
    const k = _gnKeyName_(nombre);
    if (!k) return;
    if (!map[k]) map[k] = String(id);
    else if (String(map[k]) !== String(id)) {
      ambiguous[k] = true;
    }
  });
  // Limpia claves ambiguas
  Object.keys(ambiguous).forEach(k => { delete map[k]; });
  return { map, ambiguousKeys: Object.keys(ambiguous) };
}

function _gnResolvePersonaId_(personaId, personaNombre, nameToIdMap) {
  const id = String(personaId || '').trim();
  if (id) return id;
  const nombre = String(personaNombre || '').trim();
  if (!nombre) return '';
  const k = _gnKeyName_(nombre);
  if (!k) return '';
  const mapped = nameToIdMap && nameToIdMap[k];
  if (mapped) return String(mapped);
  return `name:${k}`; // id sintético determinístico
}

function _gnGetPersonContext_() {
  // Cache de directorio de personas/usuarios (cambia poco)
  const ckey = 'gn:personCtx:v2';
  const cached = _gnCacheGet_(ckey);
  if (cached && cached.dir && cached.nameToIdMap) return cached;

  const dir = _gnBuildPersonDirectory_();
  const built = _gnBuildNameToIdMap_(dir);
  const ctx = {
    dir,
    nameToIdMap: built.map,
    ambiguousNameKeys: built.ambiguousKeys,
    lastUpdateIso: _gnNowIso_(),
  };
  _gnCachePut_(ckey, ctx, 120);
  return ctx;
}

function _gnNormalizeData_(fijos, vars, ctx) {
  ctx = ctx || _gnGetPersonContext_();
  const map = ctx.nameToIdMap || {};

  // Normaliza aportes en fijos
  (fijos || []).forEach(g => {
    const aportes = Array.isArray(g.aportes) ? g.aportes : [];
    aportes.forEach(a => {
      const rid = _gnResolvePersonaId_(a.personaId, a.personaNombre, map);
      a.personaId = rid;
      // Si tenemos nombre en directorio y el aporte no trae nombre, lo completamos
      if ((!a.personaNombre || !String(a.personaNombre).trim()) && ctx.dir && ctx.dir[rid]) {
        a.personaNombre = String(ctx.dir[rid].nombre || '').trim();
      }
    });
  });

  // Normaliza deudor/acreedor en variables
  (vars || []).forEach(v => {
    const did = _gnResolvePersonaId_(v.deudorId, v.deudorNombre, map);
    const aid = _gnResolvePersonaId_(v.acreedorId, v.acreedorNombre, map);
    v.deudorId = did;
    v.acreedorId = aid;
    if ((!v.deudorNombre || !String(v.deudorNombre).trim()) && ctx.dir && ctx.dir[did]) {
      v.deudorNombre = String(ctx.dir[did].nombre || '').trim();
    }
    if ((!v.acreedorNombre || !String(v.acreedorNombre).trim()) && ctx.dir && ctx.dir[aid]) {
      v.acreedorNombre = String(ctx.dir[aid].nombre || '').trim();
    }
  });

  return { fijos: fijos || [], vars: vars || [] };
}

function _gnBuildPersonDirectory_() {
  /**
   * Une PERSONAS + USUARIOS (si existen) en un solo directorio.
   * Estructura: { [id]: { id, nombre } }
   */
  const dir = {};

  // PERSONAS
  try {
    if (typeof listarPersonasActivas === "function") {
      (listarPersonasActivas() || []).forEach(p => {
        const id = String(p?.id_persona || p?.id || "").trim();
        const nombre = String(p?.nombre_persona || p?.nombre || "").trim();
        if (id && nombre && !dir[id]) dir[id] = { id, nombre };
      });
    }
  } catch (e) {}

  // USUARIOS
  try {
    if (typeof listarUsuariosActivos === "function") {
      (listarUsuariosActivos() || []).forEach(u => {
        const id = String(u?.id || "").trim();
        const nombre = String(u?.nombreCompleto || u?.nombre_usuario || "").trim();
        if (id && nombre && !dir[id]) dir[id] = { id, nombre };
      });
    }
  } catch (e) {}

  return dir;
}

function listarPersonasNetosActivas() {
  const dir = _gnBuildPersonDirectory_();
  const out = Object.keys(dir).map(id => dir[id]);
  out.sort((a, b) => String(a.nombre || "").localeCompare(String(b.nombre || "")));
  return out;
}

function _gnParseParams_(params) {
  params = params || {};
  const tipoRaw = String(params.tipo || "Todos").trim().toLowerCase();
  const personaId = String(params.personaId || "").trim();
  const incluirPagados = _gnBool_(params.incluirPagados, true);
  const modo = String(params.modo || "").trim().toLowerCase() || (params.desde || params.hasta ? "rango" : "mes");

  let periodo = _gnNormalizeYYYYMM_(params.periodo || "");
  let desde = _gnNormalizeISODate_(params.desde || params.fechaFrom || "");
  let hasta = _gnNormalizeISODate_(params.hasta || params.fechaTo || "");

  // Defaults: si no hay nada, usamos mes actual
  if (!periodo && !desde && !hasta) {
    const now = new Date();
    periodo = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;
  }

  if (modo === "mes") {
    if (!periodo) {
      const now = new Date();
      periodo = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;
    }
    const rr = _gnMonthToRange_(periodo);
    desde = rr.desde;
    hasta = rr.hasta;
  } else {
    if (!desde && periodo) {
      const rr = _gnMonthToRange_(periodo);
      desde = rr.desde;
    }
    if (!hasta && periodo) {
      const rr = _gnMonthToRange_(periodo);
      hasta = rr.hasta;
    }
  }

  const onlyFijos = tipoRaw === "fijos";
  const onlyVars = tipoRaw === "variables";

  if (desde && hasta && String(desde) > String(hasta)) {
    throw new Error('Rango de fechas inválido: la fecha desde es mayor que hasta.');
  }

  return { params, tipoRaw, personaId, incluirPagados, modo, periodo, desde, hasta, onlyFijos, onlyVars };
}

function _gnFetchData_(parsed, ctx) {
  const { modo, periodo, desde, hasta, personaId, incluirPagados, onlyFijos, onlyVars } = parsed;

  let fijos = [];
  let vars = [];

  // ========== Fijos ==========
  if (!onlyVars && typeof listarGastosFijos === "function") {
    const p = {};
    if (modo === "mes" && periodo) p.periodo = periodo;
    else {
      if (desde) p.venceFrom = desde;
      if (hasta) p.venceTo = hasta;
    }
    if (personaId) p.personaId = personaId;
    fijos = listarGastosFijos(p) || [];
    if (!incluirPagados) {
      fijos = fijos.filter(g => _gnNum_(g.pendiente) > 0.000001);
    }
  }

  // ========== Variables ==========
  if (!onlyFijos && typeof listarGastosVariables === "function") {
    const p = {};
    if (desde) p.fechaFrom = desde;
    if (hasta) p.fechaTo = hasta;
    vars = listarGastosVariables(p) || [];

    if (personaId) {
      vars = vars.filter(v => {
        const d = String(v.deudorId || "").trim();
        const a = String(v.acreedorId || "").trim();
        return d === personaId || a === personaId;
      });
    }

    if (!incluirPagados) {
      vars = vars.filter(v => _gnNum_(v.pendiente) > 0.000001);
    }
  }

  const norm = _gnNormalizeData_(fijos, vars, ctx);
  fijos = norm.fijos;
  vars = norm.vars;

  return { fijos, vars };
}

function _gnComputeCards_(fijos, vars, parsed) {
  const totalFijos = fijos.reduce((acc, g) => acc + _gnNum_(g.totalFactura), 0);
  const pagadoFijos = fijos.reduce((acc, g) => acc + _gnNum_(g.aportado), 0);
  const pendienteFijos = fijos.reduce((acc, g) => acc + _gnNum_(g.pendiente), 0);

  const totalVars = vars.reduce((acc, v) => acc + _gnNum_(v.monto), 0);
  const pagadoVars = vars.reduce((acc, v) => acc + _gnNum_(v.abonado), 0);
  const pendienteVars = vars.reduce((acc, v) => acc + _gnNum_(v.pendiente), 0);

  const totalGastos = totalFijos + totalVars;
  const totalPagado = pagadoFijos + pagadoVars;
  const totalPendiente = pendienteFijos + pendienteVars;

  // Saldos (variables pendientes)
  const netByPersona = {};
  vars.forEach(v => {
    const pend = _gnNum_(v.pendiente);
    if (pend <= 0.000001) return;
    const deudorId = String(v.deudorId || "").trim();
    const acreedorId = String(v.acreedorId || "").trim();
    if (deudorId) netByPersona[deudorId] = (netByPersona[deudorId] || 0) - pend;
    if (acreedorId) netByPersona[acreedorId] = (netByPersona[acreedorId] || 0) + pend;
  });

  let aFavorCount = 0, aFavorMonto = 0;
  let enContraCount = 0, enContraMonto = 0;

  if (parsed.personaId) {
    const net = _gnNum_(netByPersona[parsed.personaId] || 0);
    if (net > 0.000001) { aFavorCount = 1; aFavorMonto = net; }
    else if (net < -0.000001) { enContraCount = 1; enContraMonto = Math.abs(net); }
  } else {
    Object.keys(netByPersona).forEach(pid => {
      const net = _gnNum_(netByPersona[pid]);
      if (net > 0.000001) { aFavorCount += 1; aFavorMonto += net; }
      else if (net < -0.000001) { enContraCount += 1; enContraMonto += Math.abs(net); }
    });
  }

  return {
    totalGastos,
    totalPagado,
    totalPendiente,
    aFavorCount,
    aFavorMonto,
    enContraCount,
    enContraMonto,
    lastUpdateIso: _gnNowIso_(),
  };
}

function _gnComputeBalance_(fijos, vars, parsed, ctx) {
  ctx = ctx || _gnGetPersonContext_();
  const dir = ctx.dir || {};

  // Inicializa con directorio
  const sums = {};
  Object.keys(dir).forEach(id => {
    sums[id] = {
      personaId: id,
      personaNombre: dir[id].nombre,
      fijosAsignados: 0,
      fijosPagados: 0,
      variablesDeudor: 0,
      variablesAcreedor: 0,
      pagosRecibidos: 0,
      pagosHechos: 0,
    };
  });

  const ensure = (id, nombreHint) => {
    const pid = String(id || "").trim();
    if (!pid) return null;
    if (!sums[pid]) {
      sums[pid] = {
        personaId: pid,
        personaNombre: String(nombreHint || "(Sin nombre)").trim() || "(Sin nombre)",
        fijosAsignados: 0,
        fijosPagados: 0,
        variablesDeudor: 0,
        variablesAcreedor: 0,
        pagosRecibidos: 0,
        pagosHechos: 0,
      };
    } else if (nombreHint && (!sums[pid].personaNombre || sums[pid].personaNombre === "(Sin nombre)")) {
      sums[pid].personaNombre = String(nombreHint).trim();
    }
    return sums[pid];
  };

  // Fijos: aportes
  fijos.forEach(g => {
    const aportes = Array.isArray(g.aportes) ? g.aportes : [];
    aportes.forEach(a => {
      const pid = String(a.personaId || "").trim();
      if (!pid) return;
      const row = ensure(pid, a.personaNombre);
      if (!row) return;
      const m = _gnNum_(a.monto);
      row.fijosAsignados += m;
      if (a.pagado) row.fijosPagados += m;
    });
  });

  // Variables: pendientes + abonos
  vars.forEach(v => {
    const pend = _gnNum_(v.pendiente);
    const abono = _gnNum_(v.abonado);

    const did = String(v.deudorId || "").trim();
    const aid = String(v.acreedorId || "").trim();

    if (did) {
      const row = ensure(did, v.deudorNombre);
      if (row) {
        row.variablesDeudor += pend;
        row.pagosHechos += abono;
      }
    }
    if (aid) {
      const row = ensure(aid, v.acreedorNombre);
      if (row) {
        row.variablesAcreedor += pend;
        row.pagosRecibidos += abono;
      }
    }
  });

  // Lista
  let out = Object.keys(sums).map(id => {
    const r = sums[id];
    const saldoNeto = _gnNum_(r.variablesAcreedor) - _gnNum_(r.variablesDeudor);
    return {
      ...r,
      fijosPendiente: Math.max(0, _gnNum_(r.fijosAsignados) - _gnNum_(r.fijosPagados)),
      saldoNeto,
    };
  });

  // Si se selecciona persona, solo devolvemos esa.
  if (parsed.personaId) {
    out = out.filter(r => String(r.personaId) === String(parsed.personaId));
  }

  // Orden: saldo desc, luego nombre asc
  out.sort((a, b) => {
    const sa = _gnNum_(a.saldoNeto);
    const sb = _gnNum_(b.saldoNeto);
    if (sa !== sb) return sb - sa;
    return String(a.personaNombre || "").localeCompare(String(b.personaNombre || ""));
  });

  return out;
}

/**
 * Endpoint unificado para UI (fase 2): cards + balance en 1 sola llamada.
 */
function getGastosNetosDashboard(params) {
  const parsed = _gnParseParams_(params);
  const ctx = _gnGetPersonContext_();
  const { fijos, vars } = _gnFetchData_(parsed, ctx);

  const cards = _gnComputeCards_(fijos, vars, parsed);
  const balance = _gnComputeBalance_(fijos, vars, parsed, ctx);

  return {
    ok: true,
    cards,
    balance,
    filtrosAplicados: {
      modo: parsed.modo,
      periodo: parsed.periodo || "",
      desde: parsed.desde || "",
      hasta: parsed.hasta || "",
      tipo: parsed.params.tipo || "Todos",
      personaId: parsed.personaId || "",
      incluirPagados: parsed.incluirPagados,
    }
  };
}

function getGastosNetosBalance(params) {
  const parsed = _gnParseParams_(params);
  const ctx = _gnGetPersonContext_();
  const { fijos, vars } = _gnFetchData_(parsed, ctx);
  const balance = _gnComputeBalance_(fijos, vars, parsed, ctx);
  return { ok: true, balance };
}

/**
 * Resumen para cards de Gastos Netos.
 * @param {{modo?:("mes"|"rango"),periodo?:string,desde?:string,hasta?:string,tipo?:string,personaId?:string,incluirPagados?:boolean}} params
 */

function getGastosNetosResumen(params) {
  try {
    const parsed = _gnParseParams_(params);
    const cacheKey = 'gn:resumen:v3:' + _gnHashKey_(JSON.stringify({
      modo: parsed.modo,
      periodo: parsed.periodo,
      desde: parsed.desde,
      hasta: parsed.hasta,
      tipo: parsed.params.tipo || 'Todos',
      personaId: parsed.personaId || '',
      incluirPagados: parsed.incluirPagados,
    }));

    const cached = _gnCacheGet_(cacheKey);
    if (cached && cached.ok) return cached;

    const ctx = _gnGetPersonContext_();
    const { fijos, vars } = _gnFetchData_(parsed, ctx);
    const cards = _gnComputeCards_(fijos, vars, parsed);
    const balancePorPersona = _gnComputeBalance_(fijos, vars, parsed, ctx);

    const out = {
      ok: true,
      cards,
      balancePorPersona,
      filtrosAplicados: {
        modo: parsed.modo,
        periodo: parsed.periodo || '',
        desde: parsed.desde || '',
        hasta: parsed.hasta || '',
        tipo: parsed.params.tipo || 'Todos',
        personaId: parsed.personaId || '',
        incluirPagados: parsed.incluirPagados,
      },
      meta: {
        personNameAmbiguous: (ctx.ambiguousNameKeys || []).length,
        lastUpdateIso: _gnNowIso_(),
      }
    };

    _gnCachePut_(cacheKey, out, 45);
    return out;
  } catch (e) {
    return { ok: false, message: String(e?.message || e) };
  }
}

/**
 * =========================
 *  FASE 3 (Detalle por persona)
 * =========================
 * - Modal por persona con tabs (Resumen / Fijos / Variables / Movimientos)
 * - Endpoint: getGastosNetosDetallePersona
 */

function _gnFmtDate_(iso) {
  const s = String(iso || "").trim();
  if (!s) return "";
  // iso expected: YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  return _gnNormalizeISODate_(s);
}

function _gnPct_(a, b) {
  const x = _gnNum_(a);
  const y = _gnNum_(b);
  const t = x + y;
  if (t <= 0) return 0;
  return Math.max(0, Math.min(100, (x / t) * 100));
}

function _gnBuildDetallePersona_(personaId, fijos, vars, ctx) {
  ctx = ctx || _gnGetPersonContext_();
  const dir = ctx.dir || {};
  const personaNombre = (dir[personaId] && dir[personaId].nombre) ? dir[personaId].nombre : "";

  // Totales
  let fijosAsignados = 0;
  let fijosPagados = 0;

  let varDebeTotal = 0;
  let varDebePagado = 0;
  let varDebePendiente = 0;

  let varFavorTotal = 0;
  let varFavorRecibido = 0;
  let varFavorPendiente = 0;

  const fijosItems = [];
  const varsItems = [];
  const movimientos = [];

  // Fijos donde participa
  (fijos || []).forEach(g => {
    const aportes = Array.isArray(g.aportes) ? g.aportes : [];
    const aporte = aportes.find(a => String(a.personaId || "").trim() === String(personaId));
    if (!aporte) return;

    const monto = _gnNum_(aporte.monto);
    const pagado = !!aporte.pagado;
    const fechaPago = _gnFmtDate_(aporte.fechaPago);

    fijosAsignados += monto;
    if (pagado) fijosPagados += monto;

    const estadoAporte = pagado ? "Pagado" : "Pendiente";

    fijosItems.push({
      gastoId: String(g.id || "").trim(),
      personaNombre: String(aporte.personaNombre || "").trim(),
      servicio: String(g.servicio || "").trim(),
      periodo: String(g.periodo || "").trim(),
      vence: _gnFmtDate_(g.vence),
      proveedor: String(g.proveedor || "").trim(),
      totalFactura: _gnNum_(g.totalFactura),
      aporteMonto: monto,
      aportePagado: pagado,
      fechaPago,
      estadoAporte,
    });

    // Movimientos
    const vence = _gnFmtDate_(g.vence) || "";
    if (vence) {
      movimientos.push({
        fecha: vence,
        tipo: "Fijo",
        titulo: `Aporte asignado • ${String(g.servicio || "Fijo")}`,
        descripcion: `${String(g.proveedor || "").trim()} • ${estadoAporte}`.trim(),
        monto,
      });
    }
    if (pagado) {
      const fp = fechaPago || vence || "";
      if (fp) {
        movimientos.push({
          fecha: fp,
          tipo: "Pago",
          titulo: `Aporte pagado • ${String(g.servicio || "Fijo")}`,
          descripcion: `${String(g.proveedor || "").trim()}`.trim(),
          monto,
        });
      }
    }
  });

  // Variables donde aparece como deudor/acreedor
  (vars || []).forEach(v => {
    const did = String(v.deudorId || "").trim();
    const aid = String(v.acreedorId || "").trim();
    const isDeudor = did === String(personaId);
    const isAcreedor = aid === String(personaId);
    if (!isDeudor && !isAcreedor) return;

    const monto = _gnNum_(v.monto);
    const abonado = _gnNum_(v.abonado);
    const pendiente = _gnNum_(v.pendiente);

    const rol = isDeudor ? "Deudor" : "Acreedor";
    const contraparte = isDeudor ? String(v.acreedorNombre || "").trim() : String(v.deudorNombre || "").trim();

    // Totales separados
    if (isDeudor) {
      varDebeTotal += monto;
      varDebePagado += abonado;
      varDebePendiente += pendiente;
    } else {
      varFavorTotal += monto;
      varFavorRecibido += abonado;
      varFavorPendiente += pendiente;
    }

    const pagos = Array.isArray(v.pagos) ? v.pagos : [];
    varsItems.push({
      gastoId: String(v.id || "").trim(),
      personaNombre: String(isDeudor ? v.deudorNombre : v.acreedorNombre || "").trim(),
      fecha: _gnFmtDate_(v.fecha),
      tipo: String(v.tipo || "").trim(),
      rol,
      contraparte,
      descripcion: String(v.descripcion || "").trim(),
      monto,
      abonado,
      pendiente,
      estado: String(v.estado || "").trim(),
      pagos: pagos.map(p => ({
        id: String(p.id || "").trim(),
        fechaPago: _gnFmtDate_(p.fechaPago),
        monto: _gnNum_(p.monto),
        nota: String(p.nota || "").trim(),
      })),
    });

    // Movimientos
    const f = _gnFmtDate_(v.fecha);
    if (f) {
      movimientos.push({
        fecha: f,
        tipo: "Variable",
        titulo: isDeudor ? `Deuda creada • Debes a ${contraparte || "(sin nombre)"}` : `Deuda creada • Te debe ${contraparte || "(sin nombre)"}`,
        descripcion: `${String(v.tipo || "").trim()} • ${String(v.descripcion || "").trim()}`.trim(),
        monto,
      });
    }

    (pagos || []).forEach(p => {
      const fp = _gnFmtDate_(p.fechaPago);
      if (!fp) return;
      movimientos.push({
        fecha: fp,
        tipo: isDeudor ? "Pago" : "Cobro",
        titulo: isDeudor ? `Pago hecho • a ${contraparte || ""}` : `Pago recibido • de ${contraparte || ""}`,
        descripcion: String(p.nota || "").trim(),
        monto: _gnNum_(p.monto),
      });
    });
  });

  // Resumen
  const fijosPendiente = Math.max(0, fijosAsignados - fijosPagados);
  const pagado = fijosPagados + varDebePagado;
  const pendiente = fijosPendiente + varDebePendiente;

  const fijoVsVar = {
    fijo: fijosAsignados,
    variable: varDebeTotal,
    pctFijo: _gnPct_(fijosAsignados, varDebeTotal),
  };
  const pagadoVsPendiente = {
    pagado,
    pendiente,
    pctPagado: _gnPct_(pagado, pendiente),
  };

  // Orden movimientos: más reciente primero
  movimientos.sort((a, b) => String(b.fecha || "").localeCompare(String(a.fecha || "")));

  // Orden fijos: vence desc
  fijosItems.sort((a, b) => String(b.vence || "").localeCompare(String(a.vence || "")));
  // Orden variables: fecha desc
  varsItems.sort((a, b) => String(b.fecha || "").localeCompare(String(a.fecha || "")));

  return {
    persona: { id: String(personaId), nombre: personaNombre },
    resumen: {
      fijosAsignados,
      fijosPagados,
      fijosPendiente,
      varDebeTotal,
      varDebePagado,
      varDebePendiente,
      varFavorTotal,
      varFavorRecibido,
      varFavorPendiente,
      pagadoVsPendiente,
      fijoVsVar,
    },
    fijos: fijosItems,
    variables: varsItems,
    movimientos,
  };
}

/**
 * Detalle por persona para modal (Fase 3).
 * @param {{personaId:string,modo?:string,periodo?:string,desde?:string,hasta?:string,tipo?:string,incluirPagados?:boolean}} params
 */
function getGastosNetosDetallePersona(params) {
  try {
    params = params || {};
    const personaId = String(params.personaId || "").trim();
    if (!personaId) return { ok: false, message: "personaId es requerido" };

    const parsed = _gnParseParams_(params);
    parsed.personaId = personaId;
    const ctx = _gnGetPersonContext_();
    const { fijos, vars } = _gnFetchData_(parsed, ctx);

    const detalle = _gnBuildDetallePersona_(personaId, fijos, vars, ctx);
    // Fallback nombre si no está en el directorio
    if (!detalle.persona.nombre) {
      const hint = (detalle.fijos[0] && detalle.fijos[0].personaNombre) || (detalle.variables[0] && detalle.variables[0].personaNombre) || "";
      if (hint) detalle.persona.nombre = String(hint).trim();
    }

    return {
      ok: true,
      ...detalle,
      filtrosAplicados: {
        modo: parsed.modo,
        periodo: parsed.periodo || "",
        desde: parsed.desde || "",
        hasta: parsed.hasta || "",
        tipo: parsed.params.tipo || "Todos",
        personaId,
        incluirPagados: parsed.incluirPagados,
      },
      lastUpdateIso: _gnNowIso_(),
    };
  } catch (e) {
    return { ok: false, message: String(e?.message || e) };
  }
}


/**
 * Plan de liquidación (netting) basado en Variables pendientes.
 * Devuelve transferencias sugeridas para minimizar la cantidad de pagos.
 * Nota: Los Gastos Fijos NO generan transferencias automáticas porque no existe un acreedor explícito.
 *
 * @param {{modo?:('mes'|'rango'),periodo?:string,desde?:string,hasta?:string,tipo?:string,personaId?:string,incluirPagados?:boolean}} params
 */
function getGastosNetosPlanLiquidacion(params) {
  try {
    const parsed = _gnParseParams_(params);

    // Si filtran solo fijos, no hay netting entre personas.
    const tipoRaw = String(parsed.params.tipo || 'Todos').trim().toLowerCase();
    if (tipoRaw === 'fijos') {
      return { ok: true, transfers: [], totalMonto: 0, note: 'Plan de liquidación aplica a Variables (deudas entre personas).' };
    }

    const cacheKey = 'gn:plan:v2:' + _gnHashKey_(JSON.stringify({
      modo: parsed.modo,
      periodo: parsed.periodo,
      desde: parsed.desde,
      hasta: parsed.hasta,
      tipo: parsed.params.tipo || 'Todos',
      personaId: String(parsed.personaId || ''),
      incluirPagados: parsed.incluirPagados,
    }));

    const cached = _gnCacheGet_(cacheKey);
    if (cached && cached.ok) return cached;

    const ctx = _gnGetPersonContext_();
    const { vars } = _gnFetchData_(parsed, ctx);

    const eps = 0.005;
    const round2 = (x) => {
      const n = Number(x || 0);
      return Math.round((n + Number.EPSILON) * 100) / 100;
    };

    const saldos = {}; // id => {id,nombre,net}
    const ensure = (id, nombreHint) => {
      const pid = String(id || '').trim();
      if (!pid) return null;
      if (!saldos[pid]) {
        const nombre = (ctx.dir && ctx.dir[pid] && ctx.dir[pid].nombre) ? ctx.dir[pid].nombre : String(nombreHint || '').trim();
        saldos[pid] = { id: pid, nombre: nombre || '(Sin nombre)', net: 0 };
      }
      return saldos[pid];
    };

    (vars || []).forEach(v => {
      const pend = Number(v && v.pendiente) || 0;
      if (pend <= eps) return;
      const did = String(v.deudorId || '').trim();
      const aid = String(v.acreedorId || '').trim();
      if (did) ensure(did, v.deudorNombre).net = round2(ensure(did, v.deudorNombre).net - pend);
      if (aid) ensure(aid, v.acreedorNombre).net = round2(ensure(aid, v.acreedorNombre).net + pend);
    });

    const creditors = [];
    const debtors = [];
    Object.keys(saldos).forEach(id => {
      const s = round2(saldos[id].net);
      if (s > eps) creditors.push({ id, nombre: saldos[id].nombre, amount: s });
      else if (s < -eps) debtors.push({ id, nombre: saldos[id].nombre, amount: Math.abs(s) });
    });

    const sortDesc = (arr) => arr.sort((a, b) => Number(b.amount || 0) - Number(a.amount || 0));
    sortDesc(creditors);
    sortDesc(debtors);

    const transfers = [];
    let guard = 0;
    const guardMax = (creditors.length + debtors.length + 10) * 50;

    while (creditors.length && debtors.length && guard < guardMax) {
      guard++;
      sortDesc(creditors);
      sortDesc(debtors);

      const c = creditors[0];
      const d = debtors[0];
      const pay = round2(Math.min(Number(c.amount || 0), Number(d.amount || 0)));
      if (pay <= eps) break;

      transfers.push({
        fromPersonaId: d.id,
        fromPersonaNombre: d.nombre,
        toPersonaId: c.id,
        toPersonaNombre: c.nombre,
        monto: pay,
      });

      c.amount = round2(Number(c.amount || 0) - pay);
      d.amount = round2(Number(d.amount || 0) - pay);
      if (c.amount <= eps) creditors.shift();
      if (d.amount <= eps) debtors.shift();
    }

    let filteredTransfers = transfers;
    const personaId = String(parsed.personaId || '').trim();
    if (personaId) {
      filteredTransfers = transfers.filter(t => String(t.fromPersonaId) === personaId || String(t.toPersonaId) === personaId);
    }

    const totalMonto = round2(filteredTransfers.reduce((acc, t) => acc + Number(t.monto || 0), 0));

    const out = {
      ok: true,
      transfers: filteredTransfers,
      totalMonto,
      note: 'Basado en Variables pendientes (deudas entre personas).',
      filtrosAplicados: {
        modo: parsed.modo,
        periodo: parsed.periodo || '',
        desde: parsed.desde || '',
        hasta: parsed.hasta || '',
        tipo: parsed.params.tipo || 'Todos',
        personaId: personaId,
      },
      lastUpdateIso: _gnNowIso_(),
    };

    _gnCachePut_(cacheKey, out, 45);
    return out;
  } catch (e) {
    return { ok: false, message: String(e?.message || e) };
  }
}
