/**
 * =========================
 *  GASTOS FIJOS (recurrentes)
 * =========================
 *
 * Objetivo:
 *  - Registrar facturas recurrentes por periodo (mes/año)
 *  - Repartir aportes por persona con montos exactos
 *  - Mantener un modelo simple y amigable para Google Sheets
 *
 * Hojas (configurables en env_()):
 *  - SH_GASTOS_FIJOS            (default: "GASTOS_FIJOS")
 *  - SH_GASTOS_FIJOS_DETALLE    (default: "DETALLE_GASTO_FIJO")
 *
 * Encabezados recomendados:
 *  GASTOS_FIJOS:
 *    id, servicio, periodo, vence, proveedor, descripcion, totalFactura, creadoPor, fechaCreacion
 *
 *  DETALLE_GASTO_FIJO:
 *    id, gastoId, personaId, personaNombre, monto, pagado, fechaPago
 */

function _gfSheetName_(key, fallback) {
  try {
    if (typeof env_ === "function") {
      const env = env_();
      const val = env && env[key];
      if (val) return String(val);
    }
  } catch (e) {}
  return String(fallback);
}

function _gfNowIso_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "America/Tegucigalpa", "yyyy-MM-dd'T'HH:mm:ss");
}

function _gfEnsureSheet_(name, headers) {
  const ss = conexion();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);

  const lastCol = sh.getLastColumn();
  const headerRange = sh.getRange(1, 1, 1, Math.max(lastCol, headers.length));
  const existing = headerRange.getValues()[0].map(v => String(v || "").trim());

  // Si está vacío, set completo
  if (existing.every(v => !v)) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    return sh;
  }

  // Agrega headers faltantes al final
  const lower = existing.map(h => h.toLowerCase());
  const toAdd = headers.filter(h => lower.indexOf(String(h).toLowerCase()) === -1);
  if (toAdd.length) {
    sh.getRange(1, existing.length + 1, 1, toAdd.length).setValues([toAdd]);
  }

  return sh;
}

function _gfRead_(sh) {
  // IMPORTANT: Usamos *DisplayValues* para evitar problemas de tipo.
  // - En Sheets, columnas como "periodo" o "vence" pueden guardarse como Date y
  //   mostrarse como "2025-12" / "2026-01-28".
  // - getValues() devolvería Date objects y al convertirlos a String(), el filtro
  //   por periodo no coincide (y la UI queda "Sin registros" aunque existan).
  const values = sh.getDataRange().getDisplayValues();

  // Puede existir solo la fila de encabezados (values.length === 1).
  if (!values || values.length === 0) return { headers: [], rows: [], idx: {} };

  const headers = (values[0] || []).map(h => String(h || "").trim());
  const idx = {};
  headers.forEach((h, i) => { if (h) idx[h] = i; });

  const rows = (values.length > 1)
    ? values.slice(1).filter(r => r.some(c => String(c || "").trim() !== ""))
    : [];

  return { headers, rows, idx };
}

// Fuerza fechas ISO (YYYY-MM-DD) a texto para evitar que Sheets lo convierta en Date.
function _gfDateText_(iso) {
  const s = String(iso || "").trim();
  if (!s) return "";
  if (/^'/.test(s)) return s;
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return "'" + s;
  return s;
}

// Fuerza periodo ISO (YYYY-MM) a texto para evitar coerción.
function _gfMonthText_(yyyyMm) {
  const s = String(yyyyMm || "").trim();
  if (!s) return "";
  if (/^'/.test(s)) return s;
  if (/^\d{4}-\d{2}$/.test(s)) return "'" + s;
  return s;
}

// Normaliza fechas a ISO (YYYY-MM-DD) para poder comparar/parsear aunque venga como dd/mm/yyyy.
function _gfNormalizeISODate_(raw) {
  const s = String(raw || "").trim();
  if (!s) return "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;

  // dd/mm/yyyy o mm/dd/yyyy
  const m1 = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m1) {
    const a = parseInt(m1[1], 10);
    const b = parseInt(m1[2], 10);
    const y = parseInt(m1[3], 10);
    // Preferimos dd/mm para es-HN
    let d = a, m = b;
    if (a <= 12 && b > 12) { // mm/dd
      m = a; d = b;
    }
    const mm = String(m).padStart(2, "0");
    const dd = String(d).padStart(2, "0");
    return `${y}-${mm}-${dd}`;
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

// Normaliza periodo a YYYY-MM aunque venga como Date, "mm/yyyy" o similar.
function _gfNormalizeYYYYMM_(raw) {
  const s = String(raw || "").trim();
  if (!s) return "";
  if (/^\d{4}-\d{2}$/.test(s)) return s;
  const m1 = s.match(/^(\d{1,2})\/(\d{4})$/); // mm/yyyy
  if (m1) {
    const mm = String(parseInt(m1[1], 10)).padStart(2, "0");
    const yy = String(parseInt(m1[2], 10));
    return `${yy}-${mm}`;
  }
  // Si viene una fecha completa, tomamos año-mes
  const iso = _gfNormalizeISODate_(s);
  if (/^\d{4}-\d{2}-\d{2}$/.test(iso)) return iso.slice(0, 7);
  // Intento final: Date parse
  const dt = new Date(s);
  if (dt.toString() !== "Invalid Date") {
    const y = dt.getFullYear();
    const m = String(dt.getMonth() + 1).padStart(2, "0");
    return `${y}-${m}`;
  }
  return s;
}

function _gfGet_(row, idx, name) {
  const i = idx[name];
  return (i === undefined) ? "" : row[i];
}

function _gfFindRowIndexById_(rows, idx, id) {
  const idCol = idx["id"]; // requerido
  if (idCol === undefined) return -1;
  const needle = String(id || "").trim();
  if (!needle) return -1;
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][idCol] || "").trim() === needle) return i;
  }
  return -1;
}

function _gfNum_(v) {
  const n = (typeof v === "number") ? v : parseFloat(String(v || "").replace(/[^0-9.\-]/g, ""));
  return isNaN(n) ? 0 : n;
}

function _gfBool_(v) {
  if (typeof v === "boolean") return v;
  const s = String(v || "").trim().toLowerCase();
  return s === "true" || s === "1" || s === "si" || s === "sí" || s === "yes";
}

/**
 * Devuelve una lista simple de personas aportantes (desde USUARIOS).
 * Nota: por ahora reutilizamos USUARIOS como "personas".
 */
function listarUsuariosActivos() {
  const all = (typeof _readUsers_ === "function") ? _readUsers_() : [];
  return all
    .filter(u => {
      const estado = String(u.estado || "").trim().toLowerCase();
      return !estado || estado === "activo";
    })
    .map(u => ({
      id: String(u.id || ""),
      nombreCompleto: String(u.nombreCompleto || ""),
      nombre_usuario: String(u.nombre_usuario || "")
    }));
}

/**
 * Lista gastos fijos con aportes.
 * @param {{servicio?:string,periodo?:string,estado?:string,venceFrom?:string,venceTo?:string,personaId?:string,q?:string}} params
 */
function listarGastosFijos(params) {
  params = params || {};
  const shName = _gfSheetName_("SH_GASTOS_FIJOS", "GASTOS_FIJOS");
  const detName = _gfSheetName_("SH_GASTOS_FIJOS_DETALLE", "DETALLE_GASTO_FIJO");

  let sh;
  try {
    sh = obtenerSheet(shName);
  } catch (e) {
    return [];
  }

  let shDet = null;
  try { shDet = obtenerSheet(detName); } catch (e) { shDet = null; }

  const main = _gfRead_(sh);
  const det = shDet ? _gfRead_(shDet) : { headers: [], rows: [], idx: {} };

  const aportesByGasto = {};
  if (shDet && det.rows.length) {
    det.rows.forEach(r => {
      const gastoId = String(_gfGet_(r, det.idx, "gastoId") || "").trim();
      if (!gastoId) return;
      if (!aportesByGasto[gastoId]) aportesByGasto[gastoId] = [];
      aportesByGasto[gastoId].push({
        id: String(_gfGet_(r, det.idx, "id") || "").trim(),
        gastoId,
        personaId: String(_gfGet_(r, det.idx, "personaId") || "").trim(),
        personaNombre: String(_gfGet_(r, det.idx, "personaNombre") || "").trim(),
        monto: _gfNum_(_gfGet_(r, det.idx, "monto")),
        pagado: _gfBool_(_gfGet_(r, det.idx, "pagado")),
        fechaPago: String(_gfGet_(r, det.idx, "fechaPago") || "").trim(),
      });
    });
  }

  const filtroServicio = String(params.servicio || "").trim().toLowerCase();
  const filtroId = String(params.id || "").trim();
  const filtroPeriodo = _gfNormalizeYYYYMM_(params.periodo);
  const filtroVenceFrom = _gfNormalizeISODate_(params.venceFrom || params.desde || params.fechaFrom || params.fechaInicio || "");
  const filtroVenceTo = _gfNormalizeISODate_(params.venceTo || params.hasta || params.fechaTo || params.fechaFin || "");
  const filtroEstado = String(params.estado || "").trim().toLowerCase();
  const filtroPersona = String(params.personaId || "").trim();
  const q = String(params.q || "").trim().toLowerCase();

  const out = [];
  main.rows.forEach(r => {
    const id = String(_gfGet_(r, main.idx, "id") || "").trim();
    if (!id) return;

    if (filtroId && id !== filtroId) return;

    const servicio = String(_gfGet_(r, main.idx, "servicio") || "").trim();
    const periodo = _gfNormalizeYYYYMM_(_gfGet_(r, main.idx, "periodo"));
    const vence = _gfNormalizeISODate_(_gfGet_(r, main.idx, "vence"));
    const proveedor = String(_gfGet_(r, main.idx, "proveedor") || "").trim();
    const descripcion = String(_gfGet_(r, main.idx, "descripcion") || "").trim();
    const totalFactura = _gfNum_(_gfGet_(r, main.idx, "totalFactura"));

    const aportes = (aportesByGasto[id] || []).slice();
    const aportado = aportes.reduce((acc, a) => acc + (a.pagado ? _gfNum_(a.monto) : 0), 0);
    const pendiente = Math.max(0, totalFactura - aportado);

    // Estado derivado
    const hoy = new Date();
    const venceDate = vence ? new Date(vence) : null;
    const vencido = venceDate && venceDate.toString() !== "Invalid Date" && venceDate < new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate());
    let estado = "Pendiente";
    if (pendiente <= 0.000001 && totalFactura > 0) estado = "Pagado";
    else if (aportado > 0.000001) estado = "Parcial";
    if (vencido && estado !== "Pagado") estado = "Vencido";

    // Filtros
    if (filtroServicio && servicio.toLowerCase() !== filtroServicio) return;
    if (filtroPeriodo && periodo !== filtroPeriodo) return;

    // Rango de fechas por VENCE (YYYY-MM-DD)
    if (filtroVenceFrom) {
      if (!vence || String(vence) < String(filtroVenceFrom)) return;
    }
    if (filtroVenceTo) {
      if (!vence || String(vence) > String(filtroVenceTo)) return;
    }

    // IMPORTANTE (UX): nunca “desaparecer” registros pagados de la tabla.
    // Si el usuario filtra por un estado distinto a "Pagado", igual incluimos los "Pagado"
    // para mantener historial visible. Si filtra por "Pagado", entonces sí mostramos solo pagados.
    if (filtroEstado) {
      const e = String(estado || "").toLowerCase();
      if (filtroEstado === "pagado") {
        if (e !== "pagado") return;
      } else {
        if (e !== filtroEstado && e !== "pagado") return;
      }
    }
    if (filtroPersona) {
      const has = aportes.some(a => String(a.personaId || "").trim() === filtroPersona);
      if (!has) return;
    }
    if (q) {
      const hay = (proveedor + " " + descripcion + " " + servicio).toLowerCase();
      if (hay.indexOf(q) === -1) return;
    }

    out.push({
      id,
      servicio,
      periodo,
      vence,
      proveedor,
      descripcion,
      totalFactura,
      aportado,
      pendiente,
      estado,
      aportes,
      creadoPor: String(_gfGet_(r, main.idx, "creadoPor") || "").trim(),
      fechaCreacion: String(_gfGet_(r, main.idx, "fechaCreacion") || "").trim(),
    });
  });

  // Orden: vence asc, proveedor asc
  out.sort((a, b) => {
    const da = a.vence ? new Date(a.vence).getTime() : 0;
    const db = b.vence ? new Date(b.vence).getTime() : 0;
    if (da !== db) return da - db;
    return String(a.proveedor || "").localeCompare(String(b.proveedor || ""));
  });

  return out;
}

/**
 * Crea o actualiza un gasto fijo.
 * @param {object} payload
 */
function guardarGastoFijo(payload) {
  payload = payload || {};
  const shName = _gfSheetName_("SH_GASTOS_FIJOS", "GASTOS_FIJOS");
  const detName = _gfSheetName_("SH_GASTOS_FIJOS_DETALLE", "DETALLE_GASTO_FIJO");

  const sh = _gfEnsureSheet_(shName, [
    "id",
    "servicio",
    "periodo",
    "vence",
    "proveedor",
    "descripcion",
    "totalFactura",
    "creadoPor",
    "fechaCreacion"
  ]);

  const shDet = _gfEnsureSheet_(detName, [
    "id",
    "gastoId",
    "personaId",
    "personaNombre",
    "monto",
    "pagado",
    "fechaPago"
  ]);

  const main = _gfRead_(sh);
  const det = _gfRead_(shDet);

  const id = String(payload.id || "").trim() || Utilities.getUuid();
  const servicio = String(payload.servicio || "").trim();
  const periodo = _gfNormalizeYYYYMM_(payload.periodo);
  const vence = _gfNormalizeISODate_(payload.vence);
  const proveedor = String(payload.proveedor || "").trim();
  const descripcion = String(payload.descripcion || "").trim();
  const totalFactura = _gfNum_(payload.totalFactura);
  const creadoPor = String(payload.creadoPor || "").trim();

  // Upsert main
  const rowIndex0 = _gfFindRowIndexById_(main.rows, main.idx, id);
  const writeRow = new Array(main.headers.length).fill("");

  const setIf = (name, value) => {
    const i = main.idx[name];
    if (i === undefined) return;
    writeRow[i] = value;
  };

  setIf("id", id);
  setIf("servicio", servicio);
  // Guardamos como texto para evitar que Sheets los convierta a Date
  // y cambie su representación.
  setIf("periodo", _gfMonthText_(periodo));
  setIf("vence", _gfDateText_(vence));
  setIf("proveedor", proveedor);
  setIf("descripcion", descripcion);
  setIf("totalFactura", totalFactura);
  setIf("creadoPor", creadoPor);
  setIf("fechaCreacion", _gfNowIso_());

  if (rowIndex0 >= 0) {
    // Actualiza manteniendo fechaCreacion previa si existe
    const oldRow = main.rows[rowIndex0];
    const oldFecha = _gfGet_(oldRow, main.idx, "fechaCreacion");
    if (oldFecha) setIf("fechaCreacion", oldFecha);
    sh.getRange(rowIndex0 + 2, 1, 1, main.headers.length).setValues([writeRow]);
  } else {
    sh.appendRow(writeRow);
  }

  // Upsert aportes
  const aportesIn = Array.isArray(payload.aportes) ? payload.aportes : [];
  const keepIds = {};

  // Índice id -> rowIndex0 (en det.rows)
  const detIdCol = det.idx["id"];
  const detGastoCol = det.idx["gastoId"];
  const detRows = det.rows;

  const findDetRow0 = (aporteId) => {
    if (detIdCol === undefined) return -1;
    const needle = String(aporteId || "").trim();
    if (!needle) return -1;
    for (let i = 0; i < detRows.length; i++) {
      if (String(detRows[i][detIdCol] || "").trim() === needle) return i;
    }
    return -1;
  };

  aportesIn.forEach(a => {
    const aId = String(a.id || "").trim() || Utilities.getUuid();
    keepIds[aId] = true;

    const personaId = String(a.personaId || "").trim();
    const personaNombre = String(a.personaNombre || "").trim();
    const monto = _gfNum_(a.monto);
    const pagado = !!a.pagado;
    const fechaPago = String(a.fechaPago || "").trim();

    const row = new Array(det.headers.length).fill("");
    const set = (name, value) => {
      const i = det.idx[name];
      if (i === undefined) return;
      row[i] = value;
    };
    set("id", aId);
    set("gastoId", id);
    set("personaId", personaId);
    set("personaNombre", personaNombre);
    set("monto", monto);
    set("pagado", pagado);
    set("fechaPago", fechaPago);

    const i0 = findDetRow0(aId);
    if (i0 >= 0) {
      shDet.getRange(i0 + 2, 1, 1, det.headers.length).setValues([row]);
    } else {
      shDet.appendRow(row);
    }
  });

  // Elimina aportes removidos (solo los del gastoId)
  if (detGastoCol !== undefined && detIdCol !== undefined) {
    const toDelete = [];
    for (let i = 0; i < detRows.length; i++) {
      const gid = String(detRows[i][detGastoCol] || "").trim();
      if (gid !== id) continue;
      const aid = String(detRows[i][detIdCol] || "").trim();
      if (aid && !keepIds[aid]) toDelete.push(i + 2); // 1-based row number
    }
    // Borrar desde abajo hacia arriba
    toDelete.sort((a, b) => b - a).forEach(rn => shDet.deleteRow(rn));
  }

  // Asegura consistencia: forzar flush para que lecturas inmediatas (desde el frontend)
  // reflejen el último registro. Esto evita el bug donde el 2do registro no aparece.
  SpreadsheetApp.flush();

  return { ok: true, id };
}

function eliminarGastoFijo(id) {
  id = String(id || "").trim();
  if (!id) return { ok: false, message: "ID vacío" };

  const shName = _gfSheetName_("SH_GASTOS_FIJOS", "GASTOS_FIJOS");
  const detName = _gfSheetName_("SH_GASTOS_FIJOS_DETALLE", "DETALLE_GASTO_FIJO");

  let sh;
  try { sh = obtenerSheet(shName); } catch (e) { return { ok: true }; }
  let shDet = null;
  try { shDet = obtenerSheet(detName); } catch (e) { shDet = null; }

  const main = _gfRead_(sh);
  const row0 = _gfFindRowIndexById_(main.rows, main.idx, id);
  if (row0 >= 0) sh.deleteRow(row0 + 2);

  if (shDet) {
    const det = _gfRead_(shDet);
    const gcol = det.idx["gastoId"];
    if (gcol !== undefined) {
      const toDelete = [];
      det.rows.forEach((r, i) => {
        if (String(r[gcol] || "").trim() === id) toDelete.push(i + 2);
      });
      toDelete.sort((a, b) => b - a).forEach(rn => shDet.deleteRow(rn));
    }
  }

  return { ok: true };
}

function marcarAportePagado(aporteId, pagado) {
  aporteId = String(aporteId || "").trim();
  if (!aporteId) return { ok: false, message: "ID de aporte vacío" };

  const detName = _gfSheetName_("SH_GASTOS_FIJOS_DETALLE", "DETALLE_GASTO_FIJO");
  let shDet;
  try { shDet = obtenerSheet(detName); } catch (e) { return { ok: false, message: "No existe hoja de detalle" }; }

  const det = _gfRead_(shDet);
  const row0 = _gfFindRowIndexById_(det.rows, det.idx, aporteId);
  if (row0 < 0) return { ok: false, message: "Aporte no encontrado" };

  const pagadoCol = det.idx["pagado"];
  const fechaCol = det.idx["fechaPago"];
  if (pagadoCol !== undefined) shDet.getRange(row0 + 2, pagadoCol + 1).setValue(!!pagado);
  if (fechaCol !== undefined) shDet.getRange(row0 + 2, fechaCol + 1).setValue(pagado ? _gfNowIso_().slice(0, 10) : "");

  SpreadsheetApp.flush();

  return { ok: true };
}

function marcarGastoFijoPagado(gastoId) {
  gastoId = String(gastoId || "").trim();
  if (!gastoId) return { ok: false, message: "ID vacío" };

  const detName = _gfSheetName_("SH_GASTOS_FIJOS_DETALLE", "DETALLE_GASTO_FIJO");
  const shDet = _gfEnsureSheet_(detName, [
    "id",
    "gastoId",
    "personaId",
    "personaNombre",
    "monto",
    "pagado",
    "fechaPago"
  ]);

  const det = _gfRead_(shDet);
  const gcol = det.idx["gastoId"];
  const pcol = det.idx["pagado"];
  const fcol = det.idx["fechaPago"];
  if (gcol === undefined || pcol === undefined) return { ok: false, message: "Encabezados faltantes" };

  const today = _gfNowIso_().slice(0, 10);
  let touched = 0;
  det.rows.forEach((r, i) => {
    if (String(r[gcol] || "").trim() !== gastoId) return;
    shDet.getRange(i + 2, pcol + 1).setValue(true);
    if (fcol !== undefined) shDet.getRange(i + 2, fcol + 1).setValue(today);
    touched++;
  });

  SpreadsheetApp.flush();

  // Si no había aportes, no inventamos montos; solo marcamos sin detalle.
  return { ok: true, updated: touched };
}
