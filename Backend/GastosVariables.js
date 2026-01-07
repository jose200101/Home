/**
 * =========================
 *  GASTOS VARIABLES (deudas)
 * =========================
 *
 * Objetivo:
 *  - Registrar gastos variados (supermercado, Amazon, envíos, préstamos, etc.)
 *  - Cada registro representa una deuda: DEUDOR le debe al ACREEDOR
 *  - Soporta pagos parciales mediante una hoja de detalle (abonos)
 *
 * Hojas (configurables en env_()):
 *  - SH_GASTOS_VARIABLES            (default: "GASTOS_VARIABLES")
 *  - SH_GASTOS_VARIABLES_DETALLE    (default: "DETALLE_GASTO_VARIABLE")
 *
 * Encabezados recomendados:
 *  GASTOS_VARIABLES:
 *    id, tipo, fecha, deudorId, deudorNombre, acreedorId, acreedorNombre,
 *    descripcion, monto, creadoPor, fechaCreacion
 *
 *  DETALLE_GASTO_VARIABLE:
 *    id, gastoId, monto, fechaPago, nota, registradoPor
 */

function _gvSheetName_(key, fallback) {
  try {
    if (typeof env_ === "function") {
      const env = env_();
      const val = env && env[key];
      if (val) return String(val);
    }
  } catch (e) {}
  return String(fallback);
}

function _gvNowIso_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "America/Tegucigalpa", "yyyy-MM-dd'T'HH:mm:ss");
}

// Fuerza fechas ISO (YYYY-MM-DD) a texto para evitar que Sheets lo convierta en Date.
function _gvDateText_(iso) {
  const s = String(iso || "").trim();
  if (!s) return "";
  if (/^'/.test(s)) return s;
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return "'" + s;
  return s;
}

function _gvEnsureSheet_(name, headers) {
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

function _gvRead_(sh) {
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

function _gvGet_(row, idx, name) {
  const i = idx[name];
  return (i === undefined) ? "" : row[i];
}

function _gvFindRowIndexById_(rows, idx, id) {
  const idCol = idx["id"];
  if (idCol === undefined) return -1;
  const needle = String(id || "").trim();
  if (!needle) return -1;
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][idCol] || "").trim() === needle) return i;
  }
  return -1;
}

function _gvNum_(v) {
  const n = (typeof v === "number") ? v : parseFloat(String(v || "").replace(/[^0-9.\-]/g, ""));
  return isNaN(n) ? 0 : n;
}

function _gvNormalizeISODate_(raw) {
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

/**
 * Lista gastos variables con pagos.
 * @param {{id?:string,tipo?:string,deudorId?:string,acreedorId?:string,fechaFrom?:string,fechaTo?:string,q?:string,estado?:string}} params
 */
function listarGastosVariables(params) {
  params = params || {};

  const shName = _gvSheetName_("SH_GASTOS_VARIABLES", "GASTOS_VARIABLES");
  const detName = _gvSheetName_("SH_GASTOS_VARIABLES_DETALLE", "DETALLE_GASTO_VARIABLE");

  let sh;
  try { sh = obtenerSheet(shName); } catch (e) { return []; }

  let shDet = null;
  try { shDet = obtenerSheet(detName); } catch (e) { shDet = null; }

  const main = _gvRead_(sh);
  const det = shDet ? _gvRead_(shDet) : { headers: [], rows: [], idx: {} };

  const pagosByGasto = {};
  if (shDet && det.rows.length) {
    det.rows.forEach(r => {
      const gastoId = String(_gvGet_(r, det.idx, "gastoId") || "").trim();
      if (!gastoId) return;
      if (!pagosByGasto[gastoId]) pagosByGasto[gastoId] = [];
      pagosByGasto[gastoId].push({
        id: String(_gvGet_(r, det.idx, "id") || "").trim(),
        gastoId,
        monto: _gvNum_(_gvGet_(r, det.idx, "monto")),
        fechaPago: _gvNormalizeISODate_(_gvGet_(r, det.idx, "fechaPago")),
        nota: String(_gvGet_(r, det.idx, "nota") || "").trim(),
        registradoPor: String(_gvGet_(r, det.idx, "registradoPor") || "").trim(),
      });
    });
  }

  // Filtros
  const filtroId = String(params.id || "").trim();
  const filtroTipo = String(params.tipo || "").trim().toLowerCase();
  const filtroDeudorId = String(params.deudorId || "").trim();
  const filtroAcreedorId = String(params.acreedorId || "").trim();
  const fechaFrom = _gvNormalizeISODate_(params.fechaFrom);
  const fechaTo = _gvNormalizeISODate_(params.fechaTo);
  const filtroEstado = String(params.estado || "").trim().toLowerCase();
  const q = String(params.q || "").trim().toLowerCase();

  const out = [];
  main.rows.forEach(r => {
    const id = String(_gvGet_(r, main.idx, "id") || "").trim();
    if (!id) return;
    if (filtroId && id !== filtroId) return;

    const tipo = String(_gvGet_(r, main.idx, "tipo") || "").trim();
    const fecha = _gvNormalizeISODate_(_gvGet_(r, main.idx, "fecha"));
    const deudorId = String(_gvGet_(r, main.idx, "deudorId") || "").trim();
    const deudorNombre = String(_gvGet_(r, main.idx, "deudorNombre") || "").trim();
    const acreedorId = String(_gvGet_(r, main.idx, "acreedorId") || "").trim();
    const acreedorNombre = String(_gvGet_(r, main.idx, "acreedorNombre") || "").trim();
    const descripcion = String(_gvGet_(r, main.idx, "descripcion") || "").trim();
    const monto = _gvNum_(_gvGet_(r, main.idx, "monto"));
    const creadoPor = String(_gvGet_(r, main.idx, "creadoPor") || "").trim();
    const fechaCreacion = String(_gvGet_(r, main.idx, "fechaCreacion") || "").trim();

    const pagos = (pagosByGasto[id] || []).slice().sort((a, b) => String(a.fechaPago || "").localeCompare(String(b.fechaPago || "")));
    const abonado = pagos.reduce((acc, p) => acc + _gvNum_(p.monto), 0);
    const pendiente = Math.max(0, monto - abonado);
    let estado = "Pendiente";
    if (monto > 0 && pendiente <= 0.000001) estado = "Pagado";
    else if (abonado > 0.000001) estado = "Parcial";

    // Aplicar filtros
    if (filtroTipo && tipo.toLowerCase() !== filtroTipo) return;
    if (filtroDeudorId && deudorId !== filtroDeudorId) return;
    if (filtroAcreedorId && acreedorId !== filtroAcreedorId) return;
    if (fechaFrom && fecha && fecha < fechaFrom) return;
    if (fechaTo && fecha && fecha > fechaTo) return;
    if (filtroEstado && estado.toLowerCase() !== filtroEstado) return;
    if (q) {
      const hay = (tipo + " " + descripcion + " " + deudorNombre + " " + acreedorNombre).toLowerCase();
      if (hay.indexOf(q) === -1) return;
    }

    out.push({
      id,
      tipo,
      fecha,
      deudorId,
      deudorNombre,
      acreedorId,
      acreedorNombre,
      descripcion,
      monto,
      creadoPor,
      fechaCreacion,
      pagos,
      abonado,
      pendiente,
      estado,
    });
  });

  // Orden: fecha desc, luego tipo
  out.sort((a, b) => {
    const fa = String(a.fecha || "");
    const fb = String(b.fecha || "");
    if (fa !== fb) return fb.localeCompare(fa);
    return String(a.tipo || "").localeCompare(String(b.tipo || ""));
  });

  return out;
}

/**
 * Crea o actualiza un gasto variable.
 * @param {object} payload
 */
function guardarGastoVariable(payload) {
  payload = payload || {};
  const shName = _gvSheetName_("SH_GASTOS_VARIABLES", "GASTOS_VARIABLES");
  const detName = _gvSheetName_("SH_GASTOS_VARIABLES_DETALLE", "DETALLE_GASTO_VARIABLE");

  const sh = _gvEnsureSheet_(shName, [
    "id",
    "tipo",
    "fecha",
    "deudorId",
    "deudorNombre",
    "acreedorId",
    "acreedorNombre",
    "descripcion",
    "monto",
    "creadoPor",
    "fechaCreacion",
  ]);

  // Aseguramos hoja de detalle (para pagos)
  _gvEnsureSheet_(detName, [
    "id",
    "gastoId",
    "monto",
    "fechaPago",
    "nota",
    "registradoPor",
  ]);

  const main = _gvRead_(sh);

  const id = String(payload.id || "").trim() || Utilities.getUuid();
  const tipo = String(payload.tipo || "").trim();
  const fecha = _gvNormalizeISODate_(payload.fecha);
  const deudorId = String(payload.deudorId || "").trim();
  const deudorNombre = String(payload.deudorNombre || "").trim();
  const acreedorId = String(payload.acreedorId || "").trim();
  const acreedorNombre = String(payload.acreedorNombre || "").trim();
  const descripcion = String(payload.descripcion || "").trim();
  const monto = _gvNum_(payload.monto);
  const creadoPor = String(payload.creadoPor || "").trim();

  const rowIndex0 = _gvFindRowIndexById_(main.rows, main.idx, id);
  const writeRow = new Array(main.headers.length).fill("");

  const setIf = (name, value) => {
    const i = main.idx[name];
    if (i === undefined) return;
    writeRow[i] = value;
  };

  setIf("id", id);
  setIf("tipo", tipo);
  setIf("fecha", _gvDateText_(fecha));
  setIf("deudorId", deudorId);
  setIf("deudorNombre", deudorNombre);
  setIf("acreedorId", acreedorId);
  setIf("acreedorNombre", acreedorNombre);
  setIf("descripcion", descripcion);
  setIf("monto", monto);
  setIf("creadoPor", creadoPor);
  setIf("fechaCreacion", _gvNowIso_());

  if (rowIndex0 >= 0) {
    const oldRow = main.rows[rowIndex0];
    const oldFecha = _gvGet_(oldRow, main.idx, "fechaCreacion");
    if (oldFecha) setIf("fechaCreacion", oldFecha);
    sh.getRange(rowIndex0 + 2, 1, 1, main.headers.length).setValues([writeRow]);
  } else {
    sh.appendRow(writeRow);
  }

  // Evita lecturas inconsistentes inmediatamente después de guardar
  // (bug: el 2do registro a veces no aparece hasta refrescar).
  SpreadsheetApp.flush();

  return { ok: true, id };
}

/**
 * Registra un pago/abono.
 * @param {{gastoId:string,monto:number|string,fechaPago:string,nota?:string,registradoPor?:string}} payload
 */
function registrarPagoGastoVariable(payload) {
  payload = payload || {};
  const gastoId = String(payload.gastoId || "").trim();
  if (!gastoId) return { ok: false, message: "gastoId vacío" };

  const detName = _gvSheetName_("SH_GASTOS_VARIABLES_DETALLE", "DETALLE_GASTO_VARIABLE");
  const shDet = _gvEnsureSheet_(detName, [
    "id",
    "gastoId",
    "monto",
    "fechaPago",
    "nota",
    "registradoPor",
  ]);

  const id = Utilities.getUuid();
  const monto = _gvNum_(payload.monto);
  const fechaPago = _gvNormalizeISODate_(payload.fechaPago);
  const nota = String(payload.nota || "").trim();
  const registradoPor = String(payload.registradoPor || "").trim();

  const row = new Array(shDet.getLastColumn()).fill("");
  // armamos por headers existentes
  const det = _gvRead_(shDet);
  const set = (name, value) => {
    const i = det.idx[name];
    if (i === undefined) return;
    row[i] = value;
  };

  set("id", id);
  set("gastoId", gastoId);
  set("monto", monto);
  set("fechaPago", _gvDateText_(fechaPago));
  set("nota", nota);
  set("registradoPor", registradoPor);

  // Ajustar tamaño a headers
  const finalRow = row.slice(0, det.headers.length);
  shDet.appendRow(finalRow);

  SpreadsheetApp.flush();

  return { ok: true, id };
}

function eliminarGastoVariable(id) {
  id = String(id || "").trim();
  if (!id) return { ok: false, message: "ID vacío" };

  const shName = _gvSheetName_("SH_GASTOS_VARIABLES", "GASTOS_VARIABLES");
  const detName = _gvSheetName_("SH_GASTOS_VARIABLES_DETALLE", "DETALLE_GASTO_VARIABLE");

  let sh;
  try { sh = obtenerSheet(shName); } catch (e) { return { ok: true }; }
  let shDet = null;
  try { shDet = obtenerSheet(detName); } catch (e) { shDet = null; }

  const main = _gvRead_(sh);
  const idx0 = _gvFindRowIndexById_(main.rows, main.idx, id);
  if (idx0 >= 0) sh.deleteRow(idx0 + 2);

  if (shDet) {
    const det = _gvRead_(shDet);
    const idCol = det.idx["id"];
    const gastoCol = det.idx["gastoId"];
    if (idCol !== undefined && gastoCol !== undefined) {
      const toDelete = [];
      for (let i = 0; i < det.rows.length; i++) {
        const gid = String(det.rows[i][gastoCol] || "").trim();
        if (gid === id) toDelete.push(i + 2);
      }
      toDelete.sort((a, b) => b - a).forEach(rn => shDet.deleteRow(rn));
    }
  }

  return { ok: true };
}
