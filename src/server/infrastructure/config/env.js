/**
 * env
 * Configuración centralizada (IDs y nombres de hojas).
 *
 * SPREADSHEET_ID es obligatorio en Script Properties de cada proyecto.
 * No se usan IDs predeterminados ni el Spreadsheet activo como respaldo.
 */
function env_() {
  const propsId = String(PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID") || "").trim();
  if (!propsId || !/^[A-Za-z0-9_-]+$/.test(propsId)) {
    throw new Error("Configura SPREADSHEET_ID en Script Properties con el ID explícito del Spreadsheet de este entorno.");
  }

  return {
    ID_DATABASE: propsId,

    // HOJA REGISTRO USUARIOS (en tu captura: "USUARIOS")
    SH_REGISTRO_USUARIOS: "USUARIOS",

    // HOJA PERSONAS (en tu captura: "PERSONAS")
    SH_PERSONAS: "PERSONAS",

    // =========================
    //  GASTOS (módulo - facturas)
    // =========================
    // Gastos fijos (facturas recurrentes por mes/año)
    SH_GASTOS_FIJOS: "GASTOS_FIJOS",
    // Detalle de aportes por persona (montos exactos)
    SH_GASTOS_FIJOS_DETALLE: "DETALLE_GASTO_FIJO",


    // =========================
    //  GASTOS VARIABLES (deudas)
    // =========================
    // Gastos variables (deudas: deudor -> acreedor)
    SH_GASTOS_VARIABLES: "GASTOS_VARIABLES",
    // Detalle de pagos/abonos por gasto
    SH_GASTOS_VARIABLES_DETALLE: "DETALLE_GASTO_VARIABLE",

    // =========================
    //  PRÉSTAMOS
    // =========================
    // Cabecera del préstamo/solicitud
    SH_PRESTAMOS: "PRESTAMOS",
    // Calendario / tabla de amortización
    SH_PRESTAMOS_CUOTAS: "PRESTAMOS_CUOTAS",
    // Pagos (se usará en Parte 2, se crea desde ya)
    SH_PRESTAMOS_PAGOS: "Prestamos_Pagos",
    // Parámetros / defaults
    SH_PARAMETROS: "PARAMETROS",

  };
}
