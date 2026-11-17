/**
 * env
 * Configuración centralizada (IDs y nombres de hojas).
 *
 * RECOMENDADO:
 *  - Define una Script Property llamada SPREADSHEET_ID (Project Settings → Script Properties)
 *    para no tocar código al cambiar de archivo.
 *  - Si prefieres, pega aquí el ID del Spreadsheet (entre /d/ y /edit en la URL).
 */
function env_() {
  const propsId = String(PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID") || "").trim();

  return {
    // ID BASE DE DATOS (Spreadsheet ID)
    // Ejemplo: "1eypazsZ5hyN8fe99OxVcZs7AoaYEv...."
    ID_DATABASE: propsId || "1eypazsZ5hyN8fe99OxVcZs7AoaYEvu21mh94hQhG12s",

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
