/** Puerta de entrada de los módulos históricos. Las implementaciones son privadas. */
function diagnosticoAcceso() {
  homeLegacyAuthorize_(arguments);
  return diagnosticoAccesoLegacy_.apply(
    null,
    Array.prototype.slice.call(arguments, 0, -1),
  );
}

function eliminarGastoFijo() {
  homeLegacyAuthorize_(arguments);
  return eliminarGastoFijoLegacy_.apply(
    null,
    Array.prototype.slice.call(arguments, 0, -1),
  );
}

function eliminarGastoVariable() {
  homeLegacyAuthorize_(arguments);
  return eliminarGastoVariableLegacy_.apply(
    null,
    Array.prototype.slice.call(arguments, 0, -1),
  );
}

function eliminarPersona() {
  homeLegacyAuthorize_(arguments);
  return eliminarPersonaLegacy_.apply(
    null,
    Array.prototype.slice.call(arguments, 0, -1),
  );
}

function eliminarPrestamo() {
  homeLegacyAuthorize_(arguments);
  return eliminarPrestamoLegacy_.apply(
    null,
    Array.prototype.slice.call(arguments, 0, -1),
  );
}

function eliminarUsuario() {
  homeLegacyAuthorize_(arguments);
  return eliminarUsuarioLegacy_.apply(
    null,
    Array.prototype.slice.call(arguments, 0, -1),
  );
}

function getGastosNetosBalance() {
  homeLegacyAuthorize_(arguments);
  return getGastosNetosBalanceLegacy_.apply(
    null,
    Array.prototype.slice.call(arguments, 0, -1),
  );
}

function getGastosNetosDashboard() {
  homeLegacyAuthorize_(arguments);
  return getGastosNetosDashboardLegacy_.apply(
    null,
    Array.prototype.slice.call(arguments, 0, -1),
  );
}

function getGastosNetosDetallePersona() {
  homeLegacyAuthorize_(arguments);
  return getGastosNetosDetallePersonaLegacy_.apply(
    null,
    Array.prototype.slice.call(arguments, 0, -1),
  );
}

function getGastosNetosPlanLiquidacion() {
  homeLegacyAuthorize_(arguments);
  return getGastosNetosPlanLiquidacionLegacy_.apply(
    null,
    Array.prototype.slice.call(arguments, 0, -1),
  );
}

function getGastosNetosResumen() {
  homeLegacyAuthorize_(arguments);
  return getGastosNetosResumenLegacy_.apply(
    null,
    Array.prototype.slice.call(arguments, 0, -1),
  );
}

function getPagoSugerido() {
  homeLegacyAuthorize_(arguments);
  return getPagoSugeridoLegacy_.apply(
    null,
    Array.prototype.slice.call(arguments, 0, -1),
  );
}

function guardarGastoFijo() {
  homeLegacyAuthorize_(arguments);
  return guardarGastoFijoLegacy_.apply(
    null,
    Array.prototype.slice.call(arguments, 0, -1),
  );
}

function guardarGastoVariable() {
  homeLegacyAuthorize_(arguments);
  return guardarGastoVariableLegacy_.apply(
    null,
    Array.prototype.slice.call(arguments, 0, -1),
  );
}

function guardarPersona() {
  homeLegacyAuthorize_(arguments);
  return guardarPersonaLegacy_.apply(
    null,
    Array.prototype.slice.call(arguments, 0, -1),
  );
}

function guardarPrestamo() {
  homeLegacyAuthorize_(arguments);
  return guardarPrestamoLegacy_.apply(
    null,
    Array.prototype.slice.call(arguments, 0, -1),
  );
}

function guardarUsuario() {
  homeLegacyAuthorize_(arguments);
  return guardarUsuarioLegacy_.apply(
    null,
    Array.prototype.slice.call(arguments, 0, -1),
  );
}

function listarGastosFijos() {
  homeLegacyAuthorize_(arguments);
  return listarGastosFijosLegacy_.apply(
    null,
    Array.prototype.slice.call(arguments, 0, -1),
  );
}

function listarGastosVariables() {
  homeLegacyAuthorize_(arguments);
  return listarGastosVariablesLegacy_.apply(
    null,
    Array.prototype.slice.call(arguments, 0, -1),
  );
}

function listarPagosPrestamo() {
  homeLegacyAuthorize_(arguments);
  return listarPagosPrestamoLegacy_.apply(
    null,
    Array.prototype.slice.call(arguments, 0, -1),
  );
}

function listarPersonas() {
  homeLegacyAuthorize_(arguments);
  return listarPersonasLegacy_.apply(
    null,
    Array.prototype.slice.call(arguments, 0, -1),
  );
}

function listarPersonasActivas() {
  homeLegacyAuthorize_(arguments);
  return listarPersonasActivasLegacy_.apply(
    null,
    Array.prototype.slice.call(arguments, 0, -1),
  );
}

function listarPersonasNetosActivas() {
  homeLegacyAuthorize_(arguments);
  return listarPersonasNetosActivasLegacy_.apply(
    null,
    Array.prototype.slice.call(arguments, 0, -1),
  );
}

function listarPrestamos() {
  homeLegacyAuthorize_(arguments);
  return listarPrestamosLegacy_.apply(
    null,
    Array.prototype.slice.call(arguments, 0, -1),
  );
}

function listarUsuarios() {
  homeLegacyAuthorize_(arguments);
  return listarUsuariosLegacy_.apply(
    null,
    Array.prototype.slice.call(arguments, 0, -1),
  );
}

function listarUsuariosActivos() {
  homeLegacyAuthorize_(arguments);
  return listarUsuariosActivosLegacy_.apply(
    null,
    Array.prototype.slice.call(arguments, 0, -1),
  );
}

function marcarAportePagado() {
  homeLegacyAuthorize_(arguments);
  return marcarAportePagadoLegacy_.apply(
    null,
    Array.prototype.slice.call(arguments, 0, -1),
  );
}

function marcarGastoFijoPagado() {
  homeLegacyAuthorize_(arguments);
  return marcarGastoFijoPagadoLegacy_.apply(
    null,
    Array.prototype.slice.call(arguments, 0, -1),
  );
}

function obtenerDetallePrestamo() {
  homeLegacyAuthorize_(arguments);
  return obtenerDetallePrestamoLegacy_.apply(
    null,
    Array.prototype.slice.call(arguments, 0, -1),
  );
}

function registrarDesembolso() {
  homeLegacyAuthorize_(arguments);
  return registrarDesembolsoLegacy_.apply(
    null,
    Array.prototype.slice.call(arguments, 0, -1),
  );
}

function registrarPagoGastoVariable() {
  homeLegacyAuthorize_(arguments);
  return registrarPagoGastoVariableLegacy_.apply(
    null,
    Array.prototype.slice.call(arguments, 0, -1),
  );
}

function registrarPagoPrestamo() {
  homeLegacyAuthorize_(arguments);
  return registrarPagoPrestamoLegacy_.apply(
    null,
    Array.prototype.slice.call(arguments, 0, -1),
  );
}
