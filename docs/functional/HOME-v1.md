# HOME v1 — Reglas funcionales aprobadas

Estado: **especificación funcional de referencia**. Implementación pendiente. HNL, dos decimales, fechas `America/Tegucigalpa`. El acceso y las contraseñas se revisarán después; no está autorizado el despliegue público con datos reales.

## Integrantes y catálogos

- Un integrante puede estar activo, en liquidación o retirado. El retiro de nuevas obligaciones es inmediato; deudas, créditos y el historial permanecen.
- Solo el administrador mantiene categorías y proveedores compartidos. Un integrante puede escribir un proveedor ocasional sin modificar el catálogo.

## Gastos

- Recurrentes: acuerdo de participantes aceptado una vez; cada cambio de participantes o reparto exige nueva aceptación. La facturación queda suspendida mientras se aprueba un cambio. Cancelar la propuesta recupera el último acuerdo aprobado.
- Facturas recurrentes asistidas: el administrador confirma el importe real, período y vencimiento; no hay generación automática de obligaciones por importes estimados.
- Ocasionales: cualquier integrante activo registra su propia compra; escoge participantes, puede reutilizar la configuración y cada afectado acepta o rechaza su parte. Un rechazo obliga a revisar y reenviar el reparto; hasta entonces no se genera deuda definitiva.
- Distribución: partes iguales, porcentaje o cantidades exactas; conservar reparto de cada factura.
- Cada factura tiene un único pagador que paga el total al proveedor; puede ser un integrante activo no participante. Registro inmediato sin aprobación administrativa previa. Los reembolsos corresponden a participantes aprobados, excepto la propia cuota del pagador.

## Obligaciones, préstamos y créditos

- Abonos parciales admitidos; el receptor confirma o rechaza cada abono. Pagos declarados sin confirmar no reducen definitivamente el saldo.
- Préstamos personales: solicitud registrada por acreedor; el deudor acepta antes de generar obligación, incluso si inicialmente carece de cuenta. Vencimiento opcional; sin intereses ni mora automática en v1.
- Correcciones administrativas que modifiquen obligaciones requieren nueva aceptación de las partes afectadas. No borrar pagos anteriores.
- Excedentes por corrección crean saldo a favor. Devolución o crédito aplicable solamente a gastos compartidos entre las mismas personas, con aceptación bilateral; sin compensación automática.
- Obligaciones, por cobrar y por pagar, se calculan sin doble contabilización. Los créditos se incluyen una sola vez en el consolidado.

## Comprobantes y privacidad

- JPG/PNG/WebP/PDF opcionales y privados en Drive; registros en Sheets guardan referencias. Un archivo no confirma la recepción de dinero.
- Facturas y comprobantes generales visibles a sus participantes, pagador y administrador; comprobantes individuales a sus dos partes y administrador.
- No hay aceptación en nombre de otro, aunque quien actúe sea administrador. Historial y correcciones auditables.

## Migración acordada

Sustitución **completa** de módulos antiguos en el resultado final. Migrar solamente saldos vigentes *conciliados* a fecha de corte, con trazabilidad de origen; conservar operaciones antiguas en un historial de solo lectura. Nunca sumar automáticamente el saldo histórico y el nuevo. No realizar migraciones ni eliminar archivos históricos hasta aprobar reconciliación y pruebas.
