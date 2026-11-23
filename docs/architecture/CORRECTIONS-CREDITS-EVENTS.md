# Correcciones y créditos de HOME v1

Los comandos se ejecutan por `homeExecute(token, {id, type, payload})`. Una operación aceptada escribe una sola fila de esquema 6 con `eventVersion: 1`, cuenta autenticada, fecha del servidor e ID de operación. Se usa el mismo `ScriptLock` y la idempotencia del ledger. El servidor comprueba sesión, rol, vínculo, versión y saldos dentro del bloqueo. Repetir la operación con el mismo ID, actor y contenido devuelve `duplicate: true`. Los importes de entrada son HNL con hasta dos decimales y las proyecciones usan centavos enteros.

| Comando | Permiso y payload |
| --- | --- |
| `CORRECTION_PROPOSE` | Administrador. `obligationId`, `amount` (puede ser cero), `reason` obligatorio y `expectedVersion` (0 para la primera). Guarda el principal histórico, el importe efectivo anterior, el propuesto y la nueva versión. |
| `CORRECTION_DECIDE` | Cada parte vinculada decide por sí misma. `correctionId`, `version`, `decision` (`accept` o `reject`) y `reason` obligatorio para rechazar. Dos aceptaciones aplican la corrección. Un rechazo la cierra. |
| `CREDIT_REFUND_DECLARE` | Quien conserva el dinero declara una devolución con `creditId`, `amount` positivo y `note` opcional. La declaración reserva el crédito. |
| `CREDIT_APPLY_PROPOSE` | Cualquiera de las dos partes propone con `creditId`, `obligationId`, `amount` positivo y `note` opcional. La propuesta expresa su consentimiento y reserva el importe en ambos saldos. |
| `CREDIT_DECIDE` | Solo la contraparte de la aplicación, o quien recibe la devolución. `operationId`, `decision` (`accept` o `reject`) y `reason` obligatorio al rechazar. |
| `homeGetState(token)` | Expone `corrections`, `credits`, `creditOperations` y sus `creditTotals`. El administrador consulta el hogar. Las partes consultan su propio historial mientras su cuenta esté activa y vinculada, incluso después de retirarse del hogar. |

Las correcciones se aplican a obligaciones bilaterales de gastos y préstamos de HOME v1. Los comandos históricos y sus obligaciones de esquema 1 conservan su comportamiento. Solo puede haber una corrección pendiente por obligación. Una reserva de abono o aplicación impide abrirla. Mientras está pendiente, se bloquean nuevos abonos y aplicaciones sobre esa obligación. Los rechazos consumen su versión. No se admiten decisiones sobre versiones obsoletas o repetidas con otro ID.

## Saldos e invariantes

- `cents` conserva el principal original. `effectiveCents` lo sustituye solo en el cálculo cuando existe una corrección aprobada. `paid` conserva la suma histórica de abonos en efectivo confirmados. `creditedCents` acumula la parte de ese efectivo reconocida como créditos de origen. `appliedCreditCents` acumula crédito confirmado contra esta obligación.
- Saldo definitivo: `max(0, efectivo - paid + creditedCents - appliedCreditCents)`, donde `efectivo` es el importe efectivo corregido o, si falta, el principal. Disponible: saldo definitivo menos `reserved` de abonos y `creditReserved` de aplicaciones. La disponibilidad nunca puede volverse negativa por una operación aceptada.
- Al aprobar, el excedente nuevo es `max(0, paid - creditedCents + appliedCreditCents - effectiveCents)`. Solo se emite si es positivo. El crédito usa el ID de la corrección y conserva obligación de origen, partes, fecha e importe. El importe corregido tiene como piso lo ya aplicado en crédito, por lo que un ajuste no convierte crédito aplicado en efectivo ni deshace una aplicación confirmada.
- Ejemplo: L 500 pagados, importe corregido L 400, crédito L 100. Si el importe vuelve a L 500, el crédito reconocido se conserva y la deuda vuelve a L 100. Si después baja a L 400 otra vez, el saldo queda en cero sin emitir otro crédito. El acumulado se conserva aunque ya se haya devuelto o aplicado ese crédito.
- Crédito definitivo: `cents - returned - applied`. Disponible: definitivo menos `reserved`. Al rechazar, se libera la reserva. Al confirmar una aplicación, la misma transición consume el crédito y reduce el saldo de la obligación. Una devolución solo reduce el crédito. Cada crédito entra una vez en el consolidado, separado de las obligaciones. Los abonos y acuerdos anteriores se conservan.

La aplicación exige que el beneficiario del crédito sea el deudor de la obligación y quien conserva el dinero sea el acreedor. La obligación debe provenir de un gasto compartido cuyo pago al proveedor ya se registró. Los préstamos pueden originar créditos tras una corrección, pero no reciben aplicaciones. No hay compensación automática.

Los eventos de esquemas 1–5 mantienen su contrato. El registro guarda cada propuesta y decisión; el replay aplica el mismo cambio una vez en orden del historial. No se crean hojas ni se reescriben filas. Tras incorporar esquema 6, cualquier reversión requiere un lector probado de esquemas 1–6 que conserve la proyección del saldo, las reservas y el historial. La versión QA 8 solo admite hasta esquema 5. Si falta un lector compatible, archivar solo la implementación QA y preservar la hoja.

Las operaciones anteriores están disponibles por RPC. Las pantallas existentes reflejan el saldo definitivo proyectado de las obligaciones. Quedan pendientes los formularios de corrección, devolución, aplicación, decisiones y la vista de créditos.
