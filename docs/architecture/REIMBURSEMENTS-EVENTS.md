# Reembolsos HOME v1

Los gastos compartidos de la fase 2.3 ya generan una obligación entre el deudor y quien pagó al proveedor. Cada abono aplica a una sola obligación. La factura, el acuerdo y el principal conservan sus datos originales.

| RPC | Datos y acceso |
| --- | --- |
| `homeExecute(token, {id, type: "REIMBURSEMENT_DECLARE", payload})` | Deudor vinculado. `payload` contiene `obligationId`, `amount` en HNL con hasta dos decimales, `date` entre el pago al proveedor y hoy y `note` opcional de hasta 500 caracteres. |
| `homeExecute(token, {id, type: "REIMBURSEMENT_DECIDE", payload})` | Acreedor vinculado. `payload` contiene `paymentId`, `decision` (`confirm` o `reject`) y `reason` de hasta 500 caracteres, obligatorio al rechazar. |
| `homeGetState(token)` | El deudor y el acreedor ven sus obligaciones y abonos. El administrador puede consultar todos. Las otras cuentas no reciben esos movimientos. |

La cuenta debe seguir activa y vinculada de forma única. El retiro del integrante permite consultar y liquidar sus obligaciones existentes mientras se conserve ese vínculo. El rol administrativo no delega declaraciones ni decisiones por terceros. El servidor toma las identidades de la sesión y del historial. Revalida la sesión, los vínculos y el saldo bajo `ScriptLock`. El mismo actor puede reintentar con el mismo ID y contenido. La repetición devuelve `duplicate: true` y no añade filas. Una decisión distinta con un ID nuevo solo procede mientras el abono esté pendiente.

Los saldos de la obligación se expresan en centavos enteros. `cents` es el principal, `paid` suma los confirmados, `reserved` suma los pendientes, `pending = cents - paid` es el saldo definitivo y `available = pending - reserved` es lo disponible para declarar. Una obligación con saldo definitivo cero tiene estado `settled`. Confirmar mueve el importe reservado a confirmado. Rechazar libera la reserva y conserva la deuda.

Cada comando añade exactamente una fila de esquema 4 con `eventVersion: 1`. La fila registra ID de operación, cuenta autenticada y hora del servidor. La declaración guarda ambas partes, fecha declarada, importe, nota y referencia de la obligación. La decisión guarda referencia del abono, acreedor, decisión y motivo. La proyección aplica las transiciones en orden y expone fechas `declaredAt` y `decidedAt`. La auditoría del administrador conserva el actor de cada evento. Los abonos de esquema 4 no intervienen en el cálculo histórico de abonos ni en la compensación anterior.

`PAYMENT`, `CONFIRM`, `REJECT` y `VOID` históricos conservan sus reglas para esquemas anteriores. No permiten modificar abonos de esquema 4. Tampoco se crean hojas ni se reescriben eventos. Los lectores aceptan los esquemas 1, 2 y 3. Los listados actuales muestran los abonos y saldos. Para esta fase, los formularios de declaración, confirmación y rechazo y el detalle de saldo disponible quedan pendientes. Se ocultaron los controles históricos que no aplican a los abonos nuevos.

Para revertir después de escribir eventos de esquema 4, conservar el historial y usar una compilación verificada que lea y proyecte los cuatro esquemas. Puede retirarse el acceso a los comandos de escritura conservando la lectura y los saldos proyectados. La versión QA 6 admite solo esquemas 1 a 3. Si falta una versión compatible, archivar únicamente la implementación QA y conservar la hoja y su respaldo. Esta fase se valida localmente y no despliega QA.
