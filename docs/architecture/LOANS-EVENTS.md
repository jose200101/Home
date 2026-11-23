# Préstamos personales HOME v1

El acreedor solicita registrar un préstamo en `HOME_EVENTOS`. Solo la aceptación expresa del deudor crea una obligación. El rechazo del deudor o la cancelación del acreedor cierra una solicitud pendiente sin crear deuda. Las decisiones son terminales y los términos originales se conservan.

| RPC | Contrato |
| --- | --- |
| `homeExecute(token, {id, type: "LOAN_REQUEST", payload})` | Acreedor activo y vinculado, inferido por el servidor. `payload` contiene `debtorId`, `amount` en HNL con hasta dos decimales, `date` (AAAA-MM-DD hasta hoy), `name` (descripción) y `due` opcional. Si existe vencimiento, debe ser igual o posterior a la fecha del préstamo. |
| `homeExecute(token, {id, type: "LOAN_DECIDE", payload})` | `payload` contiene `loanId`, `decision` (`accept`, `reject` o `cancel`) y `reason` opcional. Solo el deudor vinculado acepta o rechaza. Solo el acreedor vinculado cancela. La aceptación requiere que ambos integrantes estén activos. |
| `homeGetState(token)` | Devuelve `loans` con estados `pending`, `accepted`, `rejected` y `cancelled`. Ambas partes y el administrador pueden consultar el contrato. `due: ""` indica que no se acordó vencimiento. |
| `REIMBURSEMENT_DECLARE` y `REIMBURSEMENT_DECIDE` | Una vez aceptado, usar `obligationId` igual al `loanId`. Las fechas de abono se admiten desde la fecha del préstamo hasta hoy, incluso después del vencimiento. Aplican los mismos permisos, reservas, confirmaciones y rechazos que a los gastos compartidos. |

El deudor puede ser un integrante activo sin cuenta. El administrador vincula posteriormente una cuenta mediante el comando existente `PERSON`, que impide asignar una cuenta a dos integrantes. Los comandos vuelven a comprobar la sesión y el vínculo único bajo `ScriptLock`. Los roles administrativos no sustituyen a ninguna parte. El retiro impide aceptar nuevas obligaciones y conserva la consulta y liquidación de las anteriores mientras persista una cuenta activa vinculada.

El vencimiento es informativo. El importe no cambia por el paso del tiempo y el módulo no calcula intereses ni cargos. El préstamo aceptado referencia su obligación en `obligationId`. En la obligación, `status: "settled"` y `pending: 0` indican liquidación. La solicitud conserva el estado `accepted` y el historial de abonos sigue visible. Los préstamos están excluidos del plan de compensación histórico.

## Eventos y reversión

- Esquema 5, `eventVersion: 1`. `LOAN_REQUEST` guarda acreedor, deudor, importe, fecha, vencimiento y descripción. Su ID de operación es el ID del préstamo. `LOAN_DECIDE` guarda ID del préstamo, identidad de integrante, decisión y motivo. La fila registra la cuenta autenticada y la fecha de servidor. La aceptación proyecta una obligación con el mismo ID.
- Los abonos siguen usando esquema 4 y añaden `loanId` para identificar el origen. Los abonos de gastos mantienen su forma anterior. Tras reintentar con el mismo ID, actor y contenido, la respuesta indica `duplicate: true` y conserva una sola fila. Los saldos se reconstruyen en centavos y cada movimiento afecta únicamente a su obligación.
- Los lectores de esta fase conservan compatibilidad con esquemas 1–4. Las rutas financieras históricas siguen operando sobre sus registros y no declaran, confirman ni anulan movimientos de préstamos personales nuevos.
- Tras registrar esquema 5, conservar todos los eventos y revertir únicamente a una compilación privada con lector probado para esquemas 1–5 y abonos de préstamo. QA 7 admite hasta esquema 4. Si se necesita detener QA y falta un lector compatible alternativo, archivar únicamente la implementación QA y preservar la hoja y su respaldo.

La interfaz actual muestra obligaciones aceptadas y el historial de abonos. La lista de solicitudes, los formularios de solicitud y decisión y los controles del nuevo flujo de abonos quedan pendientes de conexión por interfaz. La vinculación mantiene la operación existente.
