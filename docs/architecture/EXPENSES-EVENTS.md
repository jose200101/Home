# Gastos HOME v1

El módulo amplía `homeExecute(token, {id, type, payload})` y `homeGetState(token)` sobre `HOME_EVENTOS`. Cada comando nuevo añade una fila con esquema 3 y `eventVersion: 1`. Las propuestas y respuestas incorporadas guardan la transición validada por el dominio de acuerdos. Los eventos de esquemas 1 y 2 conservan su contrato. Los importes se reciben como decimales HNL de hasta dos posiciones y se proyectan como centavos enteros.

| Comando | Datos y permiso |
| --- | --- |
| `EXPENSE_SERVICE` | Administrador. `resourceId` nuevo, `expectedVersion: 0`, `name`, `provider`, `mode`, `participants` y `fixedAmount` para modo exacto. Para revisar, enviar el mismo recurso y la última versión conocida, incluso si quedó cancelada. |
| `AGREEMENT_RESPOND` | Para servicios recurrentes. Cada participante envía `resourceId`, `version` y `decision` (`accept` o `reject`). |
| `AGREEMENT_CANCEL` | Cancela una revisión pendiente mediante `resourceId` y `version`, con los permisos del dominio de acuerdos. |
| `EXPENSE_INVOICE` | Administrador. `resourceId`, `version` aprobada, `period` (AAAA-MM), `amount`, `date` de emisión y `due` de vencimiento. Solo una factura por servicio y período. |
| `EXPENSE_PAY` | Integrante activo y vinculado que declara haber pagado. `chargeId`, `payerId` propio, `amount` íntegro y `date`. Puede ser externo al reparto. |
| `EXPENSE_PURCHASE` | Integrante activo y vinculado que declara su propia compra ya pagada. `resourceId`, `expectedVersion`, `name`, `provider` opcional, `date`, `amount`, `payerId` propio y reparto. Si se rechaza, el mismo recurso admite un nuevo reparto con sus datos originales de compra. |
| `EXPENSE_RESPOND` | Para compras ocasionales. Cada participante envía `resourceId`, `version` y `decision`. La última aceptación guarda el cargo pagado y el reparto exacto en el mismo evento. |

Los modos son `equal`, `percent` y `amount`. Cada participante tiene `personId` y, en los dos últimos modos, `value`. El reparto asigna el residuo de centavos en orden de participantes. Una compra nueva puede reemplazar `mode` y `participants` por `reuseResourceId` de una compra aprobada visible. Se copian las condiciones del reparto y se solicitan nuevas aceptaciones. Con montos exactos, el nuevo total debe coincidir con las cuotas copiadas.

La factura guarda el proveedor, los centavos, la versión y todas las cuotas aprobadas. Se registra como `unpaid` hasta el pago al proveedor. Una revisión suspende nuevas facturas, y cancelarla recupera el reparto y los datos de la versión anterior. El financiamiento de una factura ya registrada usa su propio reparto aunque luego cambie el acuerdo del servicio. No se generan reembolsos por una factura impaga o una propuesta pendiente o rechazada. Al activarse, cada reembolso conserva acreedor, deudor, factura, recurso y versión. La cuota del pagador queda cubierta y se omiten cuotas de cero centavos.

Las vistas devuelven `expenseResources`, acuerdos filtrados, cargos y obligaciones en la proyección existente. Un integrante ve sus facturas y las que pagó, su propia cuota y las obligaciones de las que forma parte. El retiro conserva su consulta histórica mientras siga vinculada su cuenta. El administrador conserva la vista completa. Para que una persona externa al reparto registre el pago, necesita el identificador de la factura.

Los servicios nuevos no se guardan como plantillas antiguas. `TEMPLATE`, `CHARGE`, `FUND`, `VOID` y `PAYMENT` históricos siguen operando sobre sus registros y no pueden crear, financiar, anular ni abonar los gastos nuevos. El comando de propuesta anterior no opera sobre recursos del módulo y la respuesta anterior solo se comparte para recurrentes. Los reembolsos nuevos no entran al plan de compensación anterior. La idempotencia existente comprueba actor, comando y contenido bajo `ScriptLock`.

Las pantallas actuales consultan facturas y obligaciones. El módulo oculta acciones históricas que rechazan estos registros. La creación de servicios y compras, las respuestas a acuerdos y el registro de facturas y pago completo se consumen por RPC y todavía no tienen formularios. Los abonos, devoluciones y créditos quedan para otra fase.

Para validar en QA se necesita autorización de despliegue, comprobar proyecto, propiedad `SPREADSHEET_ID` y acceso `MYSELF`, respaldar sus datos y ejecutar un escenario con cuentas ficticias usando estas operaciones RPC. Una vez que existan eventos del esquema 3, cualquier reversión exige un lector compatible con los tres esquemas. Para detener QA, archivar solo su implementación privada y conservar el historial.
