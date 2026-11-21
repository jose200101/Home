# Acuerdos versionados sobre HOME_EVENTOS

## Alcance

Integración aditiva del dominio existente `AgreementRules.js` con `HomeLedger.js` mediante
`AgreementEvents.js`. No hay nuevas hojas, migraciones, cambios de login o UI, ni generación de
obligaciones por proponer o aprobar un acuerdo. No se implementa el nuevo sistema completo de
facturas, pagos al proveedor o reembolsos de HOME v1.

## API de comandos

Se conserva `homeExecute(token, {id, type, payload})`. Los IDs nuevos deben ser strings de
12–100 caracteres alfanuméricos, guion o subrayado. No reutilizar un ID con un payload diferente.

| type | payload | Permisos |
| --- | --- | --- |
| AGREEMENT_PROPOSE | `{expectedVersion, proposal}` | Recurrente: administrador. Ocasional: integrante activo vinculado y pagador de su propia compra |
| AGREEMENT_RESPOND | `{resourceId, version, decision}` | Únicamente el integrante activo vinculado que participa en esa versión; decision es accept o reject |
| AGREEMENT_CANCEL | `{resourceId, version}` | Recurrente: administrador. Ocasional: autor de la compra o administrador |

Ejemplo de propuesta recurrente inicial, usando identificadores ficticios:

```js
homeExecute(token, {
  id: "operation-energia-v1",
  type: "AGREEMENT_PROPOSE",
  payload: {
    expectedVersion: 0,
    proposal: {
      resourceId: "servicio-energia",
      kind: "recurring",
      mode: "equal",
      participants: [{personId: "persona-a"}, {personId: "persona-b"}]
    }
  }
});
```

La propuesta usa el contrato del dominio: `resourceId`, `kind` (recurring/occasional), `mode`
(equal/percent/amount) y participantes con `personId`. En percent, `value` es un porcentaje con
hasta dos decimales; en amount es un importe decimal. `amountCents` es obligatorio para ocasionales
y repartos de monto fijo; `payerId` solo se admite en ocasionales y debe ser la persona del actor.
El pagador ocasional puede no participar en el reparto, pero en ese caso no puede aceptar por otros.

`expectedVersion` es la última versión conocida, incluyendo versiones rechazadas, reemplazadas o
canceladas; cero identifica un recurso sin acuerdos. Se vuelve a comprobar bajo ScriptLock. Dos
solicitudes basadas en la misma versión no pueden sobrescribirse. RESPOND y CANCEL deben apuntar
a la versión pendiente exacta. Reenviar una compra ocasional preserva su autoría, incluso después
de una cancelación; el administrador puede cancelarla, pero no reasignársela como pagador.

## Contrato de eventos

Columnas de HOME_EVENTOS, sin cambios: id, fecha, actor, tipo, datos.

- Los diez tipos históricos conservan `datos.schema = 1` y su interpretación anterior.
- Los tres tipos AGREEMENT_* usan `datos.schema = 2`, `eventVersion = 1`.
- `version` es la versión del acuerdo, distinta de la versión técnica del evento.
- `resourceId`, `kind` y `command` identifican el recurso y la transición de dominio.
- `actor` de la fila es siempre el ID de cuenta autenticada. `datos.actor` guarda el contexto
  histórico comprobado por el servidor: id de cuenta, personId vinculado cuando exista, active y admin.
- Fecha ISO y ID de operación los establece el mismo mecanismo de persistencia del ledger.
- `requestJson` conserva el fingerprint del payload para la idempotencia existente; nunca se devuelve
  como parte de las solicitudes visibles.

Cada transición añade una sola fila. No se reescriben eventos previos ni se guarda una copia completa
del historial por transición. La aprobación se deriva de la última aceptación individual: no genera
un evento financiero ni una segunda escritura.

El replay aplica las reglas puras usando el contexto histórico del actor y el orden del ledger.
Comprueba la correspondencia entre tipo, comando, recurso, kind, versión y actor. No revalida una
aceptación pasada con el rol o vínculo actual, pues eso cambiaría silenciosamente la historia.
Los pares tipo/esquema incompatibles, versiones técnicas desconocidas y transiciones corruptas
fallan sin añadir otra operación. Los editores directos de Sheets siguen siendo usuarios de confianza;
esta etapa no añade firma criptográfica al historial.

## Autorización y retiro

La sesión se comprueba al entrar y nuevamente después de adquirir el lock para comandos de acuerdos.
Se deriva el integrante desde `PERSON.userId === session.id`, con vínculo único y persona activa para
responder o registrar compras. Nunca se equipara el ID de cuenta con el ID de persona ni se acepta
actor/personId/userId al nivel superior del payload. Una cuenta activa sin vínculo no puede aceptar.

Un administrador solo acepta su propia cuota si también es participante vinculado. Cada respuesta
pertenece a la versión pendiente; no se transfieren aceptaciones al reenviar. Los participantes de
nuevas propuestas deben estar activos y no se aprueba un reparto que incluya a alguien retirado.
Retirar a una persona conserva decisiones, acuerdos y obligaciones históricas; impide nuevas
participaciones y vuelve no elegibles los contratos que la incluyan hasta revisar el reparto.

## Consulta y privacidad

`homeGetState(token)` conserva los campos existentes y añade `agreements`, con:

- resourceId, kind, active, pending, history y latestVersion.
- suspended: hay una modificación pendiente o rechazada por resolver.
- canActivate y canInvoice: aprobación completa, sin pendiente y con integrantes activos.

Administrador: todas las versiones. Integrante: solo versiones donde participa o es pagador
ocasional; una solicitud ajena no aparece. Se eliminan IDs de cuenta del creador y decisiones
individuales de terceros; `createdByMe` indica autoría propia. El reparto se conserva visible para
poder evaluarlo. Un participante retirado conserva la consulta de sus versiones históricas mientras
la cuenta permanezca activa y vinculada. Un vínculo ambiguo falla de forma cerrada.

Si una persona participó en el acuerdo activo, pero no en la modificación propuesta, puede consultar
su acuerdo anterior y saber que el servicio está suspendido; no recibe el nuevo reparto ni sus respuestas.

## Contrato financiero y puente de compatibilidad

`homeApprovedAgreement_(state, resourceId, kind)` devuelve una copia del contrato aprobado:
resourceId, kind, version, mode, participantes normalizados y, cuando corresponda, amountCents/payerId.
No tiene efectos secundarios. Falla si no hay aprobación completa, hay una modificación pendiente
o rechazada, o un integrante incluido ya no está activo. Cancelar una modificación conserva el
acuerdo aprobado anterior; no vuelve activo a un integrante retirado.

Para asociar un servicio anterior a acuerdos, su resourceId debe ser exactamente el ID del TEMPLATE.
El CHARGE existente:

1. Sin acuerdo asociado mantiene el comportamiento anterior; no se migran ni convierten sus plantillas.
2. Con acuerdo asociado exige el contrato recurrente aprobado y no suspendido.
3. Usa el reparto aprobado, valida cualquier override explícito y respeta el total fijo si corresponde.
4. Añade `agreementVersion` a la nueva factura, manteniendo el esquema financiero 1 y sus cálculos.

Las facturas anteriores mantienen su snapshot, sin introducirles agreementVersion ni recalcularlas.
Un cliente antiguo que envíe un reparto desactualizado para un servicio ya gestionado por acuerdos
recibe un rechazo de validación; los nombres y firmas de los endpoints no cambian. No se asocian
recursos ocasionales a IDs de plantillas recurrentes. Ningún acuerdo ocasional crea CHARGE, DEBT,
pagos o reembolsos automáticamente en esta etapa.

## Validación y límites

`tests/agreement-events.test.cjs` ejercita persistencia y replay con el ledger real y adaptadores de
servicios en memoria: no accede a Sheets ni Apps Script. Incluye reconstrucción en un contexto nuevo,
aceptación completa, rechazo/reenvío, cancelación, privacidad por versión, identidad real, retiro,
idempotencia, cambios de permisos durante la espera del lock y una intercalación determinista de
dos propuestas. No representa una prueba de estrés de concurrencia distribuida en Google.

La compatibilidad se comprueba mezclando los diez tipos históricos con acuerdos nuevos y comparando
saldos y las filas anteriores byte a byte. Los tests de dominio y módulos anteriores siguen ejecutándose.
Se conservan los límites del ledger: 40.000 caracteres JSON por operación y replay completo del historial;
esta integración no añade paginación ni compactación.

Validación local del 2026-11-21: `npm test` PASS (58/58: 42 existentes y 16 de integración),
`npm run smoke` PASS (HTML, RPC y privacidad con datos simulados) y `npm run build` PASS
(28 archivos). Los adaptadores en memoria no constituyen evidencia de ejecución real en Google.

No hay despliegue ni migración en este PR. Una futura reversión de código debe conservar un lector
compatible con esquema 2: el lector antiguo solo acepta esquema 1. No eliminar ni convertir eventos
para ocultar una incompatibilidad de rollback. La validación real de esta integración en QA se acordará
separadamente después de revisar este contrato.
