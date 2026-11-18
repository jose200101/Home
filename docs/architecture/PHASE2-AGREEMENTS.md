# Fase 2 — Acuerdos y aceptación individual

**Estado:** módulo de dominio puro, sin escrituras a Google Sheets ni cambios en QA o producción.

## Auditoría focalizada
- Reutilizar: centavos enteros, HOME_EVENTOS, ScriptLock, idempotencia y filtros por rol.
- Sustituir: TEMPLATE no tiene aceptación; CHARGE y DEBT crean obligaciones inmediatas; CONFIRM permite confirmar al administrador; VOID no exige nueva aceptación.
- Revisar antes de integrar: estado activo/en liquidación, proyecciones privadas, neteo global, versión de eventos y replay del histórico.

## Primer módulo
Archivo: src/server/domain/agreements/AgreementRules.js
Salida GAS: Backend/domain/AgreementRules.js

- PROPOSE inicia una versión y solicita aceptación.
- RESPOND acepta o rechaza solamente como integrante involucrado.
- CANCEL restaura el acuerdo anterior si la nueva versión estaba pendiente.
- Rechazo obliga a reenviar distribución nueva, con nuevas aceptaciones.
- Los acuerdos recurrentes no permiten facturar mientras su versión esté pendiente.
- Los gastos ocasionales no generan obligación hasta aprobar todos.
- Igual, porcentaje y monto fijo se validan con exactitud decimal.
- No se exponen endpoints nuevos. Las funciones terminan en guion bajo.

**Límite:** todavía no conecta HOME_EVENTOS, no guarda movimientos ni genera obligaciones. La vinculación con los eventos y los saldos será un módulo posterior con pruebas propias.

**Aportes fijos recurrentes:** cuando la factura real difiera del total fijo aceptado deberá revisarse la distribución; nunca cambiar importes silenciosamente.
