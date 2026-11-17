# HOME — Plan de sustitución y datos históricos

**Política aprobada:** sustituir por completo los módulos antiguos en la versión final, **migrar únicamente saldos vigentes conciliados** y conservar historial antiguo de solo lectura.

## Procedimiento previo a cualquier escritura

1. Crear respaldo verificable del Spreadsheet original y copia aislada de QA.
2. Elegir fecha y hora de corte; identificar datos activos e impedir escrituras anteriores durante el corte.
3. Inventariar `PERSONAS`, `USUARIOS`, `GASTOS_FIJOS`, `DETALLE_GASTO_FIJO`, `GASTOS_VARIABLES`, `DETALLE_GASTO_VARIABLE`, `PRESTAMOS`, `PRESTAMOS_CUOTAS`, `PRESTAMOS_PAGOS`, y otras pestañas detectadas.
4. Identificar qué conjuntos representan las mismas facturas o préstamos; **no sumarlos por defecto**.
5. Identificar deudor, acreedor, pagador, importes realmente entregados y pagos confirmados.
6. Elaborar archivo de conciliación por obligación y por integrante, con referencia a filas originales y justificación de cada saldo.
7. Transformar a eventos de apertura del esquema nuevo con identificador de origen único.
8. Ensayar importación idempotente en QA y comparar balances por persona, tipo y hogar.
9. Probar el nuevo historial de solo lectura y permisos. Bloquear escrituras heredadas al activar el nuevo sistema.
10. Validar plan de reversión con respaldos; no fusionar historiales divergentes ni eliminar origen.

El estado de migración es **NO EJECUTADA**. No se han validado saldos reales en esta fase. Las hojas antiguas se conservarán hasta pasar criterios de aceptación.

## Criterios de aceptación

- Los saldos de apertura concuerdan con conciliación firmada/revisada.
- No se contabiliza dos veces una misma deuda ni un mismo pago.
- Importación repetida no genera nuevos movimientos.
- Los históricos son visibles solo a quienes corresponde, sin permitir edición.
- Existe respaldo y procedimiento ensayado de reversión.
