# Transición y validación

## Referencia verificada en Drive

El 16 de noviembre de 2026 se consultaron metadatos y exclusivamente encabezados de `HOME-Cuentas`,
el Spreadsheet configurado en `env.js`. No se consultaron valores de contraseñas ni se hicieron escrituras.
La conexión disponible respondió como la cuenta principal, aunque la interfaz del usuario muestra otra
cuenta conectada con etiqueta HOME. No se afirmó haber seleccionado esa segunda cuenta mediante el conector.

Pestañas verificadas: PERSONAS, USUARIOS, GASTOS_FIJOS, DETALLE_GASTO_FIJO,
GASTOS_VARIABLES, DETALLE_GASTO_VARIABLE, PRESTAMOS, PRESTAMOS_CUOTAS, PRESTAMOS_PAGOS,
PARAMETROS, COMBUSTIBLE, GASTOS, DETALLE_GASTO, TIPO DE GASTOS y COMERCIOS.

USUARIOS usa `nombre_persona`; se corrigió el login para soportarlo además de `nombreCompleto`.
La hoja reporta zona America/Guatemala y locale es_MX. No se modificaron. El motor aplica
America/Tegucigalpa explícitamente para fechas de negocio y HNL para importes.

## Migración pendiente de conciliación

No se ejecuta una migración automática de dinero: se inspeccionaron encabezados, no saldos.
Configuración ofrece un inventario de solo lectura. No importa registros ni recalcula préstamos.
La implementación mantiene las cuentas nuevas separadas de los módulos históricos.

Antes de activar con datos reales:

1. Crear una copia independiente del Spreadsheet y conservar acceso restringido.
2. Configurar `SPREADSHEET_ID` en Script Properties del entorno de prueba, apuntando a la copia.
3. Obtener una conciliación a fecha de corte: factura, monto original, participantes, pagado y pendiente.
4. Elegir qué esquema anterior es vigente: existen GASTOS/GASTOS_VARIABLES y detalles alternativos.
   No sumarlos automáticamente: podrían representar las mismas deudas.
5. Relacionar IDs de PERSONAS con deudores, acreedores y usuarios. Resolver nombres ambiguos.
6. Determinar si aportes antiguos se entregaron al proveedor, a un encargado o al pagador.
   El detalle de fijos no permite inferir por sí solo quién adelantó la factura.
7. Conciliar préstamos con sus tasas, desembolsos y pagos. Mantener amortización histórica separada.
8. Preparar archivo de migración con identificadores de origen y saldos conciliados; revisarlo antes de escribir.
9. Importar una sola vez en una copia y comparar totales por integrante y globales, sin duplicados.
10. Desactivar escrituras del sistema anterior a la fecha de corte; mantenerlo como consulta.

El importador financiero automático queda pendiente: inventar acreedores o interpretar saldos sin
conciliación podría cambiar las cuentas. No se considera completada la migración por tener un inventario.

## Validación local reproducible

Con Node.js instalado, desde la carpeta Home:

```powershell
npm test
npm run preview
```

La vista previa escucha solo en `http://127.0.0.1:4173`, no accede a Google y usa el mismo motor
HomeLedger con servicios de Apps Script simulados. La barra identifica los datos ficticios
y permite cambiar entre administrador, Ana, Luis y María. Los cambios se pierden al reiniciar el servidor.
No poner credenciales reales en la demostración. `tools`, `tests`, `docs` y `preview` se excluyen de clasp.
La demo utiliza tokens locales y cuentas ficticias; no prueba el login real.

Pruebas automatizadas: repartos y redondeos, participantes independientes, historial de facturas,
adelantos, reservas, confirmación, rechazo, reversión, idempotencia, revocación, privacidad, permisos,
referencias inválidas, esquema inválido, deudas cruzadas y sintaxis de todo el backend.

Pruebas de navegador realizadas con datos ficticios:

- Creación de gasto por porcentaje (Ana 60 %, María 40 %).
- Factura L 100.01: reparto L 60.01 y L 40.00.
- Cuenta de Ana: no revela la deuda entre Luis y María ni enlaces administrativos históricos.
- Confirmación de recepción: abono L 100 reduce de L 300 a L 200 el saldo de María con Ana.
- Resumen, facturas y formulario de pago en viewport de 390 × 844, sin desbordamiento horizontal de página.
- Capturas de escritorio y móvil en `preview/`.

No se validaron cuotas, permisos, servicios ni escrituras en el runtime real de Apps Script.
El frontend conserva React/Babel vía CDN; completar la preparación de producción y probar
el HTML Service desplegado después de cerrar la autenticación.

## Revisión por etapas

| Etapa | Resultado actual | Falta para activación real |
| --- | --- | --- |
| Reglas funcionales | Implementadas y documentadas para el primer modelo | Revisar custodia y ajustes de facturas financiadas |
| Datos e interfaz | Historial por operaciones y vistas por integrante | Medición con volumen real |
| Seguridad | Autorización de servidor y privacidad | Proveedor de acceso, contraseñas y recuperación |
| Gastos compartidos | Catálogo, participantes, facturas, aportes | Prueba en copia de Sheets |
| Cuentas personales | Deudas simples, adelantos, abonos y confirmaciones | Acuerdo sobre confirmación administrativa |
| Consolidación | Estados de cuenta y propuesta de liquidación | Compensación automática no incluida |
| Migración y validación | Inventario y procedimiento; pruebas locales | Conciliar, implementar importador y ensayar en copia |

## Respaldo y reversión operativa

Conservar una copia del Spreadsheet y una versión del código antes del despliegue. Si la prueba falla,
detener escrituras y volver al código anterior con su copia de datos correspondiente.
No fusionar historiales divergentes a mano. Las reversiones de pagos de la aplicación son eventos;
restaurar un respaldo completo requiere revisar también movimientos posteriores al respaldo.
