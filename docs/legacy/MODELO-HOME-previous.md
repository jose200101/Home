# Home: modelo del primer lanzamiento

Un hogar por Spreadsheet; moneda HNL y fechas de negocio en America/Tegucigalpa.
Administradores configuran integrantes y gastos, registran facturas y deudas ya contraídas.
Integrantes consultan su cuenta, declaran pagos y confirman o rechazan lo que reciben.
Una persona puede existir sin usuario. El vínculo persona-usuario es explícito y único; nunca se infiere por nombre.

## Gastos y participantes

Los tipos de gasto son registros, no pantallas ni columnas nuevas. Cada uno conserva nombre, proveedor,
estado y participantes predeterminados. Se reparte por partes iguales, porcentajes o montos exactos.
Una factura guarda una copia del reparto, fecha, vencimiento y período. Editar el tipo afecta solamente
facturas futuras. Hay una factura por tipo y período en esta primera versión; para un gasto adicional
independiente se crea otro tipo. Los importes se convierten a centavos y se distribuyen los residuos de
forma determinista, siguiendo el orden de selección, entre participantes con peso positivo.

## Fuente de verdad y relaciones

Se usa `HOME_EVENTOS`, una hoja adicional, sin alterar las hojas anteriores. Sus columnas son:
`id`, `fecha`, `actor`, `tipo`, `datos`.

Cada evento contiene una operación completa en JSON. Las entidades se relacionan por identificadores:

| Evento | Entidad / relaciones |
| --- | --- |
| INIT | Hogar, nombre |
| PERSON | Persona → cuenta existente opcional |
| TEMPLATE | Tipo de gasto → personas y reparto |
| CHARGE | Factura → tipo de gasto, participantes y obligaciones |
| DEBT | Deuda → deudor, acreedor |
| PAYMENT | Pago → obligación, deudor y receptor |
| CONFIRM / REJECT | Recepción → pago |
| FUND | Pago íntegro al proveedor → factura e integrante que adelanta |
| VOID | Corrección → operación original y motivo |

Catálogos, obligaciones, pagos, saldos y auditoría se reconstruyen de este historial.
Es una decisión deliberada frente a escribir cabecera, detalle, pago y auditoría en hojas separadas:
la operación de negocio se registra con una sola escritura de fila y las vistas son derivadas.
No se ofrece ACID de una base de datos relacional. No editar ni borrar manualmente los eventos.
Los editores del Spreadsheet son usuarios de confianza: Apps Script no impide alterar sus celdas.

Cada comando utiliza ScriptLock, valida referencias y saldos antes de escribir y tiene un identificador
persistente de idempotencia. Reintentar el mismo comando no genera un segundo pago. Reutilizar su ID
con datos distintos genera un conflicto. Un fallo después de guardar se resuelve reintentando
el comando original o actualizando las cuentas. No reintentar con otro ID sin revisar el resultado.

## Facturas y dinero adelantado

1. Sin pagador: cada integrante tiene una obligación con el proveedor.
2. Con pagador al crear la factura: cubre su parte y los demás le deben su aporte.
3. Una factura sin abonos puede pasar al caso 2 con FUND, indicando quién ya la pagó completa.
4. Los pagos parciales directos al proveedor se registran por obligación, por un administrador.
5. Los pagos entre integrantes quedan por confirmar y reservan saldo para impedir sobrepagos.
6. Solo el receptor o un administrador puede confirmar o rechazar. La confirmación reduce el saldo.

No se administra dinero recaudado en custodia todavía. Entregar dinero a un encargado que aún no paga
al proveedor no equivale a pagar la factura. La interfaz lo indica; no registrarlo como pago directo.
Tampoco se aceptan sobrepagos, anticipos sin obligación, dinero futuro ni devoluciones como montos negativos.

## Préstamos y compensación

El módulo nuevo registra deudas simples sin intereses, con fecha de entrega, vencimiento y abonos.
Los préstamos con amortización, interés y mora del código anterior siguen disponibles para administradores
en el módulo histórico, con sus datos separados. No combinar ambos saldos sin conciliación.

La propuesta de liquidación usa saldos globales entre integrantes, excluye dinero por confirmar y
no ejecuta transferencias ni cancela deudas. Puede proponer que A pague a C aunque A deba a B y B a C.
Una compensación requiere acuerdo externo y registro explícito de las obligaciones originales.

## Correcciones

Personas y tipos de gasto se archivan; mantienen referencias históricas. Los pagos se revierten con
motivo y conservan su evento original. Una deuda o factura con pagos no se anula hasta resolverlos.
Una factura financiada mediante FUND no se anula ni se desfinancia desde la UI en esta versión;
requiere un procedimiento de ajuste específico. Los registros anulados permanecen en auditoría.
Para corregir importes de una factura sin movimientos se anula y se crea de nuevo.

## Autorización y autenticación pendiente

Los endpoints nuevos verifican token, usuario activo y rol actual en el servidor. Los miembros reciben
solo sus obligaciones, pagos y participación; no reciben auditoría, propuestas globales ni vínculos de cuentas ajenas.
Roles administrativos reconocidos: `Admin`, `Administrador`, `Propietario`, sin distinguir mayúsculas.
Los endpoints heredados pasan por autorización de administrador. Sus implementaciones
y los helpers de conexión terminan en `_`, para no exponerlos como RPC de Apps Script.

La autenticación anterior se conserva temporalmente por decisión del alcance de esta fase.
**Todavía compara contraseñas en texto plano y no ofrece recuperación segura. No es una versión lista
para exposición pública con datos reales.** No se añadieron contraseñas reales ni claves al repositorio.
La siguiente fase sustituirá ese mecanismo por un proveedor administrado, incluyendo recuperación,
verificación de correo, revocación de sesiones y verificación de identidad desde Apps Script.
No se implementó un hash SHA-256 simple como sustituto de un algoritmo de contraseñas.

## Límites y evolución

El historial se lee completo para reconstruir saldos. Es apropiado para probar un hogar pequeño, pero
debe medirse con el volumen real antes de desplegar. No hay paginación ni compactación todavía.
El primer modelo limita a 100 participantes por gasto y 40.000 caracteres JSON por operación.
No incluye adjuntos, notificaciones por correo, automatización de cobros ni pagos bancarios.
Los respaldos y la restauración son procedimientos operativos, no tareas automáticas implementadas.
Para crecimiento, transferencias de custodia o conciliación bancaria, conviene una base relacional.
