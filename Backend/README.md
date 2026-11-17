# Backend de Home

Google Apps Script con V8. La conexión se obtiene desde `conexion_()` y Script Properties.

- `HomeLedger.js`: motor del hogar; comandos completos en `HOME_EVENTOS`, importes en centavos.
- `HomeAccess.js`: sesión y autorización de los endpoints históricos.
- `LegacyEndpoints.js`: funciones públicas protegidas; delegan a implementaciones privadas terminadas en `_`.
- `Code.js`: HTML Service y login anterior, pendiente de reemplazo por el proveedor de identidad.
- `PersonasUsuarios`, `GastosFijos`, `GastosVariables`, `GastosNetos`, `Prestamos`: módulos anteriores.

Los endpoints de Home reciben el token explícitamente. Los históricos reciben como último argumento
`{homeToken: token}`; el wrapper del frontend lo agrega automáticamente. No se debe invocar directamente
una implementación privada desde el navegador. Todos los permisos se verifican en el servidor.

Los módulos anteriores no se mezclan automáticamente con los saldos nuevos. Ver
[modelo y límites](../docs/MODELO-HOME.md) y [transición](../docs/TRANSICION-Y-VALIDACION.md).

No se desplegó ni se migraron datos reales. El login anterior sigue usando contraseñas en texto plano;
la integración de identidad y recuperación queda pendiente y es requisito antes de exposición pública.
