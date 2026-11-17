# Home · Cuentas del hogar

Web app de Google Apps Script + Google Sheets, con React y Bootstrap en HTML Service.
Administra servicios compartidos con participantes independientes, facturas, deudas entre integrantes,
adelantos, pagos por confirmar y estados de cuenta en lempiras.

## Estado

Implementación local del primer modelo. No se desplegó, no se migraron saldos y no se modificó el Spreadsheet real.
El proveedor de autenticación y la recuperación de contraseña están pendientes por decisión del alcance actual.
El login anterior aún compara contraseñas en texto plano: cerrar esa fase antes de uso público con datos reales.

## Desarrollo

```powershell
npm test
npm run preview
```

Vista previa: http://127.0.0.1:4173. Usa datos ficticios en memoria y el mismo motor de cuentas.
La barra de demostración permite ver la aplicación como administrador o integrante.
Los módulos históricos no se simulan en el navegador local.

## Estructura

- `Backend/HomeLedger.js`: validación, historial, proyecciones, comandos y consultas autorizadas.
- `Backend/HomeAccess.js` y `LegacyEndpoints.js`: autorización de módulos anteriores.
- `Frontend/views/HomeCuentas.html`: resumen, gastos, facturas, deudas, pagos, integrantes y configuración.
- `Frontend/_HomeStyles.html`: interfaz adaptable a móvil.
- `Backend/config/env.js`: conexión configurable mediante Script Property `SPREADSHEET_ID`.
- `tests/`: pruebas de integridad y permisos con servicios simulados.
- `tools/preview.cjs`: servidor local, sin llamadas a Google.

## Preparación de Apps Script

1. Usar un proyecto y una copia de Spreadsheet de prueba, separados de los datos reales.
2. Configurar `SPREADSHEET_ID` en Script Properties; el código conserva el ID anterior como fallback.
3. Mantener nombres de archivos y carpetas de HTML; `doGet` sirve `Frontend/Index`.
4. La hoja `USUARIOS` conserva `id`, `nombre_persona` (o `nombreCompleto`), `nombre_usuario`,
   `contrasenia`, `rol` y `estado`. El acceso exige estado `Activo`.
5. Para configurar Home se requiere rol `Admin`, `Administrador` o `Propietario`.
6. Configurar el hogar, agregar integrantes, vincular cuentas y crear tipos de gasto desde la interfaz.
   Se crea `HOME_EVENTOS` con la primera operación; el módulo nuevo no escribe hojas históricas.
7. Completar autenticación y validar HTML Service antes de habilitar usuarios reales.

No ejecutar `clasp push` sobre el proyecto real sin revisar destino y diff. `.claspignore` excluye
pruebas, herramientas, documentación y capturas. No publicar la vista previa como sustituto de Apps Script.

## Documentación

- [Modelo, reglas, permisos y límites](docs/MODELO-HOME.md)
- [Transición, pruebas y pendientes de activación](docs/TRANSICION-Y-VALIDACION.md)
