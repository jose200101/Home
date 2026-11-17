# HOME — Organización de código y empaquetado de Apps Script

## Decisión de arquitectura

El código fuente vive en `src/`; `dist/` se genera y es el único directorio que `clasp` debe publicar. Apps Script no ejecuta módulos CommonJS/ESM como Node.js: los scripts GAS se publican como archivos globales. **Esta fase solo reubica fuentes; no introduce módulos import/export ni modifica lógica.**

## Estructura actual de transición

- `src/server/api/Code.js`: entrada `doGet`, helpers HTML y autenticación heredada.
- `src/server/modules/home/HomeLedger.js`: libro de eventos existente, sin cambios.
- `src/server/auth/HomeAccess.js`: comprobaciones existentes.
- `src/server/infrastructure/config/env.js` y `sheets/conexion.js`: acceso actual a Spreadsheet.
- `src/server/legacy/`: funciones y endpoints históricos preservados durante la transición.
- `src/client/Index.html`, `app/`, `styles/`, `pages/home/`, `pages/legacy/`: HTML/React.
- `src/gas/appsscript.json`: manifiesto de GAS.

Carpetas de dominio futuras se crearán conforme se implementen los módulos; no generar archivos vacíos ni funciones ficticias.

## Compilación

Ejecutar `npm run build`: `tools/build.cjs` valida la existencia de archivos y copia exactamente los bytes según `tools/build-manifest.cjs`. Ejemplo: `src/server/api/Code.js` → `dist/Backend/Code.js`, `src/client/pages/home/HomeCuentas.html` → `dist/Frontend/views/HomeCuentas.html`.

Los nombres de salida preservan temporalmente las llamadas `HtmlService.createTemplateFromFile('Frontend/Index')` y `include('views/HomeCuentas')`. No deben cambiarse a ciegas.

`npm test` y `npm run preview` compilan primero. Los tests y la vista previa leen la salida `dist/`. CI comprueba que las copias y plantillas existan. El comando `npm run smoke` levanta la vista previa, verifica HTML y RPC con usuarios ficticios, y comprueba que el integrante no reciba catálogos ni auditoría de administración.

## Despliegue aislado

1. Crear un proyecto Apps Script de pruebas **distinto del real** y una copia de Spreadsheet.
2. Verificar `SPREADSHEET_ID` en Script Properties del proyecto de pruebas y que no apunte a datos productivos.
3. Ejecutar `npm test` y `npm run build`.
4. Revisar `clasp status` y el destino activo; no ejecutar `clasp push` sobre el script real.
5. Validar en el entorno de pruebas el uso de nombres con subcarpetas, las inclusiones HTML y los RPC.
6. Solo después de aprobación explícita planificar el despliegue.

`dist/` no se versiona. Nunca editarla manualmente.

## Trabajo pendiente

Extraer motor de dominio de `HomeLedger.js` a módulos independientes, diseñar eventos v2 y sustituir interfaz heredada. Aplicar cada cambio en un commit separado y conservar capacidad de retorno.
