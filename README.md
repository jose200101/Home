# HOME — Administración financiera del hogar

Proyecto basado en **Google Apps Script V8 + Google Sheets + React (JavaScript)**. Google Drive será el almacén de comprobantes cuando se implemente ese módulo. HNL y `America/Tegucigalpa`.

## Estado

**Fase 1: reorganización estructural.** No modifica reglas financieras, despliegues ni datos reales. La implementación actual todavía conserva módulos heredados para permitir una sustitución segura posterior. El modelo funcional de HOME v1 está documentado en [Especificación HOME v1](docs/functional/HOME-v1.md).

La autenticación actual es heredada y su reemplazo se ha pospuesto. **No desplegar públicamente con datos financieros reales** antes de corregirla y validar permisos.

## Desarrollo local

Requiere Node.js 20 o posterior.

```bash
npm test
npm run preview
npm run build
```

`npm test` construye `dist/` y ejecuta las pruebas existentes más las pruebas de estructura. `npm run preview` construye `dist/` y sirve una demostración local con datos **ficticios**, sin conexión a Google. `npm run build` copia fuentes a `dist/` conservando exactamente los nombres de archivos del proyecto Apps Script vigente.

## Estructura

- `src/server/domain/`: reglas de dominio futuras (crear por módulo, en próximos commits).
- `src/server/modules/home/`: motor financiero actual (`HomeLedger.js`), pendiente de división modular.
- `src/server/legacy/`: implementación histórica en transición; no usar para nuevas reglas.
- `src/server/infrastructure/`: configuración y conexión actual a Sheets.
- `src/client/`: interfaz y vistas existentes, separadas de vistas históricas.
- `src/gas/appsscript.json`: manifiesto de Apps Script.
- `tools/build-manifest.cjs`: mapa fuente → nombre GAS conservado.
- `dist/`: salida generada, ignorada por Git, **único destino de clasp**.
- `docs/legacy/`: documentación anterior conservada como referencia histórica.

Revisar [Arquitectura y despliegue](docs/architecture/STRUCTURE.md), [Transición de datos](docs/migration/PLAN-DATOS.md) y [Mapa de pantallas](docs/ux-ui/MAP.md).

## Precauciones

No ejecutar `clasp push` sobre el script real ni migrar datos hasta validar primero una copia aislada de Apps Script y de Google Sheets. `main` debe permanecer sin cambios hasta revisar la rama y sus pruebas. No registrar credenciales ni IDs de usuarios en documentación o pruebas.
