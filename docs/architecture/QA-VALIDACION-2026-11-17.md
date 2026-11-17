# HOME: validación real QA — 17 de noviembre de 2026

## Resultado final

**PASS: 14 comprobaciones focalizadas del runtime real, más smoke de la web app publicada.**
La autorización fue concedida por el usuario. Se retiró el verificador temporal y se publicó la versión limpia 3.
No se repitieron auditorías ni baterías extensas; no se avanzó al siguiente módulo financiero.

Recomendación técnica: el refactor superó QA. Este cierre incorpora el manifiesto privado, la política
automatizada y las evidencias redactadas al PR #1, que debe permanecer abierto hasta el merge autorizado.
No autoriza desplegar producción ni migrar datos. Los IDs, enlaces de entorno, cuentas y credenciales
se conservan exclusivamente en la configuración/evidencia privada, fuera del repositorio.

## Referencia validada

- Repositorio: jose200101/Home.
- Rama: refactor/home-structure-v1.
- Commit: 7999df4123ba821ca1de41b2e277bccad477ea6f.
- PR: https://github.com/jose200101/Home/pull/1.
- Cuenta comprobada en Drive, clasp y editor: cuenta propietaria QA (verificada).
- Ajuste de fuente: src/gas/appsscript.json cambia ANYONE a MYSELF, con cambio de formato JSON.
- No se modificaron contraseñas existentes, autenticación ni funcionalidades financieras.

## Recursos QA independientes

| Recurso | ID | Enlace |
| --- | --- | --- |
| HOME-QA, Apps Script | `HOME_QA_SCRIPT_ID` | Registro privado de entorno |
| HOME-QA-Cuentas, Sheets | `HOME_QA_SPREADSHEET_ID` | Registro privado de entorno |

Ambos se crearon desde cero, privados, sin copiar el original. Sus IDs se comprobaron por lectura
y son distintos de los de producción. No se añadieron permisos de otros usuarios ni de acceso público.

La hoja contiene:

- USUARIOS: id, nombre_persona, nombre_usuario, contrasenia, rol, estado.
- Tres fixtures ficticios: qa-admin/qa_admin (Administrador activo), qa-member/qa_member
  (Integrante activo), qa-inactive/qa_inactive (Integrante inactivo).
- Credenciales de fixtures exclusivamente ficticias, omitidas de la evidencia versionada.
- PERSONAS: solo encabezados; sin datos personales.
- HOME_EVENTOS: encabezados y un único INIT qa-init-20261117, actor qa-admin, nombre HOME QA.
- No existen facturas, préstamos, deudas ni pagos QA. Los saldos visualizados son L 0.00.
- Zona de la hoja: America/Tegucigalpa; locale de conversión es_MX.

## Configuración efectiva y despliegue

- SPREADSHEET_ID = `HOME_QA_SPREADSHEET_ID`, guardada y releída
  después de recargar Configuración del proyecto HOME-QA.
- HOME_QA_LAST_RESULT contiene el resultado de pruebas sin tokens ni credenciales.
- .clasp.json local ignorado por Git: scriptId QA y rootDir dist.
- HOME_QA_SCRIPT_ID se estableció explícitamente antes de cada preflight y escritura remota.
- npm run build: HOME_BUILD_OK, 26 archivos.
- npm run qa:preflight: QA_PREFLIGHT_PASS.
- Manifiesto remoto final: V8, America/Tegucigalpa, USER_DEPLOYING, MYSELF.
- Los scopes originales se conservaron sin ampliarlos.
- Versión activa: 3; 26/26 archivos coinciden con dist después de normalizar CRLF/LF.
- HEAD y versión 3 no contienen QA_Runtime.js ni la función temporal de pruebas.

Deployment ID:
`HOME_QA_DEPLOYMENT_ID`

URL de la aplicación:
URL privada de QA, omitida de Git

Google mostró en Administrar las implementaciones: Versión 3, ejecutar como
cuenta propietaria QA (verificada) y usuarios con acceso Solo yo. Versiones 1 y 2 conservadas como historial,
sin ser la versión activa; contienen snapshots del verificador temporal.

## Resultados del runtime real

Ejecución completada: 17 de noviembre de 2026, 16:05:58 America/Tegucigalpa.

| Comprobación | Resultado |
| --- | --- |
| Destinos QA e identidad previamente comprobada | PASS |
| doGet y evaluación de plantillas HTML | PASS |
| Inclusiones HTML con subcarpetas | PASS |
| Lectura real de Sheets y fixtures | PASS |
| Token inválido rechazado | PASS |
| Cuenta ficticia inactiva rechazada | PASS |
| Sesiones ficticias y roles | PASS |
| Integrante no puede inicializar | PASS |
| Endpoints históricos deniegan al integrante | PASS |
| Inicialización administrativa | PASS |
| Idempotencia real sin segunda escritura | PASS |
| Privacidad del integrante | PASS |
| Ausencia de movimientos financieros QA | PASS |
| Cierre de sesión invalida token | PASS |

El verificador tenía guardas de ScriptApp.getScriptId(), Script Property, conexión efectiva,
fixtures exclusivamente qa-* y eventos esperados antes de cualquier escritura. Se usaron las funciones
reales de la aplicación, HtmlService, SpreadsheetApp, CacheService, LockService y PropertiesService.
La cuenta se verificó con clasp/editor, sin solicitar permisos extra de correo desde el verificador.

La inicialización registró una sola fila a las 22:05:53.071Z. El reintento del mismo comando devolvió
duplicate=true sin aumentar las filas. La lectura final de HOME_EVENTOS confirmó únicamente ese INIT.
Las sesiones del verificador se cerraron y sus tokens no se guardaron en la evidencia.

## Smoke sobre la versión limpia publicada

- La petición HTTP sin credenciales termina en accounts.google.com/v3/signin/identifier, sin HTML root de HOME.
- En el perfil Chrome de la cuenta propietaria, la URL QA abrió el login real.
- qa_admin inició sesión y cargó el hogar HOME QA ya configurado, menús administrativos y saldos L 0.00.
- Cerrar Sesión regresó al login. No se guardó la contraseña en el navegador.
- No se registraron integrantes, gastos ni movimientos financieros durante el smoke.

## Problemas encontrados y resolución

1. Autorización inicial pendiente: el usuario la concedió; no se omitieron advertencias ni se aprobaron permisos en su nombre.
2. El verificador consultaba Session.getEffectiveUser().getEmail(), que requería userinfo.email.
   Se quitó esa consulta del verificador; no se agregó el scope ni se modificó la aplicación.
3. clasp 3.3.0 no detectó una eliminación aislada: tras quitar QA_Runtime, push indicó up to date
   y una lectura mostró que la versión 2 aún lo contenía. Un cambio solo de formato del manifiesto
   activó la subida completa de 26 archivos. La versión 3 se descargó y comprobó limpia antes de activarla.
4. El navegador con varias cuentas devolvió No se encontró la página para la URL privada. Se probó
   en el perfil existente de la cuenta propietaria, con la cuenta propietaria, y la app funcionó. No se relajó MYSELF,
   no se cerraron cuentas Google ni se cambiaron configuraciones del navegador.

No queda un fallo funcional reproducido en las pruebas focalizadas. La autenticación administrada,
la recuperación segura y las demás limitaciones funcionales conocidas no se consideran resueltas por QA.

## Producción: sin escrituras de esta validación

Original Apps Script: `ORIGINAL_SCRIPT_ID`.
Original Spreadsheet: `ORIGINAL_SPREADSHEET_ID`.

Metadatos antes y al finalizar:

- Apps Script original: modifiedTime 2026-11-17T14:15:31.304Z, sin variación.
- Spreadsheet original: modifiedTime 2026-02-12T02:28:58.483Z, sin variación.

No hubo escrituras, cambios de propiedades, despliegues ni pruebas funcionales contra esos IDs.
Solo se consultó metadata. Al usar el perfil Chrome propietario se abrió una pestaña QA nueva;
no se interactuó con el diálogo del proyecto original que ya estaba abierto.
Los metadatos sin variación y los destinos registrados corroboran este alcance, pero no sustituyen
una auditoría de actividad de todos los actores de Google.

## Evidencia y reversión

Soporte local en el directorio hermano .qa-home/outputs/home-qa-20261117/:

- runtime-result.json: 14 PASS, obtenido del log real.
- runtime-pass.png: captura de la ejecución completada.
- qa-live-dashboard.png: aplicación QA publicada, autenticada con fixture.
- qa-settings.png: IDs y Script Property QA.
- readback-check.json: versión 3, 26 coincidencias, MYSELF, sin verificador.
- anonymous-check.json: redirección Google sin entregar la app.

La copia de solo lectura remote-version-3/source contiene la versión activa. El verificador de soporte
QaRuntime.js permanece solo fuera del repositorio desplegable; no forma parte de src/ ni dist/.

Para revertir o detener QA:

1. Confirmar .clasp.json y HOME_QA_SCRIPT_ID QA; nunca restaurar el ID original como destino.
2. Archivar exclusivamente el deployment QA desde Gestionar implementaciones o clasp undeploy
   con su ID exacto. No se ejecutó esta acción; QA queda disponible para revisión.
3. Conservar la hoja ficticia como evidencia; eliminarla solo bajo instrucción explícita.
4. Cualquier publicación posterior debe usar build limpio, qa:preflight y lectura del manifiesto remoto.
5. No volver a desplegar versiones 1/2 con el verificador; mantener MYSELF.
6. Si se revierte el cambio local del manifiesto, no publicar ANYONE como parte de esa reversión.

No hay cambios de producción que revertir. El PR sigue abierto y no se realizó merge.

## Evidencia versionada y política preventiva

Los resultados redactados están en [evidence/2026-11-17](evidence/2026-11-17/):

- runtime-result.json: 14 comprobaciones reales PASS; se retiraron IDs y cuenta.
- readback-check.json: versión 3 limpia, 26/26 coincidencias y MYSELF.
- anonymous-check.json: acceso sin credenciales enviado al login de Google.
- screenshot-index.json: hashes de las capturas privadas; no se publican imágenes con URLs de entorno.

La política `tools/deployment-policy.json` permite únicamente MYSELF en el manifiesto base, tanto para
webapp como para executionApi si se añade. `npm run check:deployment`, el build y GitHub Actions la
comprueban. Cambiar solo appsscript.json a ANYONE, ANYONE_ANONYMOUS, DOMAIN o dejar acceso implícito
falla antes de reemplazar dist. Ampliar el acceso exige una modificación explícita y revisada de la política
y sus pruebas; no se admite una variable de entorno o configuración local como excepción.

Esta ejecución de cierre solo incorpora archivos y validaciones locales/GitHub: no vuelve a desplegar
ni accede a Google Sheets o Apps Script. `.clasp.json` permanece ignorado y fuera del commit.


## Validación del commit de cierre

- Pruebas focalizadas de política: 5/5 PASS.
- npm test: 30/30 PASS, incluyendo el build de 26 archivos.
- npm run check:deployment: PASS.
- npm run smoke: PASS, con datos simulados y servidor local cerrado al terminar.
- git diff --check: PASS.
- GitHub Actions ejecuta política, npm test y smoke para el SHA del PR; su resultado se verifica
  después de publicar el commit, sin considerar ejecuciones anteriores como evidencia del nuevo SHA.
- .clasp.json sigue ignorado por Git y no se incluye en el commit. No hay nuevos despliegues
  ni escrituras a Google durante este cierre.
