# Web App con Login (Google Apps Script + React + Bootstrap 5)

Estructura:
- `Backend/` -> Apps Script (.gs + manifest)
- `Frontend/` -> HTML Service (React + Bootstrap + Babel)

## Qué hace
- Pantalla de login con estilo similar al video (fondo oscuro + dorado, iconos, botón "INGRESAR")
- Lee usuarios desde la pestaña `USUARIOS` en Google Sheets
- Valida `estado == "Activo"`
- Mantiene sesión simple por token (CacheService) + localStorage en el navegador
- Muestra un dashboard de ejemplo tras iniciar sesión (área protegida)

## Requisitos del Sheet
Pestaña: `USUARIOS`  
Encabezados (fila 1):
- `id`
- `nombreCompleto`
- `nombre_usuario`
- `contrasenia`
- `rol`
- `estado`

## Pasos de instalación (rápidos)
1. Crea un proyecto Apps Script (recomendado: **vinculado** al Spreadsheet).
2. Copia `Backend/Code.gs` y `Backend/appsscript.json`.
3. Si NO es vinculado al Spreadsheet, configura `CONFIG.SPREADSHEET_ID` en `Code.gs`.
4. Crea los HTML:
   - `Index` (Apps Script) / `Frontend/Index` (si usas carpetas con clasp) (de `Frontend/Index.html`)
   - `_Styles` / `Frontend/_Styles` (de `Frontend/_Styles.html`)
   - `_App` / `Frontend/_App` (de `Frontend/_App.html`)
5. Deploy como Web App y prueba.

## Nota de seguridad
Este ejemplo valida contraseñas en texto plano porque así está tu hoja. En un sistema real:
- guarda hashes (por ejemplo, SHA-256 + salt)
- agrega rate limiting y registros de auditoría
- evita exponer mensajes que permitan enumeración de usuarios


## Tip: SPREADSHEET_ID sin tocar código
Si tu proyecto NO está vinculado al Spreadsheet, puedes ir a **Project Settings → Script properties** y crear una propiedad:
- `SPREADSHEET_ID` = `<ID del Google Sheet>`

El backend ahora lo detecta automáticamente.


## Configuración DB (env / conexión)
- Backend/config/env.gs: variables de entorno (ID del Spreadsheet y nombres de hojas)
- Backend/database/conexion.gs: helpers de conexión y lectura

> Tip: también puedes definir una Script Property llamada `SPREADSHEET_ID` para no tocar el código.
