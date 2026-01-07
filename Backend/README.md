# Backend (Google Apps Script)

## Archivos
- `Code.gs`: Endpoints para login (`login`, `validateToken`, `logout`) y `doGet()`.
- `appsscript.json`: manifiesto con scopes mínimos para leer Spreadsheet.

## Configuración rápida
1. Abre Apps Script (script.google.com) y crea un proyecto (ideal: **vinculado** a tu Spreadsheet).
2. Copia el contenido de `Code.gs`.
3. (Opcional pero recomendado si NO es vinculado) En `CONFIG.SPREADSHEET_ID` pega el ID del Spreadsheet.
4. Crea archivos HTML en Apps Script usando los de la carpeta **Frontend**:
   - `Index` (Apps Script) / `Frontend/Index` (si usas carpetas con clasp)
   - `_Styles` / `Frontend/_Styles`
   - `_App` / `Frontend/_App`
5. Deploy > New deployment > Web app:
   - Execute as: **Me**
   - Who has access: **Anyone** (o tu dominio)
6. Abre la URL de la web app y prueba con un usuario en la hoja `USUARIOS`.

## Hoja esperada
Pestaña: `USUARIOS`
Encabezados (fila 1):
- `id`, `nombreCompleto`, `nombre_usuario`, `contrasenia`, `rol`, `estado`

> Nota: este ejemplo compara contraseñas en texto plano porque así está tu hoja. En producción, guarda hashes.


## Tip: SPREADSHEET_ID sin tocar código
Si tu proyecto NO está vinculado al Spreadsheet, puedes ir a **Project Settings → Script properties** y crear una propiedad:
- `SPREADSHEET_ID` = `<ID del Google Sheet>`

El backend ahora lo detecta automáticamente.
