# Frontend (React + Bootstrap 5 + Babel dentro de Apps Script)

Estos archivos están pensados para usarse con **HTML Service** en Apps Script.

## Archivos
- `Index` (Apps Script) / `Frontend/Index` (si usas carpetas con clasp): página principal que carga Bootstrap, React y Babel, e incluye parciales.
- `_Styles` / `Frontend/_Styles`: estilos para replicar el look & feel del video (fondo oscuro + dorado).
- `_App` / `Frontend/_App`: App React con Login y Dashboard placeholder.

## Cómo usar
En el editor de Apps Script:
1. Crea 3 archivos HTML con estos mismos nombres:
   - `Index` (Apps Script) / `Frontend/Index` (si usas carpetas con clasp)
   - `_Styles` / `Frontend/_Styles`
   - `_App` / `Frontend/_App`
2. Pega el contenido correspondiente en cada uno.
3. Deploy como Web App.

## Personalización rápida
- Cambia el nombre/branding en `_App` / `Frontend/_App` (texto "Wilito BarberShop" y subtítulo).
- Ajusta colores en `_Styles` / `Frontend/_Styles` (variables `--gold`, etc.).


## Tip: SPREADSHEET_ID sin tocar código
Si tu proyecto NO está vinculado al Spreadsheet, puedes ir a **Project Settings → Script properties** y crear una propiedad:
- `SPREADSHEET_ID` = `<ID del Google Sheet>`

El backend ahora lo detecta automáticamente.
