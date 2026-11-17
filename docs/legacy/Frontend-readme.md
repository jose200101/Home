# Interfaz de Home

React 18 + Bootstrap 5, con JSX compilado mediante Babel en HTML Service.

- `Index.html`: carga estilos, dependencias y vistas.
- `_App.html`: acceso anterior, sesión, navegación y puente de RPC histórico.
- `views/HomeCuentas.html`: experiencia principal del hogar.
- `_HomeStyles.html`: estilos responsivos, formularios y diálogos accesibles.
- Otras vistas: consulta y mantenimiento de los módulos anteriores por administradores.

La experiencia principal incluye Resumen, Gastos compartidos, Deudas y pagos, Integrantes y Configuración.
Los diálogos tienen revisión antes de guardar, bloqueo de envío simultáneo y mensajes de error.
Las operaciones usan IDs estables para reintentar sin duplicar movimientos.

La vista de integrante es personal. Ocultar botones es solo presentación: el backend comprueba permisos
y filtra datos antes de enviarlos al navegador. El proveedor de autenticación y la recuperación siguen pendientes.

Desde la carpeta Home, `npm run preview` inicia una demostración con datos ficticios, sin conexión a Google.
La barra permite cambiar de administrador a integrante. `npm test` valida el motor y las puertas de acceso.
Las dependencias vía CDN y el runtime de HTML Service necesitan validación final antes del despliegue.
