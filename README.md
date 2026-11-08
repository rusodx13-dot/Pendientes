# Control de Mercadería

Cuenta y orden (AgroBayer) y mercadería normal (Agrosoluciones). La página vive en GitHub Pages y guarda todo en una planilla de Google Sheets de tu Drive, a través de un Apps Script.

## Puesta en marcha (una sola vez)
1. Creá una planilla nueva en Google Drive llamada **Control de Mercadería**.
2. En la planilla: **Extensiones → Apps Script**. Borrá lo que haya y pegá el contenido de `apps-script/Code.gs`. Guardá.
3. **Implementar → Nueva implementación → Aplicación web**: *Ejecutar como* **Yo**, *Quién tiene acceso* **Cualquier usuario**. Autorizá los permisos y copiá la URL que termina en `/exec`.
4. Pegá esa URL en `config.js` (`window.MERCADERIA_API = '...'`) y subí el cambio.
5. En GitHub: **Settings → Pages →** Branch `main`, carpeta `/ (root)`.
6. Abrí la página, creá el usuario administrador e importá los dos reportes.

Las solapas de la planilla (Movimientos, Pedidos, Estado, Stock, FacturasMano) se crean solas. `Usuarios` y `Sesiones` quedan ocultas.

## Uso semanal
Importar reportes → subir `ConsultaPendientes.xls` y `facturacion_clientes_2026.xlsx` de la semana. Solo se agrega lo nuevo.

## Si cambiás `Code.gs`
Hay que volver a implementar: **Implementar → Administrar implementaciones → editar → Versión nueva**. La URL no cambia.
