# Changelog

Todas las versiones relevantes de Visor XLSX.

## [1.0.0] — 2026-10-20

### Añadido
- Visor local de `.xlsx`, `.xlsm`, `.xls` y `.csv` sin subir archivos.
- Ventana de app de escritorio mediante `launcher.vbs` (Chrome o Edge).
- Acceso directo con icono propio y arranque sin pestañas ni barra de direcciones.
- Pestañas por hoja, búsqueda de filas y exportación CSV con UTF-8.
- Inspector de celda con contenido completo y copia al portapapeles.
- Carga del rango usado completo, sin recortes de filas ni columnas.
- Fórmulas sin valor en caché visibles como fórmula.
- Demo integrada y diseño oscuro responsive.

### Corregido
- El estado vacío se superponía a la tabla por un conflicto con `hidden`.
- La escala de las celdas no era uniforme entre paneles de inventario del proyecto original (cambio no incluido en este repositorio).

### Seguridad
- Procesamiento 100 % local; sin backend, analíticas ni peticiones de red.

[1.0.0]: https://github.com/smouj/xlsx-viewer/releases/tag/v1.0.0
