# Changelog

Todas las versiones relevantes de Visor XLSX.

## En desarrollo

### Mejorado
- Cuadrícula clara de lectura con encabezados de fila y columna, selección visible e inspector multilínea.
- Recuperación de rellenos, colores, fuentes, alineación y bordes por celda en libros XLSX/XLSM.
- Se respetan los anchos, alturas y filas o columnas ocultas del libro cuando el formato los proporciona.
- Ajuste de columnas al contenido y ajuste de texto de dos líneas para celdas largas.
- Virtualización por altura real de fila para conservar el desplazamiento con filas personalizadas.
- El lector conserva los metadatos disponibles y procesa las definiciones de estilo incluidas en el paquete XLSX/XLSM.

### Añadido
- Lectura local de imágenes incrustadas y cuadros de texto a partir de las relaciones OOXML del libro.
- Gráficos de columnas, barras, líneas, áreas, circular y anillo reconstruidos en SVG; también se leen datos de los rangos de origen si el gráfico no incluye caché.
- Los objetos visuales se colocan sobre la cuadrícula, muestran una vista ampliada al pulsarlos y se pueden ocultar para despejar la hoja.
- Tarjetas informativas para imágenes u objetos incrustados con formato no compatible.

### Corregido
- Liberación de las URL temporales de las imágenes al abrir otro libro o cerrar la app.
- Los objetos visuales se ocultan durante una búsqueda para evitar que queden desalineados con las filas filtradas.

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
