# Visor XLSX

**Abre Excel en Windows como una app. Rápido, local y sin subir nada.**

[![License: MIT](https://img.shields.io/badge/license-MIT-b8f34a)](LICENSE)
[![Version](https://img.shields.io/badge/version-1.0.0-4c8bf5)](https://github.com/smouj/xlsx-viewer/releases)
[![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20Web-555)](#instalación)
[![Local-first](https://img.shields.io/badge/privacy-100%25%20local-b8f34a)](#privacidad)
[![SheetJS](https://img.shields.io/badge/SheetJS-0.20.3-ea6a1a)](https://sheetjs.com/)
[![Demo](https://img.shields.io/badge/demo-online-b8f34a)](https://smouj.github.io/xlsx-viewer/)

**Demo web:** https://smouj.github.io/xlsx-viewer/

![Visor XLSX con una hoja abierta](docs/screenshot.png)

Visor de hojas de cálculo **.xlsx**, **.xlsm**, **.xls** y **.csv** que se ejecuta en tu equipo. Sin servidor, sin cuenta, sin telemetría y sin subir el archivo a ninguna parte.

## Características

- **Ventana de app**: se abre sin pestañas ni barra de direcciones, con icono y acceso directo.
- **Todas las hojas** del libro, con pestañas y búsqueda por fila.
- **Todas las filas y columnas visibles** del rango usado; la tabla virtualiza el dibujado para responder con archivos grandes.
- **Vista de hoja clara y legible**, con anchos y alturas guardados en el libro, y filas y columnas ocultas respetadas.
- **Formato visual XLSX/XLSM**: rellenos, colores de texto, negrita, cursiva, tamaño de fuente, alineación, bordes y ajuste de texto de celda.
- **Imágenes, gráficos y cuadros de texto** incrustados en `.xlsx` y `.xlsm`, colocados sobre la hoja según su ancla en el libro; clic para ampliar.
- **Imágenes insertadas dentro de una celda**: se muestran en su celda, con texto alternativo y vista ampliada al pulsarlas.
- **Gráficos reconstruidos localmente** como SVG, con sus categorías y series guardadas o leídas desde el rango de celdas de origen.
- **Control para ocultar los objetos visuales** y despejar la hoja mientras se leen los datos.
- **Ajustar columnas** al contenido o restaurar sus anchos originales; activar el ajuste de texto para leer celdas largas.
- **Zoom de hoja del 50 % al 200 %**, con controles y atajos `Ctrl` + `+` / `Ctrl` + `-`; el desplazamiento conserva su posición.
- **Inspector de celda**: clic en cualquier celda para ver el contenido completo y copiarlo.
- **Fórmulas sin valor en caché** se muestran como fórmula en azul, sin dejar la celda vacía.
- **Exportar CSV** de la hoja activa, con UTF-8 y BOM para Excel.
- **100 % local**: funciona sin conexión y no hace peticiones de red.
- **Cero instalación**: HTML + JavaScript y la librería incluida en el repo.

## Instalación

1. Descarga el ZIP desde [Releases](https://github.com/smouj/xlsx-viewer/releases).
2. Descomprímelo en una carpeta permanente.
3. Doble clic en **Visor XLSX** (acceso directo) o en `launcher.vbs`.

No requiere administrador, instalador ni dependencias. Si mueves la carpeta, crea un acceso directo nuevo a `launcher.vbs` con `app.ico`.

### Desde el código

```bash
git clone https://github.com/smouj/xlsx-viewer.git
cd xlsx-viewer
```

Abre `index.html` en Chrome, Edge o Firefox. Para la ventana de app en Windows, ejecuta `launcher.vbs`.

## Uso

- Arrastra el archivo a la ventana o pulsa **Abrir archivo**.
- Cambia de hoja con las pestañas.
- Escribe en el buscador para filtrar filas de la hoja activa.
- Haz clic en una celda para ver su valor completo en la barra inferior.
- Usa `−`, `+` o el porcentaje para ajustar el zoom de lectura.
- **Exportar CSV** descarga la hoja activa.

| Atajo | Acción |
|---|---|
| `Ctrl+O` | Abrir archivo |
| `←` / `→` | Hoja anterior / siguiente |
| `Esc` | Limpiar la búsqueda |
| `Ctrl` + `+` / `Ctrl` + `-` | Aumentar / reducir zoom |

## Privacidad

El archivo se lee con `FileReader` y se procesa en memoria dentro del navegador. No hay backend, analíticas, cookies ni peticiones de red. El visor no modifica ni guarda el archivo original.

## Fidelidad de datos

- Se carga el rango usado completo de cada hoja, sin recortes artificiales de filas ni columnas.
- Las celdas muestran el valor formateado que guardó Excel; el inspector permite ver y copiar el contenido íntegro.
- Se respetan las dimensiones y el estado oculto de filas y columnas. El ajuste de texto muestra hasta dos líneas en la cuadrícula; el inspector conserva el texto completo.
- Se recuperan estilos de celda habituales, imágenes PNG/JPEG/GIF/BMP/WebP/SVG flotantes o insertadas dentro de celdas, cuadros de texto y gráficos de columnas, barras, líneas, áreas, circular y anillo.
- Los anclajes de los objetos visuales usan los anchos y alturas mostrados, incluidos los tamaños guardados y las filas o columnas ocultas. Las imágenes de celda se ajustan a su celda y se pueden ampliar con un clic.
- Los gráficos se vuelven a dibujar para que puedan leerse en el visor; no se conserva exactamente la paleta, tipografía, ejes ni disposición de Excel. Dispersión, radar, burbuja y área se muestran como líneas; los gráficos combinados se muestran según su primer tipo. Se utilizan las cachés del gráfico o los datos de las celdas de origen.
- Para controlar el uso de memoria, la lectura de estilos y objetos OOXML se aplica a libros XLSX/XLSM de hasta 64 MB. En archivos mayores se carga la cuadrícula y los valores, pero se omite la extracción completa de estilos, imágenes y gráficos.
- El formato condicional, tablas dinámicas, segmentadores, SmartArt, objetos OLE y gráficos 3D no se reproducen. Si un objeto incrustado no se puede mostrar, se indica en su posición con una tarjeta informativa.
- Las fórmulas muestran su último valor calculado. Si el archivo no trae valor en caché, se muestra la fórmula.
- Las celdas combinadas conservan el valor en su celda superior izquierda, igual que Excel.

## Capturas

![Pantalla inicial](docs/empty.png)

## Estructura

```
xlsx-viewer/
├─ index.html                 # interfaz y lógica del visor
├─ visuals.js                 # lectura de objetos OOXML y previsualización de gráficos
├─ launcher.vbs               # ventana de app con Chrome o Edge
├─ app.ico                    # icono del acceso directo
├─ vendor/
│  ├─ xlsx.full.min.js        # SheetJS CE 0.20.3
│  └─ SHEETJS_LICENSE.txt     # Apache-2.0
├─ docs/
│  ├─ screenshot.png
│  └─ empty.png
├─ LICENSE
└─ THIRD_PARTY_NOTICES.md
```

## Desarrollo

No hay build ni gestor de paquetes: es HTML, CSS y JavaScript sin frameworks. SheetJS está incluido en `vendor/`; el lector de imágenes y gráficos forma parte de `visuals.js` y no requiere servicios externos. Pulsa **Ver ejemplo** para ver el gráfico de muestra.

Para probar cambios, abre `index.html` y usa **Ver ejemplo**. La app incluye una hoja de demostración con dos pestañas.

## Limitaciones

- Visor de solo lectura: no edita ni guarda cambios en el archivo.
- Los archivos protegidos con contraseña no se pueden abrir.
- Hojas extremadamente grandes dependen de la memoria disponible del navegador.
- Las celdas combinadas se muestran en su celda de origen, sin repetir el valor.

## Licencia

Código del visor: [MIT](LICENSE). SheetJS Community Edition: Apache-2.0, detallado en [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).

Hecho por [SMOUJ013](https://github.com/smouj).
