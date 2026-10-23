# Koto — Procesador de Textos Minimalista

Koto es un procesador de textos ligero inspirado en Microsoft Word. Funciona directamente en el navegador, sin instalar nada y sin conexión a internet: basta con copiar esta carpeta completa y abrir `index.html`.

![Screenshot](https://github.com/jglazarte/Koto/blob/master/koto.png)

## Cómo ejecutarlo

1. Copia la carpeta completa a cualquier PC.
2. Haz doble clic en `index.html` (o ábrelo con Chrome, Edge, Firefox, etc.).

No requiere servidor ni instalación. Todo (CSS, JavaScript y librerías) está incluido localmente en la carpeta.

## Funciones

- **Archivo:** nuevo, abrir, cerrar, guardar, guardar como, exportar a PDF e imprimir.
- **Formatos de archivo:**
  - Abrir: `.docx`, `.odt`, `.html`, `.htm`, `.txt` y `.md`.
  - Guardar: `.odt` (predeterminado para documentos nuevos), `.docx`, `.html` y `.txt`.
  - Al guardar un archivo abierto se respeta su formato original. «Guardar como» permite elegir otro.
- **Formato:** fuente, tamaño, interlineado (simple, 1,5 y doble), izquierda, centrar, derecha, justificar, negrita, cursiva, subrayado, subíndice y superíndice.
- **Enlaces:** botón de enlace o `Ctrl + K`. Se indica el nombre visible y la dirección. El enlace se muestra en azul y subrayado sin cambiar la fuente, el tamaño ni el interlineado. `Ctrl + clic` abre el enlace; un clic normal permite editarlo.
- **Copia local de recuperación:** mientras hay cambios sin guardar, la aplicación guarda una copia en el navegador (a los 2 segundos de dejar de escribir y cada 30 segundos). Si se cierra la pestaña o el programa inesperadamente, al volver a abrir se ofrece recuperar esa copia. La copia se borra al guardar el archivo. No reemplaza el guardado normal.
- **Pegar sin formato:** `Ctrl + Shift + V` inserta el texto del portapapeles sin conservar negritas, colores, tablas ni otros formatos externos.
- **Buscar y reemplazar:** `Ctrl + F` abre el buscador y `Ctrl + H` lleva al campo de reemplazo. Resalta todas las coincidencias, permite ir a la anterior o siguiente y reemplazar una o todas. `Esc` cierra el buscador.
- **Modo enfoque:** `Ctrl + Shift + F` oculta las barras para concentrarse en la hoja. Se sale con `Esc` o con el botón «Salir del enfoque». Para que la aplicación abra siempre directamente en este modo, marca **«Abrir siempre»** en la pastilla superior derecha; la preferencia queda guardada en el navegador.
- **Interlineado 1,5 por defecto.**
- Hoja con ancho de página A4 (21 cm), centrada y con márgenes.
- **Barra de zoom inferior** (50 % a 400 %) con botón «Ajustar ancho» para que la hoja ocupe todo el ancho de la ventana.
- **Tabulador con sangría** (Tab). Mayús+Tab reduce la sangría.
- Exportación a PDF mediante `html2canvas` + `jsPDF` (incluidas en `vendor/`). El PDF conserva la apariencia de la hoja, pero el texto queda como imagen y no es seleccionable.
- Impresión optimizada para A4. Para obtener un PDF con texto seleccionable, usa «Imprimir» y elige «Guardar como PDF» en el navegador.

## Atajos de teclado

| Atajo | Acción |
| --- | --- |
| `Ctrl + N` | Nuevo archivo |
| `Ctrl + O` | Abrir |
| `Ctrl + S` | Guardar |
| `Ctrl + Shift + S` | Guardar como |
| `Ctrl + K` | Insertar o editar enlace |
| `Ctrl + F` | Buscar |
| `Ctrl + H` | Buscar y reemplazar |
| `Ctrl + Shift + V` | Pegar sin formato |
| `Ctrl + Shift + F` | Activar o salir del modo enfoque |
| `Esc` | Cerrar el buscador o salir del modo enfoque |
| `Ctrl + P` | Imprimir |
| `Ctrl + B` | Negrita |
| `Ctrl + I` | Cursiva |
| `Ctrl + U` | Subrayado |
| `Ctrl + Shift + L` | Alinear a la izquierda |
| `Ctrl + Shift + E` | Centrar |
| `Ctrl + Shift + R` | Alinear a la derecha |
| `Ctrl + Shift + J` | Justificar |
| `Tab` | Insertar sangría |
| `Shift + Tab` | Quitar sangría |

## Qué se conserva al abrir y guardar

En documentos `.docx` y `.odt` se conservan, dentro de lo que admite el editor:

- Párrafos y saltos de línea.
- Fuente, tamaño, color, negrita, cursiva, subrayado, subíndice y superíndice.
- Alineación, interlineado y sangría de primera línea.
- Enlaces.
- Tabulaciones.

**Limitaciones importantes:** las imágenes, tablas, notas al pie o finales, comentarios, control de cambios, encabezados, pies de página, cuadros de texto y otros elementos avanzados no se conservan al guardar desde este procesador. Al abrir un `.docx` que contenga alguno de ellos se muestra un aviso. Para documentos con ese tipo de contenido, conviene editarlos en Word/LibreOffice y usar este procesador solo con el texto.

El formato `.doc` binario (Word 97-2003) no puede interpretarse directamente en el navegador. Si el archivo es en realidad un `.docx` renombrado, se abre igualmente; si es binario, la aplicación pide guardarlo como `.docx` o `.odt`.

## Estructura

```
Procesador/
├── index.html          Interfaz principal
├── favicon.svg         Ícono de la aplicación (kanji 言)
├── css/
│   └── estilos.css     Estilos de la aplicación y de impresión
├── js/
│   ├── app.js          Lógica: formato, archivos, enlaces, búsqueda, copia local, PDF e impresión
│   └── formatos.js     Lectura y escritura de DOCX y ODT
├── vendor/
│   ├── html2canvas.min.js
│   ├── jspdf.umd.min.js
│   ├── jszip.min.js
│   └── mammoth.browser.min.js
└── README.md
```

## Notas

- En Chrome/Edge, abrir y guardar usa la API de archivos del navegador, por lo que puede sobrescribir el archivo original.
- En navegadores sin esa API (por ejemplo Firefox), «Guardar» descarga el documento y «Guardar como» muestra un diálogo propio para elegir nombre y formato.
- Un documento nuevo se guarda como `documento.odt` si no se indica otra cosa.
- El formato HTML guarda el contenido del editor, no una página completa con estilos.
- Los documentos nunca se envían a internet: todo el procesamiento y la copia local de recuperación ocurren en el navegador.
- La copia local se guarda en el almacenamiento del navegador. Si se borran los datos del sitio o se usa una ventana privada, esa copia puede desaparecer; guarda siempre el archivo en disco.
- Al recuperar una copia, el documento no queda vinculado al archivo original: «Guardar» descargará una nueva copia con el nombre recordado.
- La búsqueda no distingue mayúsculas de minúsculas.
- Conviene conservar copias de seguridad de los originales antes de sobrescribirlos.
