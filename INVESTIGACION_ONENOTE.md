# Investigación Profunda de Microsoft OneNote y Arquitectura del Clon Web (Mac & Chromebook + Google Drive)

## 1. Análisis Exhaustivo de Microsoft OneNote

Microsoft OneNote se distingue fundamentalmente de procesadores de texto lineales (como Microsoft Word o Google Docs) y de aplicaciones de notas basadas en bloques rígidos (como Notion) por su **paradigma de lienzo libre multimodal (*Freeform Infinite Canvas*)** y su motor de **tinta digital (*Digital Inking*)** de primera clase.

---

### 1.1. Funcionalidades Clave para Tomar Notas con Boli Táctil (*Stylus / Inking*)

En Microsoft OneNote (pestaña **Dibujar / Draw**), la experiencia con lápiz óptico está diseñada para imitar un cuaderno físico sin perder las ventajas de la edición digital:

| Funcionalidad en OneNote | Comportamiento Técnico y UX | Implementación en nuestro Clon Web |
| :--- | :--- | :--- |
| **Detección de Puntero y *Palm Rejection*** | Distingue automáticamente entre el lápiz óptico (`pen`), el dedo (`touch`) y el ratón/trackpad (`mouse`). Cuando el lápiz toca la pantalla, rechaza los toques accidentales de la palma de la mano. | API nativa **`PointerEvent`** (`e.pointerType`). Modo exclusivo **Solo Boli Táctil** (el dedo desplaza el lienzo mientras el stylus dibuja) o **Dibujar también con dedo/ratón**. |
| **Sensibilidad a la Presión y Suavizado** | El grosor del trazo varía dinámicamente según la fuerza ejercida sobre la punta del lápiz (`0.0` a `1.0`) y se suaviza mediante curvas vectoriales para evitar trazos poligonales. | Lectura de `e.pressure`, captura de sub-frames a 120Hz con `e.getCoalescedEvents()` e interpolación mediante **curvas cuadráticas de Bézier**. |
| **Bolígrafos (*Pens*) y Tintas Especiales** | Permite configurar varios bolígrafos favoritos con colores sólidos y efectos especiales (Arcoíris, Galaxia, Oro) y distintos grosores en milímetros. | Selector de bolígrafos personalizables, paleta de colores rápidos, selector de color libre y **Tinta Arcoíris** dinámica integrada. |
| **Marcador de Resaltado (*Highlighter*)** | Punta ancha semitransparente que se funde con el contenido inferior sin ocultar el texto ni las imágenes. | Trazado con punta biselada/redondeada, opacidad calibrada y mezcla multiplicativa (`multiply`) para que nunca tape el texto ni las fotos. |
| **Borrador de Trazos vs. Borrador de Píxeles** | OneNote ofrece el **Borrador de trazos (*Stroke Eraser*)** —tocar cualquier parte de una línea borra el trazo entero al instante— y el borrador parcial. | **Borrador de Trazos Vectorial** (elimina el trazo completo al cruzar su caja/segmentos) + **Borrador Parcial** de precisión. |
| **Selección de Lazo (*Lasso Select*)** | Permite rodear con el boli un conjunto de trazos manuscritos para moverlos juntos por el lienzo, cambiarles el color o borrarlos. | Herramienta **Lazo (`Lasso`)** con algoritmo *Point-in-Polygon (Ray-Casting)* que selecciona trazos y permite **arrastrarlos en bloque, recolorearlos o eliminarlos**. |
| **Formas Geométricas (*Ink to Shape*)** | Herramientas para trazar líneas rectas, flechas, rectángulos y círculos limpios para diagramas y mapas mentales. | Herramientas dedicadas de **Línea, Flecha, Rectángulo y Círculo/Elipse** con previsualización en tiempo real al arrastrar el stylus. |
| **Conversión Tinta a Texto / Dictado** | Reconoce notas o permite dictar/convertir apuntes en contenedores de texto editables. | Función de **Dictado por Voz (`Web Speech API`)** + creación rápida de contenedores en el lienzo. |
| **Fondos de Papel (*Rule Lines & Grid*)** | Renglones de cuaderno (estrechos/anchos) y cuadrículas matemáticas (pequeñas/grandes) que guían la escritura a mano alzada. | Selector instantáneo de **Papel en Blanco, Renglones (Pauta), Cuadrícula (Grid) y Punteado (Dot Grid)** + colores de papel pastel. |

---

### 1.2. Toma de Notas con Teclado e Imágenes en el Lienzo Libre

1. **Contenedores de Texto Flotantes (*Note Containers*)**:
   - En OneNote no estás atado a escribir desde la esquina superior izquierda: **haces clic en cualquier lugar del lienzo** y nace una caja de texto independiente.
   - Cada contenedor tiene un **asa superior de arrastre** para reubicarlo en cualquier coordenada `(x, y)`, bordes redimensionables y soporte para texto enriquecido (**Negrita, Cursiva, Subrayado, Listas, Títulos**).
   - **Etiquetas Interactivas (*OneNote Tags*)**: Casillas de verificación (**To-Do**), **Importante (⭐)**, **Pregunta (❓)** e **Idea (💡)** que se pueden insertar con un clic.
2. **Imágenes Flotantes y Anotación sobre Imágenes**:
   - En OneNote es muy común pegar diapositivas de clase, capturas de pantalla o diagramas y **escribir con el boli táctil directamente encima de la imagen** (subrayar conceptos, sacar flechas hacia los márgenes).
   - Nuestro clon soporta **3 vías de inserción de imágenes**:
     1. Botón **"Insertar Imagen"** desde el explorador de archivos.
     2. **Arrastrar y soltar (*Drag & Drop*)** imágenes directamente desde el Finder (Mac) o la app Archivos (Chromebook).
     3. **Pegar desde el portapapeles (`Cmd+V` en Mac / `Ctrl+V` en Chromebook)** tras hacer una captura de pantalla.
   - Las imágenes se pueden **mover libremente, redimensionar desde la esquina inferior derecha** y quedan situadas bajo la capa vectorial de tinta para poder **anotar encima con el boli táctil**.

---

## 2. Optimización Específica para Mac y Chromebook

| Característica | 💻 macOS (Safari / Chrome) | 💻 Chromebook (ChromeOS) |
| :--- | :--- | :--- |
| **Entrada de Boli / Puntero** | Trackpad Force Touch, Ratón, Tabletas Wacom o iPad vía **Sidecar / Continuity Sketch**. | Pantallas táctiles y lápices ópticos **USI (Universal Stylus Initiative)** con baja latencia. |
| **Atajos de Teclado** | Detección automática de tecla **`Cmd (⌘)`** (`⌘+Z` Deshacer, `⌘+Shift+Z` Rehacer, `⌘+S` Guardar en Drive, `⌘+V` Pegar imagen). | Detección automática de tecla **`Ctrl`** (`Ctrl+Z` Deshacer, `Ctrl+Y` Rehacer, `Ctrl+S` Guardar en Drive, `Ctrl+V` Pegar imagen). |
| **Resolución de Pantalla** | Escalado dinámico de Canvas con `window.devicePixelRatio` para pantallas **Retina (2x/3x)** ultra nítidas. | Prevención de gestos de navegación accidental en bordes mediante `touch-action: none` en el lienzo. |
| **Instalación (PWA)** | Instalable como Web App en el Dock de macOS. | Instalable como aplicación nativa desde ChromeOS en la estantería (*Shelf*). |

---

## 3. Integración con Credenciales de Google y Google Drive

Mientras que Microsoft OneNote utiliza cuentas MSA/Entra ID y sincroniza archivos binarios `.one` en OneDrive, nuestro clon web reemplaza esa capa por el ecosistema **Google Workspace**:

1. **Autenticación con Google Identity Services (GIS - OAuth 2.0)**:
   - Utiliza el cliente oficial `https://accounts.google.com/gsi/client` (`google.accounts.oauth2.initTokenClient`).
   - Solicita los permisos mínimos y seguros (`openid email profile` + `https://www.googleapis.com/auth/drive.file`), de modo que la aplicación solo accede a los archivos y carpetas creados por ella misma en Google Drive.
2. **Almacenamiento y Sincronización en Google Drive API v3**:
   - **Carpeta dedicada**: Crea o localiza automáticamente la carpeta **`OneNote Web Clone`** en el Google Drive del usuario.
   - **Archivo Maestro JSON Vectorial (`onenote_workspace.json`)**: Guarda todos los blocs de notas, secciones, páginas, trazos vectoriales del boli táctil, contenedores de texto e imágenes codificadas mediante peticiones `multipart/related` a `https://www.googleapis.com/upload/drive/v3/files`.
   - **Exportación Adicional**: Permite subir/exportar la página actual como **imagen PNG de alta resolución** (combinando fondo + texto + imágenes + trazos de boli) o descargar/importar el archivo `.onenote.json`.
3. **Arquitectura Dual (Modo Real OAuth + Modo Local/Demo Instantáneo)**:
   - Para que puedas **usarlo inmediatamente en tu Mac o Chromebook sin bloqueos**, la app guarda todos los cambios en tiempo real en **`IndexedDB` / `localStorage`** y ofrece un **Modo Cuenta Google Demo / Local** con un solo clic, además de un **Modal de Configuración de Google Cloud (`Client ID`)** donde puedes pegar tu propio ID de cliente OAuth 2.0 para sincronizar con tu Google Drive real en cualquier momento.
