# 📘 OneNote Web Clone (PWA) — Android, Mac & Chromebook

Clon web instalable como **PWA (Progressive Web App)** de **Microsoft OneNote**, diseñado para **Android**, **macOS** y **Chromebook**, con soporte de **boli táctil (Stylus / S-Pen / USI)**, contenedores de texto libres con teclado, imágenes redimensionables y **validación directa de credenciales de Google (sin SSO ni tokens OAuth)**.

---

## 🚀 Cómo ejecutarlo y probarlo en tu Android y Mac

En la terminal dentro de `/Users/pedropm/Desktop/cloneonenote`:

```bash
npm start
```

Esto inicia el servidor [server.py](file:///Users/pedropm/Desktop/cloneonenote/server.py) escuchando en toda tu red Wi-Fi local:
- 💻 **En tu Mac / Chromebook**: Abre **`http://localhost:8090`**
- 📱 **En tu móvil o tablet Android (en la misma red Wi-Fi)**: Abre **`http://192.168.86.204:8090`** (o pulsa el botón **📱 Android / PWA** en la barra superior para escanear el **Código QR** directamente con la cámara de tu Android).

### 📲 Cómo instalar la PWA en tu Android
1. Abre `http://192.168.86.204:8090` en **Google Chrome** desde tu Android.
2. Pulsa el botón superior **`📱 Android / PWA`** &rarr; **`📲 Instalar App Ahora`** (o abre el menú `⋮` de Chrome y selecciona **"Instalar aplicación"** / **"Añadir a pantalla de inicio"**).
3. La app se instalará con su icono morado oficial (`icon-192.png` / `icon-512.png`) y funcionará a pantalla completa y **sin conexión (offline)** gracias al Service Worker ([sw.js](file:///Users/pedropm/Desktop/cloneonenote/sw.js)).

---

## 🔑 Validación Directa de Credenciales de Google (Sin SSO y Sin Tokens OAuth)

1. Pulsa en **`Cuenta Google`** (arriba a la derecha).
2. Introduce directamente tu **Correo de Google (`@gmail.com`)** y tu **Contraseña** (con soporte de autocompletado nativo de Chrome/Android y `navigator.credentials` sin pedir `Client ID` ni abrir ventanas OAuth).
3. **Sincronización y Guardado en Google Drive sin OAuth**:
   - **☁️ Guardar en Mi Cuenta (Mac ↔ Android)**: Sincroniza tus blocs de notas, imágenes y trazos bajo tu correo `@gmail.com` entre tu Mac y tu Android a través del servidor local y la bóveda `IndexedDB`.
   - **📲 Enviar a App / Carpeta Google Drive**:
     - En **Android**, abre el menú nativo del sistema (`navigator.share`) para enviarlo directamente a la app oficial **Guardar en Drive** de tu teléfono Android.
     - En **Mac y Chromebook**, permite guardarlo directamente en tu carpeta de **Google Drive** del sistema (`showSaveFilePicker`).
