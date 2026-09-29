/**
 * app.js — Motor Principal del Clon Web de Microsoft OneNote
 * Soporta:
 * - Boli Táctil (Stylus USI en Chromebook / Trackpad, Wacom o Sidecar en Mac) con PointerEvents y getCoalescedEvents()
 * - Lienzo libre multimodal con Contenedores de Texto flotantes, Etiquetas interactivas e Imágenes redimensionables
 * - Selección de Lazo (Lasso), Borrador de Trazos vectorial, Subrayador multiplicativo, Tinta Arcoíris y Formas
 * - Autenticación con Credenciales de Google (OAuth 2.0) y Sincronización con Google Drive API v3
 */

import { GoogleDriveManager } from './drive.js?v=6';

const LOCAL_STORAGE_KEY = 'onenote_clone_workspace_v1';

// Colores característicos de las pestañas de sección de Microsoft OneNote
const SECTION_COLORS = ['#7719aa', '#2563eb', '#059669', '#d97706', '#dc2626', '#0891b2', '#db2777'];

class OneNoteCloneApp {
  constructor() {
    this.workspace = this.loadInitialWorkspace();
    this.currentTool = 'pen'; // 'text' | 'pen' | 'rainbow' | 'highlighter' | 'stroke-eraser' | 'pixel-eraser' | 'lasso' | 'pan' | 'line' | 'arrow' | 'rect' | 'ellipse'
    this.inkColor = '#1e293b';
    this.strokeSize = 3;
    this.usePressure = true;
    this.stylusOnly = false;
    this.zoomLevel = 1;
    this.isMobileSimpleMode = false;

    // Estado de dibujo y puntero
    this.isDrawing = false;
    this.activePointerId = null;
    this.currentStroke = null;
    this.shapeStartPoint = null;

    // Estado de Selección de Lazo (Lasso)
    this.lassoPolygon = [];
    this.selectedStrokeIds = new Set();
    this.lassoBoundingBox = null;
    this.isDraggingLassoSelection = false;
    this.lassoDragLastPos = null;

    // Estado de Paneo con mano o toque
    this.isPanning = false;
    this.panStartScroll = { left: 0, top: 0, clientX: 0, clientY: 0 };

    // Pila de Deshacer / Rehacer por página
    this.undoStacks = new Map();
    this.redoStacks = new Map();

    // Referencia al último contenedor de texto enfocado
    this.activeNoteEditable = null;

    // Cámara Web y Dictado por Voz
    this.cameraStream = null;
    this.speechRecognition = null;
    this.isDictating = false;

    // Inicializar Gestor de Google Drive
    this.driveManager = new GoogleDriveManager({
      onStatusChange: (user) => this.updateGoogleAuthUI(user),
      onToast: (msg) => this.showToast(msg)
    });

    this.cacheDOM();
    this.initCanvases();
    this.setupDialogLightDismissFallback();
    this.detectOperatingSystem();
    this.bindEvents();
    this.checkUrlSharedNoteOnStartup();
    this.renderAll();
    this.updateGoogleAuthUI(this.driveManager.getUser());
  }

  /* =========================================================
     DATOS INICIALES DE EJEMPLO (DEMO LISTA PARA USAR)
     ========================================================= */
  loadInitialWorkspace() {
    const saved = localStorage.getItem(LOCAL_STORAGE_KEY);
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        if (parsed && Array.isArray(parsed.notebooks) && parsed.notebooks.length > 0) {
          return parsed;
        }
      } catch (_) {
        // Usar plantilla por defecto si hay error
      }
    }

    // Crear imagen SVG de ejemplo en DataURL para mostrar anotación con boli sobre imagen
    const sampleDiagramSvg = `<svg xmlns="http://www.w3.org/2000/svg" width="420" height="210" viewBox="0 0 420 210">
      <rect width="420" height="210" rx="12" fill="#f8fafc" stroke="#cbd5e1" stroke-width="2"/>
      <text x="20" y="32" font-family="sans-serif" font-size="14" font-weight="bold" fill="#334155">Esquema de Arquitectura: OneNote Web + Google Drive</text>
      <rect x="24" y="58" width="105" height="68" rx="8" fill="#ede9fe" stroke="#7c3aed" stroke-width="2"/>
      <text x="40" y="90" font-family="sans-serif" font-size="12" font-weight="bold" fill="#5b21b6">Boli Táctil</text>
      <text x="36" y="108" font-family="sans-serif" font-size="10" fill="#6d28d9">PointerEvents</text>
      <path d="M132 92 L162 92" stroke="#64748b" stroke-width="2" marker-end="url(#arrow)"/>
      <rect x="165" y="58" width="105" height="68" rx="8" fill="#dbeafe" stroke="#2563eb" stroke-width="2"/>
      <text x="180" y="90" font-family="sans-serif" font-size="12" font-weight="bold" fill="#1e40af">Lienzo Libre</text>
      <text x="176" y="108" font-family="sans-serif" font-size="10" fill="#1d4ed8">Texto + Fotos</text>
      <path d="M273 92 L303 92" stroke="#64748b" stroke-width="2"/>
      <rect x="305" y="58" width="95" height="68" rx="8" fill="#dcfce7" stroke="#16a34a" stroke-width="2"/>
      <text x="316" y="90" font-family="sans-serif" font-size="12" font-weight="bold" fill="#166534">Google Drive</text>
      <text x="320" y="108" font-family="sans-serif" font-size="10" fill="#15803d">API v3 Sync</text>
      <text x="24" y="168" font-family="sans-serif" font-size="12" fill="#475569">💡 Puedes mover o redimensionar esta imagen y dibujar encima con el boli táctil.</text>
    </svg>`;
    const sampleDiagramDataUrl = 'data:image/svg+xml;utf8,' + encodeURIComponent(sampleDiagramSvg);

    return {
      activeNotebookId: 'nb-1',
      activeSectionId: 'sec-1',
      activePageId: 'page-1',
      notebooks: [
        {
          id: 'nb-1',
          name: '📚 Mi Bloc Principal',
          sections: [
            {
              id: 'sec-1',
              name: '🖊️ Notas y Boli Táctil',
              color: '#7719aa',
              pages: [
                {
                  id: 'page-1',
                  title: 'Bienvenido a OneNote Web (Mac & Chromebook)',
                  createdAt: new Date().toLocaleString('es-ES'),
                  isSubpage: false,
                  paperStyle: 'ruled-wide',
                  paperColor: '#ffffff',
                  blocks: [
                    {
                      id: 'blk-1',
                      type: 'text',
                      x: 88,
                      y: 125,
                      width: 440,
                      isSticky: false,
                      html: `<div><b>🚀 Guía rápida de uso en Mac y Chromebook:</b></div>
                      <div class="onenote-todo-row"><input type="checkbox" checked> <span><b>Boli Táctil (Stylus):</b> Dibuja directamente sobre el papel, texto o imágenes con sensibilidad a la presión.</span></div>
                      <div class="onenote-todo-row"><input type="checkbox"> <span><b>Teclado Libre:</b> Cambia a <i>Modo Texto (⌨️)</i> y haz clic en cualquier punto del lienzo para crear otra nota.</span></div>
                      <div class="onenote-todo-row"><input type="checkbox"> <span><b>Imágenes:</b> Sube fotos, arrástralas o pégalas con <b>⌘+V / Ctrl+V</b>.</span></div>
                      <div class="onenote-todo-row"><input type="checkbox"> <span><b>Google Drive:</b> Pulsa <i>Guardar en Drive</i> arriba a la derecha para sincronizar tus blocs.</span></div>`
                    },
                    {
                      id: 'blk-2',
                      type: 'text',
                      x: 565,
                      y: 125,
                      width: 290,
                      isSticky: true,
                      html: `<div><b>💡 Nota Adhesiva (Post-it)</b></div><div>En la pestaña <b>Dibujar</b> tienes:</div><div>• 🖍️ Subrayador translúcido</div><div>• 🪢 Selección de Lazo</div><div>• 🧽 Borrador de trazos enteros</div><div>• ▶️ Reproducción animada de tinta</div>`
                    },
                    {
                      id: 'blk-img-1',
                      type: 'image',
                      x: 88,
                      y: 355,
                      width: 430,
                      height: 215,
                      src: sampleDiagramDataUrl
                    }
                  ],
                  strokes: [
                    // Subrayado amarillo de ejemplo sobre la imagen/diagrama
                    {
                      id: 'strk-demo-1',
                      tool: 'highlighter',
                      color: '#facc15',
                      size: 20,
                      points: [
                        { x: 108, y: 382, p: 0.6 },
                        { x: 220, y: 382, p: 0.6 },
                        { x: 350, y: 382, p: 0.6 },
                        { x: 465, y: 382, p: 0.6 }
                      ]
                    },
                    // Flecha manuscrita roja indicando el bloque de Google Drive
                    {
                      id: 'strk-demo-2',
                      tool: 'pen',
                      color: '#dc2626',
                      size: 3.5,
                      points: [
                        { x: 545, y: 445, p: 0.5 },
                        { x: 590, y: 432, p: 0.65 },
                        { x: 635, y: 425, p: 0.7 },
                        { x: 675, y: 425, p: 0.6 }
                      ]
                    },
                    // Círculo azul alrededor de Google Drive en el esquema
                    {
                      id: 'strk-demo-3',
                      tool: 'ellipse',
                      color: '#1d4ed8',
                      size: 3,
                      points: [
                        { x: 382, y: 406, p: 0.5 },
                        { x: 495, y: 492, p: 0.5 }
                      ]
                    }
                  ]
                },
                {
                  id: 'page-2',
                  title: 'Cálculos y Esquemas (Cuadrícula)',
                  createdAt: new Date().toLocaleString('es-ES'),
                  isSubpage: true,
                  paperStyle: 'grid-small',
                  paperColor: '#ffffff',
                  blocks: [
                    {
                      id: 'blk-p2-1',
                      type: 'text',
                      x: 88,
                      y: 120,
                      width: 460,
                      isSticky: false,
                      html: `<div><b>📐 Página con Cuadrícula Matemática</b></div><div>Usa las herramientas de <b>Formas (╱ ↗️ ▭ ◯)</b> o el boli táctil para hacer diagramas limpios. Prueba el botón <b>▶️ Reproducir</b> en la pestaña Dibujar.</div>`
                    }
                  ],
                  strokes: []
                }
              ]
            },
            {
              id: 'sec-2',
              name: '☁️ Sincronización Google',
              color: '#059669',
              pages: [
                {
                  id: 'page-3',
                  title: 'Cómo funciona Google Drive en este Clon',
                  createdAt: new Date().toLocaleString('es-ES'),
                  isSubpage: false,
                  paperStyle: 'ruled-narrow',
                  paperColor: '#fef9c3',
                  blocks: [
                    {
                      id: 'blk-p3-1',
                      type: 'text',
                      x: 88,
                      y: 125,
                      width: 560,
                      isSticky: false,
                      html: `<div><b>⭐ Integración con Credenciales de Google y Google Drive:</b></div>
                      <div>1. Pulsa en <b>Conectar Google</b> arriba a la derecha.</div>
                      <div>2. Puedes usar el <b>⚡ Acceso Rápido (Modo Demo / Local Drive)</b> sin configurar nada, o introducir tu propio <code>Client ID</code> de Google Cloud Console para sincronizar con la API oficial de Google Drive v3.</div>
                      <div>3. También puedes exportar cualquier página como <b>Imagen PNG</b> o archivo <b>.onenote.json</b> en la pestaña <i>Vista y Papel</i>.</div>`
                    }
                  ],
                  strokes: []
                }
              ]
            }
          ]
        }
      ]
    };
  }

  /* =========================================================
     REFERENCIAS AL DOM Y DETECCIÓN DE SISTEMA OPERATIVO
     ========================================================= */
  cacheDOM() {
    this.sidebarContainer = document.getElementById('sidebarContainer');
    this.toggleSidebarBtn = document.getElementById('toggleSidebarBtn');
    this.osBadge = document.getElementById('osBadge');
    this.globalSearchInput = document.getElementById('globalSearchInput');
    this.searchResultsDropdown = document.getElementById('searchResultsDropdown');

    this.pointerIcon = document.getElementById('pointerIcon');
    this.pointerTypeText = document.getElementById('pointerTypeText');
    this.pressureBadge = document.getElementById('pressureBadge');

    this.syncDot = document.getElementById('syncDot');
    this.syncText = document.getElementById('syncText');
    this.toggleSimpleLayoutBtn = document.getElementById('toggleSimpleLayoutBtn');
    this.openGoogleDocsModalBtn = document.getElementById('openGoogleDocsModalBtn');
    this.openShareModalBtn = document.getElementById('openShareModalBtn');
    this.saveToDriveBtn = document.getElementById('saveToDriveBtn');
    this.openDriveModalBtn = document.getElementById('openDriveModalBtn');
    this.openAndroidPwaModalBtn = document.getElementById('openAndroidPwaModalBtn');
    this.topbarAvatarBadge = document.getElementById('topbarAvatarBadge');
    this.googleGIcon = document.getElementById('googleGIcon');
    this.googleUserLabel = document.getElementById('googleUserLabel');
    this.themeToggleBtn = document.getElementById('themeToggleBtn');

    // Controles de Barra Lateral
    this.notebookSelect = document.getElementById('notebookSelect');
    this.newNotebookBtn = document.getElementById('newNotebookBtn');
    this.newSectionBtn = document.getElementById('newSectionBtn');
    this.sectionsList = document.getElementById('sectionsList');
    this.newPageBtn = document.getElementById('newPageBtn');
    this.pagesList = document.getElementById('pagesList');

    // Lienzo y Capas
    this.canvasViewport = document.getElementById('canvasViewport');
    this.pageSurface = document.getElementById('pageSurface');
    this.pageTitleInput = document.getElementById('pageTitleInput');
    this.pageDateLabel = document.getElementById('pageDateLabel');
    this.blocksLayer = document.getElementById('blocksLayer');
    this.noteBlocksLayer = this.blocksLayer;
    this.inkCanvas = document.getElementById('inkCanvas');
    this.overlayCanvas = document.getElementById('overlayCanvas');

    // Controles de Cinta (Ribbon)
    this.strokeSizeRange = document.getElementById('strokeSizeRange');
    this.strokeSizeValue = document.getElementById('strokeSizeValue');
    this.customInkColorPicker = document.getElementById('customInkColorPicker');
    this.pressureSensitivityToggle = document.getElementById('pressureSensitivityToggle');
    this.stylusOnlyToggle = document.getElementById('stylusOnlyToggle');
    this.paperLinesSelect = document.getElementById('paperLinesSelect');

    this.lassoSelectionControls = document.getElementById('lassoSelectionControls');
    this.lassoCountBadge = document.getElementById('lassoCountBadge');
    this.lassoRecolorBtn = document.getElementById('lassoRecolorBtn');
    this.lassoDeleteBtn = document.getElementById('lassoDeleteBtn');

    // Modales
    this.googleDocsDialog = document.getElementById('googleDocsDialog');
    this.googleDriveDialog = document.getElementById('googleDriveDialog');
    this.androidPwaDialog = document.getElementById('androidPwaDialog');
    this.shareLinkDialog = document.getElementById('shareLinkDialog');
    this.cameraDialog = document.getElementById('cameraDialog');
    this.cameraVideo = document.getElementById('cameraVideo');
    this.toastContainer = document.getElementById('toastContainer');
  }

  detectOperatingSystem() {
    const ua = navigator.userAgent || '';
    const isAndroid = /Android/i.test(ua);
    const isMobileViewport = window.innerWidth <= 900 || isAndroid || /iPhone|iPad/i.test(ua);

    if (isAndroid) {
      this.osBadge.textContent = '📱 Android PWA · Texto + Boli + Google Docs';
    } else if (/CrOS/i.test(ua)) {
      this.osBadge.textContent = '💻 Chromebook · Stylus USI + Google Docs';
    } else if (/Macintosh|Mac OS X/i.test(ua)) {
      this.osBadge.textContent = '🍎 macOS · Boli Táctil + Google Docs';
    } else {
      this.osBadge.textContent = '💻 Web PWA · Android, Mac & Chromebook';
    }

    // En móvil / Android activar automáticamente la versión simple (texto ocupa casi todo el espacio)
    // En escritorio (Mac / Chromebook) mostrar la interfaz completa de OneNote por defecto
    if (isMobileViewport) {
      this.sidebarContainer.classList.add('collapsed');
      this.setSimpleMobileMode(true, false);
      this.currentTool = 'text';
    } else {
      this.setSimpleMobileMode(false, false);
    }

    this.registerServiceWorkerAndPwaPrompt();
  }

  setSimpleMobileMode(enabled, notify = true) {
    this.isMobileSimpleMode = Boolean(enabled);
    document.body.classList.toggle('mobile-simple-mode', this.isMobileSimpleMode);
    const labelEl = document.getElementById('simpleLayoutLabel');
    const iconEl = document.getElementById('simpleLayoutIcon');
    if (labelEl) {
      labelEl.textContent = this.isMobileSimpleMode ? 'Vista Completa' : 'Vista Simple Móvil';
    }
    if (iconEl) {
      iconEl.textContent = this.isMobileSimpleMode ? '🖥️' : '📱';
    }
    if (this.isMobileSimpleMode) {
      this.setTool('text');
    }
    if (notify) {
      this.showToast(
        this.isMobileSimpleMode
          ? '📱 Vista Simple Móvil activada: El texto ocupa casi toda la pantalla.'
          : '🖥️ Vista Completa de Escritorio activada.'
      );
    }
  }

  registerServiceWorkerAndPwaPrompt() {
    this.deferredInstallPrompt = null;

    if ('caches' in window) {
      caches.keys().then((keys) => {
        keys.forEach((key) => {
          if (key !== 'onenote-pwa-cache-v6') {
            caches.delete(key);
          }
        });
      });
    }

    if ('serviceWorker' in navigator) {
      navigator.serviceWorker
        .register('./sw.js?v=6')
        .then((reg) => {
          reg.update().catch(() => {});
        })
        .catch(() => {});
    }

    window.addEventListener('beforeinstallprompt', (e) => {
      e.preventDefault();
      this.deferredInstallPrompt = e;
      const pwaLabel = document.getElementById('pwaBtnLabel');
      if (pwaLabel) pwaLabel.textContent = 'Instalar en Android';
    });

    window.addEventListener('appinstalled', () => {
      this.deferredInstallPrompt = null;
      this.showToast('🎉 ¡Aplicación PWA instalada correctamente en tu dispositivo!');
    });
  }

  /**
   * Cumplimiento estricto de modern-web-guidance (light-dismiss-a-dialog):
   * Fallback obligatorio para navegadores sin soporte nativo de <dialog closedby="any"> (como Safari en Mac).
   */
  setupDialogLightDismissFallback() {
    const dialogs = document.querySelectorAll('dialog[closedby="any"]');
    dialogs.forEach((dialog) => {
      if (!('closedBy' in HTMLDialogElement.prototype)) {
        dialog.addEventListener('click', (event) => {
          if (event.target !== dialog) return;
          const rect = dialog.getBoundingClientRect();
          const isDialogContent =
            rect.top <= event.clientY &&
            event.clientY <= rect.top + rect.height &&
            rect.left <= event.clientX &&
            event.clientX <= rect.left + rect.width;
          if (isDialogContent) return;
          dialog.close();
        });
      }
    });
  }

  /* =========================================================
     CONFIGURACIÓN DE CANVAS DE ALTA DENSIDAD (RETINA / CHROMEBOOK)
     ========================================================= */
  initCanvases() {
    this.dpr = Math.min(window.devicePixelRatio || 1, 2);
    const width = 1600;
    const height = 2000;

    [this.inkCanvas, this.overlayCanvas].forEach((canvas) => {
      canvas.width = width * this.dpr;
      canvas.height = height * this.dpr;
      canvas.style.width = `${width}px`;
      canvas.style.height = `${height}px`;
      const ctx = canvas.getContext('2d');
      ctx.scale(this.dpr, this.dpr);
    });

    this.inkCtx = this.inkCanvas.getContext('2d');
    this.overlayCtx = this.overlayCanvas.getContext('2d');
  }

  /* =========================================================
     HELPERS DE JERARQUÍA (BLOC -> SECCIÓN -> PÁGINA)
     ========================================================= */
  getActiveNotebook() {
    return (
      this.workspace.notebooks.find((n) => n.id === this.workspace.activeNotebookId) ||
      this.workspace.notebooks[0]
    );
  }

  getActiveSection() {
    const nb = this.getActiveNotebook();
    return nb.sections.find((s) => s.id === this.workspace.activeSectionId) || nb.sections[0];
  }

  getActivePage() {
    const sec = this.getActiveSection();
    return sec.pages.find((p) => p.id === this.workspace.activePageId) || sec.pages[0];
  }

  saveLocalState() {
    this.syncDot.classList.add('saving');
    this.syncText.textContent = 'Guardando...';
    try {
      localStorage.setItem(LOCAL_STORAGE_KEY, JSON.stringify(this.workspace));
    } catch (_) {}
    this.driveManager.saveLocalPwaWorkspace(this.workspace);

    clearTimeout(this._saveTimer);
    this._saveTimer = setTimeout(() => {
      this.syncDot.classList.remove('saving');
      this.syncText.textContent = this.driveManager.isSignedIn()
        ? 'Local + Google Drive'
        : 'Guardado en PWA local';
    }, 350);

    // Auto-almacenar en Google Drive cada vez que se edita la nota
    clearTimeout(this._driveAutoSyncTimer);
    this._driveAutoSyncTimer = setTimeout(() => {
      this.driveManager.saveWorkspaceToDrive(this.workspace, { silent: true });
    }, 1500);
  }

  pushUndoSnapshot() {
    const page = this.getActivePage();
    if (!page) return;
    if (!this.undoStacks.has(page.id)) this.undoStacks.set(page.id, []);
    const stack = this.undoStacks.get(page.id);
    stack.push(JSON.stringify({ strokes: page.strokes, blocks: page.blocks }));
    if (stack.length > 35) stack.shift();
    this.redoStacks.set(page.id, []);
  }

  undo() {
    const page = this.getActivePage();
    const stack = this.undoStacks.get(page.id) || [];
    if (stack.length === 0) {
      this.showToast('No hay más acciones para deshacer.');
      return;
    }
    if (!this.redoStacks.has(page.id)) this.redoStacks.set(page.id, []);
    this.redoStacks.get(page.id).push(JSON.stringify({ strokes: page.strokes, blocks: page.blocks }));

    const prev = JSON.parse(stack.pop());
    page.strokes = prev.strokes;
    page.blocks = prev.blocks;
    this.clearLassoSelection();
    this.renderActivePage();
    this.saveLocalState();
  }

  redo() {
    const page = this.getActivePage();
    const stack = this.redoStacks.get(page.id) || [];
    if (stack.length === 0) {
      this.showToast('No hay acciones para rehacer.');
      return;
    }
    this.undoStacks.get(page.id).push(JSON.stringify({ strokes: page.strokes, blocks: page.blocks }));

    const next = JSON.parse(stack.pop());
    page.strokes = next.strokes;
    page.blocks = next.blocks;
    this.clearLassoSelection();
    this.renderActivePage();
    this.saveLocalState();
  }

  /* =========================================================
     RENDERIZADO DE LA INTERFAZ Y CAPAS DE PÁGINA
     ========================================================= */
  renderAll() {
    this.renderSidebar();
    this.renderActivePage();
    this.updateToolUI();
  }

  renderSidebar() {
    const nb = this.getActiveNotebook();
    const sec = this.getActiveSection();

    // Selector de Blocs de Notas
    this.notebookSelect.innerHTML = '';
    this.workspace.notebooks.forEach((notebook) => {
      const opt = document.createElement('option');
      opt.value = notebook.id;
      opt.textContent = notebook.name;
      if (notebook.id === nb.id) opt.selected = true;
      this.notebookSelect.appendChild(opt);
    });

    // Lista de Secciones
    this.sectionsList.innerHTML = '';
    nb.sections.forEach((section) => {
      const li = document.createElement('li');
      li.className = `section-item ${section.id === sec.id ? 'active' : ''}`;
      li.style.borderLeftColor = section.color || '#7719aa';

      const titleSpan = document.createElement('span');
      titleSpan.className = 'item-title';
      titleSpan.textContent = section.name;
      li.appendChild(titleSpan);

      const actionsDiv = document.createElement('div');
      actionsDiv.className = 'item-actions';

      const renameBtn = document.createElement('button');
      renameBtn.className = 'item-action-btn';
      renameBtn.title = 'Renombrar sección';
      renameBtn.textContent = '✏️';
      renameBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        const newName = prompt('Nuevo nombre de la sección:', section.name);
        if (newName && newName.trim()) {
          section.name = newName.trim();
          this.renderSidebar();
          this.saveLocalState();
        }
      });

      const delBtn = document.createElement('button');
      delBtn.className = 'item-action-btn';
      delBtn.title = 'Eliminar sección';
      delBtn.textContent = '🗑️';
      delBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        if (nb.sections.length <= 1) {
          this.showToast('El bloc debe tener al menos una sección.');
          return;
        }
        nb.sections = nb.sections.filter((s) => s.id !== section.id);
        this.workspace.activeSectionId = nb.sections[0].id;
        this.workspace.activePageId = nb.sections[0].pages[0].id;
        this.renderAll();
        this.saveLocalState();
      });

      actionsDiv.append(renameBtn, delBtn);
      li.appendChild(actionsDiv);

      li.addEventListener('click', () => {
        this.workspace.activeSectionId = section.id;
        this.workspace.activePageId = section.pages[0].id;
        this.clearLassoSelection();
        this.renderAll();
        this.saveLocalState();
      });

      this.sectionsList.appendChild(li);
    });

    // Lista de Páginas
    this.pagesList.innerHTML = '';
    const activePage = this.getActivePage();
    sec.pages.forEach((page) => {
      const li = document.createElement('li');
      li.className = `page-item ${page.id === activePage.id ? 'active' : ''} ${
        page.isSubpage ? 'subpage' : ''
      }`;

      const titleSpan = document.createElement('span');
      titleSpan.className = 'item-title';
      titleSpan.textContent = page.title || 'Página sin título';
      li.appendChild(titleSpan);

      const actionsDiv = document.createElement('div');
      actionsDiv.className = 'item-actions';

      const subpageBtn = document.createElement('button');
      subpageBtn.className = 'item-action-btn';
      subpageBtn.title = page.isSubpage ? 'Promover a página principal' : 'Convertir en subpágina';
      subpageBtn.textContent = page.isSubpage ? '⬅️' : '➡️';
      subpageBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        page.isSubpage = !page.isSubpage;
        this.renderSidebar();
        this.saveLocalState();
      });

      const delBtn = document.createElement('button');
      delBtn.className = 'item-action-btn';
      delBtn.title = 'Eliminar página';
      delBtn.textContent = '🗑️';
      delBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        if (sec.pages.length <= 1) {
          this.showToast('La sección debe tener al menos una página.');
          return;
        }
        sec.pages = sec.pages.filter((p) => p.id !== page.id);
        this.workspace.activePageId = sec.pages[0].id;
        this.renderAll();
        this.saveLocalState();
      });

      actionsDiv.append(subpageBtn, delBtn);
      li.appendChild(actionsDiv);

      li.addEventListener('click', () => {
        this.workspace.activePageId = page.id;
        this.clearLassoSelection();
        this.renderAll();
        this.saveLocalState();
        if (this.isMobileSimpleMode || window.innerWidth <= 900) {
          this.sidebarContainer?.classList.add('collapsed');
        }
      });

      this.pagesList.appendChild(li);
    });
  }

  renderActivePage() {
    const page = this.getActivePage();
    if (!page) return;

    if (!Array.isArray(page.blocks)) page.blocks = [];
    if (!page.blocks.some((b) => b.type === 'text')) {
      page.blocks.unshift({
        id: 'blk-main-' + page.id,
        type: 'text',
        x: 88,
        y: 125,
        width: 680,
        isSticky: false,
        html: ''
      });
    }

    this.pageTitleInput.value = page.title || '';
    this.pageDateLabel.textContent = `📅 ${page.createdAt || ''}`;

    // Estilo de papel y color
    this.pageSurface.className = `page-surface paper-${page.paperStyle || 'ruled-wide'} ${
      this.currentTool === 'text' ? 'mode-text' : this.currentTool === 'pan' ? 'mode-pan' : 'mode-draw'
    }`;
    this.pageSurface.style.setProperty('--paper-bg', page.paperColor || '#ffffff');
    this.paperLinesSelect.value = page.paperStyle || 'ruled-wide';

    document.querySelectorAll('.paper-color-btn').forEach((btn) => {
      btn.classList.toggle('active', btn.dataset.bg === (page.paperColor || '#ffffff'));
    });

    this.renderBlocksLayer(page);
    this.redrawInkCanvas();
    this.redrawOverlayCanvas();
  }

  /* =========================================================
     CONTENEDORES DE TEXTO E IMÁGENES EN EL LIENZO LIBRE
     ========================================================= */
  renderBlocksLayer(page) {
    this.blocksLayer.innerHTML = '';

    (page.blocks || []).forEach((block) => {
      if (block.type === 'text') {
        const container = document.createElement('div');
        container.className = `note-container ${block.isSticky ? 'sticky-note' : ''}`;
        container.style.left = `${block.x}px`;
        container.style.top = `${block.y}px`;
        container.style.width = `${block.width || 360}px`;
        container.dataset.id = block.id;

        const handleBar = document.createElement('div');
        handleBar.className = 'container-handle-bar';
        handleBar.innerHTML = `<span class="grip-dots">⋮⋮ Mover nota</span>`;

        const delBtn = document.createElement('button');
        delBtn.className = 'container-delete-btn';
        delBtn.title = 'Eliminar este contenedor de nota';
        delBtn.textContent = '✕';
        delBtn.addEventListener('click', (e) => {
          e.stopPropagation();
          this.pushUndoSnapshot();
          page.blocks = page.blocks.filter((b) => b.id !== block.id);
          this.renderBlocksLayer(page);
          this.saveLocalState();
        });
        handleBar.appendChild(delBtn);

        const editable = document.createElement('div');
        editable.className = 'note-editable';
        editable.contentEditable = 'true';
        editable.dataset.placeholder = 'Escribe tu nota o pega una imagen aquí...';
        editable.innerHTML = block.html || '';

        editable.addEventListener('focus', () => {
          this.activeNoteEditable = editable;
        });

        editable.addEventListener('input', () => {
          block.html = editable.innerHTML;
          this.saveLocalState();
        });

        // Soporte para casillas interactivas dentro del contenedor
        editable.addEventListener('change', (e) => {
          if (e.target && e.target.type === 'checkbox') {
            if (e.target.checked) {
              e.target.setAttribute('checked', 'checked');
            } else {
              e.target.removeAttribute('checked');
            }
            block.html = editable.innerHTML;
            this.saveLocalState();
          }
        });

        this.makeElementDraggable(container, handleBar, (newX, newY) => {
          block.x = newX;
          block.y = newY;
          this.saveLocalState();
        });

        container.append(handleBar, editable);
        this.blocksLayer.appendChild(container);
      } else if (block.type === 'image') {
        const imgWrap = document.createElement('div');
        imgWrap.className = 'image-container';
        imgWrap.style.left = `${block.x}px`;
        imgWrap.style.top = `${block.y}px`;
        imgWrap.style.width = `${block.width || 320}px`;
        imgWrap.style.height = `${block.height || 220}px`;
        imgWrap.dataset.id = block.id;

        const handleBar = document.createElement('div');
        handleBar.className = 'container-handle-bar';
        handleBar.innerHTML = `<span class="grip-dots">⋮⋮ Mover imagen</span>`;

        const delBtn = document.createElement('button');
        delBtn.className = 'container-delete-btn';
        delBtn.title = 'Eliminar imagen';
        delBtn.textContent = '✕';
        delBtn.addEventListener('click', (e) => {
          e.stopPropagation();
          this.pushUndoSnapshot();
          page.blocks = page.blocks.filter((b) => b.id !== block.id);
          this.renderBlocksLayer(page);
          this.saveLocalState();
        });
        handleBar.appendChild(delBtn);

        const img = document.createElement('img');
        img.src = block.src;
        img.alt = 'Imagen insertada en OneNote';

        const resizeHandle = document.createElement('div');
        resizeHandle.className = 'resize-handle';
        resizeHandle.title = 'Arrastrar para cambiar tamaño de la imagen';

        this.makeElementDraggable(imgWrap, handleBar, (newX, newY) => {
          block.x = newX;
          block.y = newY;
          this.saveLocalState();
        });

        this.makeElementResizable(imgWrap, resizeHandle, (newW, newH) => {
          block.width = newW;
          block.height = newH;
          this.saveLocalState();
        });

        imgWrap.append(handleBar, img, resizeHandle);
        this.blocksLayer.appendChild(imgWrap);
      }
    });
  }

  makeElementDraggable(element, handle, onMoved) {
    let dragging = false;
    let startX = 0;
    let startY = 0;
    let origLeft = 0;
    let origTop = 0;

    handle.addEventListener('pointerdown', (e) => {
      if (e.target.closest('button')) return;
      e.stopPropagation();
      dragging = true;
      handle.setPointerCapture(e.pointerId);
      startX = e.clientX;
      startY = e.clientY;
      origLeft = parseFloat(element.style.left) || 0;
      origTop = parseFloat(element.style.top) || 0;
    });

    handle.addEventListener('pointermove', (e) => {
      if (!dragging) return;
      const dx = (e.clientX - startX) / this.zoomLevel;
      const dy = (e.clientY - startY) / this.zoomLevel;
      const nextX = Math.max(16, Math.min(1420, Math.round(origLeft + dx)));
      const nextY = Math.max(95, Math.min(1880, Math.round(origTop + dy)));
      element.style.left = `${nextX}px`;
      element.style.top = `${nextY}px`;
    });

    handle.addEventListener('pointerup', (e) => {
      if (!dragging) return;
      dragging = false;
      handle.releasePointerCapture(e.pointerId);
      onMoved(parseFloat(element.style.left), parseFloat(element.style.top));
    });
  }

  makeElementResizable(element, handle, onResized) {
    let resizing = false;
    let startX = 0;
    let startY = 0;
    let origW = 0;
    let origH = 0;

    handle.addEventListener('pointerdown', (e) => {
      e.stopPropagation();
      resizing = true;
      handle.setPointerCapture(e.pointerId);
      startX = e.clientX;
      startY = e.clientY;
      origW = parseFloat(element.style.width) || 240;
      origH = parseFloat(element.style.height) || 180;
    });

    handle.addEventListener('pointermove', (e) => {
      if (!resizing) return;
      const dx = (e.clientX - startX) / this.zoomLevel;
      const dy = (e.clientY - startY) / this.zoomLevel;
      const nextW = Math.max(90, Math.round(origW + dx));
      const nextH = Math.max(70, Math.round(origH + dy));
      element.style.width = `${nextW}px`;
      element.style.height = `${nextH}px`;
    });

    handle.addEventListener('pointerup', (e) => {
      if (!resizing) return;
      resizing = false;
      handle.releasePointerCapture(e.pointerId);
      onResized(parseFloat(element.style.width), parseFloat(element.style.height));
    });
  }

  createTextContainerAt(x, y, initialHtml = '', isSticky = false) {
    const page = this.getActivePage();
    this.pushUndoSnapshot();
    const newBlock = {
      id: 'blk-' + Date.now() + '-' + Math.random().toString(36).slice(2, 6),
      type: 'text',
      x: Math.max(40, Math.min(1200, Math.round(x))),
      y: Math.max(115, Math.min(1850, Math.round(y))),
      width: isSticky ? 280 : 380,
      isSticky,
      html: initialHtml
    };
    page.blocks.push(newBlock);
    this.setTool('text');
    this.renderBlocksLayer(page);
    this.saveLocalState();

    // Enfocar inmediatamente para escribir con el teclado
    const createdEl = this.blocksLayer.querySelector(`[data-id="${newBlock.id}"] .note-editable`);
    if (createdEl) {
      createdEl.focus();
      this.activeNoteEditable = createdEl;
    }
    return newBlock;
  }

  insertImageDataUrl(dataUrl, x = 140, y = 200) {
    const img = new Image();
    img.onload = () => {
      const maxWidth = 440;
      const ratio = img.width > maxWidth ? maxWidth / img.width : 1;
      const width = Math.round(Math.max(140, img.width * ratio));
      const height = Math.round(Math.max(100, img.height * ratio));

      const page = this.getActivePage();
      this.pushUndoSnapshot();
      page.blocks.push({
        id: 'img-' + Date.now() + '-' + Math.random().toString(36).slice(2, 6),
        type: 'image',
        x,
        y,
        width,
        height,
        src: dataUrl
      });
      this.renderBlocksLayer(page);
      this.saveLocalState();
      this.showToast('🖼️ Imagen insertada. Puedes moverla, redimensionarla o escribir encima con el boli táctil.');
    };
    img.src = dataUrl;
  }

  /* =========================================================
     MOTOR DE TINTA DIGITAL (BOLI TÁCTIL / STYLUS / FORMAS)
     ========================================================= */
  getCanvasCoords(e) {
    const rect = this.inkCanvas.getBoundingClientRect();
    return {
      x: (e.clientX - rect.left) / this.zoomLevel,
      y: (e.clientY - rect.top) / this.zoomLevel,
      p: e.pressure && e.pressure > 0 ? e.pressure : 0.5
    };
  }

  updatePointerStatusPill(e) {
    const pType = e.pointerType || 'mouse';
    const pressurePct = Math.round((e.pressure || 0.5) * 100);
    if (pType === 'pen') {
      this.pointerIcon.textContent = '🖊️';
      this.pointerTypeText.textContent = 'Boli Táctil Activo';
      this.pressureBadge.textContent = `Presión: ${pressurePct}%`;
    } else if (pType === 'touch') {
      this.pointerIcon.textContent = '👆';
      this.pointerTypeText.textContent = this.stylusOnly ? 'Dedo (Desplazando)' : 'Entrada Táctil';
      this.pressureBadge.textContent = `Presión: ${pressurePct}%`;
    } else {
      this.pointerIcon.textContent = '🖱️';
      this.pointerTypeText.textContent = 'Trackpad / Ratón';
      this.pressureBadge.textContent = `Presión: ${pressurePct}%`;
    }
  }

  handlePointerDown(e) {
    this.updatePointerStatusPill(e);

    // Rechazo de palma (Palm Rejection): si "Solo Boli" está activo y el usuario toca con el dedo, desplazar el lienzo
    if (this.stylusOnly && e.pointerType === 'touch') {
      this.startPanning(e);
      return;
    }

    if (this.currentTool === 'pan') {
      this.startPanning(e);
      return;
    }

    // Selección de Lazo: comprobar si se está arrastrando la caja seleccionada
    const pt = this.getCanvasCoords(e);
    if (this.currentTool === 'lasso') {
      if (this.lassoBoundingBox && this.isPointInBox(pt, this.lassoBoundingBox)) {
        this.pushUndoSnapshot();
        this.isDraggingLassoSelection = true;
        this.lassoDragLastPos = pt;
        this.inkCanvas.setPointerCapture(e.pointerId);
        return;
      }
      this.clearLassoSelection();
      this.isDrawing = true;
      this.activePointerId = e.pointerId;
      this.lassoPolygon = [pt];
      this.inkCanvas.setPointerCapture(e.pointerId);
      return;
    }

    // Borradores
    if (this.currentTool === 'stroke-eraser' || this.currentTool === 'pixel-eraser') {
      this.pushUndoSnapshot();
      this.isDrawing = true;
      this.activePointerId = e.pointerId;
      this.inkCanvas.setPointerCapture(e.pointerId);
      this.eraseAtPoint(pt);
      return;
    }

    // Formas geométricas (Ink to Shape)
    if (['line', 'arrow', 'rect', 'ellipse'].includes(this.currentTool)) {
      this.pushUndoSnapshot();
      this.isDrawing = true;
      this.activePointerId = e.pointerId;
      this.shapeStartPoint = pt;
      this.inkCanvas.setPointerCapture(e.pointerId);
      return;
    }

    // Bolígrafo / Subrayador / Tinta Arcoíris
    if (['pen', 'rainbow', 'highlighter'].includes(this.currentTool)) {
      this.pushUndoSnapshot();
      this.isDrawing = true;
      this.activePointerId = e.pointerId;
      this.inkCanvas.setPointerCapture(e.pointerId);

      this.currentStroke = {
        id: 'strk-' + Date.now() + '-' + Math.random().toString(36).slice(2, 6),
        tool: this.currentTool,
        color: this.currentTool === 'rainbow' ? 'rainbow' : this.inkColor,
        size: this.strokeSize,
        usePressure: this.usePressure && e.pointerType === 'pen',
        points: [pt]
      };

      const page = this.getActivePage();
      page.strokes.push(this.currentStroke);
      this.redrawInkCanvas();
    }
  }

  handlePointerMove(e) {
    this.updatePointerStatusPill(e);

    if (this.isPanning) {
      this.movePanning(e);
      return;
    }

    if (this.isDraggingLassoSelection && this.lassoDragLastPos) {
      const pt = this.getCanvasCoords(e);
      const dx = pt.x - this.lassoDragLastPos.x;
      const dy = pt.y - this.lassoDragLastPos.y;
      this.moveSelectedStrokes(dx, dy);
      this.lassoDragLastPos = pt;
      return;
    }

    if (!this.isDrawing || e.pointerId !== this.activePointerId) return;

    // Captura de alta frecuencia (120Hz en Chromebooks / Mac) mediante getCoalescedEvents()
    const events = typeof e.getCoalescedEvents === 'function' ? e.getCoalescedEvents() : [e];
    const samples = events.length > 0 ? events : [e];

    if (this.currentTool === 'lasso') {
      samples.forEach((ev) => {
        this.lassoPolygon.push(this.getCanvasCoords(ev));
      });
      this.redrawOverlayCanvas();
      return;
    }

    if (this.currentTool === 'stroke-eraser' || this.currentTool === 'pixel-eraser') {
      samples.forEach((ev) => {
        this.eraseAtPoint(this.getCanvasCoords(ev));
      });
      return;
    }

    if (['line', 'arrow', 'rect', 'ellipse'].includes(this.currentTool)) {
      const currentPt = this.getCanvasCoords(e);
      this.redrawOverlayCanvas(currentPt);
      return;
    }

    if (this.currentStroke) {
      samples.forEach((ev) => {
        const pt = this.getCanvasCoords(ev);
        const last = this.currentStroke.points[this.currentStroke.points.length - 1];
        if (!last || Math.hypot(pt.x - last.x, pt.y - last.y) > 1.2) {
          this.currentStroke.points.push(pt);
        }
      });
      this.redrawInkCanvas();
    }
  }

  handlePointerUp(e) {
    if (this.isPanning) {
      this.stopPanning();
      return;
    }

    if (this.isDraggingLassoSelection) {
      this.isDraggingLassoSelection = false;
      this.lassoDragLastPos = null;
      this.saveLocalState();
      return;
    }

    if (!this.isDrawing || e.pointerId !== this.activePointerId) return;
    this.isDrawing = false;
    this.activePointerId = null;

    if (this.currentTool === 'lasso') {
      this.computeLassoSelection();
      this.lassoPolygon = [];
      this.redrawOverlayCanvas();
      return;
    }

    if (['line', 'arrow', 'rect', 'ellipse'].includes(this.currentTool) && this.shapeStartPoint) {
      const endPt = this.getCanvasCoords(e);
      const page = this.getActivePage();
      page.strokes.push({
        id: 'shape-' + Date.now(),
        tool: this.currentTool,
        color: this.inkColor,
        size: this.strokeSize,
        points: [this.shapeStartPoint, endPt]
      });
      this.shapeStartPoint = null;
      this.redrawOverlayCanvas();
      this.redrawInkCanvas();
      this.saveLocalState();
      return;
    }

    this.currentStroke = null;
    this.saveLocalState();
  }

  startPanning(e) {
    this.isPanning = true;
    this.panStartScroll = {
      left: this.canvasViewport.scrollLeft,
      top: this.canvasViewport.scrollTop,
      clientX: e.clientX,
      clientY: e.clientY
    };
  }

  movePanning(e) {
    const dx = e.clientX - this.panStartScroll.clientX;
    const dy = e.clientY - this.panStartScroll.clientY;
    this.canvasViewport.scrollLeft = this.panStartScroll.left - dx;
    this.canvasViewport.scrollTop = this.panStartScroll.top - dy;
  }

  stopPanning() {
    this.isPanning = false;
  }

  /* =========================================================
     RENDERIZADO DE TRAZOS VECTORIALES Y FORMAS EN CANVAS
     ========================================================= */
  redrawInkCanvas() {
    const ctx = this.inkCtx;
    if (!ctx) return;
    ctx.clearRect(0, 0, 1600, 2000);
    const page = this.getActivePage();
    if (!page || !Array.isArray(page.strokes)) return;

    page.strokes.forEach((stroke) => {
      this.drawSingleStroke(ctx, stroke);
    });
  }

  drawSingleStroke(ctx, stroke) {
    const pts = stroke.points || [];
    if (pts.length === 0) return;

    ctx.save();
    ctx.lineCap = stroke.tool === 'highlighter' ? 'butt' : 'round';
    ctx.lineJoin = 'round';

    if (stroke.tool === 'highlighter') {
      ctx.globalAlpha = 0.36;
      ctx.globalCompositeOperation = document.body.classList.contains('dark-theme')
        ? 'screen'
        : 'multiply';
    } else {
      ctx.globalAlpha = 1.0;
      ctx.globalCompositeOperation = 'source-over';
    }

    // Formas geométricas
    if (['line', 'arrow', 'rect', 'ellipse'].includes(stroke.tool) && pts.length >= 2) {
      this.drawGeometricShape(ctx, stroke.tool, pts[0], pts[pts.length - 1], stroke.color, stroke.size);
      ctx.restore();
      return;
    }

    // Tinta Arcoíris de OneNote
    if (stroke.tool === 'rainbow' || stroke.color === 'rainbow') {
      for (let i = 1; i < pts.length; i++) {
        const p0 = pts[i - 1];
        const p1 = pts[i];
        const hue = (i * 10) % 360;
        ctx.strokeStyle = `hsl(${hue}, 90%, 52%)`;
        const pFactor = stroke.usePressure ? 0.5 + p1.p * 0.9 : 1;
        ctx.lineWidth = Math.max(1, stroke.size * pFactor);
        ctx.beginPath();
        ctx.moveTo(p0.x, p0.y);
        ctx.lineTo(p1.x, p1.y);
        ctx.stroke();
      }
      ctx.restore();
      return;
    }

    // Trazo con sensibilidad a la presión del boli táctil
    if (stroke.usePressure && pts.length > 2) {
      ctx.strokeStyle = stroke.color;
      for (let i = 1; i < pts.length; i++) {
        const p0 = pts[i - 1];
        const p1 = pts[i];
        const pressureScale = 0.45 + (p1.p || 0.5) * 1.1;
        ctx.lineWidth = Math.max(1, stroke.size * pressureScale);
        ctx.beginPath();
        ctx.moveTo(p0.x, p0.y);
        ctx.lineTo(p1.x, p1.y);
        ctx.stroke();
      }
      ctx.restore();
      return;
    }

    // Curva cuadrática de Bézier suavizada estándar
    ctx.strokeStyle = stroke.color;
    ctx.lineWidth = stroke.size;
    ctx.beginPath();
    ctx.moveTo(pts[0].x, pts[0].y);

    if (pts.length === 1) {
      ctx.lineTo(pts[0].x + 0.5, pts[0].y + 0.5);
    } else {
      for (let i = 1; i < pts.length - 1; i++) {
        const midX = (pts[i].x + pts[i + 1].x) / 2;
        const midY = (pts[i].y + pts[i + 1].y) / 2;
        ctx.quadraticCurveTo(pts[i].x, pts[i].y, midX, midY);
      }
      const last = pts[pts.length - 1];
      ctx.lineTo(last.x, last.y);
    }
    ctx.stroke();
    ctx.restore();
  }

  drawGeometricShape(ctx, type, p1, p2, color, size) {
    ctx.strokeStyle = color;
    ctx.lineWidth = size;
    ctx.beginPath();

    if (type === 'line') {
      ctx.moveTo(p1.x, p1.y);
      ctx.lineTo(p2.x, p2.y);
      ctx.stroke();
    } else if (type === 'arrow') {
      const headLen = Math.max(12, size * 3.5);
      const angle = Math.atan2(p2.y - p1.y, p2.x - p1.x);
      ctx.moveTo(p1.x, p1.y);
      ctx.lineTo(p2.x, p2.y);
      ctx.lineTo(
        p2.x - headLen * Math.cos(angle - Math.PI / 6),
        p2.y - headLen * Math.sin(angle - Math.PI / 6)
      );
      ctx.moveTo(p2.x, p2.y);
      ctx.lineTo(
        p2.x - headLen * Math.cos(angle + Math.PI / 6),
        p2.y - headLen * Math.sin(angle + Math.PI / 6)
      );
      ctx.stroke();
    } else if (type === 'rect') {
      ctx.strokeRect(p1.x, p1.y, p2.x - p1.x, p2.y - p1.y);
    } else if (type === 'ellipse') {
      const rx = Math.abs(p2.x - p1.x) / 2;
      const ry = Math.abs(p2.y - p1.y) / 2;
      const cx = Math.min(p1.x, p2.x) + rx;
      const cy = Math.min(p1.y, p2.y) + ry;
      ctx.ellipse(cx, cy, rx, ry, 0, 0, Math.PI * 2);
      ctx.stroke();
    }
  }

  redrawOverlayCanvas(previewEndPt = null) {
    const ctx = this.overlayCtx;
    if (!ctx) return;
    ctx.clearRect(0, 0, 1600, 2000);

    // 1. Previsualización de forma geométrica en vivo
    if (this.shapeStartPoint && previewEndPt) {
      ctx.save();
      this.drawGeometricShape(
        ctx,
        this.currentTool,
        this.shapeStartPoint,
        previewEndPt,
        this.inkColor,
        this.strokeSize
      );
      ctx.restore();
    }

    // 2. Polígono de Selección de Lazo mientras se arrastra el boli
    if (this.lassoPolygon.length > 1) {
      ctx.save();
      ctx.strokeStyle = '#7719aa';
      ctx.fillStyle = 'rgba(119, 25, 170, 0.08)';
      ctx.lineWidth = 1.8;
      ctx.setLineDash([6, 4]);
      ctx.beginPath();
      ctx.moveTo(this.lassoPolygon[0].x, this.lassoPolygon[0].y);
      this.lassoPolygon.forEach((p) => ctx.lineTo(p.x, p.y));
      ctx.closePath();
      ctx.fill();
      ctx.stroke();
      ctx.restore();
    }

    // 3. Caja delimitadora de trazos seleccionados con el Lazo
    if (this.lassoBoundingBox && this.selectedStrokeIds.size > 0) {
      const b = this.lassoBoundingBox;
      ctx.save();
      ctx.strokeStyle = '#7719aa';
      ctx.fillStyle = 'rgba(119, 25, 170, 0.05)';
      ctx.lineWidth = 2;
      ctx.setLineDash([5, 5]);
      ctx.strokeRect(b.minX - 8, b.minY - 8, b.maxX - b.minX + 16, b.maxY - b.minY + 16);
      ctx.fillRect(b.minX - 8, b.minY - 8, b.maxX - b.minX + 16, b.maxY - b.minY + 16);
      ctx.restore();
    }
  }

  /* =========================================================
     BORRADOR DE TRAZOS Y BORRADOR PARCIAL VECTORIAL
     ========================================================= */
  eraseAtPoint(pt) {
    const page = this.getActivePage();
    const radius = Math.max(12, this.strokeSize * 2.2);

    if (this.currentTool === 'stroke-eraser') {
      // Borrador de Trazos OneNote: elimina el trazo entero al tocar cualquiera de sus puntos
      const beforeCount = page.strokes.length;
      page.strokes = page.strokes.filter(
        (stroke) => !stroke.points.some((p) => Math.hypot(p.x - pt.x, p.y - pt.y) <= radius)
      );
      if (page.strokes.length !== beforeCount) {
        this.redrawInkCanvas();
      }
    } else if (this.currentTool === 'pixel-eraser') {
      // Borrador parcial: divide o recorta los puntos por donde pasa la goma
      const newStrokes = [];
      let changed = false;

      page.strokes.forEach((stroke) => {
        if (['line', 'arrow', 'rect', 'ellipse'].includes(stroke.tool)) {
          const hit = stroke.points.some((p) => Math.hypot(p.x - pt.x, p.y - pt.y) <= radius);
          if (!hit) newStrokes.push(stroke);
          else changed = true;
          return;
        }

        let currentSegment = [];
        stroke.points.forEach((p) => {
          if (Math.hypot(p.x - pt.x, p.y - pt.y) <= radius) {
            changed = true;
            if (currentSegment.length > 1) {
              newStrokes.push({
                ...stroke,
                id: 'strk-' + Date.now() + '-' + Math.random().toString(36).slice(2, 6),
                points: currentSegment
              });
            }
            currentSegment = [];
          } else {
            currentSegment.push(p);
          }
        });

        if (currentSegment.length > 1) {
          newStrokes.push({
            ...stroke,
            points: currentSegment
          });
        }
      });

      if (changed) {
        page.strokes = newStrokes;
        this.redrawInkCanvas();
      }
    }
  }

  /* =========================================================
     SELECCIÓN DE LAZO (LASSO SELECT) DE ONENOTE
     ========================================================= */
  computeLassoSelection() {
    if (this.lassoPolygon.length < 3) return;
    const page = this.getActivePage();
    this.selectedStrokeIds.clear();

    page.strokes.forEach((stroke) => {
      const inside = stroke.points.some((pt) => this.isPointInPolygon(pt, this.lassoPolygon));
      if (inside) {
        this.selectedStrokeIds.add(stroke.id);
      }
    });

    if (this.selectedStrokeIds.size > 0) {
      this.updateLassoBoundingBox();
      this.lassoSelectionControls.classList.remove('hidden');
      this.lassoCountBadge.textContent = `${this.selectedStrokeIds.size} trazo(s)`;
      this.showToast(
        `🪢 ${this.selectedStrokeIds.size} trazo(s) seleccionados. Arrástralos para moverlos o usa los botones superiores.`
      );
    } else {
      this.clearLassoSelection();
    }
  }

  isPointInPolygon(point, vs) {
    let x = point.x,
      y = point.y;
    let inside = false;
    for (let i = 0, j = vs.length - 1; i < vs.length; j = i++) {
      let xi = vs[i].x,
        yi = vs[i].y;
      let xj = vs[j].x,
        yj = vs[j].y;
      let intersect = yi > y !== yj > y && x < ((xj - xi) * (y - yi)) / (yj - yi + 0.00001) + xi;
      if (intersect) inside = !inside;
    }
    return inside;
  }

  isPointInBox(pt, box) {
    return (
      pt.x >= box.minX - 10 &&
      pt.x <= box.maxX + 10 &&
      pt.y >= box.minY - 10 &&
      pt.y <= box.maxY + 10
    );
  }

  updateLassoBoundingBox() {
    const page = this.getActivePage();
    let minX = Infinity,
      minY = Infinity,
      maxX = -Infinity,
      maxY = -Infinity;

    page.strokes.forEach((s) => {
      if (!this.selectedStrokeIds.has(s.id)) return;
      s.points.forEach((p) => {
        if (p.x < minX) minX = p.x;
        if (p.y < minY) minY = p.y;
        if (p.x > maxX) maxX = p.x;
        if (p.y > maxY) maxY = p.y;
      });
    });

    this.lassoBoundingBox = minX < Infinity ? { minX, minY, maxX, maxY } : null;
  }

  moveSelectedStrokes(dx, dy) {
    const page = this.getActivePage();
    page.strokes.forEach((s) => {
      if (!this.selectedStrokeIds.has(s.id)) return;
      s.points.forEach((p) => {
        p.x += dx;
        p.y += dy;
      });
    });
    this.updateLassoBoundingBox();
    this.redrawInkCanvas();
    this.redrawOverlayCanvas();
  }

  clearLassoSelection() {
    this.selectedStrokeIds.clear();
    this.lassoBoundingBox = null;
    this.lassoSelectionControls.classList.add('hidden');
    this.redrawOverlayCanvas();
  }

  /* =========================================================
     REPRODUCCIÓN ANIMADA DE TINTA (ONENOTE INK REPLAY)
     ========================================================= */
  async replayPageInk() {
    const page = this.getActivePage();
    if (!page.strokes || page.strokes.length === 0) {
      this.showToast('No hay trazos de boli en esta página para reproducir.');
      return;
    }

    this.showToast('▶️ Reproduciendo trazos de boli táctil paso a paso...');
    const ctx = this.inkCtx;
    ctx.clearRect(0, 0, 1600, 2000);

    for (const stroke of page.strokes) {
      if (['line', 'arrow', 'rect', 'ellipse'].includes(stroke.tool)) {
        this.drawSingleStroke(ctx, stroke);
        await new Promise((r) => setTimeout(r, 120));
        continue;
      }

      const partialStroke = { ...stroke, points: [] };
      const step = Math.max(1, Math.floor(stroke.points.length / 25));
      for (let i = 0; i < stroke.points.length; i += step) {
        partialStroke.points = stroke.points.slice(0, i + step);
        this.drawSingleStroke(ctx, partialStroke);
        await new Promise((r) => setTimeout(r, 16));
      }
    }
    this.redrawInkCanvas();
  }

  /* =========================================================
     EXPORTACIÓN DE PÁGINA COMPLETA A IMAGEN PNG
     ========================================================= */
  async renderCurrentPageToDataUrl() {
    const page = this.getActivePage();
    const outCanvas = document.createElement('canvas');
    outCanvas.width = 1200;
    outCanvas.height = 1400;
    const ctx = outCanvas.getContext('2d');

    // 1. Fondo del papel
    ctx.fillStyle = page.paperColor || '#ffffff';
    ctx.fillRect(0, 0, outCanvas.width, outCanvas.height);

    // 2. Líneas de regla o cuadrícula
    ctx.strokeStyle = 'rgba(59, 130, 246, 0.22)';
    ctx.lineWidth = 1;
    if (page.paperStyle === 'ruled-wide' || page.paperStyle === 'ruled-narrow') {
      const gap = page.paperStyle === 'ruled-narrow' ? 24 : 32;
      for (let y = gap; y < outCanvas.height; y += gap) {
        ctx.beginPath();
        ctx.moveTo(0, y);
        ctx.lineTo(outCanvas.width, y);
        ctx.stroke();
      }
      ctx.strokeStyle = 'rgba(239, 68, 68, 0.28)';
      ctx.beginPath();
      ctx.moveTo(72, 0);
      ctx.lineTo(72, outCanvas.height);
      ctx.stroke();
    } else if (page.paperStyle === 'grid-small' || page.paperStyle === 'grid-large') {
      const gap = page.paperStyle === 'grid-small' ? 20 : 36;
      for (let x = 0; x < outCanvas.width; x += gap) {
        ctx.beginPath();
        ctx.moveTo(x, 0);
        ctx.lineTo(x, outCanvas.height);
        ctx.stroke();
      }
      for (let y = 0; y < outCanvas.height; y += gap) {
        ctx.beginPath();
        ctx.moveTo(0, y);
        ctx.lineTo(outCanvas.width, y);
        ctx.stroke();
      }
    }

    // 3. Título de la página
    ctx.fillStyle = '#1e293b';
    ctx.font = 'bold 26px sans-serif';
    ctx.fillText(page.title || 'Página de OneNote', 88, 56);
    ctx.font = '12px sans-serif';
    ctx.fillStyle = '#64748b';
    ctx.fillText(page.createdAt || '', 88, 78);

    // 4. Imágenes flotantes y contenedores de texto
    for (const block of page.blocks || []) {
      if (block.type === 'image' && block.src) {
        await new Promise((resolve) => {
          const img = new Image();
          img.crossOrigin = 'anonymous';
          img.onload = () => {
            ctx.drawImage(img, block.x, block.y, block.width || 300, block.height || 200);
            resolve();
          };
          img.onerror = () => resolve();
          img.src = block.src;
        });
      } else if (block.type === 'text') {
        ctx.save();
        ctx.fillStyle = block.isSticky ? '#fef08a' : 'rgba(255, 255, 255, 0.85)';
        ctx.strokeStyle = '#cbd5e1';
        ctx.fillRect(block.x, block.y, block.width || 360, 120);
        ctx.strokeRect(block.x, block.y, block.width || 360, 120);

        const tempDiv = document.createElement('div');
        tempDiv.innerHTML = block.html || '';
        const textLines = (tempDiv.innerText || '').split('\n');
        ctx.fillStyle = '#1e293b';
        ctx.font = '14px sans-serif';
        textLines.slice(0, 8).forEach((line, idx) => {
          ctx.fillText(line.slice(0, 68), block.x + 12, block.y + 26 + idx * 18);
        });
        ctx.restore();
      }
    }

    // 5. Trazos de boli táctil encima
    (page.strokes || []).forEach((s) => this.drawSingleStroke(ctx, s));

    return outCanvas.toDataURL('image/png');
  }

  /**
   * Renderiza únicamente los trazos de boli táctil sobre un lienzo recortado
   * para incrustarlos como imagen dentro de un documento de Google Docs.
   */
  renderInkOnlyToDataUrl() {
    const page = this.getActivePage();
    if (!page || !Array.isArray(page.strokes) || page.strokes.length === 0) return null;

    let minX = Infinity;
    let minY = Infinity;
    let maxX = -Infinity;
    let maxY = -Infinity;
    page.strokes.forEach((s) => {
      (s.points || []).forEach((pt) => {
        if (pt.x < minX) minX = pt.x;
        if (pt.y < minY) minY = pt.y;
        if (pt.x > maxX) maxX = pt.x;
        if (pt.y > maxY) maxY = pt.y;
      });
    });

    if (!Number.isFinite(minX) || !Number.isFinite(minY)) return null;
    const pad = 28;
    const offsetX = Math.max(0, minX - pad);
    const offsetY = Math.max(0, minY - pad);
    const width = Math.min(1800, Math.max(320, Math.ceil(maxX - offsetX + pad * 2)));
    const height = Math.min(1600, Math.max(160, Math.ceil(maxY - offsetY + pad * 2)));

    const outCanvas = document.createElement('canvas');
    outCanvas.width = width;
    outCanvas.height = height;
    const ctx = outCanvas.getContext('2d');

    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, width, height);

    ctx.save();
    ctx.translate(-offsetX, -offsetY);
    page.strokes.forEach((s) => this.drawSingleStroke(ctx, s));
    ctx.restore();

    return outCanvas.toDataURL('image/png');
  }

  /* =========================================================
     INTERACCIÓN CON GOOGLE DRIVE Y PERFIL DE USUARIO
     ========================================================= */
  updateGoogleAuthUI(user) {
    const modalName = document.getElementById('modalUserName');
    const modalEmail = document.getElementById('modalUserEmail');
    const modalAvatarBox = document.getElementById('modalAvatarBox');
    const signOutBtn = document.getElementById('signOutGoogleBtn');
    const emailInput = document.getElementById('googleEmailInput');

    if (user) {
      const initial = (user.name?.[0] || user.email?.[0] || 'G').toUpperCase();
      this.googleUserLabel.textContent = user.email.split('@')[0];
      modalName.textContent = `${user.name} (Validado)`;
      modalEmail.textContent = `${user.email} · ${user.authMethod || 'Credenciales directas'}`;
      modalAvatarBox.textContent = initial;
      signOutBtn.classList.remove('hidden');

      if (this.topbarAvatarBadge) {
        this.topbarAvatarBadge.textContent = initial;
        this.topbarAvatarBadge.classList.remove('hidden');
      }
      this.googleGIcon.classList.add('hidden');
      if (emailInput && !emailInput.value) emailInput.value = user.email;
      this.syncText.textContent = 'Google Drive Activo';
    } else {
      this.googleUserLabel.textContent = 'Credenciales Google';
      if (this.topbarAvatarBadge) this.topbarAvatarBadge.classList.add('hidden');
      this.googleGIcon.classList.remove('hidden');
      modalName.textContent = 'Sin validar';
      modalEmail.textContent =
        'Introduce tu correo de Google o usa el autocompletado del navegador (Sin tokens OAuth).';
      modalAvatarBox.textContent = 'G';
      signOutBtn.classList.add('hidden');
      this.syncText.textContent = 'Guardado';
    }

    this.renderSavedGoogleAccountsUI();
    this.refreshDriveFilesListUI();
  }

  renderSavedGoogleAccountsUI() {
    const wrap = document.getElementById('savedAccountsWrap');
    const list = document.getElementById('savedAccountsList');
    if (!wrap || !list) return;

    const accounts = this.driveManager.getSavedAccounts();
    if (!accounts || accounts.length === 0) {
      wrap.classList.add('hidden');
      return;
    }

    wrap.classList.remove('hidden');
    list.innerHTML = '';
    accounts.forEach((acc) => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'saved-account-chip';
      btn.textContent = `🔑 ${acc.email}`;
      btn.title = 'Validar inmediatamente con esta cuenta de Google guardada';
      btn.addEventListener('click', async () => {
        await this.driveManager.signInWithDirectCredentials({
          email: acc.email,
          name: acc.name
        });
        await this.driveManager.saveWorkspaceToDrive(this.workspace);
        this.refreshDriveFilesListUI();
      });
      list.appendChild(btn);
    });
  }

  async openAndroidPwaInfoModal() {
    const urlCode = document.getElementById('androidWifiUrlText');
    const qrImg = document.getElementById('androidQrImg');
    let targetUrl = window.location.href.split('#')[0].split('?')[0];
    if (window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1') {
      targetUrl = 'https://piterpaul.github.io/cloneonenote/';
    }

    try {
      const res = await fetch('/api/network-info');
      if (res.ok) {
        const info = await res.json();
        if (info && info.androidUrl) {
          targetUrl = info.androidUrl;
        }
      }
    } catch (_) {}

    if (urlCode) urlCode.textContent = targetUrl;
    if (qrImg) {
      qrImg.src = `https://api.qrserver.com/v1/create-qr-code/?size=180x180&data=${encodeURIComponent(
        targetUrl
      )}`;
    }
    this.androidPwaDialog.showModal();
  }

  buildShareNoteUrl() {
    const page = this.getActivePage();
    let baseUrl = window.location.href.split('#')[0].split('?')[0];
    if (window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1') {
      baseUrl = 'https://piterpaul.github.io/cloneonenote/';
    }
    if (!page) return baseUrl;

    try {
      const sharePayload = {
        t: page.title || 'Nota Compartida',
        ps: page.paperStyle || 'ruled-wide',
        pc: page.paperColor || '#ffffff',
        b: (page.blocks || [])
          .filter((blk) => blk.type === 'text')
          .map((blk) => ({
            x: blk.x,
            y: blk.y,
            w: blk.width,
            s: Boolean(blk.isSticky),
            h: blk.html
          })),
        st: (page.strokes || []).slice(0, 60)
      };
      const jsonStr = JSON.stringify(sharePayload);
      const b64 = btoa(encodeURIComponent(jsonStr));
      return `${baseUrl}#share=${b64}`;
    } catch (_) {
      return baseUrl;
    }
  }

  openShareModal() {
    const shareNoteUrl = this.buildShareNoteUrl();
    const noteUrlEl = document.getElementById('shareNoteUrlText');
    const qrImg = document.getElementById('shareQrImg');
    if (noteUrlEl) noteUrlEl.textContent = shareNoteUrl;
    if (qrImg) {
      qrImg.src = `https://api.qrserver.com/v1/create-qr-code/?size=180x180&data=${encodeURIComponent(
        'https://piterpaul.github.io/cloneonenote/'
      )}`;
    }
    this.shareLinkDialog?.showModal();
  }

  checkUrlSharedNoteOnStartup() {
    const hash = window.location.hash || '';
    if (!hash.startsWith('#share=')) return;
    const encoded = hash.slice('#share='.length).trim();
    if (!encoded) return;

    try {
      const jsonStr = decodeURIComponent(atob(encoded));
      const payload = JSON.parse(jsonStr);
      const section = this.getActiveSection();
      if (!section || !payload) return;

      const sharedPageId = 'page-shared-' + Date.now();
      const newPage = {
        id: sharedPageId,
        title: `🔗 ${payload.t || 'Nota Compartida'}`,
        createdAt: new Date().toLocaleString('es-ES'),
        isSubpage: false,
        paperStyle: payload.ps || 'ruled-wide',
        paperColor: payload.pc || '#ffffff',
        blocks: Array.isArray(payload.b)
          ? payload.b.map((item, idx) => ({
              id: `blk-sh-${idx}-${Date.now()}`,
              type: 'text',
              x: item.x || 88,
              y: item.y || 125,
              width: item.w || 460,
              isSticky: Boolean(item.s),
              html: item.h || ''
            }))
          : [],
        strokes: Array.isArray(payload.st) ? payload.st : []
      };

      section.pages.unshift(newPage);
      this.workspace.activePageId = sharedPageId;
      this.saveLocalState();
      history.replaceState(null, '', window.location.pathname + window.location.search);
      setTimeout(() => {
        this.showToast(`🔗 Nota compartida "${payload.t || 'Nota'}" abierta desde el enlace.`);
      }, 300);
    } catch (_) {}
  }

  async refreshDriveFilesListUI() {
    const container = document.getElementById('driveFilesListContainer');
    if (!container) return;

    const files = await this.driveManager.listDriveFiles();
    if (!files || files.length === 0) {
      container.innerHTML = `<p class="muted-text">Inicia sesión con Google SSO o pulsa <strong>"☁️ Almacenar Todo en Google Drive"</strong> para guardar tus blocs y archivos .docx.</p>`;
      return;
    }

    container.innerHTML = '';
    files.forEach((f) => {
      const row = document.createElement('div');
      row.className = 'drive-file-row';
      const dateStr = f.updatedAt ? new Date(f.updatedAt).toLocaleString('es-ES') : '';
      const kb = f.size ? `${Math.max(1, Math.round(f.size / 1024))} KB` : '';
      const icon = f.mimeType?.includes('image')
        ? '🖼️'
        : f.name?.endsWith('.docx') || f.mimeType?.includes('wordprocessingml')
        ? '📄'
        : '📘';
      const badgeText = f.isRemoteDrive ? '☁️ Google Drive' : '✓ Local + Drive';
      row.innerHTML = `
        <div>
          <strong>${icon} ${f.name}</strong>
          <div style="font-size:11px;color:var(--text-muted)">Actualizado: ${dateStr} · ${kb}</div>
        </div>
        <span style="color:#16a34a;font-weight:600;font-size:12px;">${badgeText}</span>
      `;
      container.appendChild(row);
    });
  }

  /* =========================================================
     LECTURA Y EXPORTACIÓN A .DOCX + GOOGLE DOCS EN LA PWA
     ========================================================= */
  async openGoogleDocsModal() {
    await this.renderSavedGoogleDocsListUI();
    this.googleDocsDialog?.showModal();
  }

  async saveCurrentPageAsGoogleDoc(downloadFile = true) {
    const page = this.getActivePage();
    if (!page) return;
    const inkDataUrl = this.renderInkOnlyToDataUrl();
    await this.driveManager.savePageAsGoogleDoc(page, inkDataUrl, { downloadFile });
    await this.renderSavedGoogleDocsListUI();
    await this.refreshDriveFilesListUI();
  }

  async copyAndOpenNewGoogleDoc() {
    const page = this.getActivePage();
    if (!page) return;
    const inkDataUrl = this.renderInkOnlyToDataUrl();
    const docHtml = this.driveManager.buildGoogleDocHtmlFromPage(page, inkDataUrl);

    try {
      if (navigator.clipboard && window.ClipboardItem) {
        const htmlBlob = new Blob([docHtml], { type: 'text/html' });
        const textDiv = document.createElement('div');
        textDiv.innerHTML = docHtml;
        const textBlob = new Blob([textDiv.innerText || page.title || ''], { type: 'text/plain' });
        await navigator.clipboard.write([
          new ClipboardItem({ 'text/html': htmlBlob, 'text/plain': textBlob })
        ]);
      } else {
        const textDiv = document.createElement('div');
        textDiv.innerHTML = docHtml;
        await navigator.clipboard?.writeText(textDiv.innerText || '');
      }
      this.showToast('📋 Nota copiada con formato. Pégala (Cmd+V / Ctrl+V) en Google Docs.');
    } catch (_) {
      this.showToast('🌐 Abriendo documento nuevo en Google Docs...');
    }

    await this.saveCurrentPageAsGoogleDoc(false);
    window.open('https://docs.new', '_blank', 'noopener');
  }

  applyImportedGoogleDoc(imported, targetMode) {
    if (!imported || !imported.html) {
      this.showToast('No se encontró contenido en el archivo .docx.');
      return;
    }

    const mode =
      targetMode || document.getElementById('gdocImportTargetSelect')?.value || 'append-current';

    if (mode === 'append-current' || mode === 'current') {
      const page = this.getActivePage();
      if (!Array.isArray(page.blocks)) page.blocks = [];
      let mainTextBlock = page.blocks.find((b) => b.type === 'text' && !b.isSticky);
      if (mainTextBlock) {
        mainTextBlock.html = `${mainTextBlock.html || ''}<hr style="margin:12px 0;border:none;border-top:1px solid #cbd5e1;"><div class="imported-gdoc-section"><h3>📄 ${
          imported.title || 'Documento .docx'
        }</h3>${imported.html}</div>`;
      } else {
        page.blocks.push({
          id: 'blk-' + Date.now(),
          type: 'text',
          x: 88,
          y: 125,
          width: 720,
          html: `<h3>📄 ${imported.title || 'Documento .docx'}</h3>${imported.html}`
        });
      }
      this.renderActivePage();
      this.saveLocalState();
      this.showToast(`📥 "${imported.title}.docx" cargado en la página actual.`);
    } else {
      const sec = this.getActiveSection();
      const newPage = {
        id: 'page-' + Date.now(),
        title: `📄 ${imported.title || 'Documento .docx'}`,
        createdAt: new Date().toLocaleString('es-ES'),
        isSubpage: false,
        paperStyle: 'ruled-wide',
        paperColor: '#ffffff',
        blocks: [
          {
            id: 'blk-' + Date.now(),
            type: 'text',
            x: 88,
            y: 125,
            width: 760,
            html: imported.html
          }
        ],
        strokes: []
      };
      sec.pages.push(newPage);
      this.workspace.activePageId = newPage.id;
      this.renderAll();
      this.saveLocalState();
      this.showToast(`📄 Página creada desde "${imported.title}.docx".`);
    }

    this.googleDocsDialog?.close();
  }

  async renderSavedGoogleDocsListUI() {
    const container = document.getElementById('savedGDocsListContainer');
    if (!container) return;

    const docs = await this.driveManager.listSavedGoogleDocs();
    if (!docs || docs.length === 0) {
      container.innerHTML = `<p class="muted-text">Aún no has exportado ni abierto archivos .docx en esta cuenta.</p>`;
      return;
    }

    container.innerHTML = '';
    docs.forEach((doc) => {
      const row = document.createElement('div');
      row.className = 'gdoc-item-row';
      const dateStr = doc.updatedAt ? new Date(doc.updatedAt).toLocaleString('es-ES') : '';
      const displayFilename = doc.name?.endsWith('.docx')
        ? doc.name
        : `${(doc.title || 'Nota').replace(/\s+/g, '_')}.docx`;
      row.innerHTML = `
        <div class="gdoc-item-meta">
          <strong>📄 ${doc.title || displayFilename}</strong>
          <span>${displayFilename} · ${dateStr} ${doc.isRemoteDrive ? '· ☁️ Google Drive' : '· 💾 Local + Drive'}</span>
        </div>
        <div class="gdoc-item-actions">
          <button type="button" class="secondary-btn gdoc-use-btn" style="padding:5px 10px;font-size:11.5px;">📥 Abrir en Nota</button>
          <button type="button" class="secondary-btn gdoc-dl-btn" style="padding:5px 10px;font-size:11.5px;">⬇️ .docx</button>
        </div>
      `;

      row.querySelector('.gdoc-use-btn')?.addEventListener('click', async () => {
        if (doc.isRemoteDrive && doc.driveFileId) {
          const imported = await this.driveManager.downloadAndImportFromRealGoogleDrive(
            doc.driveFileId,
            doc.name,
            doc.mimeType
          );
          if (imported) {
            this.applyImportedGoogleDoc(imported);
          }
          return;
        }
        if (doc.docxBase64 && !doc.contentHtml) {
          try {
            const parsed = this.driveManager.dataUrlToUint8Array(
              `data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,${doc.docxBase64}`
            );
            if (parsed) {
              const html = await this.driveManager.parseDocxInBrowser(parsed.bytes.buffer);
              this.applyImportedGoogleDoc({ title: doc.title || displayFilename, html });
              return;
            }
          } catch (_) {}
        }
        this.applyImportedGoogleDoc({
          title: doc.title || displayFilename.replace(/\.docx$/i, ''),
          html: doc.contentHtml || doc.content || `<div>${doc.title || ''}</div>`
        });
      });

      row.querySelector('.gdoc-dl-btn')?.addEventListener('click', async () => {
        let docxBlob = null;
        if (doc.docxBase64) {
          const parsed = this.driveManager.dataUrlToUint8Array(
            `data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,${doc.docxBase64}`
          );
          if (parsed) {
            docxBlob = new Blob([parsed.bytes], {
              type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            });
          }
        }
        if (!docxBlob) {
          docxBlob = await this.driveManager.buildDocxBlobFromPage({
            title: doc.title || 'Nota',
            createdAt: dateStr,
            blocks: [{ type: 'text', html: doc.contentHtml || doc.content || '' }],
            strokes: []
          });
        }
        const a = document.createElement('a');
        a.href = URL.createObjectURL(docxBlob);
        a.download = displayFilename;
        a.click();
      });

      container.appendChild(row);
    });
  }

  /* =========================================================
     DICTADO POR VOZ (WEB SPEECH API) Y CÁMARA WEB
     ========================================================= */
  toggleVoiceDictation() {
    const SpeechRec = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SpeechRec) {
      this.showToast('El navegador actual no soporta Web Speech API. Usa Google Chrome en Mac o Chromebook.');
      return;
    }

    const labelEl = document.getElementById('voiceDictationLabel');
    if (this.isDictating && this.speechRecognition) {
      this.speechRecognition.stop();
      this.isDictating = false;
      labelEl.textContent = 'Dictar Voz';
      return;
    }

    this.speechRecognition = new SpeechRec();
    this.speechRecognition.lang = 'es-ES';
    this.speechRecognition.continuous = false;
    this.speechRecognition.interimResults = false;

    this.speechRecognition.onstart = () => {
      this.isDictating = true;
      labelEl.textContent = '🔴 Escuchando...';
      this.showToast('🎤 Habla ahora para dictar tu nota en español...');
    };

    this.speechRecognition.onresult = (event) => {
      const transcript = event.results[0][0].transcript;
      if (this.activeNoteEditable && document.body.contains(this.activeNoteEditable)) {
        this.activeNoteEditable.innerHTML += ` <span>${transcript}</span>`;
        this.activeNoteEditable.dispatchEvent(new Event('input'));
      } else {
        this.createTextContainerAt(120, 180, `<div>🎤 ${transcript}</div>`);
      }
      this.showToast('✓ Texto dictado añadido a la nota.');
    };

    this.speechRecognition.onend = () => {
      this.isDictating = false;
      labelEl.textContent = 'Dictar Voz';
    };

    this.speechRecognition.start();
  }

  async openCameraModal() {
    try {
      this.cameraStream = await navigator.mediaDevices.getUserMedia({ video: true, audio: false });
      this.cameraVideo.srcObject = this.cameraStream;
      this.cameraDialog.showModal();
    } catch (err) {
      this.showToast('No se pudo acceder a la cámara o se denegó el permiso.');
    }
  }

  closeCameraModal() {
    if (this.cameraStream) {
      this.cameraStream.getTracks().forEach((t) => t.stop());
      this.cameraStream = null;
    }
    this.cameraDialog.close();
  }

  captureCameraPhoto() {
    if (!this.cameraStream) return;
    const tempCanvas = document.createElement('canvas');
    tempCanvas.width = this.cameraVideo.videoWidth || 640;
    tempCanvas.height = this.cameraVideo.videoHeight || 480;
    tempCanvas.getContext('2d').drawImage(this.cameraVideo, 0, 0);
    const dataUrl = tempCanvas.toDataURL('image/png');
    this.closeCameraModal();
    this.insertImageDataUrl(dataUrl, 160, 220);
  }

  /* =========================================================
     GESTIÓN DE HERRAMIENTAS Y EVENTOS
     ========================================================= */
  setTool(toolName) {
    this.currentTool = toolName;
    if (toolName !== 'lasso') {
      this.clearLassoSelection();
    }
    this.updateToolUI();
  }

  updateToolUI() {
    document.querySelectorAll('.mode-trigger').forEach((btn) => {
      btn.classList.toggle('active', btn.dataset.tool === this.currentTool);
    });

    document.querySelectorAll('.pen-preset-btn').forEach((btn) => {
      const isSameTool = btn.dataset.tool === this.currentTool;
      const isSameColor =
        btn.dataset.color === 'rainbow' ||
        btn.dataset.color?.toLowerCase() === this.inkColor.toLowerCase();
      btn.classList.toggle('active', isSameTool && isSameColor);
    });

    this.pageSurface.classList.toggle('mode-text', this.currentTool === 'text');
    this.pageSurface.classList.toggle('mode-pan', this.currentTool === 'pan');
    this.pageSurface.classList.toggle(
      'mode-draw',
      this.currentTool !== 'text' && this.currentTool !== 'pan'
    );
  }

  bindEvents() {
    // Mostrar / Ocultar Barra Lateral
    this.toggleSidebarBtn.addEventListener('click', () => {
      this.sidebarContainer.classList.toggle('collapsed');
    });

    // Modo Claro / Oscuro
    this.themeToggleBtn.addEventListener('click', () => {
      document.body.classList.toggle('dark-theme');
      const isDark = document.body.classList.contains('dark-theme');
      this.themeToggleBtn.textContent = isDark ? '☀️' : '🌙';
      this.redrawInkCanvas();
    });

    // Pestañas del Ribbon
    document.querySelectorAll('.ribbon-tab').forEach((tabBtn) => {
      tabBtn.addEventListener('click', () => {
        document.querySelectorAll('.ribbon-tab').forEach((b) => b.classList.remove('active'));
        document.querySelectorAll('.ribbon-panel').forEach((p) => p.classList.remove('active'));
        tabBtn.classList.add('active');
        const panel = document.getElementById(`tab-${tabBtn.dataset.tab}`);
        if (panel) panel.classList.add('active');
      });
    });

    // Botones de Modo (Texto, Lazo, Mano, Borradores, Formas)
    document.querySelectorAll('.mode-trigger').forEach((btn) => {
      btn.addEventListener('click', () => {
        this.setTool(btn.dataset.tool);
      });
    });

    // Presets de Bolígrafos y Subrayadores
    document.querySelectorAll('.pen-preset-btn').forEach((btn) => {
      btn.addEventListener('click', () => {
        const tool = btn.dataset.tool;
        const color = btn.dataset.color;
        const size = Number(btn.dataset.size || 3);
        if (color && color !== 'rainbow') {
          this.inkColor = color;
          this.customInkColorPicker.value = color;
        }
        this.strokeSize = size;
        this.strokeSizeRange.value = String(size);
        this.strokeSizeValue.textContent = String(size);
        this.setTool(tool);
      });
    });

    // Paleta de Colores de Tinta
    document.querySelectorAll('.color-swatch').forEach((swatch) => {
      swatch.addEventListener('click', () => {
        document.querySelectorAll('.color-swatch').forEach((s) => s.classList.remove('active'));
        swatch.classList.add('active');
        this.inkColor = swatch.dataset.color;
        this.customInkColorPicker.value = this.inkColor;
        if (this.currentTool === 'text' || this.currentTool === 'stroke-eraser') {
          this.setTool('pen');
        } else {
          this.updateToolUI();
        }
      });
    });

    this.customInkColorPicker.addEventListener('input', (e) => {
      this.inkColor = e.target.value;
      if (this.currentTool === 'text') this.setTool('pen');
      this.updateToolUI();
    });

    this.strokeSizeRange.addEventListener('input', (e) => {
      this.strokeSize = Number(e.target.value);
      this.strokeSizeValue.textContent = String(this.strokeSize);
    });

    this.pressureSensitivityToggle.addEventListener('change', (e) => {
      this.usePressure = e.target.checked;
    });

    this.stylusOnlyToggle.addEventListener('change', (e) => {
      this.stylusOnly = e.target.checked;
      this.showToast(
        this.stylusOnly
          ? '✋ Rechazo de palma activado: El boli dibuja y el dedo desplaza la página.'
          : '🖊️ Modo libre: Puedes dibujar con boli táctil, dedo, ratón o trackpad.'
      );
    });

    // Eventos de Boli Táctil / PointerEvents sobre el Canvas
    this.inkCanvas.addEventListener('pointerdown', (e) => this.handlePointerDown(e));
    this.inkCanvas.addEventListener('pointermove', (e) => this.handlePointerMove(e));
    this.inkCanvas.addEventListener('pointerup', (e) => this.handlePointerUp(e));
    this.inkCanvas.addEventListener('pointercancel', (e) => this.handlePointerUp(e));

    // Cambio automático a Boli Táctil cuando se toca la página con un Stylus físico en modo Texto
    this.pageSurface.addEventListener(
      'pointerdown',
      (e) => {
        this.updatePointerStatusPill(e);
        if (e.pointerType === 'pen' && this.currentTool === 'text') {
          this.setTool('pen');
          this.showToast('🖊️ Boli táctil detectado: Modo dibujo activado automáticamente.');
        }
      },
      { capture: true }
    );

    // Clic en el lienzo en Modo Texto -> Crear nuevo contenedor de nota en esas coordenadas
    this.pageSurface.addEventListener('click', (e) => {
      if (this.currentTool !== 'text') return;
      if (
        e.target.closest('.note-container') ||
        e.target.closest('.image-container') ||
        e.target.closest('.page-header-zone')
      ) {
        return;
      }
      const rect = this.pageSurface.getBoundingClientRect();
      const x = (e.clientX - rect.left) / this.zoomLevel;
      const y = (e.clientY - rect.top) / this.zoomLevel;
      this.createTextContainerAt(x, y, '');
    });

    // Controles de Selección de Lazo
    this.lassoRecolorBtn.addEventListener('click', () => {
      if (this.selectedStrokeIds.size === 0) return;
      this.pushUndoSnapshot();
      const page = this.getActivePage();
      page.strokes.forEach((s) => {
        if (this.selectedStrokeIds.has(s.id)) {
          s.color = this.inkColor;
          if (s.tool === 'rainbow') s.tool = 'pen';
        }
      });
      this.redrawInkCanvas();
      this.saveLocalState();
      this.showToast('🎨 Color aplicado a los trazos seleccionados.');
    });

    this.lassoDeleteBtn.addEventListener('click', () => {
      if (this.selectedStrokeIds.size === 0) return;
      this.pushUndoSnapshot();
      const page = this.getActivePage();
      page.strokes = page.strokes.filter((s) => !this.selectedStrokeIds.has(s.id));
      this.clearLassoSelection();
      this.redrawInkCanvas();
      this.saveLocalState();
    });

    // Deshacer, Rehacer, Reproducir y Limpiar Tinta
    document.getElementById('quickUndoBtn').addEventListener('click', () => this.undo());
    document.getElementById('quickRedoBtn').addEventListener('click', () => this.redo());
    document.getElementById('replayInkBtn').addEventListener('click', () => this.replayPageInk());
    document.getElementById('clearInkBtn').addEventListener('click', () => {
      const page = this.getActivePage();
      if (!page.strokes || page.strokes.length === 0) return;
      this.pushUndoSnapshot();
      page.strokes = [];
      this.clearLassoSelection();
      this.redrawInkCanvas();
      this.saveLocalState();
      this.showToast('🗑️ Trazos de boli eliminados de la página.');
    });

    // Pestaña Inicio: Formato de Texto y Etiquetas OneNote
    document.getElementById('addTextBoxBtn').addEventListener('click', () => {
      this.createTextContainerAt(110 + Math.random() * 80, 170 + Math.random() * 80, '');
    });

    document.querySelectorAll('.format-btn[data-cmd]').forEach((btn) => {
      btn.addEventListener('click', () => {
        document.execCommand(btn.dataset.cmd, false, null);
      });
    });

    document.getElementById('fontSizeSelect').addEventListener('change', (e) => {
      document.execCommand('fontSize', false, e.target.value);
    });

    document.querySelectorAll('.tag-btn').forEach((btn) => {
      btn.addEventListener('click', () => {
        const tag = btn.dataset.tag;
        let snippet = '';
        if (tag === 'todo') {
          snippet = `<div class="onenote-todo-row"><input type="checkbox"> <span>Nueva tarea...</span></div>`;
        } else if (tag === 'important') {
          snippet = `<div>⭐ <b>Importante:</b> </div>`;
        } else if (tag === 'question') {
          snippet = `<div>❓ <b>Pregunta:</b> </div>`;
        } else if (tag === 'idea') {
          snippet = `<div>💡 <b>Idea:</b> </div>`;
        }

        if (this.activeNoteEditable && document.body.contains(this.activeNoteEditable)) {
          this.activeNoteEditable.innerHTML += snippet;
          this.activeNoteEditable.dispatchEvent(new Event('input'));
        } else {
          this.createTextContainerAt(120, 180, snippet);
        }
      });
    });

    document.getElementById('voiceDictationBtn').addEventListener('click', () => {
      this.toggleVoiceDictation();
    });

    // Pestaña Insertar: Imágenes, Cámara, Post-it, Tabla, Fecha
    const imageInput = document.getElementById('imageFileInput');
    document.getElementById('insertImageBtn').addEventListener('click', () => imageInput.click());
    imageInput.addEventListener('change', (e) => {
      Array.from(e.target.files || []).forEach((file, idx) => {
        const reader = new FileReader();
        reader.onload = (ev) =>
          this.insertImageDataUrl(ev.target.result, 120 + idx * 30, 180 + idx * 30);
        reader.readAsDataURL(file);
      });
      imageInput.value = '';
    });

    document.getElementById('openCameraModalBtn').addEventListener('click', () =>
      this.openCameraModal()
    );
    document.getElementById('closeCameraDialogBtn').addEventListener('click', () =>
      this.closeCameraModal()
    );
    document.getElementById('capturePhotoBtn').addEventListener('click', () =>
      this.captureCameraPhoto()
    );

    document.getElementById('insertStickyNoteBtn').addEventListener('click', () => {
      this.createTextContainerAt(
        520,
        160,
        `<div><b>📌 Nota Rápida</b></div><div>Escribe aquí tu recordatorio...</div>`,
        true
      );
    });

    document.getElementById('insertTableBtn').addEventListener('click', () => {
      const tableHtml = `<div><b>📊 Tabla de Datos</b></div><table><tr><th>Concepto</th><th>Detalle</th><th>Estado</th></tr><tr><td>Elemento 1</td><td>Descripción</td><td>✅</td></tr><tr><td>Elemento 2</td><td>Descripción</td><td>⏳</td></tr></table>`;
      this.createTextContainerAt(140, 220, tableHtml);
    });

    document.getElementById('insertTimestampBtn').addEventListener('click', () => {
      const stamp = `<div>🕒 <i>${new Date().toLocaleString('es-ES')}</i></div>`;
      if (this.activeNoteEditable && document.body.contains(this.activeNoteEditable)) {
        this.activeNoteEditable.innerHTML += stamp;
        this.activeNoteEditable.dispatchEvent(new Event('input'));
      } else {
        this.createTextContainerAt(120, 160, stamp);
      }
    });

    // Arrastrar y Soltar Imágenes (Drag & Drop) y Pegar con Cmd+V / Ctrl+V
    this.pageSurface.addEventListener('dragover', (e) => e.preventDefault());
    this.pageSurface.addEventListener('drop', (e) => {
      e.preventDefault();
      const files = Array.from(e.dataTransfer?.files || []).filter((f) =>
        f.type.startsWith('image/')
      );
      if (files.length === 0) return;
      const rect = this.pageSurface.getBoundingClientRect();
      const x = (e.clientX - rect.left) / this.zoomLevel;
      const y = (e.clientY - rect.top) / this.zoomLevel;
      files.forEach((file, i) => {
        const reader = new FileReader();
        reader.onload = (ev) => this.insertImageDataUrl(ev.target.result, x + i * 24, y + i * 24);
        reader.readAsDataURL(file);
      });
    });

    window.addEventListener('paste', (e) => {
      const items = Array.from(e.clipboardData?.items || []);
      const imgItem = items.find((item) => item.type.startsWith('image/'));
      if (imgItem) {
        e.preventDefault();
        const file = imgItem.getAsFile();
        const reader = new FileReader();
        reader.onload = (ev) => this.insertImageDataUrl(ev.target.result, 150, 210);
        reader.readAsDataURL(file);
      }
    });

    // Pestaña Vista: Líneas de papel, Color de página, Zoom y Exportación
    this.paperLinesSelect.addEventListener('change', (e) => {
      const page = this.getActivePage();
      page.paperStyle = e.target.value;
      this.renderActivePage();
      this.saveLocalState();
    });

    document.querySelectorAll('.paper-color-btn').forEach((btn) => {
      btn.addEventListener('click', () => {
        const page = this.getActivePage();
        page.paperColor = btn.dataset.bg;
        this.renderActivePage();
        this.saveLocalState();
      });
    });

    document.getElementById('zoomInBtn').addEventListener('click', () => this.setZoom(this.zoomLevel + 0.1));
    document.getElementById('zoomOutBtn').addEventListener('click', () => this.setZoom(this.zoomLevel - 0.1));
    document.getElementById('zoomResetBtn').addEventListener('click', () => this.setZoom(1));

    document.getElementById('exportPngBtn').addEventListener('click', async () => {
      const dataUrl = await this.renderCurrentPageToDataUrl();
      const a = document.createElement('a');
      a.href = dataUrl;
      a.download = `${(this.getActivePage().title || 'Nota_OneNote').replace(/\s+/g, '_')}.png`;
      a.click();
      this.showToast('🖼️ Página exportada como imagen PNG.');
    });

    document.getElementById('uploadPngDriveBtn').addEventListener('click', async () => {
      const dataUrl = await this.renderCurrentPageToDataUrl();
      await this.driveManager.uploadPagePngToDrive(this.getActivePage().title, dataUrl);
      this.refreshDriveFilesListUI();
    });

    document.getElementById('exportJsonBtn').addEventListener('click', () => {
      const blob = new Blob([JSON.stringify(this.workspace, null, 2)], {
        type: 'application/json'
      });
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = 'mis_blocs_onenote.json';
      a.click();
      this.showToast('💾 Copia de seguridad .onenote.json descargada.');
    });

    const importJsonInput = document.getElementById('importJsonInput');
    document.getElementById('importJsonBtn').addEventListener('click', () =>
      importJsonInput.click()
    );
    importJsonInput.addEventListener('change', (e) => {
      const file = e.target.files?.[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = (ev) => {
        try {
          const parsed = JSON.parse(ev.target.result);
          if (parsed && Array.isArray(parsed.notebooks)) {
            this.workspace = parsed;
            this.renderAll();
            this.saveLocalState();
            this.showToast('📂 Blocs de notas importados correctamente.');
          }
        } catch (_) {
          this.showToast('Archivo JSON inválido.');
        }
      };
      reader.readAsText(file);
    });

    // Gestión de Blocs, Secciones y Páginas
    this.notebookSelect.addEventListener('change', (e) => {
      this.workspace.activeNotebookId = e.target.value;
      const nb = this.getActiveNotebook();
      this.workspace.activeSectionId = nb.sections[0].id;
      this.workspace.activePageId = nb.sections[0].pages[0].id;
      this.renderAll();
      this.saveLocalState();
    });

    this.newNotebookBtn.addEventListener('click', () => {
      const name = prompt('Nombre del nuevo Bloc de Notas:', '📘 Nuevo Bloc de Notas');
      if (!name || !name.trim()) return;
      const newNb = {
        id: 'nb-' + Date.now(),
        name: name.trim(),
        sections: [
          {
            id: 'sec-' + Date.now(),
            name: 'General',
            color: SECTION_COLORS[Math.floor(Math.random() * SECTION_COLORS.length)],
            pages: [
              {
                id: 'page-' + Date.now(),
                title: 'Primera página',
                createdAt: new Date().toLocaleString('es-ES'),
                isSubpage: false,
                paperStyle: 'ruled-wide',
                paperColor: '#ffffff',
                blocks: [],
                strokes: []
              }
            ]
          }
        ]
      };
      this.workspace.notebooks.push(newNb);
      this.workspace.activeNotebookId = newNb.id;
      this.workspace.activeSectionId = newNb.sections[0].id;
      this.workspace.activePageId = newNb.sections[0].pages[0].id;
      this.renderAll();
      this.saveLocalState();
    });

    this.newSectionBtn.addEventListener('click', () => {
      const name = prompt('Nombre de la nueva sección:', 'Nueva Sección');
      if (!name || !name.trim()) return;
      const nb = this.getActiveNotebook();
      const newSec = {
        id: 'sec-' + Date.now(),
        name: name.trim(),
        color: SECTION_COLORS[nb.sections.length % SECTION_COLORS.length],
        pages: [
          {
            id: 'page-' + Date.now(),
            title: 'Página sin título',
            createdAt: new Date().toLocaleString('es-ES'),
            isSubpage: false,
            paperStyle: 'ruled-wide',
            paperColor: '#ffffff',
            blocks: [],
            strokes: []
          }
        ]
      };
      nb.sections.push(newSec);
      this.workspace.activeSectionId = newSec.id;
      this.workspace.activePageId = newSec.pages[0].id;
      this.renderAll();
      this.saveLocalState();
    });

    this.newPageBtn.addEventListener('click', () => {
      const sec = this.getActiveSection();
      const newPage = {
        id: 'page-' + Date.now(),
        title: `Página ${sec.pages.length + 1}`,
        createdAt: new Date().toLocaleString('es-ES'),
        isSubpage: false,
        paperStyle: 'ruled-wide',
        paperColor: '#ffffff',
        blocks: [],
        strokes: []
      };
      sec.pages.push(newPage);
      this.workspace.activePageId = newPage.id;
      this.renderAll();
      this.saveLocalState();
      this.pageTitleInput.focus();
      this.pageTitleInput.select();
    });

    this.pageTitleInput.addEventListener('input', (e) => {
      const page = this.getActivePage();
      page.title = e.target.value;
      this.renderSidebar();
      this.saveLocalState();
    });

    // Buscador Global Instantáneo
    this.globalSearchInput.addEventListener('input', (e) => {
      const q = e.target.value.trim().toLowerCase();
      if (!q) {
        this.searchResultsDropdown.classList.add('hidden');
        return;
      }

      const matches = [];
      this.workspace.notebooks.forEach((nb) => {
        nb.sections.forEach((sec) => {
          sec.pages.forEach((pg) => {
            const titleMatch = (pg.title || '').toLowerCase().includes(q);
            const bodyMatch = (pg.blocks || []).some((b) =>
              (b.html || '').toLowerCase().includes(q)
            );
            if (titleMatch || bodyMatch) {
              matches.push({ nb, sec, pg });
            }
          });
        });
      });

      this.searchResultsDropdown.innerHTML = '';
      if (matches.length === 0) {
        this.searchResultsDropdown.innerHTML = `<div class="search-result-item">Sin resultados para "${q}"</div>`;
      } else {
        matches.forEach(({ nb, sec, pg }) => {
          const item = document.createElement('div');
          item.className = 'search-result-item';
          item.innerHTML = `
            <div><strong>📄 ${pg.title || 'Sin título'}</strong></div>
            <div class="search-result-meta">${nb.name} &rsaquo; ${sec.name}</div>
          `;
          item.addEventListener('click', () => {
            this.workspace.activeNotebookId = nb.id;
            this.workspace.activeSectionId = sec.id;
            this.workspace.activePageId = pg.id;
            this.searchResultsDropdown.classList.add('hidden');
            this.globalSearchInput.value = '';
            this.renderAll();
          });
          this.searchResultsDropdown.appendChild(item);
        });
      }
      this.searchResultsDropdown.classList.remove('hidden');
    });

    // Modal y Botones de Validación con Google SSO y Almacenamiento en Google Drive
    this.saveToDriveBtn.addEventListener('click', async () => {
      if (!this.driveManager.isSignedIn()) {
        this.renderSavedGoogleAccountsUI();
        this.googleDriveDialog.showModal();
        document.getElementById('googleEmailInput')?.focus();
        return;
      }
      await this.driveManager.saveWorkspaceToDrive(this.workspace);
      this.refreshDriveFilesListUI();
    });

    this.openDriveModalBtn.addEventListener('click', () => {
      this.renderSavedGoogleAccountsUI();
      this.refreshDriveFilesListUI();
      this.googleDriveDialog.showModal();
    });

    document.getElementById('closeDriveDialogBtn').addEventListener('click', () => {
      this.googleDriveDialog.close();
    });

    const directLoginForm = document.getElementById('directGoogleLoginForm');
    if (directLoginForm) {
      directLoginForm.addEventListener('submit', async (e) => {
        e.preventDefault();
        const email = document.getElementById('googleEmailInput')?.value || '';
        const password = document.getElementById('googlePasswordInput')?.value || '';
        const ok = await this.driveManager.signInWithDirectCredentials({ email, password });
        if (ok) {
          await this.driveManager.saveWorkspaceToDrive(this.workspace);
          this.refreshDriveFilesListUI();
        }
      });
    }

    const browserAutoCredsBtn = document.getElementById('browserAutoCredsBtn');
    if (browserAutoCredsBtn) {
      browserAutoCredsBtn.addEventListener('click', async () => {
        const email = document.getElementById('googleEmailInput')?.value || '';
        if (email) {
          await this.driveManager.signInWithDirectCredentials({ email });
        } else {
          await this.driveManager.signInFromBrowserCredentialStore();
        }
        await this.driveManager.saveWorkspaceToDrive(this.workspace);
        this.refreshDriveFilesListUI();
      });
    }

    document.getElementById('signOutGoogleBtn').addEventListener('click', () => {
      this.driveManager.signOut();
    });

    document.getElementById('syncNowModalBtn').addEventListener('click', async () => {
      await this.driveManager.saveWorkspaceToDrive(this.workspace);
      this.refreshDriveFilesListUI();
    });

    document.getElementById('loadFromDriveModalBtn').addEventListener('click', async () => {
      const restored = await this.driveManager.loadWorkspaceFromDrive();
      if (restored && Array.isArray(restored.notebooks)) {
        this.workspace = restored;
        this.renderAll();
        this.saveLocalState();
      }
    });

    const exportSystemDriveBtn = document.getElementById('exportSystemDriveBtn');
    if (exportSystemDriveBtn) {
      exportSystemDriveBtn.addEventListener('click', async () => {
        await this.driveManager.exportDirectToSystemGoogleDrive(this.workspace);
      });
    }

    document.getElementById('exportPagePngDriveModalBtn').addEventListener('click', async () => {
      const dataUrl = await this.renderCurrentPageToDataUrl();
      await this.driveManager.uploadPagePngToDrive(this.getActivePage().title, dataUrl);
      this.refreshDriveFilesListUI();
    });

    // Modal y Botones de Instalación PWA / QR para Android
    if (this.openAndroidPwaModalBtn) {
      this.openAndroidPwaModalBtn.addEventListener('click', () => {
        this.openAndroidPwaInfoModal();
      });
    }

    const closeAndroidBtn = document.getElementById('closeAndroidPwaDialogBtn');
    if (closeAndroidBtn) {
      closeAndroidBtn.addEventListener('click', () => {
        this.androidPwaDialog.close();
      });
    }

    const triggerPwaInstallBtn = document.getElementById('triggerPwaInstallBtn');
    if (triggerPwaInstallBtn) {
      triggerPwaInstallBtn.addEventListener('click', async () => {
        if (this.deferredInstallPrompt) {
          this.deferredInstallPrompt.prompt();
          await this.deferredInstallPrompt.userChoice;
          this.deferredInstallPrompt = null;
        } else {
          this.showToast(
            '📲 En Chrome (Android/Mac/Chromebook): abre el menú ⋮ del navegador y pulsa "Instalar aplicación" o "Añadir a pantalla de inicio".'
          );
        }
      });
    }

    const copyAndroidUrlBtn = document.getElementById('copyAndroidUrlBtn');
    if (copyAndroidUrlBtn) {
      copyAndroidUrlBtn.addEventListener('click', () => {
        const text = document.getElementById('androidWifiUrlText')?.textContent || '';
        navigator.clipboard?.writeText(text);
        this.showToast(`📋 URL copiada: ${text}`);
      });
    }

    // Modal y Botones de Compartir Enlace (x20web / Web HTTPS / Nota Codificada)
    if (this.openShareModalBtn) {
      this.openShareModalBtn.addEventListener('click', () => {
        this.openShareModal();
      });
    }

    document.getElementById('closeShareDialogBtn')?.addEventListener('click', () => {
      this.shareLinkDialog?.close();
    });

    document.getElementById('copyX20UrlBtn')?.addEventListener('click', () => {
      const text = document.getElementById('x20CorpUrlText')?.textContent || '';
      navigator.clipboard?.writeText(text);
      this.showToast(`📋 Enlace x20web copiado: ${text}`);
    });

    document.getElementById('copyPublicWebUrlBtn')?.addEventListener('click', () => {
      const text = document.getElementById('publicWebUrlText')?.textContent || '';
      navigator.clipboard?.writeText(text);
      this.showToast(`📋 Enlace Web Público copiado: ${text}`);
    });

    document.getElementById('copyShareNoteUrlBtn')?.addEventListener('click', () => {
      const text = document.getElementById('shareNoteUrlText')?.textContent || '';
      navigator.clipboard?.writeText(text);
      this.showToast('📋 Enlace directo con la nota actual copiado al portapapeles.');
    });

    document.getElementById('nativeShareUrlBtn')?.addEventListener('click', async () => {
      const shareUrl = document.getElementById('shareNoteUrlText')?.textContent || '';
      const pageTitle = this.getActivePage()?.title || 'Nota de OneNote Clone';
      if (navigator.share) {
        try {
          await navigator.share({
            title: pageTitle,
            text: `Mira esta nota en OneNote Web Clone: "${pageTitle}"`,
            url: shareUrl
          });
          return;
        } catch (_) {}
      }
      navigator.clipboard?.writeText(shareUrl);
      this.showToast('📋 Enlace copiado para compartir.');
    });

    // Conmutador Vista Simple Móvil (Texto a pantalla completa)
    if (this.toggleSimpleLayoutBtn) {
      this.toggleSimpleLayoutBtn.addEventListener('click', () => {
        this.setSimpleMobileMode(!this.isMobileSimpleMode, true);
      });
    }

    // Botones de la Barra Rápida Móvil (PWA Simplificada)
    document.getElementById('mobileInsertImageBtn')?.addEventListener('click', () => {
      document.getElementById('imageFileInput')?.click();
    });

    document.getElementById('mobileTodoBtn')?.addEventListener('click', () => {
      const snippet = `<div class="onenote-todo-row"><input type="checkbox"> <span>Nueva tarea...</span></div>`;
      if (this.activeNoteEditable && document.body.contains(this.activeNoteEditable)) {
        this.activeNoteEditable.innerHTML += snippet;
        this.activeNoteEditable.dispatchEvent(new Event('input'));
      } else {
        const firstEditable = this.noteBlocksLayer.querySelector('.note-editable');
        if (firstEditable) {
          firstEditable.innerHTML += snippet;
          firstEditable.dispatchEvent(new Event('input'));
        } else {
          this.createTextContainerAt(16, 110, snippet);
        }
      }
    });

    document.getElementById('mobileVoiceBtn')?.addEventListener('click', () => {
      this.toggleVoiceDictation();
    });

    document.getElementById('mobileSaveGDocBtn')?.addEventListener('click', () => {
      this.saveCurrentPageAsGoogleDoc(true);
    });

    document.getElementById('mobileImportGDocBtn')?.addEventListener('click', () => {
      this.openGoogleDocsModal();
    });

    // Modal y Botones de Google Docs (Coger y Grabar en formato Google Docs)
    if (this.openGoogleDocsModalBtn) {
      this.openGoogleDocsModalBtn.addEventListener('click', () => {
        this.openGoogleDocsModal();
      });
    }

    document.getElementById('closeGoogleDocsDialogBtn')?.addEventListener('click', () => {
      this.googleDocsDialog?.close();
    });

    document.getElementById('ribbonQuickSaveGDocBtn')?.addEventListener('click', () => {
      this.saveCurrentPageAsGoogleDoc(true);
    });

    document.getElementById('ribbonSaveGDocBtn')?.addEventListener('click', () => {
      this.saveCurrentPageAsGoogleDoc(true);
    });

    document.getElementById('ribbonImportGDocBtn')?.addEventListener('click', () => {
      this.openGoogleDocsModal();
    });

    document.getElementById('saveCurrentAsGDocBtn')?.addEventListener('click', () => {
      this.saveCurrentPageAsGoogleDoc(true);
    });

    document.getElementById('copyAndOpenNewGDocBtn')?.addEventListener('click', () => {
      this.copyAndOpenNewGoogleDoc();
    });

    const gdocFileInput = document.getElementById('gdocFileInput');
    document.getElementById('pickGDocFileBtn')?.addEventListener('click', () => {
      gdocFileInput?.click();
    });

    if (gdocFileInput) {
      gdocFileInput.addEventListener('change', async (e) => {
        const file = e.target.files?.[0];
        if (!file) return;
        const imported = await this.driveManager.importGoogleDocFile(file);
        if (imported) {
          this.applyImportedGoogleDoc(imported);
          await this.renderSavedGoogleDocsListUI();
        }
        gdocFileInput.value = '';
      });
    }

    document.getElementById('importGDocUrlBtn')?.addEventListener('click', async () => {
      const urlInput = document.getElementById('gdocUrlInput');
      const rawUrl = urlInput?.value || '';
      const imported = await this.driveManager.importGoogleDocFromUrl(rawUrl);
      if (imported) {
        if (urlInput) urlInput.value = '';
        this.applyImportedGoogleDoc(imported);
        await this.renderSavedGoogleDocsListUI();
      }
    });

    // Atajos de Teclado para Mac (Cmd ⌘) y Chromebook (Ctrl)
    window.addEventListener('keydown', (e) => {
      const isCmdOrCtrl = e.metaKey || e.ctrlKey;
      if (isCmdOrCtrl && e.key.toLowerCase() === 's') {
        e.preventDefault();
        this.driveManager.saveWorkspaceToDrive(this.workspace);
      } else if (isCmdOrCtrl && e.key.toLowerCase() === 'z') {
        if (document.activeElement?.classList.contains('note-editable')) return;
        e.preventDefault();
        if (e.shiftKey) this.redo();
        else this.undo();
      } else if (isCmdOrCtrl && e.key.toLowerCase() === 'y') {
        if (document.activeElement?.classList.contains('note-editable')) return;
        e.preventDefault();
        this.redo();
      } else if (
        (e.key === 'Delete' || e.key === 'Backspace') &&
        this.selectedStrokeIds.size > 0 &&
        !document.activeElement?.classList.contains('note-editable') &&
        document.activeElement !== this.pageTitleInput
      ) {
        e.preventDefault();
        this.lassoDeleteBtn.click();
      }
    });
  }

  setZoom(nextZoom) {
    this.zoomLevel = Math.max(0.6, Math.min(1.8, Number(nextZoom.toFixed(2))));
    this.pageSurface.style.transform = `scale(${this.zoomLevel})`;
    document.getElementById('zoomResetBtn').textContent = `${Math.round(this.zoomLevel * 100)}%`;
  }

  showToast(message) {
    const toast = document.createElement('div');
    toast.className = 'toast';
    toast.textContent = message;
    this.toastContainer.appendChild(toast);
    setTimeout(() => {
      toast.remove();
    }, 3400);
  }
}

window.addEventListener('DOMContentLoaded', () => {
  window.oneNoteApp = new OneNoteCloneApp();
});
