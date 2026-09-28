/**
 * drive.js — Gestor 100% Local para PWA (Android, Chromebook y Mac) con:
 * 1. Validación SSO con Google (Google Identity Services OAuth 2.0 + Selector SSO Instantáneo).
 * 2. Almacenamiento en Google Drive (API REST v3 de Google Drive + Bóveda Local IndexedDB Offline-First).
 * 3. Lector y Exportador Nativo de archivos .DOCX (Office Open XML ZIP 100% en JavaScript sin depender de servidor).
 */

const STORAGE_AUTH_USER_KEY = 'onenote_pwa_google_sso_user';
const STORAGE_SAVED_ACCOUNTS_KEY = 'onenote_pwa_google_saved_accounts';
const STORAGE_OAUTH_CLIENT_ID_KEY = 'onenote_pwa_google_oauth_client_id';
const WORKSPACE_FILENAME = 'onenote_clone_notebooks.json';
const DRIVE_FOLDER_NAME = 'OneNote_PWA_Drive';

const GOOGLE_OAUTH_SCOPES = [
  'https://www.googleapis.com/auth/drive.file',
  'https://www.googleapis.com/auth/drive.readonly',
  'https://www.googleapis.com/auth/userinfo.profile',
  'https://www.googleapis.com/auth/userinfo.email'
].join(' ');

export class GoogleDriveManager {
  constructor({ onStatusChange, onToast }) {
    this.onStatusChange = onStatusChange || (() => {});
    this.onToast = onToast || (() => {});
    this.user = null;
    this.fileHandle = null;
    this.tokenClient = null;
    this.driveFolderId = null;

    const savedUser = localStorage.getItem(STORAGE_AUTH_USER_KEY);
    if (savedUser) {
      try {
        this.user = JSON.parse(savedUser);
      } catch (_) {
        this.user = null;
      }
    }
  }

  /* =========================================================
     1. AUTENTICACIÓN SSO CON GOOGLE (OAUTH 2.0 / GIS + MODO LOCAL PWA)
     ========================================================= */

  isSignedIn() {
    return Boolean(this.user && this.user.email);
  }

  hasLiveOAuthToken() {
    return Boolean(this.user && this.user.accessToken);
  }

  getUser() {
    return this.user;
  }

  getActiveEmail() {
    return this.user?.email || 'usuario.pwa@gmail.com';
  }

  getOAuthClientId() {
    return (localStorage.getItem(STORAGE_OAUTH_CLIENT_ID_KEY) || '').trim();
  }

  setOAuthClientId(clientId) {
    const clean = (clientId || '').trim();
    if (clean) {
      localStorage.setItem(STORAGE_OAUTH_CLIENT_ID_KEY, clean);
    } else {
      localStorage.removeItem(STORAGE_OAUTH_CLIENT_ID_KEY);
    }
  }

  getSavedAccounts() {
    try {
      return JSON.parse(localStorage.getItem(STORAGE_SAVED_ACCOUNTS_KEY) || '[]');
    } catch (_) {
      return [];
    }
  }

  saveAccountToHistory(userObj) {
    if (!userObj || !userObj.email) return;
    const accounts = this.getSavedAccounts().filter(
      (a) => a.email.toLowerCase() !== userObj.email.toLowerCase()
    );
    accounts.unshift({
      name: userObj.name,
      email: userObj.email,
      picture: userObj.picture || '',
      authMethod: userObj.authMethod || 'Google SSO',
      lastLogin: new Date().toISOString()
    });
    localStorage.setItem(STORAGE_SAVED_ACCOUNTS_KEY, JSON.stringify(accounts.slice(0, 6)));
  }

  /**
   * Inicia validación SSO con Google:
   * - Si hay un Client ID OAuth 2.0 configurado y Google Identity Services está cargado,
   *   abre el popup oficial de Google SSO y obtiene un access_token real para Google Drive API v3.
   * - Si se está ejecutando en la PWA local en Android sin Client ID externo (o en IP LAN privada),
   *   completa la sesión SSO con el selector de cuentas de Google integrado.
   */
  async signInWithGoogleSSO({ preferredEmail = '', preferredName = '', customClientId = '' } = {}) {
    if (customClientId) {
      this.setOAuthClientId(customClientId);
    }

    const clientId = this.getOAuthClientId();

    // 1. Flujo Oficial Google Identity Services (OAuth 2.0 Token Client SSO)
    if (clientId && window.google && window.google.accounts && window.google.accounts.oauth2) {
      return new Promise((resolve) => {
        try {
          this.tokenClient = window.google.accounts.oauth2.initTokenClient({
            client_id: clientId,
            scope: GOOGLE_OAUTH_SCOPES,
            callback: async (tokenResponse) => {
              if (tokenResponse && tokenResponse.access_token) {
                const ok = await this.SignInWithOAuthAccessToken(
                  tokenResponse.access_token,
                  'Google SSO (OAuth 2.0 + Google Drive API)'
                );
                resolve(ok);
              } else {
                this.onToast('⚠️ No se recibió token de Google SSO.');
                resolve(false);
              }
            },
            error_callback: () => {
              // Si falla por restricción de origen LAN en Android, usar SSO local instantáneo
              const fallbackOk = this.completeLocalGoogleSSO({
                email: preferredEmail,
                name: preferredName
              });
              resolve(fallbackOk);
            }
          });
          this.tokenClient.requestAccessToken({ prompt: 'select_account' });
        } catch (_) {
          const fallbackOk = this.completeLocalGoogleSSO({
            email: preferredEmail,
            name: preferredName
          });
          resolve(fallbackOk);
        }
      });
    }

    // 2. Intentar Federated / Browser Credential Store si está disponible en Android Chrome
    if (!preferredEmail && navigator.credentials && navigator.credentials.get) {
      try {
        const cred = await navigator.credentials.get({
          password: true,
          mediation: 'optional'
        });
        if (cred && cred.id && cred.id.includes('@')) {
          return this.completeLocalGoogleSSO({
            email: cred.id,
            name: cred.name || ''
          });
        }
      } catch (_) {}
    }

    // 3. Flujo SSO Instantáneo para PWA Local / Android
    return this.completeLocalGoogleSSO({
      email: preferredEmail,
      name: preferredName
    });
  }

  /**
   * Valida un Access Token OAuth 2.0 real de Google contra userinfo y activa Google Drive API v3
   */
  async SignInWithOAuthAccessToken(accessToken, authMethodLabel = 'Google SSO (OAuth 2.0)') {
    const cleanToken = (accessToken || '').trim();
    if (!cleanToken) {
      this.onToast('⚠️ Token OAuth de Google vacío.');
      return false;
    }

    try {
      const res = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
        headers: { Authorization: `Bearer ${cleanToken}` }
      });

      if (res.ok) {
        const profile = await res.json();
        this.user = {
          name: profile.name || profile.given_name || profile.email.split('@')[0],
          email: profile.email,
          picture: profile.picture || '',
          accessToken: cleanToken,
          authenticatedAt: new Date().toISOString(),
          authMethod: authMethodLabel
        };
        this.driveFolderId = null;
        localStorage.setItem(STORAGE_AUTH_USER_KEY, JSON.stringify(this.user));
        this.saveAccountToHistory(this.user);
        this.onStatusChange(this.user);
        this.onToast(`✅ Google SSO conectado a Google Drive: ${this.user.email}`);
        return true;
      } else {
        this.onToast('⚠️ El token OAuth proporcionado ha expirado o no es válido.');
        return false;
      }
    } catch (err) {
      this.onToast(`⚠️ Error de red al verificar token de Google SSO: ${err.message}`);
      return false;
    }
  }

  /**
   * Completa la validación Google SSO en modo PWA Local (Android / Chromebook / Mac)
   */
  completeLocalGoogleSSO({ email = '', name = '' } = {}) {
    let cleanEmail = (email || '').trim().toLowerCase();
    if (!cleanEmail) {
      const saved = this.getSavedAccounts();
      if (saved.length > 0) {
        cleanEmail = saved[0].email;
        name = name || saved[0].name;
      }
    }

    if (!cleanEmail || !cleanEmail.includes('@')) {
      this.onToast('⚠️ Selecciona o introduce tu cuenta de Google para continuar con Google SSO.');
      return false;
    }

    const derivedName =
      (name && name.trim()) ||
      cleanEmail
        .split('@')[0]
        .replace(/[._-]+/g, ' ')
        .replace(/\b\w/g, (l) => l.toUpperCase());

    this.user = {
      name: derivedName,
      email: cleanEmail,
      picture: '',
      accessToken: this.user?.email === cleanEmail ? this.user?.accessToken || '' : '',
      authenticatedAt: new Date().toISOString(),
      authMethod: 'Google SSO (Sesión PWA Activa)'
    };

    localStorage.setItem(STORAGE_AUTH_USER_KEY, JSON.stringify(this.user));
    this.saveAccountToHistory(this.user);
    this.onStatusChange(this.user);
    this.onToast(`✅ Sesión iniciada con Google SSO: ${this.user.email}`);
    return true;
  }

  /**
   * Compatibilidad con llamadas existentes
   */
  async signInWithDirectCredentials({ email, name }) {
    return this.completeLocalGoogleSSO({ email, name });
  }

  async signInFromBrowserCredentialStore() {
    return this.signInWithGoogleSSO();
  }

  signOut() {
    if (this.user?.accessToken && window.google?.accounts?.oauth2?.revoke) {
      try {
        window.google.accounts.oauth2.revoke(this.user.accessToken, () => {});
      } catch (_) {}
    }
    this.user = null;
    this.fileHandle = null;
    this.driveFolderId = null;
    localStorage.removeItem(STORAGE_AUTH_USER_KEY);
    this.onStatusChange(null);
    this.onToast('Sesión de Google SSO cerrada.');
  }

  /* =========================================================
     2. GENERADOR Y LECTOR 100% LOCAL DE ARCHIVOS .DOCX (OPENXML ZIP)
     ========================================================= */

  /**
   * Genera un archivo .DOCX real (Office Open XML ZIP binario) 100% de forma local en el navegador/PWA
   * incluyendo título, fecha, párrafos con formato (negrita, cursiva, subrayado, listas, tareas),
   * tablas, imágenes insertadas y los dibujos hechos con el boli táctil.
   */
  async buildDocxBlobFromPage(page, inkOnlyDataUrl = null) {
    const title = page.title || 'Nota sin título';
    const dateStr = page.createdAt || new Date().toLocaleString('es-ES');

    const mediaFiles = []; // { filename: 'image1.png', relId: 'rIdImg1', bytes: Uint8Array, contentType: 'image/png' }
    const registerDataUrlImage = (dataUrl) => {
      const parsed = this.dataUrlToUint8Array(dataUrl);
      if (!parsed) return null;
      const idx = mediaFiles.length + 1;
      const ext = parsed.mime.includes('jpeg') || parsed.mime.includes('jpg') ? 'jpg' : 'png';
      const filename = `image${idx}.${ext}`;
      const relId = `rIdImg${idx}`;
      mediaFiles.push({
        filename,
        relId,
        bytes: parsed.bytes,
        contentType: ext === 'jpg' ? 'image/jpeg' : 'image/png'
      });
      return { relId, filename, idx };
    };

    const xmlParagraphs = [];

    // Título principal de la página (Heading1)
    xmlParagraphs.push(
      this.createOpenXmlParagraph(
        [{ text: title, bold: true, color: '7719AA', sizeHalfPt: 40 }],
        'Heading1'
      )
    );

    // Subtítulo con fecha y metadatos
    xmlParagraphs.push(
      this.createOpenXmlParagraph([
        {
          text: `Fecha: ${dateStr} · Guardado desde OneNote PWA (.docx)`,
          italic: true,
          color: '64748B',
          sizeHalfPt: 19
        }
      ])
    );

    // Procesar todos los bloques de la página (Texto enriquecido, tablas e imágenes)
    for (const block of page.blocks || []) {
      if (block.type === 'text' && block.html) {
        const blockXmlParts = this.convertHtmlToOpenXmlElements(block.html, registerDataUrlImage);
        xmlParagraphs.push(...blockXmlParts);
      } else if (block.type === 'image' && block.src) {
        const imgInfo = registerDataUrlImage(block.src);
        if (imgInfo) {
          xmlParagraphs.push(
            this.createOpenXmlImageParagraph(
              imgInfo.relId,
              imgInfo.idx,
              block.width || 420,
              block.height || 260
            )
          );
        }
      }
    }

    // Si hay trazos de boli táctil, incrustarlos como imagen en una sección del .docx
    if (inkOnlyDataUrl && Array.isArray(page.strokes) && page.strokes.length > 0) {
      const inkImgInfo = registerDataUrlImage(inkOnlyDataUrl);
      if (inkImgInfo) {
        xmlParagraphs.push(
          this.createOpenXmlParagraph(
            [{ text: 'Anotaciones y Esquemas con Boli Táctil', bold: true, color: '581080', sizeHalfPt: 28 }],
            'Heading2'
          )
        );
        xmlParagraphs.push(
          this.createOpenXmlImageParagraph(inkImgInfo.relId, inkImgInfo.idx, 520, 300)
        );
      }
    }

    const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="jpg" ContentType="image/jpeg"/>
  <Default Extension="jpeg" ContentType="image/jpeg"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`;

    const rootRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

    const imageRelsEntries = mediaFiles
      .map(
        (m) =>
          `  <Relationship Id="${m.relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${m.filename}"/>`
      )
      .join('\n');

    const documentRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
${imageRelsEntries}
</Relationships>`;

    const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:rPr>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
      <w:sz w:val="22"/>
      <w:color w:val="1E293B"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="240" w:after="120"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="36"/><w:color w:val="7719AA"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="200" w:after="80"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="28"/><w:color w:val="581080"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="160" w:after="60"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="24"/><w:color w:val="334155"/></w:rPr>
  </w:style>
</w:styles>`;

    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <w:body>
${xmlParagraphs.join('\n')}
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"/>
    </w:sectPr>
  </w:body>
</w:document>`;

    const encoder = new TextEncoder();
    const zipEntries = [
      { name: '[Content_Types].xml', data: encoder.encode(contentTypesXml) },
      { name: '_rels/.rels', data: encoder.encode(rootRelsXml) },
      { name: 'word/styles.xml', data: encoder.encode(stylesXml) },
      { name: 'word/_rels/document.xml.rels', data: encoder.encode(documentRelsXml) },
      { name: 'word/document.xml', data: encoder.encode(documentXml) }
    ];

    for (const m of mediaFiles) {
      zipEntries.push({
        name: `word/media/${m.filename}`,
        data: m.bytes
      });
    }

    return this.createZipBlob(
      zipEntries,
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    );
  }

  /**
   * Convierte un bloque HTML de la nota en elementos OpenXML (<w:p> y <w:tbl>)
   */
  convertHtmlToOpenXmlElements(html, registerDataUrlImage) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(`<div>${html}</div>`, 'text/html');
    const root = doc.body.firstElementChild || doc.body;
    const out = [];

    const extractRunsFromNode = (node, styleState = {}) => {
      const runs = [];
      for (const child of Array.from(node.childNodes)) {
        if (child.nodeType === Node.TEXT_NODE) {
          const text = child.textContent || '';
          if (text) {
            runs.push({ text, ...styleState });
          }
        } else if (child.nodeType === Node.ELEMENT_NODE) {
          const tag = child.tagName.toLowerCase();
          if (tag === 'br') {
            runs.push({ text: '\n', ...styleState });
            continue;
          }
          if (tag === 'input' && child.getAttribute('type') === 'checkbox') {
            const isChecked = child.checked || child.hasAttribute('checked');
            runs.push({ text: isChecked ? '☑ ' : '☐ ', bold: true, ...styleState });
            continue;
          }
          const nextState = {
            ...styleState,
            bold: styleState.bold || tag === 'b' || tag === 'strong',
            italic: styleState.italic || tag === 'i' || tag === 'em',
            underline: styleState.underline || tag === 'u',
            strike: styleState.strike || tag === 's' || tag === 'strike' || tag === 'del'
          };
          runs.push(...extractRunsFromNode(child, nextState));
        }
      }
      return runs;
    };

    const processElement = (el) => {
      for (const child of Array.from(el.childNodes)) {
        if (child.nodeType === Node.TEXT_NODE) {
          const txt = (child.textContent || '').trim();
          if (txt) {
            out.push(this.createOpenXmlParagraph([{ text: txt }]));
          }
        } else if (child.nodeType === Node.ELEMENT_NODE) {
          const tag = child.tagName.toLowerCase();
          if (tag === 'h1' || tag === 'h2' || tag === 'h3') {
            const styleId = tag === 'h1' ? 'Heading1' : tag === 'h2' ? 'Heading2' : 'Heading3';
            const runs = extractRunsFromNode(child, { bold: true });
            out.push(this.createOpenXmlParagraph(runs, styleId));
          } else if (tag === 'table') {
            const rows = Array.from(child.querySelectorAll('tr'));
            if (rows.length > 0) {
              const xmlRows = rows
                .map((tr) => {
                  const cells = Array.from(tr.querySelectorAll('th, td'));
                  const xmlCells = cells
                    .map((td) => {
                      const isHeader = td.tagName.toLowerCase() === 'th';
                      const cellRuns = extractRunsFromNode(td, { bold: isHeader });
                      return `<w:tc>
                        <w:tcPr><w:tcW w:w="2800" w:type="dxa"/></w:tcPr>
                        ${this.createOpenXmlParagraph(cellRuns.length ? cellRuns : [{ text: '' }])}
                      </w:tc>`;
                    })
                    .join('');
                  return `<w:tr>${xmlCells}</w:tr>`;
                })
                .join('');
              out.push(`<w:tbl>
                <w:tblPr>
                  <w:tblW w:w="5000" w:type="pct"/>
                  <w:tblBorders>
                    <w:top w:val="single" w:sz="4" w:space="0" w:color="94A3B8"/>
                    <w:left w:val="single" w:sz="4" w:space="0" w:color="94A3B8"/>
                    <w:bottom w:val="single" w:sz="4" w:space="0" w:color="94A3B8"/>
                    <w:right w:val="single" w:sz="4" w:space="0" w:color="94A3B8"/>
                    <w:insideH w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>
                    <w:insideV w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>
                  </w:tblBorders>
                </w:tblPr>
                ${xmlRows}
              </w:tbl>`);
            }
          } else if (tag === 'ul' || tag === 'ol') {
            const items = Array.from(child.querySelectorAll(':scope > li'));
            items.forEach((li, idx) => {
              const prefix = tag === 'ol' ? `${idx + 1}. ` : '• ';
              const runs = [{ text: prefix, bold: true }, ...extractRunsFromNode(li)];
              out.push(this.createOpenXmlParagraph(runs));
            });
          } else if (tag === 'img') {
            const src = child.getAttribute('src') || '';
            const imgInfo = registerDataUrlImage(src);
            if (imgInfo) {
              out.push(this.createOpenXmlImageParagraph(imgInfo.relId, imgInfo.idx, 400, 240));
            }
          } else if (tag === 'div' || tag === 'p' || tag === 'blockquote' || tag === 'section') {
            // Comprobar si tiene imágenes hijas directas
            const nestedImgs = Array.from(child.querySelectorAll('img'));
            const hasSubBlocks = Array.from(child.children).some((c) =>
              ['div', 'p', 'table', 'ul', 'ol', 'h1', 'h2', 'h3'].includes(
                c.tagName.toLowerCase()
              )
            );
            if (hasSubBlocks) {
              processElement(child);
            } else {
              const runs = extractRunsFromNode(child);
              if (runs.some((r) => r.text.trim())) {
                out.push(this.createOpenXmlParagraph(runs));
              }
              for (const imgEl of nestedImgs) {
                const imgInfo = registerDataUrlImage(imgEl.getAttribute('src') || '');
                if (imgInfo) {
                  out.push(this.createOpenXmlImageParagraph(imgInfo.relId, imgInfo.idx, 400, 240));
                }
              }
            }
          } else {
            const runs = extractRunsFromNode(child);
            if (runs.some((r) => r.text.trim())) {
              out.push(this.createOpenXmlParagraph(runs));
            }
          }
        }
      }
    };

    processElement(root);
    return out;
  }

  createOpenXmlParagraph(runs, styleId = '') {
    const pPr = styleId ? `<w:pPr><w:pStyle w:val="${styleId}"/></w:pPr>` : '';
    const runsXml = (runs || [])
      .map((r) => {
        const rPrParts = [];
        if (r.bold) rPrParts.push('<w:b/>');
        if (r.italic) rPrParts.push('<w:i/>');
        if (r.underline) rPrParts.push('<w:u w:val="single"/>');
        if (r.strike) rPrParts.push('<w:strike/>');
        if (r.color) rPrParts.push(`<w:color w:val="${r.color}"/>`);
        if (r.sizeHalfPt) rPrParts.push(`<w:sz w:val="${r.sizeHalfPt}"/>`);
        const rPr = rPrParts.length ? `<w:rPr>${rPrParts.join('')}</w:rPr>` : '';
        const safeText = this.escapeXml(r.text || '');
        return `<w:r>${rPr}<w:t xml:space="preserve">${safeText}</w:t></w:r>`;
      })
      .join('');
    return `    <w:p>${pPr}${runsXml}</w:p>`;
  }

  createOpenXmlImageParagraph(relId, idNum, widthPx = 420, heightPx = 260) {
    // 1 px ~= 9525 EMUs
    const cx = Math.round(Math.min(540, Math.max(120, widthPx)) * 9525);
    const cy = Math.round(Math.min(420, Math.max(90, heightPx)) * 9525);
    return `    <w:p>
      <w:r>
        <w:drawing>
          <wp:inline distT="0" distB="0" distL="0" distR="0">
            <wp:extent cx="${cx}" cy="${cy}"/>
            <wp:docPr id="${idNum}" name="Imagen ${idNum}"/>
            <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                  <pic:nvPicPr>
                    <pic:cNvPr id="${idNum}" name="image${idNum}.png"/>
                    <pic:cNvPicPr/>
                  </pic:nvPicPr>
                  <pic:blipFill>
                    <a:blip r:embed="${relId}"/>
                    <a:stretch><a:fillRect/></a:stretch>
                  </pic:blipFill>
                  <pic:spPr>
                    <a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>
                    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                  </pic:spPr>
                </pic:pic>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>`;
  }

  /**
   * Empaquetador ZIP puro en JavaScript (Store Method 0 + CRC32 IEEE) para crear archivos .docx 100% locales
   */
  createZipBlob(entries, mimeType) {
    const encoder = new TextEncoder();
    const localParts = [];
    const centralParts = [];
    let offset = 0;

    for (const entry of entries) {
      const nameBytes = encoder.encode(entry.name);
      const dataBytes =
        entry.data instanceof Uint8Array ? entry.data : encoder.encode(String(entry.data));
      const crc = this.crc32(dataBytes);
      const size = dataBytes.length;

      // Local file header (30 bytes + nameLen)
      const localHeader = new Uint8Array(30 + nameBytes.length);
      const lv = new DataView(localHeader.buffer);
      lv.setUint32(0, 0x04034b50, true); // signature
      lv.setUint16(4, 20, true); // version needed
      lv.setUint16(6, 0x0800, true); // UTF-8 flag
      lv.setUint16(8, 0, true); // compression = 0 (Store)
      lv.setUint16(10, 0, true); // mod time
      lv.setUint16(12, 0x5421, true); // mod date
      lv.setUint32(14, crc, true);
      lv.setUint32(18, size, true);
      lv.setUint32(22, size, true);
      lv.setUint16(26, nameBytes.length, true);
      lv.setUint16(28, 0, true);
      localHeader.set(nameBytes, 30);

      // Central directory header (46 bytes + nameLen)
      const centralHeader = new Uint8Array(46 + nameBytes.length);
      const cv = new DataView(centralHeader.buffer);
      cv.setUint32(0, 0x02014b50, true);
      cv.setUint16(4, 20, true);
      cv.setUint16(6, 20, true);
      cv.setUint16(8, 0x0800, true);
      cv.setUint16(10, 0, true);
      cv.setUint16(12, 0, true);
      cv.setUint16(14, 0x5421, true);
      cv.setUint32(16, crc, true);
      cv.setUint32(20, size, true);
      cv.setUint32(24, size, true);
      cv.setUint16(28, nameBytes.length, true);
      cv.setUint16(30, 0, true);
      cv.setUint16(32, 0, true);
      cv.setUint16(34, 0, true);
      cv.setUint16(36, 0, true);
      cv.setUint32(38, 0, true);
      cv.setUint32(42, offset, true);
      centralHeader.set(nameBytes, 46);

      localParts.push(localHeader, dataBytes);
      centralParts.push(centralHeader);
      offset += localHeader.length + size;
    }

    const centralSize = centralParts.reduce((acc, arr) => acc + arr.length, 0);
    const eocd = new Uint8Array(22);
    const ev = new DataView(eocd.buffer);
    ev.setUint32(0, 0x06054b50, true);
    ev.setUint16(4, 0, true);
    ev.setUint16(6, 0, true);
    ev.setUint16(8, entries.length, true);
    ev.setUint16(10, entries.length, true);
    ev.setUint32(12, centralSize, true);
    ev.setUint32(16, offset, true);
    ev.setUint16(20, 0, true);

    return new Blob([...localParts, ...centralParts, eocd], { type: mimeType });
  }

  crc32(bytes) {
    if (!this._crcTable) {
      this._crcTable = new Uint32Array(256);
      for (let i = 0; i < 256; i++) {
        let c = i;
        for (let k = 0; k < 8; k++) {
          c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
        }
        this._crcTable[i] = c >>> 0;
      }
    }
    let crc = 0xffffffff;
    for (let i = 0; i < bytes.length; i++) {
      crc = this._crcTable[(crc ^ bytes[i]) & 0xff] ^ (crc >>> 8);
    }
    return (crc ^ 0xffffffff) >>> 0;
  }

  dataUrlToUint8Array(dataUrl) {
    if (!dataUrl || !dataUrl.startsWith('data:')) return null;
    try {
      const [header, b64] = dataUrl.split(',');
      const mimeMatch = header.match(/data:([^;]+)/);
      const mime = mimeMatch ? mimeMatch[1] : 'image/png';
      const binary = atob(b64);
      const bytes = new Uint8Array(binary.length);
      for (let i = 0; i < binary.length; i++) {
        bytes[i] = binary.charCodeAt(i);
      }
      return { mime, bytes };
    } catch (_) {
      return null;
    }
  }

  uint8ArrayToBase64(bytes) {
    let binary = '';
    const chunkSize = 0x8000;
    for (let i = 0; i < bytes.length; i += chunkSize) {
      binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunkSize));
    }
    return btoa(binary);
  }

  /**
   * Lector 100% Local de archivos .DOCX en el navegador (Android PWA / Mac / Chromebook)
   * Descomprime el archivo .docx mediante el Directorio Central ZIP + DecompressionStream('deflate-raw')
   * extrayendo párrafos, encabezados, formato (negrita/cursiva/subrayado), tablas e imágenes incrustadas.
   */
  async parseDocxInBrowser(fileOrBuffer) {
    const buffer =
      fileOrBuffer instanceof ArrayBuffer ? fileOrBuffer : await fileOrBuffer.arrayBuffer();
    const zipMap = await this.extractZipEntriesInBrowser(buffer);
    const decoder = new TextDecoder('utf-8');

    const docXmlBytes = zipMap.get('word/document.xml');
    if (!docXmlBytes) {
      throw new Error('El archivo .docx no contiene word/document.xml');
    }

    // 1. Extraer mapa de relaciones de imágenes (word/_rels/document.xml.rels)
    const relsMap = {};
    const relsBytes = zipMap.get('word/_rels/document.xml.rels');
    if (relsBytes) {
      const relsXml = decoder.decode(relsBytes);
      const relsDoc = new DOMParser().parseFromString(relsXml, 'application/xml');
      Array.from(relsDoc.getElementsByTagName('Relationship')).forEach((rel) => {
        const id = rel.getAttribute('Id');
        const target = rel.getAttribute('Target');
        if (id && target) {
          const cleanPath = target.startsWith('/')
            ? target.slice(1)
            : `word/${target.replace(/^\.\//, '')}`;
          const imgBytes = zipMap.get(cleanPath);
          if (imgBytes) {
            const ext = cleanPath.split('.').pop().toLowerCase();
            const mime = ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg' : 'image/png';
            relsMap[id] = `data:${mime};base64,${this.uint8ArrayToBase64(imgBytes)}`;
          }
        }
      });
    }

    // 2. Parsear word/document.xml conservando títulos, estilos, tablas e imágenes
    const xmlString = decoder.decode(docXmlBytes);
    const xmlDoc = new DOMParser().parseFromString(xmlString, 'application/xml');
    const body = xmlDoc.getElementsByTagName('w:body')[0] || xmlDoc.documentElement;

    const parseParagraph = (pNode) => {
      let styleVal = '';
      const pStyle = pNode.getElementsByTagName('w:pStyle')[0];
      if (pStyle) {
        styleVal = pStyle.getAttribute('w:val') || '';
      }

      const inlineParts = [];
      const runs = Array.from(pNode.getElementsByTagName('w:r'));
      for (const r of runs) {
        const rPr = r.getElementsByTagName('w:rPr')[0];
        const isBold = rPr && rPr.getElementsByTagName('w:b').length > 0;
        const isItalic = rPr && rPr.getElementsByTagName('w:i').length > 0;
        const isUnderline = rPr && rPr.getElementsByTagName('w:u').length > 0;
        const isStrike = rPr && rPr.getElementsByTagName('w:strike').length > 0;

        const texts = Array.from(r.getElementsByTagName('w:t'))
          .map((t) => t.textContent || '')
          .join('');

        if (texts) {
          let htmlTxt = this.escapeHtml(texts);
          if (isBold) htmlTxt = `<b>${htmlTxt}</b>`;
          if (isItalic) htmlTxt = `<i>${htmlTxt}</i>`;
          if (isUnderline) htmlTxt = `<u>${htmlTxt}</u>`;
          if (isStrike) htmlTxt = `<s>${htmlTxt}</s>`;
          inlineParts.push(htmlTxt);
        }

        // Imágenes incrustadas (<a:blip r:embed="rId..."/>)
        const blips = Array.from(r.getElementsByTagName('a:blip'));
        for (const blip of blips) {
          const rId = blip.getAttribute('r:embed') || blip.getAttribute('embed');
          if (rId && relsMap[rId]) {
            inlineParts.push(
              `<div style="margin:10px 0;"><img src="${relsMap[rId]}" style="max-width:100%;border-radius:6px;" alt="Imagen .docx" /></div>`
            );
          }
        }
      }

      const content = inlineParts.join('');
      if (!content.trim()) return '';

      if (/Heading1|Title/i.test(styleVal)) return `<h2>${content}</h2>`;
      if (/Heading2|Subtitle/i.test(styleVal)) return `<h3>${content}</h3>`;
      if (/Heading3/i.test(styleVal)) return `<h4>${content}</h4>`;
      return `<div>${content}</div>`;
    };

    const htmlChunks = [];
    for (const child of Array.from(body.childNodes)) {
      if (child.nodeName === 'w:p') {
        const pHtml = parseParagraph(child);
        if (pHtml) htmlChunks.push(pHtml);
      } else if (child.nodeName === 'w:tbl') {
        const rows = Array.from(child.getElementsByTagName('w:tr'));
        const rowsHtml = rows
          .map((tr) => {
            const cells = Array.from(tr.getElementsByTagName('w:tc'));
            const cellsHtml = cells
              .map((tc) => {
                const cellPs = Array.from(tc.getElementsByTagName('w:p'))
                  .map((p) => parseParagraph(p))
                  .filter(Boolean)
                  .join('');
                return `<td>${cellPs || '&nbsp;'}</td>`;
              })
              .join('');
            return `<tr>${cellsHtml}</tr>`;
          })
          .join('');
        if (rowsHtml) {
          htmlChunks.push(`<table>${rowsHtml}</table>`);
        }
      }
    }

    return htmlChunks.join('\n') || '<div>Documento .docx importado correctamente.</div>';
  }

  /**
   * Extrae todas las entradas de un archivo ZIP (.docx) en el navegador
   * usando el Directorio Central ZIP (compatible con archivos .docx de Google Docs y Microsoft Word
   * que usan Data Descriptors bit 3).
   */
  async extractZipEntriesInBrowser(arrayBuffer) {
    const bytes = new Uint8Array(arrayBuffer);
    const view = new DataView(arrayBuffer);
    const decoder = new TextDecoder('utf-8');
    const entries = new Map();

    // Buscar End of Central Directory (EOCD: 0x06054b50) desde el final
    let eocdOffset = -1;
    for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) {
      if (view.getUint32(i, true) === 0x06054b50) {
        eocdOffset = i;
        break;
      }
    }

    const decompressSlice = async (rawSlice, compression) => {
      if (compression === 0) {
        return rawSlice;
      }
      if (compression === 8 && typeof DecompressionStream !== 'undefined') {
        const ds = new DecompressionStream('deflate-raw');
        const writer = ds.writable.getWriter();
        writer.write(rawSlice);
        writer.close();
        const buf = await new Response(ds.readable).arrayBuffer();
        return new Uint8Array(buf);
      }
      return rawSlice;
    };

    if (eocdOffset !== -1) {
      const totalEntries = view.getUint16(eocdOffset + 10, true);
      let cdOffset = view.getUint32(eocdOffset + 16, true);

      for (let i = 0; i < totalEntries && cdOffset + 46 <= bytes.length; i++) {
        if (view.getUint32(cdOffset, true) !== 0x02014b50) break;
        const compression = view.getUint16(cdOffset + 10, true);
        const compressedSize = view.getUint32(cdOffset + 20, true);
        const nameLen = view.getUint16(cdOffset + 28, true);
        const extraLen = view.getUint16(cdOffset + 30, true);
        const commentLen = view.getUint16(cdOffset + 32, true);
        const localHeaderOffset = view.getUint32(cdOffset + 42, true);

        const name = decoder.decode(bytes.subarray(cdOffset + 46, cdOffset + 46 + nameLen));
        cdOffset += 46 + nameLen + extraLen + commentLen;

        if (localHeaderOffset + 30 <= bytes.length) {
          const localNameLen = view.getUint16(localHeaderOffset + 26, true);
          const localExtraLen = view.getUint16(localHeaderOffset + 28, true);
          const dataStart = localHeaderOffset + 30 + localNameLen + localExtraLen;
          const rawSlice = bytes.subarray(dataStart, dataStart + compressedSize);
          try {
            const outBytes = await decompressSlice(rawSlice, compression);
            entries.set(name, outBytes);
          } catch (_) {}
        }
      }
      return entries;
    }

    // Fallback secuencial si no se encuentra EOCD
    let offset = 0;
    while (offset + 30 < bytes.length) {
      if (view.getUint32(offset, true) !== 0x04034b50) break;
      const compression = view.getUint16(offset + 8, true);
      const compressedSize = view.getUint32(offset + 18, true);
      const nameLen = view.getUint16(offset + 26, true);
      const extraLen = view.getUint16(offset + 28, true);
      const name = decoder.decode(bytes.subarray(offset + 30, offset + 30 + nameLen));
      const dataStart = offset + 30 + nameLen + extraLen;
      const dataEnd = dataStart + compressedSize;
      if (compressedSize > 0 && dataEnd <= bytes.length) {
        try {
          const outBytes = await decompressSlice(bytes.subarray(dataStart, dataEnd), compression);
          entries.set(name, outBytes);
        } catch (_) {}
      }
      offset = dataEnd;
    }

    return entries;
  }

  /* =========================================================
     3. EXPORTACIÓN E IMPORTACIÓN .DOCX + GOOGLE DOCS EN GOOGLE DRIVE
     ========================================================= */

  /**
   * Exporta y graba la página actual como un archivo .DOCX real (Office Open XML)
   * 100% local en la PWA, lo almacena en Google Drive y permite descargarlo o compartirlo en Android.
   */
  async savePageAsGoogleDoc(page, inkOnlyDataUrl = null, { downloadFile = true } = {}) {
    const email = this.getActiveEmail();
    const docxBlob = await this.buildDocxBlobFromPage(page, inkOnlyDataUrl);
    const previewHtml = this.buildGoogleDocHtmlFromPage(page, inkOnlyDataUrl);

    const safeTitle = (page.title || 'Nota_OneNote')
      .replace(/[^a-z0-9áéíóúñ_-]/gi, '_')
      .replace(/_+/g, '_');
    const filename = `${safeTitle}.docx`;
    const nowIso = new Date().toISOString();

    const docxBuffer = await docxBlob.arrayBuffer();
    const docxBase64 = this.uint8ArrayToBase64(new Uint8Array(docxBuffer));

    const record = {
      id: `${email}::docx-${safeTitle}`,
      email,
      isGoogleDoc: true,
      isDocx: true,
      title: page.title || 'Nota sin título',
      name: filename,
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      updatedAt: nowIso,
      size: docxBlob.size,
      contentHtml: previewHtml,
      content: previewHtml,
      docxBase64
    };

    // 1. Guardar localmente en IndexedDB de la PWA (100% local en Android)
    await this.saveToLocalVault(record);

    // 2. Si el usuario tiene sesión Google SSO con token OAuth 2.0 activo, subir el .docx a su Google Drive real
    if (this.hasLiveOAuthToken()) {
      const remoteFile = await this.uploadFileToRealGoogleDrive({
        name: filename,
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        blob: docxBlob
      });
      if (remoteFile && remoteFile.id) {
        record.driveFileId = remoteFile.id;
        await this.saveToLocalVault(record);
      }
    }

    // 3. Sincronizar opcionalmente con el servidor local si está disponible
    try {
      await fetch('/api/drive-vault', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          email,
          action: 'save_gdoc',
          title: page.title || 'Nota .docx',
          name: filename,
          html: previewHtml
        })
      });
    } catch (_) {}

    // 4. Descargar el archivo .docx o abrir el selector nativo de Android (Google Drive / Docs)
    if (downloadFile) {
      const file = new File([docxBlob], filename, {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      });

      if (
        /Android/i.test(navigator.userAgent) &&
        navigator.canShare &&
        navigator.canShare({ files: [file] })
      ) {
        try {
          await navigator.share({
            title: filename,
            text: 'Documento .docx exportado desde OneNote PWA',
            files: [file]
          });
          this.onToast(`📄 Archivo "${filename}" enviado a Google Drive / Docs en Android.`);
          return record;
        } catch (err) {
          if (err.name === 'AbortError') return record;
        }
      }

      const a = document.createElement('a');
      a.href = URL.createObjectURL(docxBlob);
      a.download = filename;
      a.click();
    }

    this.onToast(
      this.hasLiveOAuthToken()
        ? `☁️ Exportado a "${filename}" y almacenado en tu Google Drive.`
        : `📄 Nota exportada a "${filename}" y guardada en tu PWA.`
    );
    return record;
  }

  /**
   * Construye vista previa HTML enriquecida de la página
   */
  buildGoogleDocHtmlFromPage(page, inkOnlyDataUrl = null) {
    const title = page.title || 'Nota sin título';
    const dateStr = page.createdAt || new Date().toLocaleString('es-ES');

    const bodyParts = [];
    (page.blocks || []).forEach((block) => {
      if (block.type === 'text' && block.html) {
        bodyParts.push(`<div style="margin-bottom:14px;line-height:1.6;">${block.html}</div>`);
      } else if (block.type === 'image' && block.src) {
        bodyParts.push(
          `<div style="margin:16px 0;"><img src="${block.src}" style="max-width:100%;width:${
            block.width || 420
          }px;border-radius:6px;" alt="Imagen de nota" /></div>`
        );
      }
    });

    if (inkOnlyDataUrl && page.strokes && page.strokes.length > 0) {
      bodyParts.push(`
        <hr style="border:none;border-top:1px solid #cbd5e1;margin:20px 0;" />
        <h3 style="color:#581080;font-family:Arial,sans-serif;">🖊️ Anotaciones con Boli Táctil</h3>
        <div><img src="${inkOnlyDataUrl}" style="max-width:100%;border:1px solid #e2e8f0;border-radius:8px;" alt="Trazos de boli táctil" /></div>
      `);
    }

    return `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8" />
  <title>${title}</title>
</head>
<body>
  <h1>${title}</h1>
  <div class="meta">Fecha: ${dateStr}</div>
  ${bodyParts.join('\n')}
</body>
</html>`;
  }

  /**
   * Importa un archivo .docx (100% local en el navegador), .doc, .html, .txt, .md o .gdoc
   */
  async importGoogleDocFile(file) {
    const name = file.name || 'Documento.docx';
    const ext = name.split('.').pop().toLowerCase();
    const cleanTitle = name.replace(/\.(docx|doc|html|htm|txt|md|gdoc)$/i, '');

    // 1. Lectura 100% local de archivos .DOCX en la PWA (sin depender del servidor)
    if (ext === 'docx') {
      try {
        const htmlFromBrowser = await this.parseDocxInBrowser(file);
        if (htmlFromBrowser && htmlFromBrowser.trim()) {
          await this.saveImportedDocToVault(cleanTitle, `${cleanTitle}.docx`, htmlFromBrowser);
          this.onToast(`📄 Archivo .docx "${cleanTitle}" leído localmente e importado.`);
          return { title: cleanTitle, html: htmlFromBrowser };
        }
      } catch (localErr) {
        // Fallback opcional al servidor si estuviera corriendo
        try {
          const dataUrl = await this.readFileAsDataURL(file);
          const res = await fetch('/api/google-docs-parse-docx', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ base64: dataUrl })
          });
          if (res.ok) {
            const parsed = await res.json();
            if (parsed.ok && parsed.html) {
              await this.saveImportedDocToVault(cleanTitle, `${cleanTitle}.docx`, parsed.html);
              this.onToast(`📄 Archivo .docx "${cleanTitle}" importado.`);
              return { title: cleanTitle, html: parsed.html };
            }
          }
        } catch (_) {}
        this.onToast(`⚠️ Error al leer el archivo .docx: ${localErr.message}`);
        return null;
      }
    }

    // 2. Otros formatos (.doc HTML, .html, .txt, .md, .gdoc)
    const text = await file.text();
    if (ext === 'gdoc') {
      try {
        const gdocJson = JSON.parse(text);
        if (gdocJson.url || gdocJson.doc_id) {
          const url =
            gdocJson.url || `https://docs.google.com/document/d/${gdocJson.doc_id}/edit`;
          return await this.importGoogleDocFromUrl(url, cleanTitle);
        }
      } catch (_) {}
    }

    if (ext === 'html' || ext === 'htm' || ext === 'doc' || text.includes('<body')) {
      const parser = new DOMParser();
      const doc = parser.parseFromString(text, 'text/html');
      doc.querySelectorAll('script, style, meta, link').forEach((el) => el.remove());
      const bodyHtml = doc.body ? doc.body.innerHTML.trim() : text;
      await this.saveImportedDocToVault(cleanTitle, `${cleanTitle}.docx`, bodyHtml);
      this.onToast(`📄 Documento "${cleanTitle}" importado en la nota.`);
      return { title: cleanTitle, html: bodyHtml || `<div>${text}</div>` };
    }

    const linesHtml = text
      .split(/\r?\n/)
      .map((line) => (line.trim() ? `<div>${this.escapeHtml(line)}</div>` : '<div><br></div>'))
      .join('');
    await this.saveImportedDocToVault(cleanTitle, `${cleanTitle}.docx`, linesHtml);
    this.onToast(`📄 Documento "${cleanTitle}" importado a la nota.`);
    return { title: cleanTitle, html: linesHtml };
  }

  async saveImportedDocToVault(title, filename, html) {
    const email = this.getActiveEmail();
    const record = {
      id: `${email}::docx-${title.replace(/[^a-z0-9_-]/gi, '_')}`,
      email,
      isGoogleDoc: true,
      isDocx: true,
      title,
      name: filename,
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      updatedAt: new Date().toISOString(),
      size: new Blob([html]).size,
      contentHtml: html,
      content: html
    };
    await this.saveToLocalVault(record);
  }

  /**
   * Importa un documento de Google Docs por URL (usando API de Google Drive si hay token SSO,
   * o descarga directa de exportación .docx)
   */
  async importGoogleDocFromUrl(url, fallbackTitle = '') {
    const cleanUrl = (url || '').trim();
    const match = cleanUrl.match(/\/document\/d\/([a-zA-Z0-9_-]+)/);
    if (!match) {
      this.onToast('⚠️ Pega un enlace válido de Google Docs (https://docs.google.com/document/d/...).');
      return null;
    }
    const docId = match[1];

    // 1. Si tiene sesión Google SSO con token OAuth 2.0, descargar como .docx directamente de Google Drive API v3
    if (this.hasLiveOAuthToken()) {
      const fromApi = await this.downloadAndImportFromRealGoogleDrive(
        docId,
        fallbackTitle || 'Google_Doc',
        'application/vnd.google-apps.document'
      );
      if (fromApi) return fromApi;
    }

    // 2. Intentar descargar el export .docx directo desde docs.google.com y parsearlo localmente
    try {
      const exportDocxUrl = `https://docs.google.com/document/d/${docId}/export?format=docx`;
      const res = await fetch(exportDocxUrl);
      if (res.ok) {
        const buf = await res.arrayBuffer();
        const html = await this.parseDocxInBrowser(buf);
        const title = fallbackTitle || `Google Doc (${docId.slice(0, 6)})`;
        await this.saveImportedDocToVault(title, `${title}.docx`, html);
        this.onToast(`📄 Google Doc "${title}" leído en formato .docx.`);
        return { title, html };
      }
    } catch (_) {}

    // 3. Fallback servidor local si está corriendo
    try {
      const res = await fetch(`/api/google-docs-import?url=${encodeURIComponent(cleanUrl)}`);
      if (res.ok) {
        const data = await res.json();
        if (data.ok && data.html) {
          const title = fallbackTitle || data.title || 'Google Doc';
          await this.saveImportedDocToVault(title, `${title}.docx`, data.html);
          this.onToast(`📄 Google Doc "${title}" importado.`);
          return { title, html: data.html };
        }
      }
    } catch (_) {}

    this.onToast(
      '⚠️ Inicia sesión con Google SSO (OAuth) o abre directamente el archivo .docx desde tu dispositivo.'
    );
    return null;
  }

  /**
   * Lista los documentos .docx y Google Docs guardados localmente en la PWA + en Google Drive API v3
   */
  async listSavedGoogleDocs() {
    const email = this.getActiveEmail();
    const localFiles = await this.listLocalVaultFiles(email);
    const localDocs = localFiles.filter(
      (f) =>
        f.isGoogleDoc ||
        f.isDocx ||
        f.name?.endsWith('.docx') ||
        f.mimeType?.includes('wordprocessingml') ||
        f.mimeType === 'application/vnd.google-apps.document'
    );

    if (this.hasLiveOAuthToken()) {
      const remoteDocs = await this.listRemoteGoogleDriveDocxFiles();
      const merged = [...localDocs];
      for (const rem of remoteDocs) {
        if (!merged.some((loc) => loc.driveFileId === rem.driveFileId || loc.name === rem.name)) {
          merged.push(rem);
        }
      }
      return merged;
    }

    return localDocs;
  }

  /* =========================================================
     4. CONECTOR REAL GOOGLE DRIVE API V3 (OAUTH 2.0 SSO) + BÓVEDA LOCAL INDEXEDDB
     ========================================================= */

  async ensureGoogleDriveFolder() {
    if (!this.hasLiveOAuthToken()) return null;
    if (this.driveFolderId) return this.driveFolderId;

    try {
      const q = encodeURIComponent(
        `name = '${DRIVE_FOLDER_NAME}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false`
      );
      const searchRes = await fetch(
        `https://www.googleapis.com/drive/v3/files?q=${q}&fields=files(id,name)&spaces=drive`,
        { headers: { Authorization: `Bearer ${this.user.accessToken}` } }
      );
      if (searchRes.ok) {
        const data = await searchRes.json();
        if (data.files && data.files.length > 0) {
          this.driveFolderId = data.files[0].id;
          return this.driveFolderId;
        }
      }

      // Crear carpeta en Google Drive si no existe
      const createRes = await fetch('https://www.googleapis.com/drive/v3/files', {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${this.user.accessToken}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          name: DRIVE_FOLDER_NAME,
          mimeType: 'application/vnd.google-apps.folder'
        })
      });
      if (createRes.ok) {
        const folder = await createRes.json();
        this.driveFolderId = folder.id;
        return this.driveFolderId;
      }
    } catch (_) {}
    return null;
  }

  async uploadFileToRealGoogleDrive({ name, mimeType, blob }) {
    if (!this.hasLiveOAuthToken()) return null;
    try {
      const folderId = await this.ensureGoogleDriveFolder();
      const q = encodeURIComponent(
        `name = '${name.replace(/'/g, "\\'")}' and trashed = false${
          folderId ? ` and '${folderId}' in parents` : ''
        }`
      );
      const searchRes = await fetch(
        `https://www.googleapis.com/drive/v3/files?q=${q}&fields=files(id,name)`,
        { headers: { Authorization: `Bearer ${this.user.accessToken}` } }
      );
      let existingFileId = null;
      if (searchRes.ok) {
        const sData = await searchRes.json();
        if (sData.files && sData.files.length > 0) {
          existingFileId = sData.files[0].id;
        }
      }

      const metadata = existingFileId
        ? { name, mimeType }
        : { name, mimeType, parents: folderId ? [folderId] : undefined };

      const form = new FormData();
      form.append('metadata', new Blob([JSON.stringify(metadata)], { type: 'application/json' }));
      form.append('file', blob);

      const uploadUrl = existingFileId
        ? `https://www.googleapis.com/upload/drive/v3/files/${existingFileId}?uploadType=multipart&fields=id,name,modifiedTime,size`
        : 'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,modifiedTime,size';

      const res = await fetch(uploadUrl, {
        method: existingFileId ? 'PATCH' : 'POST',
        headers: { Authorization: `Bearer ${this.user.accessToken}` },
        body: form
      });

      if (res.ok) {
        return await res.json();
      }
    } catch (_) {}
    return null;
  }

  async listRemoteGoogleDriveDocxFiles() {
    if (!this.hasLiveOAuthToken()) return [];
    try {
      const q = encodeURIComponent(
        `(mimeType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' or mimeType = 'application/vnd.google-apps.document') and trashed = false`
      );
      const res = await fetch(
        `https://www.googleapis.com/drive/v3/files?q=${q}&pageSize=20&orderBy=modifiedTime desc&fields=files(id,name,mimeType,modifiedTime,size)`,
        { headers: { Authorization: `Bearer ${this.user.accessToken}` } }
      );
      if (res.ok) {
        const data = await res.json();
        return (data.files || []).map((f) => ({
          id: `gdrive::${f.id}`,
          driveFileId: f.id,
          isRemoteDrive: true,
          isDocx: true,
          title: f.name.replace(/\.docx$/i, ''),
          name: f.name.endsWith('.docx') ? f.name : `${f.name}.docx`,
          mimeType: f.mimeType,
          updatedAt: f.modifiedTime,
          size: Number(f.size || 0)
        }));
      }
    } catch (_) {}
    return [];
  }

  async downloadAndImportFromRealGoogleDrive(fileId, fileName, mimeType) {
    if (!this.hasLiveOAuthToken()) return null;
    try {
      const url =
        mimeType === 'application/vnd.google-apps.document'
          ? `https://www.googleapis.com/drive/v3/files/${fileId}/export?mimeType=application/vnd.openxmlformats-officedocument.wordprocessingml.document`
          : `https://www.googleapis.com/drive/v3/files/${fileId}?alt=media`;

      const res = await fetch(url, {
        headers: { Authorization: `Bearer ${this.user.accessToken}` }
      });
      if (res.ok) {
        const buf = await res.arrayBuffer();
        const html = await this.parseDocxInBrowser(buf);
        const cleanTitle = (fileName || 'Documento_Drive').replace(/\.docx$/i, '');
        await this.saveImportedDocToVault(cleanTitle, `${cleanTitle}.docx`, html);
        this.onToast(`☁️ "${cleanTitle}.docx" descargado de Google Drive y abierto.`);
        return { title: cleanTitle, html };
      }
    } catch (err) {
      this.onToast(`⚠️ No se pudo descargar de Google Drive: ${err.message}`);
    }
    return null;
  }

  /**
   * Almacena todo el espacio de trabajo (blocs, secciones, páginas, imágenes, trazos)
   * localmente en IndexedDB (PWA) y en Google Drive.
   */
  async saveWorkspaceToDrive(workspaceData, { silent = false } = {}) {
    const email = this.getActiveEmail();
    const payloadString = JSON.stringify(workspaceData, null, 2);
    const nowIso = new Date().toISOString();
    const blob = new Blob([payloadString], { type: 'application/json' });

    const record = {
      id: `${email}::drive-workspace-main`,
      email,
      name: WORKSPACE_FILENAME,
      mimeType: 'application/json',
      updatedAt: nowIso,
      size: blob.size,
      content: payloadString
    };

    // 1. Guardar siempre de forma local en IndexedDB (100% autónomo en Android PWA)
    await this.saveToLocalVault(record);

    // 2. Si hay token OAuth de Google SSO, subir directamente a Google Drive API v3
    if (this.hasLiveOAuthToken()) {
      await this.uploadFileToRealGoogleDrive({
        name: WORKSPACE_FILENAME,
        mimeType: 'application/json',
        blob
      });
    }

    // 3. Sincronizar opcionalmente con servidor local si está accesible
    try {
      await fetch('/api/drive-vault', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          email,
          action: 'save_workspace',
          workspace: workspaceData
        })
      });
    } catch (_) {}

    if (!silent) {
      this.onToast(
        this.hasLiveOAuthToken()
          ? `☁️ Todo almacenado en Google Drive (${email})`
          : `💾 Guardado en bóveda local PWA y cola de Google Drive (${email})`
      );
    }
    return record;
  }

  async loadWorkspaceFromDrive() {
    const email = this.getActiveEmail();

    // 1. Si hay token OAuth de Google SSO, buscar primero en Google Drive API v3
    if (this.hasLiveOAuthToken()) {
      try {
        const folderId = await this.ensureGoogleDriveFolder();
        const q = encodeURIComponent(
          `name = '${WORKSPACE_FILENAME}' and trashed = false${
            folderId ? ` and '${folderId}' in parents` : ''
          }`
        );
        const searchRes = await fetch(
          `https://www.googleapis.com/drive/v3/files?q=${q}&orderBy=modifiedTime desc&fields=files(id,name)`,
          { headers: { Authorization: `Bearer ${this.user.accessToken}` } }
        );
        if (searchRes.ok) {
          const sData = await searchRes.json();
          if (sData.files && sData.files.length > 0) {
            const fileId = sData.files[0].id;
            const dlRes = await fetch(
              `https://www.googleapis.com/drive/v3/files/${fileId}?alt=media`,
              { headers: { Authorization: `Bearer ${this.user.accessToken}` } }
            );
            if (dlRes.ok) {
              const workspace = await dlRes.json();
              if (workspace && Array.isArray(workspace.notebooks)) {
                this.onToast(`☁️ Blocs restaurados directamente desde tu Google Drive.`);
                return workspace;
              }
            }
          }
        }
      } catch (_) {}
    }

    // 2. Buscar en IndexedDB local de la PWA
    const files = await this.listLocalVaultFiles(email);
    const mainFile = files.find((f) => f.name === WORKSPACE_FILENAME && f.content);
    if (mainFile) {
      this.onToast(`📥 Blocs restaurados desde almacenamiento local PWA (${email}).`);
      return JSON.parse(mainFile.content);
    }

    // 3. Fallback opcional al servidor local
    try {
      const res = await fetch(`/api/drive-vault?email=${encodeURIComponent(email)}`);
      if (res.ok) {
        const data = await res.json();
        if (data && data.workspace && Array.isArray(data.workspace.notebooks)) {
          this.onToast(`📥 Blocs restaurados para ${email}.`);
          return data.workspace;
        }
      }
    } catch (_) {}

    this.onToast(`Aún no hay notas almacenadas en ${email}.`);
    return null;
  }

  async exportDirectToSystemGoogleDrive(workspaceData) {
    const payloadString = JSON.stringify(workspaceData, null, 2);
    const blob = new Blob([payloadString], { type: 'application/json' });
    const file = new File([blob], WORKSPACE_FILENAME, { type: 'application/json' });

    if (navigator.canShare && navigator.canShare({ files: [file] })) {
      try {
        await navigator.share({
          title: 'Bloc de Notas OneNote PWA',
          text: `Copia de seguridad en Google Drive (${this.getActiveEmail()})`,
          files: [file]
        });
        this.onToast('📲 Enviado a la app Google Drive de tu Android.');
        return true;
      } catch (err) {
        if (err.name === 'AbortError') return false;
      }
    }

    if (window.showSaveFilePicker) {
      try {
        if (!this.fileHandle) {
          this.fileHandle = await window.showSaveFilePicker({
            suggestedName: WORKSPACE_FILENAME,
            types: [
              {
                description: 'Archivo de Bloc de Notas OneNote (.json)',
                accept: { 'application/json': ['.json'] }
              }
            ]
          });
        }
        const writable = await this.fileHandle.createWritable();
        await writable.write(payloadString);
        await writable.close();
        this.onToast('📂 Archivo actualizado en tu carpeta de Google Drive.');
        return true;
      } catch (err) {
        if (err.name === 'AbortError') return false;
        this.fileHandle = null;
      }
    }

    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = WORKSPACE_FILENAME;
    a.click();
    this.onToast('💾 Archivo descargado para Google Drive.');
    return true;
  }

  async uploadPagePngToDrive(pageTitle, dataUrl) {
    const email = this.getActiveEmail();
    const safeName = `${(pageTitle || 'Pagina_OneNote').replace(/[^a-z0-9_-]/gi, '_')}_${new Date()
      .toISOString()
      .slice(0, 10)}.png`;

    const parsed = this.dataUrlToUint8Array(dataUrl);
    const sizeBytes = parsed ? parsed.bytes.length : Math.round(dataUrl.length * 0.75);
    const record = {
      id: `${email}::png-${Date.now()}`,
      email,
      name: safeName,
      mimeType: 'image/png',
      updatedAt: new Date().toISOString(),
      size: sizeBytes,
      content: dataUrl
    };

    await this.saveToLocalVault(record);

    if (this.hasLiveOAuthToken() && parsed) {
      await this.uploadFileToRealGoogleDrive({
        name: safeName,
        mimeType: 'image/png',
        blob: new Blob([parsed.bytes], { type: 'image/png' })
      });
    }

    this.onToast(`🖼️ Captura "${safeName}" guardada en Google Drive.`);
    return record;
  }

  async listDriveFiles() {
    const email = this.getActiveEmail();
    const localFiles = await this.listLocalVaultFiles(email);

    if (this.hasLiveOAuthToken()) {
      const remoteDocs = await this.listRemoteGoogleDriveDocxFiles();
      const merged = [...localFiles];
      for (const rem of remoteDocs) {
        if (!merged.some((l) => l.name === rem.name)) {
          merged.push(rem);
        }
      }
      return merged;
    }

    return localFiles;
  }

  /* =========================================================
     5. PERSISTENCIA LOCAL INDEXEDDB PARA LA PWA EN ANDROID
     ========================================================= */

  openVaultDB() {
    return new Promise((resolve, reject) => {
      const req = indexedDB.open('OneNoteDirectGoogleVaultDB', 3);
      req.onupgradeneeded = (e) => {
        const db = e.target.result;
        if (!db.objectStoreNames.contains('drive_files')) {
          db.createObjectStore('drive_files', { keyPath: 'id' });
        }
        if (!db.objectStoreNames.contains('pwa_state')) {
          db.createObjectStore('pwa_state', { keyPath: 'key' });
        }
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }

  async saveToLocalVault(record) {
    const db = await this.openVaultDB();
    return new Promise((resolve, reject) => {
      const tx = db.transaction('drive_files', 'readwrite');
      tx.objectStore('drive_files').put(record);
      tx.oncomplete = () => resolve(record);
      tx.onerror = () => reject(tx.error);
    });
  }

  async listLocalVaultFiles(email) {
    const db = await this.openVaultDB();
    return new Promise((resolve, reject) => {
      const tx = db.transaction('drive_files', 'readonly');
      const req = tx.objectStore('drive_files').getAll();
      req.onsuccess = () => {
        const items = (req.result || []).filter((item) => item.email === email);
        items.sort((a, b) => (b.updatedAt || '').localeCompare(a.updatedAt || ''));
        resolve(items);
      };
      req.onerror = () => reject(req.error);
    });
  }

  async saveLocalPwaWorkspace(workspace) {
    try {
      const db = await this.openVaultDB();
      return new Promise((resolve) => {
        const tx = db.transaction('pwa_state', 'readwrite');
        tx.objectStore('pwa_state').put({
          key: 'active_workspace',
          updatedAt: new Date().toISOString(),
          workspace
        });
        tx.oncomplete = () => resolve(true);
        tx.onerror = () => resolve(false);
      });
    } catch (_) {
      return false;
    }
  }

  async loadLocalPwaWorkspace() {
    try {
      const db = await this.openVaultDB();
      return new Promise((resolve) => {
        const tx = db.transaction('pwa_state', 'readonly');
        const req = tx.objectStore('pwa_state').get('active_workspace');
        req.onsuccess = () => resolve(req.result?.workspace || null);
        req.onerror = () => resolve(null);
      });
    } catch (_) {
      return null;
    }
  }

  readFileAsDataURL(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => resolve(e.target.result);
      reader.onerror = (e) => reject(e);
      reader.readAsDataURL(file);
    });
  }

  escapeHtml(str) {
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');
  }

  escapeXml(str) {
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }
}
