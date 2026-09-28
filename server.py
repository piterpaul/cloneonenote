#!/usr/bin/env python3
"""
server.py — Servidor PWA, Sincronización Directa y Conector Google Docs para OneNote Web Clone
- Sirve la PWA en toda la red local (0.0.0.0:8090) para Android, Mac y Chromebook.
- /api/network-info: Devuelve IP Wi-Fi local para el código QR de Android.
- /api/drive-vault: Sincroniza blocs y documentos en formato Google Docs asociados al correo @gmail.com sin OAuth.
- /api/google-docs-import: Importa documentos directamente desde enlaces de Google Docs (docs.google.com/document/d/...).
- /api/google-docs-parse-docx: Extrae texto, formato e imágenes de archivos .docx de Google Docs / Word.
"""

import base64
import html
import http.server
import io
import json
import os
import re
import socket
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
import zipfile
from datetime import datetime

PORT = int(os.environ.get("PORT", 8090))
VAULT_FILE = os.path.join(os.path.dirname(__file__), "google_drive_vault.json")


def get_local_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"


def load_vault():
    if not os.path.exists(VAULT_FILE):
        return {}
    try:
        with open(VAULT_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_vault(data):
    with open(VAULT_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def parse_docx_bytes_to_html(docx_bytes):
    """Extrae párrafos, títulos, tablas e imágenes desde un archivo .docx exportado de Google Docs."""
    out_parts = []
    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as zf:
        # Cargar relaciones para imágenes si existen
        rels_map = {}
        if "word/_rels/document.xml.rels" in zf.namelist():
            try:
                rels_xml = zf.read("word/_rels/document.xml.rels")
                rels_root = ET.fromstring(rels_xml)
                for rel in rels_root:
                    r_id = rel.attrib.get("Id")
                    target = rel.attrib.get("Target", "")
                    if r_id and target:
                        rels_map[r_id] = target
            except Exception:
                pass

        if "word/document.xml" not in zf.namelist():
            return "<p>Archivo .docx sin contenido de texto estándar.</p>"

        xml_content = zf.read("word/document.xml")
        root = ET.fromstring(xml_content)
        ns = {
            "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
            "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        }

        body = root.find("w:body", ns)
        if body is None:
            return "<p>Documento vacío.</p>"

        for elem in body:
            tag_name = elem.tag.split("}")[-1]
            if tag_name == "p":
                texts = []
                is_bold = False
                for r_elem in elem.findall(".//w:r", ns):
                    b_elem = r_elem.find("w:rPr/w:b", ns)
                    run_bold = b_elem is not None
                    for t_elem in r_elem.findall(".//w:t", ns):
                        if t_elem.text:
                            escaped = html.escape(t_elem.text)
                            texts.append(f"<b>{escaped}</b>" if run_bold else escaped)

                # Buscar imágenes incrustadas en el párrafo
                for blip in elem.findall(".//a:blip", ns):
                    embed_id = blip.attrib.get(f"{{{ns['r']}}}embed")
                    if embed_id and embed_id in rels_map:
                        media_path = "word/" + rels_map[embed_id].lstrip("/")
                        if media_path in zf.namelist():
                            img_bytes = zf.read(media_path)
                            ext = media_path.split(".")[-1].lower()
                            mime = "image/jpeg" if ext in ("jpg", "jpeg") else "image/png"
                            b64 = base64.b64encode(img_bytes).decode("ascii")
                            out_parts.append(
                                f'<div><img src="data:{mime};base64,{b64}" style="max-width:100%;border-radius:6px;margin:6px 0;" /></div>'
                            )

                line_html = "".join(texts).strip()
                if line_html:
                    style_elem = elem.find("w:pPr/w:pStyle", ns)
                    style_val = style_elem.attrib.get(f"{{{ns['w']}}}val", "") if style_elem is not None else ""
                    if "Heading1" in style_val or "Title" in style_val:
                        out_parts.append(f"<h3><b>{line_html}</b></h3>")
                    elif "Heading2" in style_val:
                        out_parts.append(f"<h4><b>{line_html}</b></h4>")
                    else:
                        out_parts.append(f"<div>{line_html}</div>")

            elif tag_name == "tbl":
                rows_html = []
                for tr in elem.findall(".//w:tr", ns):
                    cells_html = []
                    for tc in tr.findall(".//w:tc", ns):
                        cell_text = "".join(t.text for t in tc.findall(".//w:t", ns) if t.text)
                        cells_html.append(f"<td>{html.escape(cell_text)}</td>")
                    if cells_html:
                        rows_html.append("<tr>" + "".join(cells_html) + "</tr>")
                if rows_html:
                    out_parts.append("<table>" + "".join(rows_html) + "</table>")

    return "\n".join(out_parts) if out_parts else "<p>Documento importado sin texto.</p>"


def fetch_google_doc_by_url(doc_url):
    """Descarga un documento de Google Docs a partir de su enlace compartido."""
    match = re.search(r"/document/d/([a-zA-Z0-9_-]+)", doc_url)
    if not match:
        raise ValueError("No se encontró un ID válido de Google Docs en la URL.")
    doc_id = match.group(1)

    export_txt_url = f"https://docs.google.com/document/d/{doc_id}/export?format=txt"
    req = urllib.request.Request(
        export_txt_url,
        headers={"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)"}
    )
    with urllib.request.urlopen(req, timeout=8) as resp:
        raw_text = resp.read().decode("utf-8", errors="replace")

    lines = [line.strip() for line in raw_text.splitlines()]
    non_empty = [l for l in lines if l]
    title = non_empty[0][:70] if non_empty else f"Google Doc ({doc_id[:8]})"
    html_lines = [
        f"<div>{html.escape(line)}</div>" if line else "<div><br></div>"
        for line in lines
    ]
    return {
        "docId": doc_id,
        "title": title,
        "html": "\n".join(html_lines)
    }


class OneNotePWAHandler(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.end_headers()

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == "/api/network-info":
            ip = get_local_ip()
            payload = {
                "ip": ip,
                "port": PORT,
                "androidUrl": f"http://{ip}:{PORT}",
                "localUrl": f"http://localhost:{PORT}"
            }
            self.send_json(payload)
            return

        if parsed.path == "/api/drive-vault":
            qs = urllib.parse.parse_qs(parsed.query)
            email = (qs.get("email", ["default@gmail.com"])[0] or "default@gmail.com").lower().strip()
            vault = load_vault()
            account_data = vault.get(email, {"files": [], "workspace": None, "gdocs": []})
            self.send_json(account_data)
            return

        if parsed.path == "/api/google-docs-import":
            qs = urllib.parse.parse_qs(parsed.query)
            url = qs.get("url", [""])[0]
            try:
                doc_data = fetch_google_doc_by_url(url)
                self.send_json({"ok": True, **doc_data})
            except Exception as e:
                self.send_json({"ok": False, "error": str(e)})
            return

        return super().do_GET()

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)

        if parsed.path == "/api/google-docs-parse-docx":
            length = int(self.headers.get("Content-Length", 0))
            raw = self.rfile.read(length).decode("utf-8") if length > 0 else "{}"
            body = json.loads(raw)
            b64_data = body.get("base64", "")
            if "," in b64_data:
                b64_data = b64_data.split(",", 1)[1]
            try:
                docx_bytes = base64.b64decode(b64_data)
                parsed_html = parse_docx_bytes_to_html(docx_bytes)
                self.send_json({"ok": True, "html": parsed_html})
            except Exception as e:
                self.send_json({"ok": False, "error": str(e)})
            return

        if parsed.path == "/api/drive-vault":
            length = int(self.headers.get("Content-Length", 0))
            raw = self.rfile.read(length).decode("utf-8") if length > 0 else "{}"
            body = json.loads(raw)

            email = (body.get("email") or "default@gmail.com").lower().strip()
            action = body.get("action", "save_workspace")
            vault = load_vault()
            if email not in vault:
                vault[email] = {"files": [], "workspace": None, "gdocs": []}
            if "gdocs" not in vault[email]:
                vault[email]["gdocs"] = []

            now_iso = datetime.now().isoformat()

            if action == "save_workspace":
                workspace = body.get("workspace")
                vault[email]["workspace"] = workspace
                size_bytes = len(json.dumps(workspace, ensure_ascii=False).encode("utf-8"))
                files = [f for f in vault[email]["files"] if f.get("id") != "drive-workspace-main"]
                files.insert(0, {
                    "id": "drive-workspace-main",
                    "name": "onenote_clone_notebooks.json",
                    "mimeType": "application/json",
                    "updatedAt": now_iso,
                    "size": size_bytes
                })
                vault[email]["files"] = files[:30]
                save_vault(vault)
                self.send_json({"ok": True, "files": vault[email]["files"], "updatedAt": now_iso})
                return

            if action == "save_gdoc":
                doc_title = body.get("title") or "Nota de OneNote"
                doc_html = body.get("html") or ""
                safe_name = re.sub(r"[^a-zA-Z0-9_-]+", "_", doc_title).strip("_") or "Documento"
                filename = f"{safe_name}.gdoc.doc"
                size_bytes = len(doc_html.encode("utf-8"))
                gdoc_id = "gdoc-" + str(int(datetime.now().timestamp() * 1000))

                gdoc_entry = {
                    "id": gdoc_id,
                    "title": doc_title,
                    "name": filename,
                    "mimeType": "application/vnd.google-apps.document",
                    "updatedAt": now_iso,
                    "size": size_bytes,
                    "html": doc_html
                }
                vault[email]["gdocs"].insert(0, gdoc_entry)
                vault[email]["gdocs"] = vault[email]["gdocs"][:25]

                vault[email]["files"].insert(0, {
                    "id": gdoc_id,
                    "name": filename,
                    "mimeType": "application/vnd.google-apps.document",
                    "updatedAt": now_iso,
                    "size": size_bytes
                })
                vault[email]["files"] = vault[email]["files"][:30]
                save_vault(vault)
                self.send_json({"ok": True, "gdoc": gdoc_entry, "files": vault[email]["files"]})
                return

            if action == "save_png":
                name = body.get("name", f"Pagina_{now_iso[:10]}.png")
                size_bytes = int(body.get("size", 0))
                files = vault[email]["files"]
                files.insert(0, {
                    "id": "png-" + str(int(datetime.now().timestamp() * 1000)),
                    "name": name,
                    "mimeType": "image/png",
                    "updatedAt": now_iso,
                    "size": size_bytes
                })
                vault[email]["files"] = files[:30]
                save_vault(vault)
                self.send_json({"ok": True, "files": vault[email]["files"]})
                return

        self.send_response(404)
        self.end_headers()

    def send_json(self, data):
        encoded = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(encoded)))
        self.end_headers()
        self.wfile.write(encoded)


if __name__ == "__main__":
    local_ip = get_local_ip()
    print("\n" + "=" * 68)
    print("🚀 ONENOTE WEB CLONE (PWA) + GOOGLE DOCS — ANDROID, MAC & CHROMEBOOK")
    print("=" * 68)
    print(f"💻 Abrir en este Mac / Chromebook : http://localhost:{PORT}")
    print(f"📱 Abrir en tu Android (misma WiFi): http://{local_ip}:{PORT}")
    print("=" * 68 + "\n")

    server = http.server.ThreadingHTTPServer(("0.0.0.0", PORT), OneNotePWAHandler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nServidor detenido.")
        server.server_close()
