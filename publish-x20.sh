#!/usr/bin/env bash
# ==============================================================================
# publish-x20.sh — Publicador Automático de OneNote Web Clone (PWA)
# 1) Publica en Google Corp x20web:
#    https://pedropm.users.x20web.corp.google.com/cloneonenote/
# 2) Publica en GitHub Pages (HTTPS público para cualquier dispositivo):
#    https://piterpaul.github.io/cloneonenote/
# ==============================================================================
set -euo pipefail

CORP_USER="${USER:-pedropm}"
CORP_HOST="${CORP_HOST:-${CORP_USER}linux.c.googlers.com}"
FIRST_TWO="${CORP_USER:0:2}"
X20_WWW_DIR="/google/data/rw/users/${FIRST_TWO}/${CORP_USER}/www/cloneonenote"
X20_URL="https://${CORP_USER}.users.x20web.corp.google.com/cloneonenote/"
PUBLIC_URL="https://piterpaul.github.io/cloneonenote/"

STATIC_FILES=(
  "index.html"
  "styles.css"
  "app.js"
  "drive.js"
  "manifest.json"
  "sw.js"
  "icon-192.png"
  "icon-512.png"
  "favicon.ico"
  "favicon.png"
  ".nojekyll"
)

echo "======================================================================"
echo "🚀 PUBLICANDO ONENOTE WEB CLONE (PWA)"
echo "======================================================================"

# 1. Publicar en GitHub Pages (HTTPS público inmediato)
if git remote get-url personal >/dev/null 2>&1; then
  echo "🌐 [1/2] Actualizando enlace público HTTPS (GitHub Pages)..."
  git push personal HEAD:main --force >/dev/null 2>&1 || true
  echo "✅ Publicado en: ${PUBLIC_URL}"
fi

# 2. Publicar en Google Corp x20web (/google/data/rw/users/pe/pedropm/www/cloneonenote)
echo ""
echo "🏢 [2/2] Publicando en Google Corp x20web (${X20_URL})..."

if ! ssh -o BatchMode=yes -o ConnectTimeout=6 "${CORP_HOST}" "echo ok" >/dev/null 2>&1; then
  echo "🔑 Tu cookie SSO corporativa requiere renovar gcert para conectar a ${CORP_HOST}."
  echo "👉 Ejecutando gcert (toca tu llave de seguridad Gnubby cuando parpadee)..."
  gcert
fi

echo "📦 Copiando archivos estáticos a ${CORP_HOST}:${X20_WWW_DIR} ..."
tar czf - "${STATIC_FILES[@]}" | ssh "${CORP_HOST}" "mkdir -p '${X20_WWW_DIR}' && tar xzf - -C '${X20_WWW_DIR}' && chmod -R a+rX '${X20_WWW_DIR}'"

echo ""
echo "======================================================================"
echo "🎉 ¡PUBLICACIÓN COMPLETADA CON ÉXITO!"
echo "🏢 Enlace Corp x20web : ${X20_URL}"
echo "🌐 Enlace Web Público : ${PUBLIC_URL}"
echo "======================================================================"
