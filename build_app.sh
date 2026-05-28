#!/bin/bash
# ──────────────────────────────────────────────────────────────────
# Script de build PyInstaller pour Gandalf OSINT (macOS / Linux)
# Usage: bash build_app.sh
# ──────────────────────────────────────────────────────────────────

set -e
cd "$(dirname "$0")"

echo "==> Installation des dépendances..."
pip install yt-dlp pywebview pyinstaller openpyxl pillow --quiet

if [ -n "$GANDALF_VERSION" ]; then
  printf "%s" "$GANDALF_VERSION" > version.txt
elif git describe --tags --exact-match HEAD >/dev/null 2>&1; then
  git describe --tags --exact-match HEAD > version.txt
else
  printf "%s" "dev" > version.txt
fi

ICON_FLAG=""
if [ -f "icon.icns" ]; then
  ICON_FLAG="--icon icon.icns"
fi

echo "==> Build PyInstaller..."
pyinstaller \
  --onefile \
  --windowed \
  --name "GandalfOSINT" \
  $ICON_FLAG \
  --collect-all yt_dlp \
  --collect-all webview \
  --add-data "app.html:." \
  --add-data "version.txt:." \
  --add-data "gandalf.gif:." \
  -y \
  downloader_app.py

echo ""
echo "✅ Build terminé !"
echo "   Exécutable : dist/GandalfOSINT"
