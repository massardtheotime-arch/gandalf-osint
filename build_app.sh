#!/bin/bash
# ──────────────────────────────────────────────────────────────────
# Script de build PyInstaller pour Gandalf OSINT (macOS / Linux)
# Usage: bash build_app.sh
# ──────────────────────────────────────────────────────────────────

set -e
cd "$(dirname "$0")"

echo "==> Installation des dépendances..."
pip install yt-dlp pywebview pyinstaller openpyxl pillow --quiet

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
  --add-data "gandalf.gif:." \
  -y \
  downloader_app.py

echo ""
echo "✅ Build terminé !"
echo "   Exécutable : dist/GandalfOSINT"
