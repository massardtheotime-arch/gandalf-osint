#!/bin/bash
# ──────────────────────────────────────────────────────────────────
# Script de build PyInstaller pour VideoDownloader Pro
# Usage: bash build_app.sh
# ──────────────────────────────────────────────────────────────────

set -e
cd "$(dirname "$0")"

echo "==> Installation des dépendances..."
pip install yt-dlp pyinstaller --quiet

echo "==> Build PyInstaller..."
pyinstaller \
  --onefile \
  --windowed \
  --name "VideoDownloader" \
  --collect-all yt_dlp \
  downloader_app.py

echo ""
echo "✅ Build terminé !"
echo "   Exécutable : dist/VideoDownloader"
