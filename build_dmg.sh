#!/bin/bash
# ──────────────────────────────────────────────────────────────────
# Build GandalfOSINT.dmg — macOS, ffmpeg intégré
# Usage: bash build_dmg.sh
# ──────────────────────────────────────────────────────────────────
set -e
cd "$(dirname "$0")"

APP_NAME="GandalfOSINT"
SPEC="GandalfOSINT.spec"
DMG_OUT="${APP_NAME}.dmg"
VOLUME_NAME="Gandalf OSINT"

# ── 1. Vérifications ──────────────────────────────────────────────
echo "==> Vérification de ffmpeg..."
FFMPEG_PATH="/opt/homebrew/bin/ffmpeg"
FFPROBE_PATH="/opt/homebrew/bin/ffprobe"

if [ ! -f "$FFMPEG_PATH" ]; then
  echo "❌  ffmpeg introuvable à $FFMPEG_PATH"
  echo "    Installe-le via : brew install ffmpeg"
  exit 1
fi

echo "    ffmpeg : $FFMPEG_PATH  ✓"
echo "    ffprobe: $FFPROBE_PATH  ✓"

# ── 2. Mise à jour des dépendances Python ─────────────────────────
echo ""
echo "==> Mise à jour yt-dlp et pyinstaller..."
pip3 install -q --upgrade yt-dlp pyinstaller pywebview openpyxl

# ── 3. Nettoyage des builds précédents ───────────────────────────
echo ""
echo "==> Nettoyage..."
rm -rf build dist

# ── 4. Build PyInstaller (.app) ───────────────────────────────────
echo ""
echo "==> Build PyInstaller (.app)..."
pyinstaller "$SPEC" --noconfirm

APP_PATH="dist/${APP_NAME}.app"
if [ ! -d "$APP_PATH" ]; then
  echo "❌  .app introuvable : $APP_PATH"
  exit 1
fi
echo "    ✓ $APP_PATH créé"

# ── 5. Création du DMG ────────────────────────────────────────────
echo ""
echo "==> Création du DMG..."
rm -f "$DMG_OUT"

create-dmg \
  --volname "$VOLUME_NAME" \
  --volicon "icon.icns" \
  --window-pos 200 120 \
  --window-size 660 400 \
  --icon-size 128 \
  --icon "${APP_NAME}.app" 180 170 \
  --hide-extension "${APP_NAME}.app" \
  --app-drop-link 480 170 \
  --background "dmg_background.png" \
  "$DMG_OUT" \
  "$APP_PATH" 2>/dev/null || \
create-dmg \
  --volname "$VOLUME_NAME" \
  --volicon "icon.icns" \
  --window-pos 200 120 \
  --window-size 560 340 \
  --icon-size 128 \
  --icon "${APP_NAME}.app" 150 160 \
  --hide-extension "${APP_NAME}.app" \
  --app-drop-link 400 160 \
  "$DMG_OUT" \
  "$APP_PATH"

echo ""
echo "✅  DMG prêt : $(pwd)/$DMG_OUT"
echo "    Taille : $(du -sh "$DMG_OUT" | cut -f1)"
