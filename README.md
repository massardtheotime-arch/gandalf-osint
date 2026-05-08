# Gandalf OSINT

> *You Shall Not Buffer*

Application de téléchargement vidéo orientée OSINT, packagée en app macOS standalone.

## Plateformes supportées

YouTube · TikTok · Instagram · Twitter/X · Facebook · Vimeo · Dailymotion · Telegram

## Fonctionnalités

- **Téléchargement par liens** — colle un ou plusieurs liens, analyse les qualités disponibles et choisis par vidéo
- **Import Excel (.xlsx)** — importe une liste de liens avec métadonnées et renomme automatiquement chaque vidéo : `key_publication_date_description_location`
- **Transcodage** :
  - Apple ProRes 422 LT (`.mov`) — post-production macOS / DaVinci Resolve
  - MP4 H.264 (`.mp4`) — compatible Adobe Premiere Pro
  - Aucun — conserve le fichier source

## Prérequis

```bash
pip install yt-dlp pillow openpyxl pyinstaller
brew install ffmpeg
```

## Lancer en développement

```bash
python3 downloader_app.py
```

## Compiler l'app macOS

```bash
python3 -m PyInstaller \
  --onefile --windowed \
  --name "GandalfOSINT" \
  --collect-all yt_dlp \
  --add-data "gandalf.gif:." \
  downloader_app.py
```

L'app est générée dans `dist/GandalfOSINT.app`.

## Fait par

[Théotime Massard](https://youtube.com/@theotimemassard?si=FFIVqWw1d5wMWeXq)
