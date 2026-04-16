@echo off
REM ──────────────────────────────────────────────────────────────────
REM Script de build PyInstaller pour Gandalf OSINT (Windows)
REM Usage: build_app.bat
REM ──────────────────────────────────────────────────────────────────

cd /d "%~dp0"

echo ==> Installation des dependances...
pip install yt-dlp pywebview pyinstaller openpyxl pillow --quiet

echo ==> Generation de l'icone Windows...
python generate_icon.py

echo ==> Build PyInstaller...
pyinstaller ^
  --onefile ^
  --windowed ^
  --name GandalfOSINT ^
  --icon icon.ico ^
  --collect-all yt_dlp ^
  --collect-all webview ^
  --add-data "app.html;." ^
  --add-data "gandalf.gif;." ^
  -y ^
  downloader_app.py

echo.
echo Build termine !
echo    Executable : dist\GandalfOSINT.exe
pause
