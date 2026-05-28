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

if not "%GANDALF_VERSION%"=="" (
  <nul set /p="%GANDALF_VERSION%" > version.txt
) else (
  git describe --tags --exact-match HEAD > version.txt 2>nul
  if errorlevel 1 <nul set /p="dev" > version.txt
)

echo ==> Build PyInstaller...
pyinstaller ^
  --onefile ^
  --windowed ^
  --name GandalfOSINT ^
  --icon icon.ico ^
  --collect-all yt_dlp ^
  --collect-all webview ^
  --add-data "app.html;." ^
  --add-data "version.txt;." ^
  --add-data "gandalf.gif;." ^
  -y ^
  downloader_app.py

echo.
echo Build termine !
echo    Executable : dist\GandalfOSINT.exe
pause
