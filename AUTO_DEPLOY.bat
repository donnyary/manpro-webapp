@echo off
color 0B
cls
echo.
echo  ╔══════════════════════════════════════════════════════════════╗
echo  ║                                                              ║
echo  ║   🔄 MANPRO AUTO DEPLOY WATCHER                             ║
echo  ║   Monitor perubahan & auto-push ke GitHub                    ║
echo  ║                                                              ║
echo  ║   ⚠️  JANGAN TUTUP WINDOW INI!                               ║
echo  ║   Tekan Ctrl+C untuk berhenti                                ║
echo  ║                                                              ║
echo  ╚══════════════════════════════════════════════════════════════╝
echo.

cd /d "%~dp0"

echo 📁 Folder: %CD%
echo.
echo 🚀 Memulai auto-deploy watcher...
echo.
echo ════════════════════════════════════════════════════════════════
echo.

python auto_deploy.py

echo.
echo ════════════════════════════════════════════════════════════════
echo 👋 Auto-deploy watcher dihentikan.
echo.
pause
