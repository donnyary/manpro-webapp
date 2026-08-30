@echo off
color 0B
cls
echo.
echo  ╔══════════════════════════════════════════════════════════════╗
echo  ║                                                              ║
echo  ║   🚀 MANPRO AUTO DEPLOY                                     ║
echo  ║   Push kode ke GitHub ^& auto-deploy ke Railway               ║
echo  ║                                                              ║
echo  ╚══════════════════════════════════════════════════════════════╝
echo.

cd /d "%~dp0"

echo 📋 Perubahan yang terdeteksi:
echo ────────────────────────────────
git status --short
echo.

echo 🚀 Siap deploy ke production?
echo ────────────────────────────────
set /p confirm="Ketik 'DEPLOY' untuk konfirmasi: "

if /i not "%confirm%"=="DEPLOY" (
    echo.
    echo ❌ Dibatalkan.
    pause
    exit /b
)

echo.
echo ────────────────────────────────────────────────────────────
echo 🚀 DEPLOY DIMULAI...
echo ────────────────────────────────────────────────────────────
echo.

echo 📦 Step 1/3: Git Add...
git add .
echo ✅ File ditambahkan
echo.

echo 💾 Step 2/3: Git Commit...
git commit -m "Deploy: %date% %time%"
echo ✅ Commit berhasil
echo.

echo ⬆️  Step 3/3: Push ke GitHub...
git push
echo.

echo ════════════════════════════════════════════════════════════════
echo.
echo 🎉 DEPLOY BERHASIL!
echo.
echo 📋 Yang terjadi:
echo    ✅ Kode di-push ke GitHub
echo    ✅ Railway akan auto-deploy dalam 1-2 menit
echo.
echo 🌐 URL Aplikasi:
echo    https://manpro-webapp-production.up.railway.app
echo.
echo ════════════════════════════════════════════════════════════════
echo.
pause
