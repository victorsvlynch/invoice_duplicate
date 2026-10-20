@echo off
chcp 65001 > nul
setlocal EnableDelayedExpansion

echo ============================================================
echo  🚀 AUTO-PUSH: Invoice Duplicate Verifier (Batch Runner)
echo ============================================================

cd /d "%~dp0\..\.."

set "COMMIT_MSG=%~1"

if "%COMMIT_MSG%"=="" (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0auto_push.ps1"
    goto end
)

echo.
echo 📦 Preparando cambios con git add...
git add .

echo 📋 Verificando estado de Git...
git status -s

echo 📝 Creando commit con mensaje: "%COMMIT_MSG%"
git commit -m "%COMMIT_MSG%"
if errorlevel 1 (
    echo.
    echo ℹ️ No habia cambios pendientes para commitear o la operacion se cancelo.
    goto end
)

echo.
echo 🚀 Intentando subir cambios al remoto...
git push origin main 2>nul
if errorlevel 1 (
    git push origin master 2>nul
    if errorlevel 1 (
        echo ❌ No se pudo realizar el push directamente. Verifique la conexion o el remoto configurado.
    ) else (
        echo ✅ Push realizado exitosamente a la rama master!
    )
) else (
    echo ✅ Push realizado exitosamente a la rama main!
)

:end
echo.
pause
