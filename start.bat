@echo off
echo ======================================================
echo  Iniciando Verificador de Facturas Duplicadas
echo ======================================================
echo.

docker --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Docker no esta instalado o no se encuentra en el PATH.
    echo Por favor inicia Docker Desktop y vuelve a intentar.
    pause
    exit /b 1
)

echo Construyendo y levantando contenedor Docker...
docker compose up --build -d

if %errorlevel% equ 0 (
    echo.
    echo ======================================================
    echo  Contenedor iniciado correctamente!
    echo  Abriendo navegador en http://localhost:8504 ...
    echo ======================================================
    timeout /t 3 >nul
    start http://localhost:8504
) else (
    echo [ERROR] Hubo un problema al iniciar Docker Compose.
    pause
)
