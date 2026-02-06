@echo off
REM Sistema de Pedidos de Compra - Vivero Aranjuez V2
REM Script de ejecucion simple (doble clic para ejecutar)

echo ============================================
echo SISTEMA DE PEDIDOS DE COMPRA - VIVERO V2
echo ============================================
echo.

REM Cambiar al directorio del script
cd /d "%~dp0"

REM Verificar si existe Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python no esta instalado o no esta en el PATH.
    echo Por favor, instala Python desde https://python.org
    pause
    exit /b 1
)

echo Python encontrado.
echo.
echo Ejecutando sistema de pedidos...
echo.

REM Ejecutar el script principal
REM Usar --status para ver estado, o ejecutar normalmente
python main.py %*

echo.
echo ============================================
echo Ejecucion finalizada.
echo ============================================
pause
