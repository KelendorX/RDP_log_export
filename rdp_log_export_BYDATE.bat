@echo off
chcp 65001 > nul
echo ========================================
echo    RDP Connections Log Export - BY DATE
echo ========================================
echo.

cd /d "%~dp0"

set /p INPUT_DATE=Enter date (YYYY-MM-DD) or press Enter for today: 

if "%INPUT_DATE%"=="" (
    powershell -ExecutionPolicy Bypass -File "RDP_Log_ByDate.ps1"
) else (
    powershell -ExecutionPolicy Bypass -File "RDP_Log_ByDate.ps1" -Date "%INPUT_DATE%"
)

echo.
echo ========================================