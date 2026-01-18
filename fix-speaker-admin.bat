@echo off
:: Check for admin rights and request elevation if needed
net session >nul 2>&1
if %errorLevel% == 0 (
    echo Running with Administrator rights...
    goto :run
) else (
    echo Requesting Administrator rights...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

:run
echo ========================================
echo Speaker Fix Tool (Administrator Mode)
echo ========================================
echo.
echo Running diagnostics with full permissions...
echo.
powershell.exe -ExecutionPolicy Bypass -File "%~dp0fix-speaker.ps1"
pause
