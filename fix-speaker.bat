@echo off
echo ========================================
echo Speaker Fix Tool Launcher
echo ========================================
echo.
echo This will run the speaker diagnostic and fix script.
echo Some fixes may require Administrator rights.
echo.
pause

powershell.exe -ExecutionPolicy Bypass -File "%~dp0fix-speaker.ps1"

pause
