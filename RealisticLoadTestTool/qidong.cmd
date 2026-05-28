@echo off
:: Check admin rights; if not admin, relaunch as admin
net session >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo Requesting admin rights...
    powershell -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

:: Run PowerShell script as admin
PowerShell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Start-RealisticLoadTest.ps1"
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Script failed with code: %ERRORLEVEL%
    pause
)
