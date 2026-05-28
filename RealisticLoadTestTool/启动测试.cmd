@echo off
set LOG=%~dp0error_log.txt
echo 启动时间: %date% %time% > "%LOG%"
PowerShell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Start-RealisticLoadTest.ps1" >> "%LOG%" 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo. >> "%LOG%"
    echo 退出码: %ERRORLEVEL% >> "%LOG%"
    echo.
    echo 脚本出错，错误已写入 error_log.txt，请把该文件发给管理员。
    pause
)
