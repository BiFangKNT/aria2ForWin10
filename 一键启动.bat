@echo off
chcp 65001

:: 检查是否以管理员权限运行
NET SESSION >NUL 2>&1
IF %ERRORLEVEL% NEQ 0 (
    ECHO 请求管理员权限...
    powershell -Command "Start-Process -FilePath '%~dpnx0' -Verb RunAs"
    EXIT /B
)

:: 运行 pwsh
:: 获取当前脚本所在的完整路径
set "SCRIPT_PATH=%~dp0setup.ps1"

:: 检查 PowerShell 脚本是否存在
if not exist "%SCRIPT_PATH%" (
    echo PowerShell脚本文件不存在: %SCRIPT_PATH%
    pause
    exit /b 1
)

:: 使用完整路径运行 PowerShell 脚本
pwsh -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_PATH%"
if %ERRORLEVEL% NEQ 0 (
    echo PowerShell脚本执行失败
    pause
    exit /b 1
)

pause
