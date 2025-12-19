@echo off
setlocal enabledelayedexpansion

set "current_dir=%~dp0"
set "script_name=DomainAdmin.ps1"

:: 检查脚本是否存在
if not exist "%current_dir%%script_name%" (
    echo 错误：未找到 %script_name%
    echo 请确保批处理文件和PowerShell脚本在同一目录
    pause
    exit /b 1
)

:: 权限检查
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo 正在请求管理员权限...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

:: 执行脚本
echo 请稍候。。。。正在加载 %script_name%...域控账号管理工具
cd /d "%current_dir%"

:: powershell.exe -ExecutionPolicy Bypass -File "%script_name%"
start powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -Command "& .\%script_name% | Out-Null"

if %errorLevel% equ 0 (
    echo 脚本执行成功！
) else (
    echo 脚本执行失败！错误代码: %errorLevel%
)

exit 1
