@echo off
chcp 936 >nul 2>&1
title 自动识别点击工具
cd /d "%~dp0"
if exist "%~dp0python312\pythonw.exe" (
    start "" "%~dp0python312\pythonw.exe" "%~dp0main.pyw"
    exit /b 0
)
echo [错误] 未找到便携 Python（python312\pythonw.exe），请确认解压完整。
pause