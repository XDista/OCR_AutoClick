@echo off
chcp 936 >nul 2>&1
title 一键生成便携版
cd /d "%~dp0"
echo ==============================================
echo   OCR_AutoClick 便携版 一键生成
echo ==============================================
echo.
echo   即将执行：下载便携Python - 安装依赖 - 打包zip
echo   全程可能需要 10~30 分钟（视网络而定），请勿关闭窗口。
echo.
where pwsh >nul 2>&1
if %errorlevel%==0 (
    pwsh -NoProfile -ExecutionPolicy Bypass -File "%~dp0生成便携版.ps1" %*
) else (
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0生成便携版.ps1" %*
)
echo.
echo ==============================================
echo   脚本执行结束。
echo ==============================================
pause