@echo off
chcp 65001 > nul 2>&1  :: 设置编码为UTF-8，避免中文乱码
title 自动识别点击工具 - 依赖库安装脚本
echo ==============================================
echo 自动识别点击工具 - 依赖库安装脚本
echo ==============================================
echo.

:: 升级pip到最新版本
echo [1/3] 升级pip工具...
python -m pip install --upgrade pip -i https://pypi.tuna.tsinghua.edu.cn/simple
if errorlevel 1 (
    echo 警告：pip升级失败，可能影响后续依赖安装！
    echo 请检查Python环境是否正确配置（已添加到系统环境变量）
    pause
    exit /b 1
)
echo.

:: 安装核心必装依赖
echo [2/3] 安装核心依赖库...
python -m pip install ^
opencv-python ^
numpy ^
pygetwindow ^
psutil ^
pywin32 ^
pyautogui ^
pillow ^
plyer ^
pynput ^
-i https://pypi.tuna.tsinghua.edu.cn/simple
if errorlevel 1 (
    echo 错误：核心依赖安装失败！
    echo 可能原因：网络问题 / Python版本过低（需3.7+）
    pause
    exit /b 1
)
echo.

:: 可选：安装额外兼容性依赖
echo [3/3] 安装可选兼容性依赖...
python -m pip install pyperclip -i https://pypi.tuna.tsinghua.edu.cn/simple
if errorlevel 1 (
    echo 警告：可选兼容性依赖安装失败！
    echo 不影响核心功能，仅可能影响部分系统的显示适配
)
echo.

echo ==============================================
echo 依赖库安装完成！
echo ==============================================
pause