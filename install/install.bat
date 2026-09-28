@echo off
chcp 65001 > nul 2>&1
title 自动识别点击工具 - 一键安装脚本

echo ==============================================
echo   自动识别点击工具 - 一键安装脚本
echo ==============================================
echo.

:: ========================================
:: Step 1: 升级 pip
:: ========================================
echo [1/4] 升级 pip 到最新版...
python -m pip install --upgrade pip -i https://pypi.tuna.tsinghua.edu.cn/simple
if errorlevel 1 (
    echo [警告] pip 升级失败，不影响后续安装。请先检查本机 Python 是否已经创建环境变量。
)
echo.

:: ========================================
:: Step 2: 安装全部依赖
:: ========================================
echo [2/4] 安装全部依赖包（requirements.txt）...
python -m pip install -r ..\requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
if errorlevel 1 (
    echo [错误] 依赖安装失败，请检查网络连接或 Python 环境。
    pause
    exit /b 1
)
echo.

:: ========================================
:: Step 3: 修复 onnxruntime 冲突
:: rapidocr-onnxruntime 的依赖链会自动拉取 onnxruntime（纯CPU版），
:: 与 onnxruntime-directml（DirectML GPU版）产生模块覆盖冲突，
:: 高版本 CPU 版会覆盖低版本 DirectML 版，导致 DmlExecutionProvider 不可用。
:: 此处同时卸载两个版本，再重装 onnxruntime-directml，确保干净环境。
:: ========================================
echo [3/4] 修复 onnxruntime 冲突...
echo         卸载 onnxruntime（CPU版）及 onnxruntime-directml（GPU版）...
python -m pip uninstall onnxruntime onnxruntime-directml -y 2>nul
echo         重装 onnxruntime-directml（DirectML GPU版）...
python -m pip install onnxruntime-directml -i https://pypi.tuna.tsinghua.edu.cn/simple
if errorlevel 1 (
    echo [警告] onnxruntime-directml 重装失败，请手动安装。
) else (
    echo [完成] onnxruntime-directml 已重装，DirectML GPU 加速可用。
)
echo.

:: ========================================
:: Step 4: 清理 GPU 缓存
:: ========================================
echo [4/4] 清理 GPU 缓存，下次启动将重新检测...
if exist "..\gpu.json" (
    del "..\gpu.json"
    echo [完成] 已删除 GPU 缓存文件 gpu.json。
) else (
    echo [信息] GPU 缓存文件不存在，跳过。
)
echo.

echo ==============================================
echo   安装全部完成！
echo   现在可以运行 main.pyw 启动程序。初次启动或将存在问题，重启即可。
echo ==============================================
echo.
pause