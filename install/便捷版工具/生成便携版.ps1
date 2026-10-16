# ============================================================
#  OCR_AutoClick 便携版一键生成脚本
# ------------------------------------------------------------
#  用法：把本脚本和「一键生成便携版.bat」放到项目根目录，双击 bat 即可。
#        生成的 zip 在项目上一级目录。
#
#  产出：OCR_AutoClick_vX.Y.Z_便携版.zip（解压后双击 启动.bat 即用）
#
#  做了什么：
#    1. 下载可重定位的便携 Python 3.12（python-build-standalone）
#    2. 修正 requirements.txt 编码（UTF-16 -> UTF-8）
#    3. 安装全部依赖（清华源，PIP_NO_CACHE_DIR 规避 pip 缓存 bug）
#    4. 修复 onnxruntime CPU/GPU 冲突（仅保留 directml 版）
#    5. 内置 VC++ 运行库（msvcp140/vcomp140 等，免目标机装运行库）
#    6. 生成 启动.bat
#    7. 清理 pip 临时目录 / __pycache__ / PDB 等
#    8. 用 7-Zip 打包成 zip
# ============================================================

[CmdletBinding()]
param(
    # 便携 Python 版本（python-build-standalone 的 3.12.x 可重定位构建）
    [string]$PythonVer = "3.12.7",
    [string]$PyTag    = "20241016",
    # pip 镜像源（国内默认清华源）
    [string]$Mirror   = "https://pypi.tuna.tsinghua.edu.cn/simple",
    # 强制重建（删除已有 python312 重新下载安装）
    [switch]$Force
)

$ErrorActionPreference = "Stop"
$Root = $PSScriptRoot

# 关键：规避 pip 自带 cachecontrol 组件的缓存 bug（会卡死），
# 用环境变量可传递到构建子进程。
$env:PIP_NO_CACHE_DIR = "1"

function Write-Step([string]$m) { Write-Host "`n[$((Get-Date).ToString('HH:mm:ss'))] $m" -ForegroundColor Cyan }
function Write-Err([string]$m)   { Write-Host "  [错误] $m" -ForegroundColor Red }
function Write-OK([string]$m)    { Write-Host "  [完成] $m" -ForegroundColor Green }

# ============================================================
# 0. 前置检查：7-Zip 与 tar
# ============================================================
Write-Step "检查前置工具..."
$seven = @("$env:ProgramFiles\7-Zip\7z.exe", "${env:ProgramFiles(x86)}\7-Zip\7z.exe") |
         Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
if (-not $seven) {
    Write-Err "未找到 7-Zip，请先安装：https://www.7-zip.org/"
    exit 1
}
$tar = Get-Command tar -ErrorAction SilentlyContinue
if (-not $tar) {
    Write-Err "未找到 tar（Windows 10 及以上内置），请确认系统版本。"
    exit 1
}
Write-OK "7-Zip: $seven"

# pip 执行封装：用 Start-Process + 文件重定向，输出到日志、返回退出码
function Invoke-Pip {
    param([string]$PipArgs, [string]$LogName)
    $out = Join-Path $env:TEMP $LogName
    $err = "$out.err"
    Remove-Item $out, $err -ErrorAction SilentlyContinue
    $p = Start-Process -FilePath $PythonExe -ArgumentList $PipArgs -WorkingDirectory $Root `
        -RedirectStandardOutput $out -RedirectStandardError $err -Wait -PassThru -NoNewWindow
    if ($p.ExitCode -ne 0) {
        Write-Err "pip 失败 (exit=$($p.ExitCode))"
        Write-Host "  --- 错误日志尾部 ---" -ForegroundColor Red
        Get-Content $err -Tail 15 -ErrorAction SilentlyContinue
        exit 1
    }
    Get-Content $out -Tail 3 -ErrorAction SilentlyContinue | ForEach-Object { Write-Host "  $_" }
}

# ============================================================
# 1~4. 构建环境（仅首次或 -Force 时执行）
# ============================================================
$PyDir     = Join-Path $Root "python312"
$PythonExe = Join-Path $PyDir "python.exe"
$needBuild = $Force -or -not (Test-Path $PythonExe)

if ($needBuild) {
    # ---------- 1. 下载并部署便携 Python ----------
    Write-Step "下载便携 Python $PythonVer ..."
    if (Test-Path $PyDir) { Remove-Item $PyDir -Recurse -Force }
    $pyUrl = "https://github.com/astral-sh/python-build-standalone/releases/download/$PyTag/cpython-$PythonVer%2B$PyTag-x86_64-pc-windows-msvc-install_only.tar.gz"
    $archive = Join-Path $env:TEMP "cpython-$PythonVer.tar.gz"
    Write-Host "  下载: $pyUrl"
    Invoke-WebRequest -Uri $pyUrl -OutFile $archive -UseBasicParsing
    Write-Host "  解压..."
    tar -xzf $archive -C $Root
    Remove-Item $archive -Force -ErrorAction SilentlyContinue
    $extracted = Join-Path $Root "python"
    if (-not (Test-Path (Join-Path $extracted "python.exe"))) {
        Write-Err "解压结构异常，未找到 python\python.exe"
        exit 1
    }
    Rename-Item $extracted "python312"
    & $PythonExe --version
    Write-OK "便携 Python 已就位"

    # ---------- 2. 修正 requirements.txt 编码 ----------
    Write-Step "检查 requirements.txt 编码..."
    $req = Join-Path $Root "requirements.txt"
    if (-not (Test-Path $req)) { Write-Err "缺少 requirements.txt"; exit 1 }
    $bom = [System.IO.File]::ReadAllBytes($req)[0..1]
    if ($bom[0] -eq 0xFF -and $bom[1] -eq 0xFE) {
        $txt = [System.IO.File]::ReadAllText($req, [System.Text.Encoding]::Unicode)
        [System.IO.File]::WriteAllText($req, $txt, (New-Object System.Text.UTF8Encoding($false)))
        Write-OK "requirements.txt 已从 UTF-16 转为 UTF-8"
    } else {
        Write-OK "requirements.txt 已是 UTF-8"
    }

    # ---------- 3. 安装依赖 ----------
    Write-Step "安装构建工具 (setuptools/wheel)..."
    Invoke-Pip "-m pip install setuptools wheel -i $Mirror --progress-bar off --disable-pip-version-check" "ocr_step3.log"

    Write-Step "安装全部依赖 (requirements.txt)，需下载 torch 等大包，请耐心等待..."
    Invoke-Pip "-m pip install -r requirements.txt -i $Mirror --progress-bar off --disable-pip-version-check" "ocr_step4.log"
    Write-OK "依赖安装完成"

    # ---------- 4. 修复 onnxruntime 冲突 ----------
    Write-Step "修复 onnxruntime 冲突（卸载 CPU 版，保留 directml 版）..."
    Invoke-Pip "-m pip uninstall onnxruntime onnxruntime-directml -y --disable-pip-version-check" "ocr_step5a.log"
    Invoke-Pip "-m pip install onnxruntime-directml -i $Mirror --progress-bar off --disable-pip-version-check" "ocr_step5b.log"
    Write-OK "onnxruntime-directml 已就绪"
} else {
    Write-Step "检测到已有 python312，跳过下载与依赖安装（加 -Force 可强制重建）"
}

# ============================================================
# 5. 内置 VC++ 运行库（目标机可能没装 VC++ Redistributable）
# ============================================================
Write-Step "内置 VC++ 运行库..."
# 8 个 VC++ 运行库 + 3 个 MFC 运行库（win32ui/dde 依赖 mfc140u.dll，否则目标机报 DLL load failed）
$vcDlls = @('msvcp140.dll','msvcp140_1.dll','msvcp140_2.dll','vcomp140.dll','concrt140.dll','vccorlib140.dll','vcruntime140.dll','vcruntime140_1.dll','mfc140u.dll','mfc140chs.dll','mfc140enu.dll')
$copied = 0
$missed = @()
foreach ($d in $vcDlls) {
    $src = "C:\Windows\System32\$d"
    if (Test-Path $src) {
        Copy-Item $src (Join-Path $PyDir $d) -Force
        $copied++
    } else {
        $missed += $d
    }
}
if ($missed.Count -gt 0) {
    Write-Err "以下运行库 DLL 在本机 System32 中未找到，目标机可能缺少运行库：$($missed -join ', ')"
} else {
    Write-OK "已内置 $copied 个运行库 DLL（共 $($vcDlls.Count) 个，无缺失）"
}

# ============================================================
# 6. 生成启动器
# ============================================================
Write-Step "生成启动器 启动.bat ..."
$launcher = @'
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
'@
# 关键：批处理必须 GBK(936) + CRLF 换行，否则中文 Windows 下 cmd 误读中文多字节字符、行被切断
$launcherCrlf = $launcher -replace "`r?`n", "`r`n"
[System.IO.File]::WriteAllText((Join-Path $Root "启动.bat"), $launcherCrlf, [System.Text.Encoding]::GetEncoding(936))

# 读版本号
$ver = "1.0.0"
$mainpyw = Join-Path $Root "main.pyw"
if (Test-Path $mainpyw) {
    $m = [regex]::Match((Get-Content $mainpyw -Raw), 'APP_VERSION\s*=\s*"([^"]+)"')
    if ($m.Success) { $ver = $m.Groups[1].Value }
}

Write-OK "已生成 启动.bat（版本 v$ver）"

# ============================================================
# 7. 清理
# ============================================================
Write-Step "清理缓存与临时文件..."
$count = 0
Get-ChildItem $Root -Directory | Where-Object { $_.Name -match '^pip-' } | ForEach-Object { Remove-Item $_.FullName -Recurse -Force -EA SilentlyContinue; $count++ }
Get-ChildItem $Root -Recurse -Directory -Filter '__pycache__' -EA SilentlyContinue | ForEach-Object { Remove-Item $_.FullName -Recurse -Force -EA SilentlyContinue; $count++ }
Get-ChildItem $Root -Recurse -File -Filter '*.pyc' -EA SilentlyContinue | ForEach-Object { Remove-Item $_.FullName -Force -EA SilentlyContinue; $count++ }
Get-ChildItem $PyDir -Recurse -File -Filter '*.pdb' -EA SilentlyContinue | ForEach-Object { Remove-Item $_.FullName -Force -EA SilentlyContinue; $count++ }
Remove-Item (Join-Path $PyDir "include") -Recurse -Force -EA SilentlyContinue
Remove-Item (Join-Path $Root "gpu.json") -Force -EA SilentlyContinue
Write-OK "清理完成（删除 $count 个缓存项）"

# ============================================================
# 8. 打包
# ============================================================
Write-Step "打包为 zip（7-Zip 多线程，请耐心等待）..."
$zip = Join-Path (Split-Path $Root -Parent) "OCR_AutoClick_v${ver}_便携版.zip"
if (Test-Path $zip) { Remove-Item $zip -Force }
# 关键：7-Zip 用 Start-Process + 文件重定向，避免 PowerShell 嵌套管道导致输出捕获挂起
$zipOut = Join-Path $env:TEMP "7z_pack.log"
$zipErr = Join-Path $env:TEMP "7z_pack.err"
Remove-Item $zipOut, $zipErr -ErrorAction SilentlyContinue
$zipArgs = "a -tzip -mx=7 -mmt=on `"$zip`" `"$Root`""
$zp = Start-Process -FilePath $seven -ArgumentList $zipArgs -RedirectStandardOutput $zipOut -RedirectStandardError $zipErr -Wait -PassThru -NoNewWindow
if ($zp.ExitCode -ne 0) {
    Write-Err "7-Zip 打包失败 (exit=$($zp.ExitCode))"
    Write-Host "  --- 7-Zip 输出尾部 ---" -ForegroundColor Red
    Get-Content $zipOut -Tail 8 -ErrorAction SilentlyContinue
    Get-Content $zipErr -Tail 8 -ErrorAction SilentlyContinue
    exit 1
}
if (-not (Test-Path $zip)) {
    Write-Err "打包失败，未生成 zip"
    exit 1
}
Get-Content $zipOut -Tail 3 -ErrorAction SilentlyContinue | ForEach-Object { Write-Host "  $_" }

# ============================================================
# 完成
# ============================================================
$sizeMB = [math]::Round((Get-Item $zip).Length / 1MB, 1)
Write-Host ""
Write-Host "==============================================" -ForegroundColor Green
Write-Host "  便携版生成完成！" -ForegroundColor Green
Write-Host "  文件: $zip" -ForegroundColor Green
Write-Host "  大小: $sizeMB MB" -ForegroundColor Green
Write-Host "==============================================" -ForegroundColor Green
Write-Host "  发给他人后：解压 -> 双击 启动.bat 即可使用。" -ForegroundColor Green
