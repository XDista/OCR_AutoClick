"""GPU 硬件加速模块

支持 NVIDIA / AMD / Intel 三家 GPU：
  - OpenCV OpenCL     → 图像预处理加速（三家通用，零额外依赖）
  - ONNX DirectML     → OCR 识别加速（三家通用，需 pip install onnxruntime-directml）

用法：
    from gpu_accelerator import GPUAccelerator

    gpu = GPUAccelerator()
    gpu.enable_opencv_accel()          # 启用 OpenCV OpenCL
    print(gpu.summary())               # 打印 GPU 信息

    # OCR 预处理（自动走 GPU 或 CPU）
    processed = gpu.preprocess_image(bgr_image, grayscale=True, threshold=127)
"""

import codecs
import json
import locale
import os
import re
import shutil
import subprocess
import tempfile
import time

# —— GPU 缓存文件路径 ——
def _get_gpu_cache_path():
    """获取 gpu.json 路径（项目根目录）"""
    return os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "gpu.json")


# —— 兼容性常量 ——
# gpu.json 结构版本：旧版本文件没有该字段，读取时会做兼容处理。
_CACHE_SCHEMA_VERSION = 2
# CPU-only（未发现 GPU）结果的缓存时间；避免每次启动重复探测，同时允许插上 eGPU 后自动恢复。
_NEGATIVE_CACHE_TTL = 24 * 60 * 60
_CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)
_GPU_VENDOR_KEYWORDS = ("nvidia", "amd", "ati", "intel", "radeon")
_GPU_NAME_KEYWORDS = (
    "nvidia", "geforce", "quadro", "tesla",
    "amd", "radeon", "firepro",
    "intel", "iris", "uhd", "hd graphics", "arc",
)


def _decode_console(data):
    """解码控制台命令输出（兼容 OEM 代码页 / UTF-8）"""
    if isinstance(data, str):
        return data
    if not data:
        return ""

    encodings = [
        "utf-8-sig",
        locale.getpreferredencoding(False),
        "mbcs",
        "gbk",
        "cp936",
        "latin-1",
    ]
    for encoding in encodings:
        if not encoding:
            continue
        try:
            return data.decode(encoding)
        except (LookupError, UnicodeDecodeError):
            continue
    return data.decode("utf-8", errors="replace")


def _decode_text_file(raw):
    """解码文本文件（DXDiag 在中文系统上通常输出 UTF-16）"""
    if not raw:
        return ""
    if raw.startswith(codecs.BOM_UTF16_LE) or raw.startswith(codecs.BOM_UTF16_BE):
        return raw.decode("utf-16", errors="replace")
    if raw.startswith(codecs.BOM_UTF8):
        return raw.decode("utf-8-sig", errors="replace")
    # 无 BOM 的 UTF-16：ASCII 字符之间会大量出现 0x00
    if len(raw) >= 4 and raw.count(b"\x00") > len(raw) // 4:
        try:
            return raw.decode("utf-16-le", errors="replace")
        except Exception:
            pass
    for encoding in ("utf-8", locale.getpreferredencoding(False), "gbk", "cp936"):
        if not encoding:
            continue
        try:
            return raw.decode(encoding)
        except (LookupError, UnicodeDecodeError):
            continue
    return raw.decode("utf-8", errors="replace")


def _infer_gpu_vendor(name):
    """根据 GPU 名称推断厂商"""
    lower = (name or "").lower()
    if "nvidia" in lower or "geforce" in lower or "quadro" in lower or "tesla" in lower:
        return "nvidia"
    if "amd" in lower or "radeon" in lower or "ati" in lower or "firepro" in lower:
        return "amd"
    if "intel" in lower or "iris" in lower or "uhd" in lower or "hd graphics" in lower or "arc" in lower:
        return "intel"
    return "unknown"


def _is_real_gpu_name(vendor, name):
    """判断一行 WMI/WMIC 记录是否代表真实 GPU（排除虚拟显示器）"""
    vendor_lower = (vendor or "").lower()
    name_lower = (name or "").lower()
    if not name_lower:
        return False
    if any(k in vendor_lower for k in _GPU_VENDOR_KEYWORDS):
        return True
    # 部分驱动 AdapterCompatibility 为空或为 "Advanced Micro Devices, Inc."
    return any(k in name_lower for k in _GPU_NAME_KEYWORDS)


def _filter_gpu_pairs(pairs):
    """过滤并去重 (厂商, 名称) 列表，保持系统枚举顺序"""
    names = []
    seen = set()
    for vendor, name in pairs:
        name = (name or "").strip().strip("\x00")
        if not _is_real_gpu_name(vendor, name):
            continue
        key = name.lower()
        if key in seen:
            continue
        seen.add(key)
        names.append(name)
    return names


def _parse_vendor_name_rows(text):
    """解析 WMIC 的“厂商 + 多个空格 + GPU 名称”表格

    兼容两种列顺序（WMIC 在不同系统语言/版本下列顺序可能不同）：
        AdapterCompatibility  Name
        Name                  AdapterCompatibility
    """
    if not text:
        return []

    lines = [line.rstrip() for line in text.splitlines() if line.strip()]
    if not lines:
        return []

    header_idx = None
    header = ""
    for i, line in enumerate(lines):
        lower = line.lower()
        if "adaptercompatibility" in lower and "name" in lower:
            header_idx = i
            header = line
            break

    pairs = []
    if header_idx is not None:
        # WMIC 的列是定宽填充的，用表头位置切列比按空格切更可靠，
        # 也能正确处理 GPU 名称里本身含多个空格、以及列顺序不同的情况。
        lower_header = header.lower()
        name_col = lower_header.find("name")
        vendor_col = lower_header.find("adaptercompatibility")
        if name_col >= 0 and vendor_col >= 0 and name_col != vendor_col:
            for line in lines[header_idx + 1:]:
                if line.strip().lower().startswith("adaptercompatibility"):
                    continue
                if name_col < vendor_col:
                    name = line[name_col:vendor_col].strip()
                    vendor = line[vendor_col:].strip()
                else:
                    vendor = line[vendor_col:name_col].strip()
                    name = line[name_col:].strip()
                if name:
                    pairs.append((vendor, name))
            return pairs

    # 没有可识别的表头（或列位置异常）时，退回按连续空格切分。
    start = (header_idx + 1) if header_idx is not None else 0
    for line in lines[start:]:
        parts = re.split(r"\s{2,}", line.strip(), maxsplit=1)
        if len(parts) != 2:
            pairs.append(("", parts[0]))
            continue
        left, right = parts[0].strip(), parts[1].strip()
        left_vendor = any(k in left.lower() for k in _GPU_VENDOR_KEYWORDS)
        right_vendor = any(k in right.lower() for k in _GPU_VENDOR_KEYWORDS)
        if left_vendor and not right_vendor:
            pairs.append((left, right))
        elif right_vendor and not left_vendor:
            pairs.append((right, left))
        else:
            pairs.append((left, right))
    return pairs

# —— OpenCV  ——
try:
    import cv2
    _OPENCV_AVAILABLE = True
except ImportError:
    _OPENCV_AVAILABLE = False

# —— ONNX Runtime（按需安装） ——
_ONNX_PROVIDERS = []


def _try_import_onnx():
    """检测 ONNX Runtime 及可用的执行提供器"""
    global _ONNX_PROVIDERS
    if _ONNX_PROVIDERS:
        return _ONNX_PROVIDERS

    try:
        import onnxruntime as ort
        available = ort.get_available_providers()

        priority = [
            "DmlExecutionProvider",        # DirectML — Windows 三家通用
            "OpenVINOExecutionProvider",   # Intel OpenVINO
            "ROCMExecutionProvider",       # AMD ROCm
            "CUDAExecutionProvider"       # NVIDIA CUDA
        ]
        for p in priority:
            if p in available:
                _ONNX_PROVIDERS.append(p)
    except ImportError:
        pass

    return _ONNX_PROVIDERS


def _get_gpu_names_via_wmi():
    """通过 WMI COM 获取 GPU 名称（Windows 7+，兼容 Windows 11 24H2+）

    Windows 11 24H2 起 WMIC 默认不再安装，因此把 WMI COM 作为首选方案。
    pywin32 已包含在 requirements.txt 中；导入失败时由后续方案接管。
    """
    if os.name != "nt":
        return []

    try:
        import win32com.client

        # 注意：不要在这里显式调用 CoUninitialize()。
        # pywin32 的 Dispatch 会处理线程 COM 初始化；提前反初始化会导致
        # WMI 返回的 COM 对象在释放时触发 “Win32 exception occurred
        # releasing IUnknown” 错误。
        locator = win32com.client.Dispatch("WbemScripting.SWbemLocator")
        service = locator.ConnectServer(".", r"root\cimv2")
        rows = service.ExecQuery(
            "SELECT Name, AdapterCompatibility FROM Win32_VideoController"
        )
        pairs = []
        for item in rows:
            pairs.append((
                getattr(item, "AdapterCompatibility", "") or "",
                getattr(item, "Name", "") or "",
            ))
        return _filter_gpu_pairs(pairs)
    except Exception:
        return []


def _get_gpu_names_via_powershell():
    """通过 PowerShell CIM 获取 GPU 名称（WMI COM 不可用时的后备）"""
    if os.name != "nt":
        return []

    powershell = shutil.which("powershell")
    if not powershell:
        system_root = os.environ.get("SystemRoot", r"C:\Windows")
        candidate = os.path.join(
            system_root, "System32", "WindowsPowerShell", "v1.0", "powershell.exe"
        )
        if os.path.exists(candidate):
            powershell = candidate
    if not powershell:
        return []

    script = (
        "$OutputEncoding = [Console]::OutputEncoding = "
        "[System.Text.UTF8Encoding]::new($false); "
        "Get-CimInstance Win32_VideoController | "
        "Select-Object Name,AdapterCompatibility | ConvertTo-Json -Compress"
    )
    try:
        result = subprocess.run(
            [powershell, "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
             "-Command", script],
            capture_output=True, timeout=20, creationflags=_CREATE_NO_WINDOW,
        )
        output = _decode_console(result.stdout).strip()
        if not output:
            return []
        data = json.loads(output)
        if isinstance(data, dict):
            data = [data]
        pairs = [
            (item.get("AdapterCompatibility", "") or "", item.get("Name", "") or "")
            for item in data if isinstance(item, dict)
        ]
        return _filter_gpu_pairs(pairs)
    except Exception:
        return []


def _get_gpu_names_via_wmic():
    """通过 WMIC 获取 GPU 名称列表（仅用于 Windows 10 及更早系统的后备）

    Windows 11 24H2+ 默认没有 wmic.exe，因此该函数失败是预期行为，
    上层会继续尝试其他检测方式。

    :return: list[str]，GPU 名称列表，失败时返回空列表
    """
    if os.name != "nt":
        return []

    wmic = shutil.which("wmic")
    if not wmic:
        system_root = os.environ.get("SystemRoot", r"C:\Windows")
        candidate = os.path.join(system_root, "System32", "wbem", "wmic.exe")
        if os.path.exists(candidate):
            wmic = candidate
    if not wmic:
        return []

    try:
        result = subprocess.run(
            [wmic, "path", "Win32_VideoController", "get", "Name,AdapterCompatibility"],
            capture_output=True, timeout=10, creationflags=_CREATE_NO_WINDOW,
        )
        output = _decode_console(result.stdout)
    except Exception:
        return []

    return _filter_gpu_pairs(_parse_vendor_name_rows(output))


def _get_gpu_names():
    """按可靠性顺序获取真实 GPU 名称列表

    顺序：WMI COM（Win7+，无 WMIC 依赖）→ PowerShell CIM → WMIC（旧系统）。
    任一方式得到结果即返回，避免在正常系统上执行慢速的 DXDiag。
    """
    for getter in (
        _get_gpu_names_via_wmi,
        _get_gpu_names_via_powershell,
        _get_gpu_names_via_wmic,
    ):
        try:
            names = getter()
        except Exception:
            names = []
        if names:
            return names
    return []


def _get_opencl_device_info():
    """通过 OpenCV 获取默认 OpenCL GPU 设备

    OpenCV 5.x 没有旧的 ``cv2.ocl.getPlatfoms`` API，应使用
    ``cv2.ocl.Device_getDefault()``。同时排除纯 CPU 的 OpenCL 设备。

    :return: {"name": str, "vendor": str} 或 None
    """
    if not _OPENCV_AVAILABLE:
        return None

    try:
        if not cv2.ocl.haveOpenCL():
            return None
        cv2.ocl.setUseOpenCL(True)

        get_default = getattr(cv2.ocl, "Device_getDefault", None)
        if get_default is None:
            return None
        device = get_default()
        name = (device.name() or "").strip()
        if not name:
            return None

        try:
            if device.type() == getattr(cv2.ocl, "DEVICE_TYPE_CPU", 2):
                return None
        except Exception:
            pass

        vendor = _infer_gpu_vendor(name)
        if vendor == "unknown":
            try:
                vendor = {
                    getattr(cv2.ocl, "DEVICE_VENDOR_NVIDIA", 1): "nvidia",
                    getattr(cv2.ocl, "DEVICE_VENDOR_INTEL", 2): "intel",
                    getattr(cv2.ocl, "DEVICE_VENDOR_AMD", 3): "amd",
                }.get(device.vendorID(), "unknown")
            except Exception:
                pass
        return {"name": name, "vendor": vendor}
    except Exception:
        return None


# ==================== GPU 检测 ====================

class GPUAccelerator:
    """跨厂商 GPU 硬件加速器"""

    def __init__(self):
        self._detected = False
        self.gpu_info = {"vendor": "unknown", "name": "未检测", "devices": []}
        self._opencv_ocl_ready = False
        self._onnx_ready = False
        self._onnx_providers = []

    # ---------- 检测 ----------

    def detect(self, force_refresh=False):
        """执行 GPU 检测（按需调用，避免启动卡顿）

        优先从 gpu.json 缓存读取，避免每次启动都执行 dxdiag/wmic 等重操作。
        缓存无效时才执行实时检测。

        :param force_refresh: 强制实时检测（忽略缓存）。用于"检测GPU"按钮。
        :return: self（链式调用）
        """
        if self._detected and not force_refresh:
            return self

        if force_refresh:
            global _ONNX_PROVIDERS
            _ONNX_PROVIDERS = []
            self._onnx_providers = []
            self._onnx_ready = False

        # 优先从缓存加载，避免启动时执行 dxdiag（10~30秒）
        if not force_refresh:
            cached = self.load_from_cache()
            if cached:
                self._load_from_cache_data(cached)
                self._detected = True
                return self

        try:
            self.gpu_info = self._detect_gpus()
        except Exception:
            self.gpu_info = {"vendor": "cpu", "name": "CPU (GPU 检测失败)", "devices": []}

        try:
            self._onnx_ready = bool(_try_import_onnx())
            self._onnx_providers = _ONNX_PROVIDERS
        except Exception:
            self._onnx_ready = False
            self._onnx_providers = []

        try:
            self.enable_opencv_accel()
        except Exception:
            self._opencv_ocl_ready = False

        self._detected = True
        self._save_to_cache()
        return self

    def _load_from_cache_data(self, cached):
        """从缓存数据恢复 GPU 检测状态"""
        self.gpu_info = {
            "vendor": cached.get("vendor", "unknown"),
            "name": cached.get("name", "未知"),
            "devices": cached.get("devices", []),
        }
        self._opencv_ocl_ready = cached.get("opencv_ocl_ready", False)
        self._onnx_providers = cached.get("onnx_providers", [])
        self._onnx_ready = bool(self._onnx_providers)
        if self._opencv_ocl_ready and _OPENCV_AVAILABLE:
            try:
                cv2.ocl.setUseOpenCL(True)
            except Exception:
                self._opencv_ocl_ready = False

    def _save_to_cache(self):
        """将 GPU 检测结果保存到项目根目录 gpu.json"""
        cache_path = _get_gpu_cache_path()

        # 安全获取设备列表，即使检测异常也不影响缓存写入。
        # 把已检测到的 GPU 名称传进去，避免 WMI/OpenCL 在两次调用之间
        # 偶发失败时出现“devices 有 GPU、device_list 却只有 CPU”的不一致缓存。
        try:
            device_list = GPUAccelerator._build_device_list_raw(
                fallback_names=self.gpu_info.get("devices") or None
            )
        except Exception:
            device_list = [
                {"index": -1, "name": "CPU", "device": "cpu", "backend": "cpu"},
            ]

        data = {
            "schema_version": _CACHE_SCHEMA_VERSION,
            "detected_at": time.time(),
            "vendor": self.gpu_info.get("vendor", "unknown"),
            "name": self.gpu_info.get("name", "未知"),
            "devices": self.gpu_info.get("devices", []),
            "opencv_ocl_ready": self._opencv_ocl_ready,
            "onnx_providers": list(self._onnx_providers),
            "device_list": device_list,
        }
        try:
            os.makedirs(os.path.dirname(cache_path), exist_ok=True)
            with open(cache_path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    @staticmethod
    def load_from_cache():
        """从 gpu.json 读取缓存的 GPU 检测结果

        兼容旧版 gpu.json（没有 schema_version）：
        - 旧缓存里已有 devices 时继续复用，避免升级后强制重扫；
        - 旧缓存是“未检测到 GPU”（devices 为空）时视为失效，
          因为旧版本在 Windows 11 24H2 上会误判并写入这种缓存。
        - 新版 CPU-only 缓存只保留 24 小时，方便插拔 eGPU 后自动恢复。

        :return: dict 或 None（缓存不存在或无效时）
        """
        try:
            cache_path = _get_gpu_cache_path()
            if not os.path.exists(cache_path):
                return None
            with open(cache_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if not isinstance(data, dict):
                return None
            if not isinstance(data.get("device_list"), list):
                return None

            version = data.get("schema_version")
            devices = data.get("devices")
            if version != _CACHE_SCHEMA_VERSION:
                # 旧缓存：只有确实记录了 GPU，且设备列表里真的有非 CPU 设备时才复用。
                # 旧版可能因 WMIC 缺失出现 devices 有值、device_list 却只有 CPU 的情况。
                if not devices:
                    return None
                has_gpu_device = any(
                    isinstance(item, dict) and item.get("device") not in (None, "", "cpu")
                    for item in data["device_list"]
                )
                return data if has_gpu_device else None

            if not devices:
                detected_at = data.get("detected_at")
                if not isinstance(detected_at, (int, float)):
                    return None
                if time.time() - detected_at > _NEGATIVE_CACHE_TTL:
                    return None

            return data
        except Exception:
            pass
        return None

    def _detect_gpus(self):
        """检测系统 GPU（Windows 优先）"""
        info = {"vendor": "unknown", "name": "未知", "devices": []}

        if os.name != "nt":
            return self._detect_gpus_linux(info)

        # 方式1：WMI COM / PowerShell / WMIC（Win7~Win11 24H2+ 均覆盖）
        try:
            names = _get_gpu_names()
            if names:
                info["devices"] = list(names)
                info["name"] = names[0]
                info["vendor"] = _infer_gpu_vendor(names[0])
                return info
        except Exception:
            pass

        # 方式2：OpenCV OpenCL 探测（无需 WMIC，速度快，作为 WMI 的后备）
        opencl_device = _get_opencl_device_info()
        if opencl_device:
            info["opencl_available"] = True
            info["devices"] = [opencl_device["name"]]
            info["name"] = opencl_device["name"]
            info["vendor"] = opencl_device["vendor"]
            return info

        # 方式3：DXDiag（较慢，10~30秒，仅在前两种方式都无结果时使用）
        info = self._detect_gpus_via_dxdiag(info)
        if info["devices"]:
            if info.get("vendor") in (None, "", "unknown"):
                info["vendor"] = _infer_gpu_vendor(info["devices"][0])
            return info

        info["vendor"] = "cpu"
        info["name"] = "CPU (未检测到 GPU)"
        return info

    def _detect_gpus_via_dxdiag(self, info):
        """通过 DXDiag 获取 GPU 信息（兼容 UTF-16 与中英文输出）"""
        tmp_file = os.path.join(
            tempfile.gettempdir(),
            "dxdiag_gpu_{}_{}.txt".format(os.getpid(), int(time.time())),
        )
        try:
            subprocess.run(
                ["dxdiag", "/t", tmp_file],
                capture_output=True, timeout=60, creationflags=_CREATE_NO_WINDOW,
            )
            if not os.path.exists(tmp_file):
                return info
            with open(tmp_file, "rb") as f:
                raw = f.read()
            return self._parse_dxdiag(_decode_text_file(raw), info)
        except Exception:
            return info
        finally:
            try:
                os.remove(tmp_file)
            except OSError:
                pass

    def _detect_gpus_linux(self, info):
        """Linux GPU 检测"""
        # lspci
        try:
            result = subprocess.run(
                ["lspci"], capture_output=True, text=True, timeout=5
            )
            for line in result.stdout.splitlines():
                lower = line.lower()
                if "vga" in lower or "3d" in lower or "display" in lower:
                    if "nvidia" in lower:
                        info["vendor"] = "nvidia"
                    elif "amd" in lower or "radeon" in lower:
                        info["vendor"] = "amd"
                    elif "intel" in lower:
                        info["vendor"] = "intel"
                    info["devices"].append(line.strip())
                    info["name"] = line.strip()
        except Exception:
            pass
        return info

    def _parse_dxdiag(self, content, info):
        """解析 DXDiag 输出（兼容 UTF-16 解码后的中英文标签）"""
        if not content:
            return info

        # 中文系统 DXDiag 的“Card name”可能是“显卡名称/顯示卡名稱”等，
        # 且原解析器依赖后面的 Dedicated Memory 行，一旦本地化就完全失效。
        card_name_pattern = re.compile(
            r"^\s*(?:card\s+name|显卡名称|顯示卡名稱|"
            r"显示卡名称|图形卡名称|圖形卡名稱)\s*[:：]\s*(.+?)\s*$",
            re.IGNORECASE,
        )

        devices = []
        seen = set()
        for line in content.splitlines():
            match = card_name_pattern.match(line)
            if not match:
                continue
            name = match.group(1).strip().strip("\x00")
            if not name or name.lower() in seen:
                continue
            seen.add(name.lower())
            devices.append(name)

        if devices:
            info["devices"] = devices
            info["name"] = devices[0]
            info["vendor"] = _infer_gpu_vendor(devices[0])

        return info

    def _parse_wmic(self, output, info):
        """解析 WMIC 输出（列为 厂商 + 名称，顺序自动识别）"""
        names = _filter_gpu_pairs(_parse_vendor_name_rows(output))
        if not names:
            return info

        info["devices"] = names
        info["name"] = names[0]
        info["vendor"] = _infer_gpu_vendor(names[0])
        return info

    # ---------- OpenCV OpenCL 加速 ----------

    def enable_opencv_accel(self):
        """启用 OpenCV OpenCL 加速

        三家 GPU 均通过 OpenCL 支持，cv2 图像操作自动走 GPU。
        影响函数：cv2.cvtColor / cv2.resize / cv2.threshold 等
        """
        if not _OPENCV_AVAILABLE:
            return False

        try:
            have_ocl = cv2.ocl.haveOpenCL()
            if have_ocl:
                cv2.ocl.setUseOpenCL(True)
                # 验证
                test = cv2.ocl.useOpenCL()
                self._opencv_ocl_ready = test
                return test
        except Exception:
            pass

        self._opencv_ocl_ready = False
        return False

    def disable_opencv_accel(self):
        """停用 OpenCV OpenCL 加速，释放 OpenCL GPU 资源"""
        if not _OPENCV_AVAILABLE:
            return
        try:
            cv2.ocl.setUseOpenCL(False)
            self._opencv_ocl_ready = False
        except Exception:
            pass

    @property
    def opencv_ocl_enabled(self):
        return self._opencv_ocl_ready

    # ---------- 图像预处理（带 GPU 路径） ----------

    def preprocess_image(self, img_bgr, *, grayscale=True, scale=1.0,
                         threshold_enabled=False, threshold_value=127):
        """GPU 加速的图像预处理

        内部自动选择：
          - GPU 可用 → OpenCV OpenCL 路径
          - GPU 不可用 → CPU 路径（无额外开销）

        参数与 OCREngine._preprocess_image 保持一致。
        """
        img = img_bgr

        if grayscale:
            if len(img.shape) == 3:
                img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        else:
            if len(img.shape) == 3:
                img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)

        if scale != 1.0:
            h, w = img.shape[:2]
            img = cv2.resize(img, (int(w * scale), int(h * scale)),
                             interpolation=cv2.INTER_CUBIC)

        if threshold_enabled:
            _, img = cv2.threshold(img, threshold_value, 255, cv2.THRESH_BINARY)

        return img

    # ---------- ONNX Runtime ----------

    @property
    def onnx_available(self):
        return self._onnx_ready

    @property
    def onnx_providers(self):
        return list(self._onnx_providers)

    def best_onnx_provider(self):
        """返回最佳的 ONNX 执行提供器"""
        if self._onnx_providers:
            return self._onnx_providers[0]
        return None

    # ---------- 摘要 ----------

    def summary(self):
        """打印 GPU 加速状态摘要"""
        lines = [
            f"GPU 厂商: {self.gpu_info.get('vendor', 'unknown').upper()}",
            f"GPU 型号: {self.gpu_info.get('name', '未知')}",
        ]
        if self.gpu_info.get("devices"):
            lines.append(f"设备列表: {', '.join(self.gpu_info['devices'])}")

        lines.append(f"OpenCV OpenCL: {'✓ 已启用' if self._opencv_ocl_ready else '✗ 不可用'}")

        if self._onnx_providers:
            providers_str = " → ".join(self._onnx_providers)
            lines.append(f"ONNX Runtime:  已安装 (执行提供器: {providers_str})")
        else:
            lines.append(f"ONNX Runtime:  未安装（可选，用于 OCR 识别 GPU 加速）")
            lines.append(f"               安装: pip install onnxruntime-directml")

        return "\n".join(lines)

    # ---------- 加速设备列表 ----------

    @staticmethod
    def detect_accel_devices(force_refresh=False):
        """检测可用的加速 GPU 设备列表（DirectML + CPU）

        优先从 gpu.json 缓存读取，除非 force_refresh=True。

        :param force_refresh: 是否强制重新检测（忽略缓存）
        :return: list[dict]，每个 dict 含:
            - "index": GPU 索引 (0, 1, ...)
            - "name": GPU 名称
            - "device": 设备字符串 ("cpu", "dml:0", ...)
            - "backend": 后端类型 ("cpu", "dml")
        """
        return GPUAccelerator._build_device_list(force_refresh=force_refresh)

    @staticmethod
    def _build_device_list(force_refresh=False):
        """构建所有可用加速设备列表

        优先从 gpu.json 缓存读取，不存在或 force_refresh=True 时执行实时检测。
        """
        if not force_refresh:
            cached = GPUAccelerator.load_from_cache()
            if cached is not None:
                return cached["device_list"]
        return GPUAccelerator._build_device_list_raw()

    @staticmethod
    def _build_device_list_raw(fallback_names=None):
        """实时检测所有可用加速设备列表（DirectML + CPU）

        DirectML 设备通过 ONNX Runtime 判断是否可用，GPU 名称依次通过
        WMI COM / PowerShell / WMIC / OpenCL 获取，兼容 Windows 11 24H2+。

        :param fallback_names: 可选的已检测 GPU 名称列表，系统枚举失败时使用
        """
        devices = [
            {"index": -1, "name": "CPU", "device": "cpu", "backend": "cpu"},
        ]

        # —— DirectML（通过 ONNX Runtime + 系统 GPU 枚举） ——
        try:
            import onnxruntime as ort
            if "DmlExecutionProvider" in ort.get_available_providers():
                gpu_names = _get_gpu_names()
                if not gpu_names and fallback_names:
                    gpu_names = [
                        name for name in fallback_names
                        if name and _is_real_gpu_name("", name)
                    ]
                if not gpu_names:
                    # 最后兜底：WMI/WMIC 都不可用时，OpenCL 仍可能拿到 GPU 名称
                    opencl_device = _get_opencl_device_info()
                    if opencl_device:
                        gpu_names = [opencl_device["name"]]

                dml_index = 0
                for gpu_name in gpu_names:
                    try:
                        devices.append({
                            "index": dml_index,
                            "name": f"DirectML: {gpu_name}",
                            "device": f"dml:{dml_index}",
                            "backend": "dml",
                        })
                        dml_index += 1
                    except Exception:
                        pass
        except Exception:
            pass

        return devices


# ==================== 全局单例 ====================

_gpu_accelerator: GPUAccelerator | None = None


def get_gpu_accelerator() -> GPUAccelerator:
    """获取全局 GPU 加速器单例"""
    global _gpu_accelerator
    if _gpu_accelerator is None:
        _gpu_accelerator = GPUAccelerator()
    return _gpu_accelerator