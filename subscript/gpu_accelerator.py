"""GPU 硬件加速模块

支持 NVIDIA / AMD / Intel 三家 GPU：
  - OpenCV OpenCL     → 图像预处理加速（三家通用，零额外依赖）
  - ONNX DirectML     → OCR 识别加速（三家通用，需 pip install onnxruntime-directml）
  - ONNX CUDA         → NVIDIA 专用（需 pip install onnxruntime-gpu）

用法：
    from gpu_accelerator import GPUAccelerator

    gpu = GPUAccelerator()
    gpu.enable_opencv_accel()          # 启用 OpenCV OpenCL
    print(gpu.summary())               # 打印 GPU 信息

    # OCR 预处理（自动走 GPU 或 CPU）
    processed = gpu.preprocess_image(bgr_image, grayscale=True, threshold=127)
"""

import subprocess
import sys
import os
import json

# —— GPU 缓存文件路径 ——
def _get_gpu_cache_path():
    """获取 gpu.json 路径（项目根目录）"""
    return os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "gpu.json")

# —— OpenCV（项目中已有） ——
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
            "CUDAExecutionProvider",       # NVIDIA CUDA
            "OpenVINOExecutionProvider",   # Intel OpenVINO
            "ROCMExecutionProvider",       # AMD ROCm
        ]
        for p in priority:
            if p in available:
                _ONNX_PROVIDERS.append(p)
    except ImportError:
        pass

    return _ONNX_PROVIDERS


def _get_gpu_names_via_wmic():
    """通过 WMIC 获取系统 GPU 名称列表（Windows）

    不依赖任何第三方库，仅使用 subprocess + WMIC。
    用于替代 torch_directml.device_name()。

    :return: list[str]，GPU 名称列表，失败时返回空列表
    """
    if sys.platform != "win32":
        return []

    try:
        result = subprocess.run(
            ["wmic", "path", "Win32_VideoController", "get", "Name,AdapterCompatibility"],
            capture_output=True, text=True, timeout=10,
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        output = result.stdout
    except Exception:
        return []

    # 解析 WMIC 输出，提取 GPU 名称
    # WMIC 实际输出列顺序为：AdapterCompatibility  Name（厂商在前，名称在后）
    lines = [l.strip() for l in output.splitlines() if l.strip()]
    header_idx = -1
    for i, line in enumerate(lines):
        if "AdapterCompatibility" in line and "Name" in line:
            header_idx = i
            break

    if header_idx < 0:
        return []

    names = []
    for i in range(header_idx + 1, len(lines)):
        # WMIC 输出格式：厂商名称 + 多个空格 + GPU 名称
        # 例："NVIDIA                NVIDIA GeForce RTX 4090"
        parts = lines[i].rsplit("  ", 1)
        vendor = parts[0].strip() if len(parts) > 1 else ""
        name = parts[1].strip() if len(parts) > 1 else lines[i].strip()

        # 仅保留真实 GPU 厂商（排除虚拟显示适配器）
        vendor_lower = vendor.lower()
        is_real_gpu = any(
            kw in vendor_lower for kw in ("nvidia", "amd", "ati", "intel", "radeon")
        )
        if name and is_real_gpu:
            names.append(name)

    return names


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
            if cached and cached.get("devices"):
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

        # 安全获取设备列表，即使 torch 驱动异常也不影响缓存写入
        try:
            device_list = GPUAccelerator._build_device_list_raw()
        except Exception:
            device_list = [
                {"index": -1, "name": "CPU", "device": "cpu", "backend": "cpu"},
            ]

        data = {
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

        :return: dict 或 None（缓存不存在或无效时）
        """
        try:
            cache_path = _get_gpu_cache_path()
            if not os.path.exists(cache_path):
                return None
            with open(cache_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if "device_list" in data and isinstance(data["device_list"], list):
                return data
        except Exception:
            pass
        return None

    def _detect_gpus(self):
        """检测系统 GPU（Windows 优先）"""
        info = {"vendor": "unknown", "name": "未知", "devices": []}

        if sys.platform != "win32":
            return self._detect_gpus_linux(info)

        # 方式1：WMIC（较快，通常1~3秒）
        try:
            result = subprocess.run(
                ["wmic", "path", "Win32_VideoController", "get", "Name,AdapterCompatibility"],
                capture_output=True, text=True, timeout=10, creationflags=subprocess.CREATE_NO_WINDOW
            )
            info = self._parse_wmic(result.stdout, info)
            if info["devices"]:
                return info
        except Exception:
            pass

        # 方式2：DXDiag（较慢，10~30秒，仅在 WMIC 无结果时作为后备）
        try:
            result = subprocess.run(
                ["dxdiag", "/t", os.path.join(os.environ.get("TEMP", "."), "dxdiag_gpu.txt")],
                capture_output=True, timeout=30, creationflags=subprocess.CREATE_NO_WINDOW
            )
            tmp_file = os.path.join(os.environ.get("TEMP", "."), "dxdiag_gpu.txt")
            if os.path.exists(tmp_file):
                with open(tmp_file, "r", encoding="utf-8", errors="ignore") as f:
                    content = f.read()
                info = self._parse_dxdiag(content, info)
                try:
                    os.remove(tmp_file)
                except OSError:
                    pass
                if info["devices"]:
                    return info
        except Exception:
            pass

        # 方式3：OpenCV OpenCL 探测
        if _OPENCV_AVAILABLE:
            try:
                if cv2.ocl.haveOpenCL():
                    cv2.ocl.setUseOpenCL(True)
                    info["opencl_available"] = True
                    # 尝试获取 OpenCL 设备名
                    try:
                        platforms = cv2.ocl.getPlatfoms()
                    except Exception:
                        platforms = []
                    if not info["devices"] and platforms:
                        for plat in platforms:
                            try:
                                for dev in plat.getDevices():
                                    info["devices"].append(dev.name())
                            except Exception:
                                pass
            except Exception:
                pass

        if not info["devices"]:
            info["vendor"] = "cpu"
            info["name"] = "CPU (未检测到 GPU)"

        return info

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
        """解析 DXDiag 输出"""
        devices = []
        vendor = "unknown"
        in_display = False
        current_name = ""

        for line in content.splitlines():
            if "Display Devices" in line:
                in_display = True
                continue
            if in_display:
                if line.startswith("------------"):
                    continue
                if line.startswith("Card name:"):
                    current_name = line.split(":", 1)[1].strip()
                if line.startswith("Dedicated Memory:"):
                    if current_name:
                        devices.append(current_name)
                        current_name = ""
                if line.strip() == "" and devices:
                    break

        if devices:
            info["devices"] = devices
            name = devices[0].lower()
            if "nvidia" in name or "nv" in name:
                info["vendor"] = "nvidia"
                info["name"] = devices[0]
            elif "amd" in name or "radeon" in name or "ati" in name:
                info["vendor"] = "amd"
                info["name"] = devices[0]
            elif "intel" in name or "uhd" in name or "iris" in name or "arc" in name:
                info["vendor"] = "intel"
                info["name"] = devices[0]

        return info

    def _parse_wmic(self, output, info):
        """解析 WMIC 输出"""
        lines = [l.strip() for l in output.splitlines() if l.strip()]
        header_idx = -1
        for i, line in enumerate(lines):
            if "AdapterCompatibility" in line and "Name" in line:
                header_idx = i
                break

        for i in range(header_idx + 1, len(lines)):
            parts = lines[i].rsplit("  ", 1)
            name_part = parts[0] if len(parts) > 1 else lines[i]
            vendor_part = parts[1].strip() if len(parts) > 1 else ""

            name_lower = name_part.lower()
            if vendor_part:
                vendor = vendor_part.lower()
                if "nvidia" in vendor:
                    info["vendor"] = "nvidia"
                elif "amd" in vendor or "ati" in vendor:
                    info["vendor"] = "amd"
                elif "intel" in vendor:
                    info["vendor"] = "intel"
            else:
                if "nvidia" in name_lower:
                    info["vendor"] = "nvidia"
                elif "amd" in name_lower or "radeon" in name_lower:
                    info["vendor"] = "amd"
                elif "intel" in name_lower:
                    info["vendor"] = "intel"

            info["devices"].append(name_part.strip())
            if not info["name"] or info["name"] == "未知":
                info["name"] = name_part.strip()

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
          - GPU 不可用 → CPU 路径（无额外开销，只是少了一个 setUseOpenCL(True)）

        参数与 OCREngine._preprocess_image 保持一致。
        """
        img = img_bgr

        if not self._opencv_ocl_ready:
            self.enable_opencv_accel()

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

    # ---------- CUDA 设备列表 ----------

    @staticmethod
    def detect_cuda_devices(force_refresh=False):
        """检测可用的 CUDA GPU 设备列表

        优先从 gpu.json 缓存读取，除非 force_refresh=True。

        :param force_refresh: 是否强制重新检测（忽略缓存）
        :return: list[dict]，每个 dict 含:
            - "index": GPU 索引 (0, 1, ...)
            - "name": GPU 名称
            - "device": 设备字符串 ("cpu", "cuda:0", "dml:0", ...)
            - "backend": 后端类型 ("cpu", "cuda", "dml")
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
    def _build_device_list_raw():
        """实时检测所有可用加速设备列表（CUDA + DirectML + CPU）

        DirectML 设备可通过 ONNX Runtime + WMIC 检测（无需 torch-directml）。
        注意：torch 可能因驱动缺失、硬件异常等原因抛出 RuntimeError
        等非 ImportError 异常，因此这里捕获所有 Exception。
        """
        devices = [
            {"index": -1, "name": "CPU", "device": "cpu", "backend": "cpu"},
        ]

        # —— CUDA（NVIDIA 专用） ——
        try:
            import torch
            if torch.cuda.is_available():
                for i in range(torch.cuda.device_count()):
                    try:
                        devices.append({
                            "index": i,
                            "name": f"NVIDIA {torch.cuda.get_device_name(i)}",
                            "device": f"cuda:{i}",
                            "backend": "cuda",
                        })
                    except Exception:
                        pass
        except Exception:
            pass

        # —— DirectML（通过 ONNX Runtime + WMIC，无需 torch-directml） ——
        try:
            import onnxruntime as ort
            if "DmlExecutionProvider" in ort.get_available_providers():
                # 通过 WMIC 获取系统 GPU 名称（Windows 系统级 API，无第三方依赖）
                gpu_names = _get_gpu_names_via_wmic()

                # 排除已在 CUDA 分支中列出的 GPU（避免 NVIDIA 卡重复出现）
                cuda_names = {
                    d["name"].replace("NVIDIA ", "") for d in devices
                    if d.get("backend") == "cuda"
                }

                dml_index = 0
                for gpu_name in gpu_names:
                    # 模糊去重：CUDA 名和 WMIC 名可能略有差异
                    if any(cn in gpu_name or gpu_name in cn for cn in cuda_names if cn):
                        continue
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

    @staticmethod
    def get_torch_device(device_str):
        """将设备字符串转为 PyTorch device 对象

        :param device_str: "cpu" / "cuda:0" ...
        :return: torch.device
        """
        import torch
        return torch.device(device_str)


# ==================== 全局单例 ====================

_gpu_accelerator: GPUAccelerator | None = None


def get_gpu_accelerator() -> GPUAccelerator:
    """获取全局 GPU 加速器单例"""
    global _gpu_accelerator
    if _gpu_accelerator is None:
        _gpu_accelerator = GPUAccelerator()
    return _gpu_accelerator