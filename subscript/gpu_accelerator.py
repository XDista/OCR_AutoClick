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

    def detect(self):
        """执行 GPU 检测（按需调用，避免启动卡顿）

        优先从 gpu.json 缓存读取，避免每次启动都执行 dxdiag/wmic 等重操作。
        缓存无效时才执行实时检测。

        :return: self（链式调用）
        """
        if self._detected:
            return self

        # 优先从缓存加载，避免启动时执行 dxdiag（10~30秒）
        cached = self.load_from_cache()
        if cached and cached.get("devices"):
            self._load_from_cache_data(cached)
            self._detected = True
            return self

        self._detected = True

        self.gpu_info = self._detect_gpus()
        self._onnx_ready = bool(_try_import_onnx())
        self._onnx_providers = _ONNX_PROVIDERS
        self.enable_opencv_accel()
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
        try:
            cache_path = _get_gpu_cache_path()
            data = {
                "vendor": self.gpu_info.get("vendor", "unknown"),
                "name": self.gpu_info.get("name", "未知"),
                "devices": self.gpu_info.get("devices", []),
                "opencv_ocl_ready": self._opencv_ocl_ready,
                "onnx_providers": list(self._onnx_providers),
                "device_list": GPUAccelerator._build_device_list_raw(),
            }
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
        """实时检测所有可用加速设备列表（CUDA + DirectML + CPU）"""
        devices = [
            {"index": -1, "name": "CPU", "device": "cpu", "backend": "cpu"},
        ]

        # —— CUDA（NVIDIA 专用） ——
        try:
            import torch
            if torch.cuda.is_available():
                for i in range(torch.cuda.device_count()):
                    devices.append({
                        "index": i,
                        "name": f"NVIDIA {torch.cuda.get_device_name(i)}",
                        "device": f"cuda:{i}",
                        "backend": "cuda",
                    })
        except ImportError:
            pass

        # —— DirectML（NVIDIA / AMD / Intel 三家通用，需 pip install torch-directml） ——
        try:
            import torch_directml
            dml_count = torch_directml.device_count()
            for i in range(dml_count):
                dml_name = torch_directml.device_name(i)
                devices.append({
                    "index": i,
                    "name": f"DirectML: {dml_name}",
                    "device": f"dml:{i}",
                    "backend": "dml",
                })
        except ImportError:
            pass

        return devices

    @staticmethod
    def get_torch_device(device_str):
        """将设备字符串转为 PyTorch device 对象

        :param device_str: "cpu" / "cuda:0" / "dml:0" ...
        :return: torch.device
        """
        import torch
        if device_str.startswith("dml"):
            gpu_id = int(device_str.split(":")[1]) if ":" in device_str else 0
            try:
                import torch_directml
                return torch_directml.device(gpu_id)
            except ImportError:
                raise ImportError(
                    "DirectML 设备需要安装 torch-directml，请执行:"
                    " pip install torch-directml"
                )
        return torch.device(device_str)


# ==================== 全局单例 ====================

_gpu_accelerator: GPUAccelerator | None = None


def get_gpu_accelerator() -> GPUAccelerator:
    """获取全局 GPU 加速器单例"""
    global _gpu_accelerator
    if _gpu_accelerator is None:
        _gpu_accelerator = GPUAccelerator()
    return _gpu_accelerator