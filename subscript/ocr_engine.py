"""OCR 文字识别引擎（基于 RapidOCR + ONNX Runtime，三家 GPU 通用）

GPU 加速方案：
  - ONNX DirectML    → NVIDIA / AMD / Intel 三家通用（需 pip install onnxruntime-directml）
  - CPU              → 默认，无需额外安装

用法：
    from ocr_engine import OCREngine
    engine = OCREngine()
    engine.configure(device="dml:0")
    results = engine.recognize(screenshot_bgr)
"""

import os
import json
import gc
import threading

from gpu_accelerator import get_gpu_accelerator

# ==================== 模块级缓存的 RapidOCR Reader ====================

_ocr_reader = None          # RapidOCR 实例
_ocr_reader_device = None   # 当前实例对应的设备字符串
_ocr_reader_providers = None  # 实际使用的 ONNX providers
_ocr_reader_params = None   # 当前实例对应的 RapidOCR 库参数（dict）
_reader_lock = threading.Lock()  # 保护 _ocr_reader 的创建/销毁


# RapidOCR 库默认参数（与 rapidocr_onnxruntime/config.yaml 保持一致）
DEFAULT_RAPIDOCR_PARAMS = {
    "rec_batch_num": 6,
    "cls_batch_num": 6,
    "det_limit_type": "min",
    "det_limit_side_len": 736,
    "use_cls": True,
}


_ONNX_AVAILABLE_PROVIDERS = None


def _detect_onnx_providers():
    """检测 ONNX Runtime 可用执行提供器（缓存结果）"""
    global _ONNX_AVAILABLE_PROVIDERS
    if _ONNX_AVAILABLE_PROVIDERS is not None:
        return _ONNX_AVAILABLE_PROVIDERS
    try:
        import onnxruntime as ort
        _ONNX_AVAILABLE_PROVIDERS = ort.get_available_providers()
    except ImportError:
        _ONNX_AVAILABLE_PROVIDERS = []
    return _ONNX_AVAILABLE_PROVIDERS


def _release_onnx_gpu_memory():
    """释放 ONNX Runtime 占用的 GPU 显存"""
    try:
        import onnxruntime as ort
        providers = ort.get_available_providers()
        if "DmlExecutionProvider" in providers:
            gc.collect()
    except Exception:
        pass


def _get_reader(device="cpu", rapidocr_params=None):
    """获取或创建 RapidOCR Reader 实例（线程安全，高频调用优化）

    RapidOCR 原生支持 ONNX Runtime 的 DirectML / CPU 后端。
    通过参数 det_use_dml / cls_use_dml / rec_use_dml 启用 DirectML。

    模型 ~30MB，首次使用时自动下载缓存。
    切换设备或 RapidOCR 库参数会销毁旧 Reader 并新建。

    :param device: "cpu" / "dml:0" ...
    :param rapidocr_params: 可选 dict，覆盖 RapidOCR 库默认参数
        （rec_batch_num / cls_batch_num / det_limit_type /
          det_limit_side_len / use_cls）
    :return: RapidOCR 实例
    """
    global _ocr_reader, _ocr_reader_device, _ocr_reader_providers, _ocr_reader_params

    params = dict(DEFAULT_RAPIDOCR_PARAMS)
    if rapidocr_params:
        params.update(rapidocr_params)

    # 快速路径：无锁检查（覆盖 99% 的调用）
    reader = _ocr_reader
    if (reader is not None and _ocr_reader_device == device
            and _ocr_reader_params == params):
        return reader

    with _reader_lock:
        if (_ocr_reader is not None and _ocr_reader_device == device
                and _ocr_reader_params == params):
            return _ocr_reader

        if _ocr_reader is not None:
            del _ocr_reader
            _ocr_reader = None
            _ocr_reader_device = None
            _ocr_reader_providers = None
            _ocr_reader_params = None
            gc.collect()
            _release_onnx_gpu_memory()

        try:
            from rapidocr_onnxruntime import RapidOCR
        except ImportError:
            raise ImportError(
                "RapidOCR 未安装，请执行:\n"
                "  pip install rapidocr-onnxruntime\n"
                "GPU 加速需额外安装:\n"
                "  NVIDIA/AMD/Intel: pip install onnxruntime-directml"
            )

        available = _detect_onnx_providers()
        base_kwargs = {"text_score": 0.0, **params}

        if device.startswith("dml"):
            if "DmlExecutionProvider" in available:
                dml_device_id = int(device.split(":")[1]) if ":" in device else 0
                base_kwargs.update(
                    det_use_dml=True, cls_use_dml=True, rec_use_dml=True
                )

                from rapidocr_onnxruntime.utils.infer_engine import OrtInferSession
                _original_get_ep_list = OrtInferSession._get_ep_list

                def _patched_get_ep_list(self):
                    ep_list = _original_get_ep_list(self)
                    for i, (ep, opts) in enumerate(ep_list):
                        if ep == "DmlExecutionProvider":
                            ep_list[i] = (ep, {**opts, "device_id": dml_device_id})
                    return ep_list

                OrtInferSession._get_ep_list = _patched_get_ep_list
                try:
                    _ocr_reader = RapidOCR(**base_kwargs)
                finally:
                    OrtInferSession._get_ep_list = _original_get_ep_list
                _ocr_reader_providers = _read_actual_providers(_ocr_reader)
                _ocr_reader_device = device
                _ocr_reader_params = params
                return _ocr_reader
            else:
                print(
                    "[RapidOCR] DirectML 不可用，请安装 onnxruntime-directml，"
                    "暂时回退到 CPU。",
                    flush=True,
                )

        _ocr_reader = RapidOCR(**base_kwargs)
        _ocr_reader_providers = _read_actual_providers(_ocr_reader)
        _ocr_reader_device = device
        _ocr_reader_params = params
        return _ocr_reader


def _read_actual_providers(reader):
    """从已创建的 Reader 读取实际使用的 ONNX 执行提供器"""
    providers_set = set()
    for attr_name in ("text_det", "text_cls", "text_rec"):
        model = getattr(reader, attr_name, None)
        if model is None:
            continue
        # OrtInferSession 放在 .infer 属性上（text_rec 可能例外）
        engine = getattr(model, "infer", None)
        if engine is None:
            continue
        session = getattr(engine, "session", None)
        if session is not None:
            try:
                for p in session.get_providers():
                    providers_set.add(p)
            except Exception:
                pass
    return sorted(providers_set) if providers_set else ["CPUExecutionProvider"]


def get_reader_providers():
    """返回当前缓存的 Reader 实际使用的 ONNX 执行提供器"""
    global _ocr_reader_providers
    return _ocr_reader_providers or ["CPUExecutionProvider"]


def is_reader_ready(device="cpu", rapidocr_params=None):
    """判断指定设备的 Reader 是否已缓存（无需重新加载）

    供 GUI / 测试流程在调用 _get_reader 前判断是否需要提示"正在加载模型"。
    :param device: "cpu" / "dml:0" ...
    :param rapidocr_params: 与 _get_reader 相同含义
    :return: bool
    """
    global _ocr_reader, _ocr_reader_device, _ocr_reader_params
    params = dict(DEFAULT_RAPIDOCR_PARAMS)
    if rapidocr_params:
        params.update(rapidocr_params)
    return (_ocr_reader is not None and _ocr_reader_device == device
            and _ocr_reader_params == params)


def cleanup_ocr_reader():
    """显式释放模块级缓存的 OCR Reader（退出时调用）"""
    global _ocr_reader, _ocr_reader_device, _ocr_reader_providers, _ocr_reader_params
    with _reader_lock:
        if _ocr_reader is not None:
            del _ocr_reader
            _ocr_reader = None
            _ocr_reader_device = None
            _ocr_reader_providers = None
            _ocr_reader_params = None
            gc.collect()
            _release_onnx_gpu_memory()


# ==================== 图像预处理 ====================

# 小图面积阈值：低于该值视为"小图"，预处理直接走 CPU，
# 避免 OpenCL kernel 启动/同步的固定开销（小图时该开销反而大于计算本身）
_SMALL_IMAGE_AREA = 100_000


def preprocess_image(img_bgr, *, grayscale=True, scale=1.0,
                     threshold_enabled=False, threshold_value=127):
    """GPU/CPU 自适应的 OCR 图像预处理

    - 复用全局 GPUAccelerator 单例（避免每次重复 detect / 读 gpu.json）
    - 小图（面积 < _SMALL_IMAGE_AREA）直接走 CPU 路径
    - 大图走 OpenCL 加速路径
    """
    gpu = get_gpu_accelerator()
    grayscale_flag = grayscale or threshold_enabled

    h, w = img_bgr.shape[:2]
    if w * h < _SMALL_IMAGE_AREA:
        return _preprocess_cpu(
            img_bgr,
            grayscale=grayscale_flag,
            scale=scale,
            threshold_enabled=threshold_enabled,
            threshold_value=threshold_value,
        )

    gpu.detect()  # 大图走 GPU 前确保检测完成（单例内部有缓存，成本极低）
    return gpu.preprocess_image(
        img_bgr,
        grayscale=grayscale_flag,
        scale=scale,
        threshold_enabled=threshold_enabled,
        threshold_value=threshold_value,
    )


def _preprocess_cpu(img_bgr, *, grayscale, scale, threshold_enabled, threshold_value):
    """CPU 图像预处理

    小图走 OpenCL 会因固定开销负优化，这里用 numpy 直接在 CPU 上计算，
    且不切换全局 OpenCL 状态，线程安全。
    """
    import numpy as np
    import cv2

    img = img_bgr
    if grayscale:
        if img.ndim == 3:
            # BGR → 灰度：Y = 0.114B + 0.587G + 0.299R（与 OpenCV BGR2GRAY 一致）
            img = np.round(
                img[..., 0].astype(np.float32) * 0.114
                + img[..., 1].astype(np.float32) * 0.587
                + img[..., 2].astype(np.float32) * 0.299
            ).astype(np.uint8)
    else:
        if img.ndim == 3:
            img = img[..., ::-1].copy()  # BGR → RGB

    if scale != 1.0:
        h, w = img.shape[:2]
        img = cv2.resize(img, (int(w * scale), int(h * scale)),
                         interpolation=cv2.INTER_CUBIC)

    if threshold_enabled:
        img = np.where(img > threshold_value, 255, 0).astype(np.uint8)

    return img


# ==================== OCR 引擎核心 ====================

class OCREngine:
    """OCR 文字识别引擎（基于 RapidOCR + ONNX Runtime）

    职责：
    1. 对截图执行 OCR，返回识别到的所有文字
    2. 加载 task_ocr/ 目录下的 JSON 任务文件
    3. 根据识别结果匹配 targets，返回匹配到的 actions 列表

    使用示例：
        engine = OCREngine()
        engine.configure(language=['ch_sim', 'en'], confidence_threshold=0.5)

        task_config = engine.load_ocr_task("demo_ocr_task")

        screenshot_bgr = capture_window(...)
        actions, matched_text = engine.match(screenshot_bgr, task_config["TASK1"])

        if actions:
            for action in actions:
                execute_action(window, action["type"], action["params"])
    """

    # RapidOCR 默认模型支持的语言（PP-OCRv4 中英文）
    SUPPORTED_LANGUAGES = {
        "ch_sim", "ch_tra", "en", "ch", "ch_en",
        "chi_sim", "chi_tra", "eng",
    }

    # 语言码映射：Tesseract / EasyOCR 格式 → 内部使用（RapidOCR 模型选择）
    _LANG_MAP = {
        "chi_sim": "ch_sim", "chi_tra": "ch_tra",
        "eng": "en", "jpn": "ja", "kor": "ko",
        "fra": "fr", "fre": "fr", "deu": "de", "ger": "de",
        "spa": "es", "tha": "th", "vie": "vi",
        "ch_sim": "ch_sim", "ch_tra": "ch_tra",
        "en": "en", "ja": "ja", "ko": "ko",
        "fr": "fr", "de": "de", "es": "es", "th": "th", "vi": "vi",
    }

    def __init__(self):
        self._language = ['ch_sim', 'en']
        self._confidence_threshold = 0.5
        self._device = "cpu"
        self._preprocess = {
            "grayscale": False,
            "threshold_enabled": False,
            "threshold_value": 127,
            "scale": 1.0,
        }
        self._rapidocr_params = dict(DEFAULT_RAPIDOCR_PARAMS)
        self._gpu = get_gpu_accelerator()
        self._gpu.detect()  # 缓存优先，首次走后备检测，后续从 gpu.json 读取

    # ---------- 配置 ----------

    def configure(self, language=None, confidence_threshold=None, device=None,
                  preprocess=None, rapidocr_params=None):
        """配置 OCR 参数

        :param language: 语言代码列表，如 ['ch_sim', 'en']
                         兼容旧 EasyOCR / Tesseract 格式，自动转换
        :param confidence_threshold: 置信度阈值 (0~1)
        :param device: "cpu" / "dml:0" ...
        :param preprocess: 图像预处理设置 dict
        :param rapidocr_params: RapidOCR 库参数 dict（rec_batch_num / cls_batch_num /
                                det_limit_type / det_limit_side_len / use_cls）
        """
        global _ocr_reader, _ocr_reader_device, _ocr_reader_providers, _ocr_reader_params
        if language is not None:
            self._language = self._normalize_language(language)
            _ocr_reader = None
            _ocr_reader_device = None
            _ocr_reader_providers = None
            _ocr_reader_params = None
        if confidence_threshold is not None:
            self._confidence_threshold = confidence_threshold
        if device is not None and device != self._device:
            self._device = device
            _ocr_reader = None
            _ocr_reader_device = None
            _ocr_reader_providers = None
            _ocr_reader_params = None
        if preprocess is not None:
            self._preprocess.update(preprocess)
        if rapidocr_params is not None:
            self._rapidocr_params.update(rapidocr_params)
            _ocr_reader = None
            _ocr_reader_device = None
            _ocr_reader_providers = None
            _ocr_reader_params = None

    @staticmethod
    def _normalize_language(lang):
        """将各种语言格式统一为列表

        支持:
          - ['ch_sim', 'en']          → ['ch_sim', 'en']
          - "ch_sim,en"               → ['ch_sim', 'en']
          - "chi_sim+eng"             → ['ch_sim', 'en']
          - ["chi_sim", "eng"]        → ['ch_sim', 'en']
        """
        if isinstance(lang, list):
            codes = lang
        elif isinstance(lang, str):
            if '+' in lang:
                codes = lang.split('+')
            elif ',' in lang:
                codes = lang.split(',')
            else:
                codes = [lang]
        else:
            return ['ch_sim', 'en']

        return [OCREngine._LANG_MAP.get(c.strip(), c.strip()) for c in codes]

    def get_config(self):
        """返回当前配置"""
        return {
            "language": self._language,
            "confidence_threshold": self._confidence_threshold,
            "device": self._device,
            "preprocess": dict(self._preprocess),
            "rapidocr": dict(self._rapidocr_params),
        }

    def gpu_summary(self, force_refresh=False):
        """返回 GPU 加速状态摘要

        :param force_refresh: 强制重新检测 GPU（忽略缓存和 _detected 标志）
        """
        self._gpu.detect(force_refresh=force_refresh)
        lines = self._gpu.summary().split("\n")
        # 追加 ONNX 运行时提供器信息
        providers = _detect_onnx_providers()
        if providers:
            gpu_providers = [
                p for p in providers
                if p != "CPUExecutionProvider" and p != "AzureExecutionProvider"
            ]
            if gpu_providers:
                lines.append(
                    f"ONNX 可用后端: {' → '.join(gpu_providers)}"
                )

        # 追加当前 Reader 的实际 provider
        actual = get_reader_providers()
        if actual:
            backend_str = " + ".join(actual)
            if "CPUExecutionProvider" in actual and len(actual) > 1:
                backend_str += " (GPGPU 加速)"
            elif actual == ["CPUExecutionProvider"]:
                backend_str += " (无 GPU 加速)"
            lines.append(f"当前 OCR 后端: {backend_str}")

        return "\n".join(lines)

    # ---------- 图像预处理 ----------

    def _preprocess_image(self, img_bgr):
        """对截图进行 OCR 前的预处理（小图走 CPU，大图走 OpenCL）"""
        return preprocess_image(
            img_bgr,
            grayscale=self._preprocess.get("grayscale", True),
            scale=self._preprocess.get("scale", 1.0),
            threshold_enabled=self._preprocess.get("threshold_enabled", False),
            threshold_value=self._preprocess.get("threshold_value", 127),
        )

    # ---------- OCR 识别 ----------

    def recognize(self, screenshot_bgr):
        """对截图执行 OCR，返回所有识别到的文字信息

        :param screenshot_bgr: BGR 格式 numpy 数组（或已预处理图像，见 recognize_preprocessed）
        :return: list[dict]，每个 dict 包含:
            - "text": 识别文字
            - "confidence": 置信度 (0~100 整数)
            - "bbox": [x, y, w, h] 文字区域
        """
        processed = self._preprocess_image(screenshot_bgr)
        return self.recognize_preprocessed(processed)

    def recognize_preprocessed(self, processed):
        """对已预处理的图像执行 OCR（跳过重复预处理，供多 target 共享推理结果）

        :param processed: 已经过 _preprocess_image 处理的 numpy 数组
        :return: list[dict]，与 recognize() 相同的输出格式
        """
        reader = _get_reader(self._device, rapidocr_params=self._rapidocr_params)
        ocr_result, _elapse = reader(processed)
        return self._parse_ocr_result(ocr_result)

    def _parse_ocr_result(self, ocr_result):
        """将 RapidOCR 原始输出转换为统一的 dict 列表"""
        output = []
        if ocr_result is None:
            return output

        for box, text, score in ocr_result:
            if text is None:
                continue
            if score < self._confidence_threshold:
                continue
            text = text.strip()
            if not text:
                continue

            xs = [p[0] for p in box]
            ys = [p[1] for p in box]
            x, y = int(min(xs)), int(min(ys))
            w, h = int(max(xs) - x), int(max(ys) - y)

            output.append({
                "text": text,
                "confidence": int(score * 100),
                "bbox": [x, y, w, h],
            })

        return output

    def recognize_text_only(self, screenshot_bgr):
        """便捷方法：只返回识别到的纯文本列表"""
        return [r["text"] for r in self.recognize(screenshot_bgr)]

    # ---------- 截图区域裁剪 ----------

    def _crop_region(self, screenshot_bgr, region):
        """从截图中裁剪指定区域"""
        if region is None:
            return screenshot_bgr

        x, y, w, h = region
        h_img, w_img = screenshot_bgr.shape[:2]

        x = max(0, x)
        y = max(0, y)
        w = min(w, w_img - x)
        h = min(h, h_img - y)

        if w <= 0 or h <= 0:
            return screenshot_bgr

        return screenshot_bgr[y:y + h, x:x + w]

    # ---------- 匹配 ----------

    def _match_text(self, ocr_texts, target_text, match_mode):
        """检查目标文字是否出现在 OCR 结果中"""
        if match_mode == "exact":
            return any(t == target_text for t in ocr_texts)
        elif match_mode == "regex":
            import re
            try:
                pattern = re.compile(target_text)
                return any(pattern.search(t) for t in ocr_texts)
            except re.error:
                return False
        else:
            return any(target_text in t for t in ocr_texts)

    def match(self, screenshot_bgr, task_config):
        """对截图执行 OCR，按 targets 依次匹配，返回首个命中的 actions

        :return: (actions, matched_desc) 或 (None, None)
        """
        targets = task_config.get("targets", [])
        if not targets:
            return None, None

        for target in targets:
            region = target.get("region", None)
            target_text = target.get("text", "")
            match_mode = target.get("match_mode", "contains")

            if not target_text:
                continue

            cropped = self._crop_region(screenshot_bgr, region)
            ocr_texts = self.recognize_text_only(cropped)

            is_match = self._match_text(ocr_texts, target_text, match_mode)
            reverse = target.get("reverse_match", False)
            effective_match = (not reverse and is_match) or (
                reverse and not is_match
            )

            if effective_match:
                return target.get("actions", []), target.get("desc", "")

        return None, None

    # ---------- 任务文件加载 ----------

    def load_ocr_task(self, task_name, task_ocr_dir=None):
        """加载 task_ocr 目录下的 JSON 任务文件"""
        if task_ocr_dir is None:
            task_ocr_dir = os.path.join(
                os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "task_ocr"
            )
        task_path = os.path.join(task_ocr_dir, f"{task_name}.json")
        if not os.path.exists(task_path):
            raise FileNotFoundError(f"OCR 任务文件不存在: {task_path}")

        with open(task_path, "r", encoding="utf-8") as f:
            task_config = json.load(f)

        return task_config

    def list_ocr_tasks(self, task_ocr_dir=None):
        """列出 task_ocr 目录下所有可用的 OCR 任务名称"""
        if task_ocr_dir is None:
            task_ocr_dir = os.path.join(
                os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "task_ocr"
            )
        if not os.path.exists(task_ocr_dir):
            return []

        return sorted([
            f[:-5] for f in os.listdir(task_ocr_dir)
            if f.endswith(".json")
        ])

    def apply_ocr_settings(self, task_config):
        """从任务配置中读取 ocr_settings 并应用"""
        settings = task_config.get("ocr_settings", {})
        if settings:
            lang = settings.get("language")
            conf = settings.get("confidence_threshold")
            pre = settings.get("preprocess")
            self.configure(language=lang, confidence_threshold=conf,
                           preprocess=pre)