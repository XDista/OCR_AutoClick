"""
OCR 测试模块 - 独立于配置文件，线程安全

用法:
    from ocr_test import OCRTest

    test = OCRTest(
        hwnd=hwnd,
        device="dml:0",
        language="ch_sim,en",
        confidence_threshold=0.5,
        screenshot_mode="Win32GUI",
        adb_device_serial="",
        on_progress=lambda msg: print(msg),
        on_complete=lambda result: print(result),
    )
    test.start()        # 开始测试（后台线程）
    test.cancel()       # 取消测试

completion 回调收到 dict:
    {
        "status": "done" / "cancelled" / "error",
        "duration_ms": 1234,
        "texts": [{"text": "...", "confidence": 95, "bbox": [x,y,w,h]}, ...],
        "text_count": 5,
    }
"""

import gc
import threading
import time
from gpu_accelerator import GPUAccelerator
from screenshot_capture import (
    capture_window_win32gui,
    capture_window_memory,
    capture_window_printwindow,
    capture_adb_screenshot,
)


class OCRTest:
    """独立的 OCR 测试工具

    - 不依赖 task_ocr / task_worker 等任务配置文件
    - 截图复用 screenshot_capture（ADB 模式会读取 config.ini）
    - 预处理通过 GPUAccelerator 获得 OpenCL 加速
    - 线程安全：通过 threading.Lock 保护状态，threading.Event 取消
    - 回调 on_progress / on_complete 在后台线程调用，GUI 使用者需自行 root.after
    """

    STATUS_IDLE = "idle"
    STATUS_RUNNING = "running"
    STATUS_DONE = "done"
    STATUS_CANCELLED = "cancelled"
    STATUS_ERROR = "error"

    def __init__(
        self,
        hwnd,
        *,
        device="cpu",
        language="ch_sim,en",
        confidence_threshold=0.5,
        screenshot_mode="Win32GUI",
        adb_device_serial="",
        grayscale=True,
        threshold_enabled=False,
        threshold_value=127,
        on_progress=None,
        on_complete=None,
    ):
        """
        :param hwnd: 目标窗口句柄（int），ADB 模式时可为 0
        :param device: 加速设备，如 "cpu" / "cuda:0" / "dml:1"
        :param language: 识别语言，如 "ch_sim,en"
        :param confidence_threshold: 置信度阈值 0.0~1.0
        :param screenshot_mode: 截图模式 "Win32GUI" / "Win32Memory" / "PrintWindow" / "ADB"
        :param adb_device_serial: ADB 设备序列号（仅 ADB 模式使用）
        :param grayscale: 是否灰度化
        :param threshold_enabled: 是否二值化
        :param threshold_value: 二值化阈值 0~255
        :param on_progress: 进度回调 callable(str)
        :param on_complete: 完成回调 callable(dict)
        """
        self._hwnd = hwnd
        self._device = device
        self._language = language
        self._confidence_threshold = confidence_threshold
        self._screenshot_mode = screenshot_mode
        self._adb_device_serial = adb_device_serial
        self._grayscale = grayscale
        self._threshold_enabled = threshold_enabled
        self._threshold_value = threshold_value
        self._on_progress = on_progress
        self._on_complete = on_complete

        self._lock = threading.Lock()
        self._cancel_event = threading.Event()
        self._status = self.STATUS_IDLE

    # ---------- public ----------

    @property
    def status(self):
        with self._lock:
            return self._status

    @property
    def is_running(self):
        return self.status == self.STATUS_RUNNING

    def start(self):
        """启动 OCR 测试（后台线程），重复调用会被忽略"""
        with self._lock:
            if self._status == self.STATUS_RUNNING:
                return False
            self._status = self.STATUS_RUNNING
            self._cancel_event.clear()

        thread = threading.Thread(target=self._run, daemon=True)
        thread.start()
        return True

    def cancel(self):
        """请求取消当前测试"""
        with self._lock:
            if self._status != self.STATUS_RUNNING:
                return
        self._cancel_event.set()

    # ---------- internal ----------

    def _emit_progress(self, msg):
        if self._on_progress:
            try:
                self._on_progress(str(msg))
            except Exception:
                pass

    def _emit_complete(self, result):
        if self._on_complete:
            try:
                self._on_complete(result)
            except Exception:
                pass

    def _is_cancelled(self):
        return self._cancel_event.is_set()

    def _finish(self, status, result):
        with self._lock:
            self._status = status
        if result is None:
            result = {"status": status, "duration_ms": 0, "texts": [], "text_count": 0}
        else:
            result["status"] = status
        self._emit_complete(result)

    def _run(self):
        """后台主流程（在工作线程中执行）"""
        start_time = time.perf_counter()

        try:
            # ---- 开始 ----
            self._emit_progress("OCR 测试开始")
            self._emit_progress(
                f"  设备: {self._device}"
                f"  |  语言: {self._language}"
                f"  |  置信度阈值: {self._confidence_threshold}"
            )
            self._emit_progress("  正在加载模型（首次较慢），请稍候 ...")

            if self._is_cancelled():
                self._finish(self.STATUS_CANCELLED, None)
                return

            # ---- 截图 ----
            self._emit_progress("  正在截取窗口 ...")
            screenshot = _capture(
                hwnd=self._hwnd,
                mode=self._screenshot_mode,
                adb_device_serial=self._adb_device_serial,
            )
            if screenshot is None:
                self._emit_progress("  ❌ 截图失败")
                self._finish(self.STATUS_ERROR, None)
                return

            if self._is_cancelled():
                self._finish(self.STATUS_CANCELLED, None)
                return

            self._emit_progress(
                f"  截图尺寸: {screenshot.shape[1]}x{screenshot.shape[0]}"
            )

            # ---- 预处理 ----
            processed = _preprocess(
                screenshot,
                grayscale=self._grayscale,
                threshold_enabled=self._threshold_enabled,
                threshold_value=self._threshold_value,
            )

            # ---- OCR ----
            self._emit_progress("  OCR 识别中 ...")
            from ocr_engine import _get_reader
            reader = _get_reader(self._device)
            ocr_result, _elapse = reader(processed)

            # ---- 整理结果 ----
            texts = []
            if ocr_result:
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

                    texts.append({
                        "text": text,
                        "confidence": int(score * 100),
                        "bbox": [x, y, w, h],
                    })

            duration_ms = round((time.perf_counter() - start_time) * 1000)

            result = {
                "duration_ms": duration_ms,
                "texts": texts,
                "text_count": len(texts),
            }
            self._finish(self.STATUS_DONE, result)

        except ImportError as e:
            self._emit_progress(f"  ❌ 依赖缺失: {e}")
            self._finish(self.STATUS_ERROR, None)
        except Exception:
            import traceback
            self._emit_progress(f"  ❌ 异常:\n{traceback.format_exc()}")
            self._finish(self.STATUS_ERROR, None)
        finally:
            try:
                del screenshot
            except Exception:
                pass
            try:
                del processed
            except Exception:
                pass
            gc.collect()


# ==================== 图像预处理 ====================

def _preprocess(img_bgr, *, grayscale, threshold_enabled, threshold_value):
    """图像预处理（使用 GPUAccelerator 获得 OpenCL 加速）"""
    gpu = GPUAccelerator()
    gpu.detect()
    return gpu.preprocess_image(
        img_bgr,
        grayscale=grayscale or threshold_enabled,
        threshold_enabled=threshold_enabled,
        threshold_value=threshold_value,
    )


# ==================== 截图（复用 screenshot_capture） ====================

_CAPTURE_DISPATCH = {
    "Win32GUI": capture_window_win32gui,
    "Win32Memory": capture_window_memory,
    "PrintWindow": capture_window_printwindow,
    "ADB": capture_adb_screenshot,
}


def _capture(hwnd, mode, adb_device_serial=""):
    import win32gui
    fn = _CAPTURE_DISPATCH.get(mode)
    if fn is None:
        raise ValueError(f"未知的截图模式: {mode}（支持: {', '.join(_CAPTURE_DISPATCH)}）")
    if not win32gui.IsWindow(hwnd) and mode != "ADB":
        return None
    if mode == "ADB":
        return fn(device_serial=adb_device_serial)
    return fn(hwnd)