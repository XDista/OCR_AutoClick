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

import threading
import time

import cv2
import numpy as np


class OCRTest:
    """独立的 OCR 测试工具

    - 不依赖任何配置文件（config.ini / task_ocr / MAIN_CONFIG_PATH 等）
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
        grayscale=True,
        threshold_enabled=False,
        threshold_value=127,
        on_progress=None,
        on_complete=None,
    ):
        """
        :param hwnd: 目标窗口句柄（int）
        :param device: 加速设备，如 "cpu" / "cuda:0" / "dml:1"
        :param language: 识别语言，如 "ch_sim,en"
        :param confidence_threshold: 置信度阈值 0.0~1.0
        :param screenshot_mode: 截图模式 "Win32GUI" / "Win32Memory" / "PrintWindow"
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
            screenshot = _capture(hwnd=self._hwnd, mode=self._screenshot_mode)
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


# ==================== 图像预处理 ====================

def _preprocess(img_bgr, *, grayscale, threshold_enabled, threshold_value):
    img = img_bgr
    if grayscale:
        if len(img.shape) == 3:
            img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    else:
        if len(img.shape) == 3:
            img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)

    if threshold_enabled:
        if len(img.shape) == 3:
            img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        _, img = cv2.threshold(img, threshold_value, 255, cv2.THRESH_BINARY)

    return img


# ==================== 截图（自包含，不依赖 app / config） ====================

_CAPTURE_REGISTRY = {}


def _register(name):
    def dec(fn):
        _CAPTURE_REGISTRY[name] = fn
        return fn
    return dec


def _capture(hwnd, mode):
    fn = _CAPTURE_REGISTRY.get(mode)
    if fn is None:
        raise ValueError(f"未知的截图模式: {mode}（支持: {', '.join(_CAPTURE_REGISTRY)}）")
    return fn(hwnd)


@_register("Win32GUI")
def _capture_win32gui(hwnd):
    """pyautogui 屏幕坐标截图（需要窗口可见且不被遮挡）"""
    import pyautogui
    import win32gui
    from window_utils import get_window_client_rect, is_window_visible

    try:
        if not win32gui.IsWindow(hwnd) or not is_window_visible(hwnd):
            return None

        left, top, width, height = get_window_client_rect(hwnd)
        if width <= 0 or height <= 0:
            return None

        pil_img = pyautogui.screenshot(region=(left, top, width, height))
        return cv2.cvtColor(np.array(pil_img), cv2.COLOR_RGB2BGR)
    except Exception:
        return None


@_register("Win32Memory")
def _capture_memory(hwnd):
    """从窗口 DC 读取像素（不依赖窗口位置，但对部分应用无效）"""
    import win32gui
    import win32ui
    import win32con
    from window_utils import get_window_client_rect

    try:
        if not win32gui.IsWindow(hwnd):
            return None

        wr = win32gui.GetWindowRect(hwnd)
        left, top, width, height = get_window_client_rect(hwnd)
        if width <= 0 or height <= 0:
            return None

        ox = left - wr[0]
        oy = top - wr[1]
        if ox < 0 or oy < 0:
            return None

        hdc_win = win32gui.GetDC(hwnd)
        hdc_mem = win32ui.CreateDCFromHandle(hdc_win)
        mem_dc = hdc_mem.CreateCompatibleDC()

        bmp = win32ui.CreateBitmap()
        bmp.CreateCompatibleBitmap(hdc_mem, width, height)
        mem_dc.SelectObject(bmp)

        mem_dc.BitBlt(
            (0, 0), (width, height),
            hdc_mem, (ox, oy), win32con.SRCCOPY,
        )

        data = bmp.GetBitmapBits(True)
        img = np.frombuffer(data, dtype="uint8")
        img.shape = (height, width, 4)

        win32gui.DeleteObject(bmp.GetHandle())
        mem_dc.DeleteDC()
        hdc_mem.DeleteDC()
        win32gui.ReleaseDC(hwnd, hdc_win)

        return cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
    except Exception:
        return None


@_register("PrintWindow")
def _capture_printwindow(hwnd):
    """PrintWindow API 截图（支持被遮挡窗口和硬件加速应用）"""
    import ctypes
    from ctypes import wintypes
    import win32gui
    import win32ui
    from window_utils import get_window_client_rect

    try:
        if not win32gui.IsWindow(hwnd):
            return None

        left, top, width, height = get_window_client_rect(hwnd)
        if width <= 0 or height <= 0:
            return None

        hdc_win = win32gui.GetDC(hwnd)
        hdc_mem = win32ui.CreateDCFromHandle(hdc_win)
        mem_dc = hdc_mem.CreateCompatibleDC()

        bmp = win32ui.CreateBitmap()
        bmp.CreateCompatibleBitmap(hdc_mem, width, height)
        mem_dc.SelectObject(bmp)

        user32 = ctypes.WinDLL("user32.dll")
        LRESULT = wintypes.LONG

        user32.PrintWindow.argtypes = [wintypes.HWND, wintypes.HDC, wintypes.UINT]
        user32.PrintWindow.restype = wintypes.BOOL

        user32.SendMessageW.argtypes = [
            wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM,
        ]
        user32.SendMessageW.restype = LRESULT

        PW_CLIENTONLY = 0x00000001
        PW_RENDERFULLCONTENT = 0x00000002
        WM_PAINT = 0x000F
        WM_PRINT = 0x0317
        PRF_CLIENT = 0x00000004

        user32.SendMessageW(hwnd, WM_PAINT, 0, 0)
        user32.SendMessageW(hwnd, WM_PRINT, mem_dc.GetSafeHdc(), PRF_CLIENT)

        flags = PW_CLIENTONLY | PW_RENDERFULLCONTENT
        _ = user32.PrintWindow(hwnd, mem_dc.GetSafeHdc(), flags)

        data = bmp.GetBitmapBits(True)
        img = np.frombuffer(data, dtype="uint8")
        img.shape = (height, width, 4)

        win32gui.DeleteObject(bmp.GetHandle())
        mem_dc.DeleteDC()
        hdc_mem.DeleteDC()
        win32gui.ReleaseDC(hwnd, hdc_win)

        return cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
    except Exception:
        return None