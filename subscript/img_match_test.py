"""
模板匹配基准测试模块 - 线程安全

用法:
    from img_match_test import ImgMatchTest

    test = ImgMatchTest(
        hwnd=hwnd,
        screenshot_mode="Win32GUI",
        adb_device_serial="",
        match_step=0.05,
        threshold=0.9,
        ref_image_path="refs/example/example.jpg",
        on_progress=lambda msg: print(msg),
        on_complete=lambda result: print(result),
    )
    test.start()        # 开始测试（后台线程）
    test.cancel()       # 取消测试

completion 回调收到 dict:
    {
        "status": "done" / "cancelled" / "error",
        "duration_ms": 1234,
        "is_match": True,
        "similarity": 0.9876,
        "ref_path": "refs/example/example.jpg",
        "screenshot_size": "1920x1080",
    }
"""

import os
import threading
import time
import cv2

from project_paths import REFS_DIR
from screenshot_capture import capture_window
from image_matcher import template_match


class _WindowProxy:
    """供 capture_window 使用的轻量窗口代理"""
    __slots__ = ("_hWnd",)

    def __init__(self, hwnd):
        self._hWnd = hwnd


class ImgMatchTest:
    """独立的模板匹配基准测试工具

    - 截图直接调用 screenshot_capture.capture_window（与 task_worker 完全一致）
    - 匹配直接调用 image_matcher.template_match（与 task_worker 完全一致）
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
        screenshot_mode="Win32GUI",
        adb_device_serial="",
        match_step=0.05,
        threshold=0.9,
        ref_image_path=None,
        on_progress=None,
        on_complete=None,
    ):
        """
        :param hwnd: 目标窗口句柄（int），ADB 模式时可为 0
        :param screenshot_mode: 截图模式 "Win32GUI" / "Win32Memory" / "PrintWindow" / "ADB"
        :param adb_device_serial: ADB 设备序列号（仅 ADB 模式使用）
        :param match_step: 多尺度匹配步长（0.001~0.5）
        :param threshold: 匹配阈值 0.0~1.0
        :param ref_image_path: 参考图路径，默认使用 refs/example/example.jpg
        :param on_progress: 进度回调 callable(str)
        :param on_complete: 完成回调 callable(dict)
        """
        self._hwnd = hwnd
        self._screenshot_mode = screenshot_mode
        self._adb_device_serial = adb_device_serial
        self._match_step = max(0.001, min(0.5, match_step))
        self._threshold = threshold
        self._ref_image_path = ref_image_path or os.path.join(REFS_DIR, "example", "example.jpg")
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
        """启动测试（后台线程），重复调用会被忽略"""
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
            result = {
                "status": status,
                "duration_ms": 0,
                "is_match": False,
                "similarity": 0.0,
                "ref_path": self._ref_image_path,
                "screenshot_size": "",
            }
        else:
            result["status"] = status
        self._emit_complete(result)

    def _run(self):
        """后台主流程（在工作线程中执行），与实际任务完全相同的逻辑"""
        start_time = time.perf_counter()

        try:
            # ---- 开始 ----
            self._emit_progress("模板匹配基准测试开始")
            self._emit_progress(
                f"  截图模式: {self._screenshot_mode}"
                f"  |  匹配步长: {self._match_step}"
                f"  |  阈值: {self._threshold}"
            )
            self._emit_progress(f"  参考图: {self._ref_image_path}")

            if self._is_cancelled():
                self._finish(self.STATUS_CANCELLED, None)
                return

            # ---- 检查参考图是否存在 ----
            if not os.path.exists(self._ref_image_path):
                self._emit_progress(f"  ❌ 参考图不存在: {self._ref_image_path}")
                self._finish(self.STATUS_ERROR, None)
                return

            # ---- 截图（直接调用 screenshot_capture.capture_window，与 task_worker 完全一致）----
            self._emit_progress("  正在截取窗口 ...")
            screenshot = capture_window(
                _WindowProxy(self._hwnd),
                screenshot_mode=self._screenshot_mode,
                adb_device_serial=self._adb_device_serial,
            )

            if screenshot is None:
                self._emit_progress("  ❌ 截图失败")
                self._finish(self.STATUS_ERROR, None)
                return

            if self._is_cancelled():
                self._finish(self.STATUS_CANCELLED, None)
                return

            screenshot_h, screenshot_w = screenshot.shape[:2]
            screenshot_size = f"{screenshot_w}x{screenshot_h}"
            self._emit_progress(f"  截图尺寸: {screenshot_size}")

            # ---- 转灰度（与 task_worker 完全一致）----
            screenshot_gray = cv2.cvtColor(screenshot, cv2.COLOR_BGR2GRAY)
            self._emit_progress("  截图已转为灰度图")

            # ---- 模板匹配（直接调用 image_matcher.template_match，与 task_worker 完全一致）----
            self._emit_progress("  正在进行模板匹配 ...")
            match_start = time.perf_counter()

            is_match, similarity = template_match(
                screenshot_gray,
                self._ref_image_path,
                threshold=self._threshold,
                match_step=self._match_step,
            )

            match_elapsed_ms = round((time.perf_counter() - match_start) * 1000)
            total_elapsed_ms = round((time.perf_counter() - start_time) * 1000)

            self._emit_progress(
                f"  匹配完成 | 相似度: {similarity:.4f}"
                f"  |  是否匹配: {'✅ 是' if is_match else '❌ 否'}"
                f"  |  匹配耗时: {match_elapsed_ms} ms"
                f"  |  总耗时: {total_elapsed_ms} ms"
            )

            del screenshot
            del screenshot_gray

            result = {
                "duration_ms": total_elapsed_ms,
                "match_elapsed_ms": match_elapsed_ms,
                "is_match": is_match,
                "similarity": round(float(similarity), 6),
                "ref_path": self._ref_image_path,
                "screenshot_size": screenshot_size,
            }
            self._finish(self.STATUS_DONE, result)

        except ImportError as e:
            self._emit_progress(f"  ❌ 依赖缺失: {e}")
            self._finish(self.STATUS_ERROR, None)
        except Exception:
            import traceback
            self._emit_progress(f"  ❌ 异常:\n{traceback.format_exc()}")
            self._finish(self.STATUS_ERROR, None)