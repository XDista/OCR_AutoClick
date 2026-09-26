import time
import configparser
import psutil
import pygetwindow as gw
import win32gui
import win32con
import win32process
from project_paths import MAIN_CONFIG_PATH
from app_registry import get_app


def get_window_client_rect(hwnd):
    """获取窗口客户区的屏幕绝对坐标和尺寸
    :return: client_left, client_top, client_width, client_height
    """
    client_rect = win32gui.GetClientRect(hwnd)
    client_width = client_rect[2] - client_rect[0]
    client_height = client_rect[3] - client_rect[1]
    client_left, client_top = win32gui.ClientToScreen(hwnd, (client_rect[0], client_rect[1]))
    return client_left, client_top, client_width, client_height


def get_all_visible_windows_simple():
    """获取窗口列表（标题+进程名+句柄，用于下拉列表）"""
    windows = []

    def callback(hwnd, extra):
        if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowTextLength(hwnd) > 0:
            title = win32gui.GetWindowText(hwnd)
            try:
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                process = psutil.Process(pid)
                process_name = process.name()
            except Exception:
                process_name = "未知进程"
            display_text = f"{title[:50]} ({process_name})"
            windows.append({
                "display": display_text,
                "hwnd": hwnd,
                "title": title,
                "process_name": process_name,
            })

    win32gui.EnumWindows(callback, None)
    return windows


def get_target_window_from_config():
    """从配置读取默认窗口（优先模糊匹配）"""
    config = configparser.ConfigParser()
    config.read(MAIN_CONFIG_PATH, encoding="utf-8")
    target_program = config["GENERAL"]["target_program_name"]
    target_title = config["GENERAL"]["target_window_title"]
    target_hwnd = config["GENERAL"]["target_window_hwnd"]

    if target_hwnd and target_hwnd.isdigit():
        hwnd = int(target_hwnd)
        if win32gui.IsWindow(hwnd) and win32gui.IsWindowVisible(hwnd):
            try:
                return gw.Window(hwnd)
            except Exception:
                pass

    if target_title:
        for win in gw.getAllWindows():
            if win32gui.IsWindowVisible(win._hWnd) and target_title in win.title:
                return win

    if target_program:
        pid_list = []
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                if proc.info['name'] and target_program.lower() in proc.info['name'].lower():
                    pid_list.append(proc.info['pid'])
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue

        win_list = []

        def enum_windows_callback(hwnd, extra):
            if win32gui.IsWindowVisible(hwnd):
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                if pid in extra["pid_list"]:
                    try:
                        win = gw.Window(hwnd)
                        win_list.append(win)
                    except Exception:
                        pass
            return True

        win32gui.EnumWindows(enum_windows_callback, {"pid_list": pid_list})
        if win_list:
            return win_list[0]

    return None


def is_window_visible(hwnd):
    try:
        return (win32gui.IsWindowVisible(hwnd)
                and win32gui.GetWindowText(hwnd) != "")
    except Exception:
        return False


def window_relative_to_screen(window, x, y):
    return win32gui.ClientToScreen(window._hWnd, (x, y))


def bring_window_to_front(window):
    try:
        hwnd = window._hWnd
        if not win32gui.IsWindow(hwnd):
            raise RuntimeError("窗口已关闭")

        for _ in range(3):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd)
            if win32gui.GetForegroundWindow() == hwnd:
                time.sleep(0.2)
                return
        raise RuntimeError("窗口激活失败，可能被其他程序阻塞")
    except Exception as e:
        app = get_app()
        try:
            if app:
                app.log(f"⚠️ 窗口置顶失败：{e}")
        except Exception:
            pass
        raise