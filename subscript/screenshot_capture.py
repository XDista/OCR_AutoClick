import os
import subprocess
import ctypes
from ctypes import wintypes
import numpy as np
import cv2
import pyautogui
import win32gui
import win32con
import win32ui
from PIL import Image

from project_paths import BASE_DIR
from window_utils import get_window_client_rect, is_window_visible
from app_registry import get_app


# ========== ADB工具函数 ==========

def get_adb_executable_path():
    """根据配置获取ADB可执行文件路径
    :return: (是否有效, adb路径/错误信息)
    """
    import configparser
    main_config = configparser.ConfigParser()
    main_config.read(os.path.join(BASE_DIR, "config.ini"), encoding="utf-8")
    adb_enabled = main_config["ADBConfig"].get("adb_enabled", "1")
    adb_usage_mode = main_config["ADBConfig"].get("adb_usage_mode", "1")
    adb_position = main_config["ADBConfig"].get("adb_position", "").strip()

    if adb_enabled != "1":
        return False, "ADB功能已禁用（adb_enabled=0）"

    if adb_usage_mode == "1":
        adb_path = os.path.join(BASE_DIR, "adb", "platform-tools", "adb.exe")
        if os.path.exists(adb_path):
            return True, adb_path
        else:
            return False, f"模式1 - ADB文件不存在：{adb_path}（请检查脚本目录下的adb子文件夹）"

    elif adb_usage_mode == "2":
        if not adb_position:
            return False, "模式2 - 未配置自定义ADB路径（adb_position为空）"
        if os.path.exists(adb_position) and adb_position.endswith("adb.exe"):
            return True, adb_position
        else:
            return False, f"模式2 - 自定义ADB路径无效：{adb_position}"

    elif adb_usage_mode == "3":
        try:
            result = subprocess.run(
                "adb version", shell=True, check=True,
                stdout=subprocess.PIPE, stderr=subprocess.PIPE, encoding="utf-8",
            )
            return True, "adb（系统环境变量）"
        except subprocess.CalledProcessError as e:
            return False, f"模式3 - 系统环境变量中ADB执行失败：{e.stderr}"
        except FileNotFoundError:
            return False, "模式3 - 系统环境变量中未找到ADB"

    elif adb_usage_mode == "4":
        where_alas = main_config["GENERAL"].get("where_alas", "").strip()
        if not where_alas:
            return False, "模式4 - 未配置Alas路径（where_alas为空）"
        if not os.path.exists(where_alas):
            return False, f"模式4 - Alas路径不存在：{where_alas}"
        alas_dir = os.path.dirname(where_alas)
        adb_path = os.path.join(alas_dir, "toolkit", "Lib", "site-packages", "adbutils", "binaries", "adb.exe")
        if os.path.exists(adb_path):
            return True, adb_path
        else:
            return False, f"模式4 - Alas附带ADB不存在：{adb_path}"

    else:
        return False, f"无效的ADB使用模式：{adb_usage_mode}（仅支持1/2/3/4）"


def validate_adb_environment():
    """验证ADB环境是否可用"""
    import configparser
    main_config = configparser.ConfigParser()
    main_config.read(os.path.join(BASE_DIR, "config.ini"), encoding="utf-8")
    adb_usage_mode = main_config["ADBConfig"].get("adb_usage_mode", "1")

    is_valid, msg = get_adb_executable_path()
    if not is_valid:
        return False, msg

    adb_path = msg if is_valid else ""
    try:
        cmd = f'"{adb_path}" version' if adb_usage_mode != "3" else "adb version"
        result = subprocess.run(
            cmd, shell=True, check=True,
            stdout=subprocess.PIPE, stderr=subprocess.PIPE, encoding="utf-8", timeout=10,
        )
        version_info = result.stdout.strip().split("\n")[0] if result.stdout else "未知版本"
        return True, f"ADB环境验证通过：{adb_path if adb_usage_mode != 3 else '系统环境变量'} | {version_info}"
    except Exception as e:
        return False, f"ADB可执行但版本检查失败：{str(e)}"


# ========== 截图函数 ==========

def capture_adb_screenshot(device_serial=""):
    """通过ADB获取设备截图，转为OpenCV BGR格式"""
    app = get_app()
    adb_valid, adb_path = get_adb_executable_path()
    if not adb_valid:
        app.log(f"ADB截图失败：{adb_path}")
        return None

    adb_cmd_prefix = f'"{adb_path}"'
    if device_serial.strip():
        adb_cmd_prefix += f' -s {device_serial.strip()}'

    temp_png = os.path.join(BASE_DIR, "adb_screenshot_temp.png")
    if os.path.exists(temp_png):
        os.remove(temp_png)

    adb_screenshot_cmd = f'{adb_cmd_prefix} exec-out screencap -p > "{temp_png}"'
    try:
        result = subprocess.run(
            adb_screenshot_cmd, shell=True, check=True,
            stdout=subprocess.PIPE, stderr=subprocess.PIPE, encoding="utf-8", timeout=10,
        )
        if not os.path.exists(temp_png) or os.path.getsize(temp_png) == 0:
            app.log(f"ADB截图命令执行成功，但临时文件为空：{temp_png}")
            return None

        screenshot_pil = Image.open(temp_png).convert('RGB')
        img = cv2.cvtColor(np.array(screenshot_pil), cv2.COLOR_RGB2BGR)
        os.remove(temp_png)
        app.log(f"ADB截图成功（设备：{device_serial or '默认'}）：{img.shape[1]}x{img.shape[0]}")
        return img
    except subprocess.CalledProcessError as e:
        app.log(f"ADB截图命令执行失败：{e.stderr}")
        return None
    except subprocess.TimeoutExpired:
        app.log(f"ADB截图命令超时（10秒）")
        return None
    except Exception as e:
        app.log(f"ADB截图处理失败：{str(e)}")
        return None


def capture_window_memory(hwnd):
    """从窗口内存直接读取像素（Win32 Memory模式）

    :param hwnd: 窗口句柄
    :return: OpenCV BGR图像 / None（失败）
    """
    app = get_app()
    try:
        window_rect = win32gui.GetWindowRect(hwnd)
        window_left, window_top = window_rect[0], window_rect[1]

        client_left, client_top, client_width, client_height = get_window_client_rect(hwnd)
        if client_width <= 0 or client_height <= 0:
            app.log(f"窗口客户区尺寸无效：{client_width}x{client_height}")
            return None

        offset_x = client_left - window_left
        offset_y = client_top - window_top
        if offset_x < 0 or offset_y < 0:
            app.log(f"客户区偏移异常：offset_x={offset_x}, offset_y={offset_y}")
            return None

        hdc_window = win32gui.GetDC(hwnd)
        hdc_mem = win32ui.CreateDCFromHandle(hdc_window)
        mem_dc = hdc_mem.CreateCompatibleDC()

        bmp = win32ui.CreateBitmap()
        bmp.CreateCompatibleBitmap(hdc_mem, client_width, client_height)
        mem_dc.SelectObject(bmp)

        mem_dc.BitBlt(
            (0, 0), (client_width, client_height),
            hdc_mem, (offset_x, offset_y), win32con.SRCCOPY,
        )

        signed_ints_array = bmp.GetBitmapBits(True)
        img = np.frombuffer(signed_ints_array, dtype='uint8')
        img.shape = (client_height, client_width, 4)
        img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)

        app.log(f"win32memory截图成功：{client_width}x{client_height} | 窗口内偏移：({offset_x},{offset_y})")
        return img

    except Exception as e:
        app.log(f"win32memory截图失败：{str(e)}")
        return None

    finally:
        try:
            if 'bmp' in locals():
                win32gui.DeleteObject(bmp.GetHandle())
            if 'mem_dc' in locals():
                mem_dc.DeleteDC()
            if 'hdc_mem' in locals():
                hdc_mem.DeleteDC()
            if 'hdc_window' in locals() and hwnd:
                win32gui.ReleaseDC(hwnd, hdc_window)
        except Exception as release_e:
            app.log(f"win32memory资源释放警告：{str(release_e)}")


def capture_window_printwindow(hwnd):
    """使用PrintWindow API截图（支持被遮挡窗口+硬件加速应用）

    :param hwnd: 窗口句柄
    :return: OpenCV BGR图像 / None（失败）
    """
    app = get_app()
    try:
        client_left, client_top, client_width, client_height = get_window_client_rect(hwnd)
        if client_width <= 0 or client_height <= 0:
            app.log("PrintWindow截图失败：客户区尺寸无效")
            return None

        hdc_window = win32gui.GetDC(hwnd)
        hdc_mem = win32ui.CreateDCFromHandle(hdc_window)
        mem_dc = hdc_mem.CreateCompatibleDC()

        bmp = win32ui.CreateBitmap()
        bmp.CreateCompatibleBitmap(hdc_mem, client_width, client_height)
        mem_dc.SelectObject(bmp)

        user32 = ctypes.WinDLL('user32.dll')
        LRESULT = wintypes.LONG

        user32.PrintWindow.argtypes = [wintypes.HWND, wintypes.HDC, wintypes.UINT]
        user32.PrintWindow.restype = wintypes.BOOL

        user32.SendMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
        user32.SendMessageW.restype = LRESULT

        PW_CLIENTONLY = 0x00000001
        PW_RENDERFULLCONTENT = 0x00000002
        WM_PAINT = 0x000F
        WM_PRINT = 0x0317
        PRF_CLIENT = 0x00000004

        user32.SendMessageW(hwnd, WM_PAINT, 0, 0)
        user32.SendMessageW(hwnd, WM_PRINT, mem_dc.GetSafeHdc(), PRF_CLIENT)

        flags = PW_CLIENTONLY | PW_RENDERFULLCONTENT
        result = user32.PrintWindow(hwnd, mem_dc.GetSafeHdc(), flags)

        if not result:
            app.log("PrintWindow截图失败：API调用失败")
            win32gui.DeleteObject(bmp.GetHandle())
            mem_dc.DeleteDC()
            hdc_mem.DeleteDC()
            win32gui.ReleaseDC(hwnd, hdc_window)
            return None

        signed_ints_array = bmp.GetBitmapBits(True)
        img = np.frombuffer(signed_ints_array, dtype='uint8')
        img.shape = (client_height, client_width, 4)
        img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)

        win32gui.DeleteObject(bmp.GetHandle())
        mem_dc.DeleteDC()
        hdc_mem.DeleteDC()
        win32gui.ReleaseDC(hwnd, hdc_window)

        app.log(f"PrintWindow截图成功：{client_width}x{client_height}")
        return img

    except Exception as e:
        app.log(f"PrintWindow截图失败：{str(e)}")
        return None


def capture_window_win32gui(hwnd):
    """Win32GUI模式截图：基于pyautogui截取窗口客户区

    :param hwnd: 窗口句柄
    :return: OpenCV BGR图像 / None（失败）
    """
    app = get_app()
    try:
        if not win32gui.IsWindow(hwnd) or not is_window_visible(hwnd):
            app.log("win32gui截图失败：窗口无效/不可见/无标题")
            return None

        client_left, client_top, client_width, client_height = get_window_client_rect(hwnd)

        if client_width <= 0 or client_height <= 0:
            app.log(f"截图失败：窗口客户区尺寸无效：{client_width}x{client_height}")
            return None

        screenshot_pil = pyautogui.screenshot(region=(client_left, client_top, client_width, client_height))
        img = cv2.cvtColor(np.array(screenshot_pil), cv2.COLOR_RGB2BGR)

        app.log(f"win32gui截图成功（客户区）：{img.shape[1]}x{img.shape[0]}")
        return img

    except Exception as e:
        app.log(f"win32gui截图失败：{str(e)}")
        return None


def capture_window(window, screenshot_mode="Win32GUI", adb_device_serial=""):
    """截图统一入口"""
    app = get_app()
    hdc_window = None
    hdc_memdc = None
    hbitmap = None
    hwnd = None

    try:
        if screenshot_mode == "ADB":
            img = capture_adb_screenshot(device_serial=adb_device_serial)
            if img is None:
                app.log("ADB截图模式：截图失败")
                return None
            return img

        elif screenshot_mode == "Win32GUI":
            if not window or not win32gui.IsWindow(window._hWnd):
                app.log("win32gui截图失败：窗口句柄无效/已关闭")
                return None
            hwnd = window._hWnd
            img = capture_window_win32gui(hwnd)
            return img

        elif screenshot_mode == "Win32Memory":
            if not window or not win32gui.IsWindow(window._hWnd):
                app.log("win32memory截图失败：窗口句柄无效/已关闭")
                return None
            hwnd = window._hWnd
            if not win32gui.IsWindowVisible(hwnd) or win32gui.GetWindowText(hwnd) == "":
                app.log("win32memory截图失败：窗口不可见/无标题")
                return None
            img = capture_window_memory(hwnd)
            return img

        elif screenshot_mode == "PrintWindow":
            if not window or not win32gui.IsWindow(window._hWnd):
                app.log("PrintWindow截图失败：窗口句柄无效/已关闭")
                return None
            hwnd = window._hWnd
            img = capture_window_printwindow(hwnd)
            return img

        else:
            app.log(f"无效的截图模式：{screenshot_mode}（仅支持Win32GUI/PrintWindow/Win32Memory/ADB）")
            return None
    except Exception as e:
        app.log(f"截图失败：{str(e)}")
        return None
    finally:
        if hbitmap and hdc_memdc:
            try:
                hdc_memdc.SelectObject(None)
            except Exception:
                pass
        if hbitmap:
            try:
                win32gui.DeleteObject(hbitmap.GetHandle())
            except Exception:
                app.log("警告：位图对象删除失败")
        if hdc_memdc:
            try:
                hdc_memdc.DeleteDC()
            except Exception:
                pass
        if hdc_window and hwnd:
            try:
                win32gui.ReleaseDC(hwnd, hdc_window)
            except Exception:
                pass