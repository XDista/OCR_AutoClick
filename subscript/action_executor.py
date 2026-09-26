import os
import time
import datetime
import subprocess

import win32gui
import win32con
import pyautogui

from app_registry import get_app
from window_utils import get_window_client_rect, bring_window_to_front, is_window_visible


def send_windows_notification(title, message):
    """发送Windows系统通知"""
    from plyer import notification
    try:
        notification.notify(
            title=title,
            message=message,
            app_name="自动点击工具",
            timeout=15,
        )
    except Exception as e:
        print(f"Toast通知失败，使用备用弹窗：{e}")
        win32gui.MessageBox(
            0, message, title,
            win32con.MB_ICONINFORMATION | win32con.MB_OK | win32con.MB_TOPMOST,
        )


def auto_click(window, x, y, times=1, interval=0.5, click_mode="sendmessage"):
    """支持双模式的自动点击函数"""
    if not window or not is_window_visible(window._hWnd):
        raise RuntimeError("目标窗口不可见或已关闭")

    hwnd = window._hWnd

    if click_mode == "sendmessage":
        def send_mouse_down():
            l_param = y << 16 | x
            win32gui.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, l_param)

        def send_mouse_up():
            l_param = y << 16 | x
            win32gui.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, l_param)

        for _ in range(times):
            send_mouse_down()
            time.sleep(0.05)
            send_mouse_up()
            time.sleep(interval)

    elif click_mode == "pyautogui":
        client_left, client_top, _, _ = get_window_client_rect(hwnd)
        screen_x = client_left + x
        screen_y = client_top + y

        bring_window_to_front(window)
        time.sleep(0.1)

        for _ in range(times):
            pyautogui.moveTo(screen_x, screen_y, duration=0.05)
            pyautogui.click(screen_x, screen_y)
            time.sleep(interval)


def execute_action(window, action_type, params, stop_flag=False, click_mode="sendmessage", action_config=None):
    """执行具体动作

    :param action_config: 预读取的配置字典，包含 adb_device_serial, adb_usage_mode 等
    """
    app = get_app()

    if app and app.stop_flag:
        return "已触发停止指令，终止动作执行", True

    if action_type == "click":
        if len(params) >= 2:
            x, y = int(params[0]), int(params[1])
            auto_click(window, x, y, times=1, interval=0.1, click_mode=click_mode)
            return f"执行点击: ({x},{y}) | 模式: {click_mode}"
        else:
            return "点击动作参数不足"

    elif action_type == "sleep":
        if len(params) >= 1:
            t = float(params[0])
            app.log(f"  - 执行等待: {t}秒")
            start_time = time.monotonic()
            elapsed = 0.0
            app.log(f"【Sleep开始】预期等待{t}秒 | 开始时间: {datetime.datetime.now()}")
            while elapsed < t:
                if app and app.stop_flag:
                    app.log(f"【Sleep中断】预期{t}秒，实际等待{elapsed:.2f}秒")
                    return ("等待被中断（用户停止）", True, True)
                sleep_chunk = min(0.1, t - elapsed)
                time.sleep(sleep_chunk)
                elapsed = time.monotonic() - start_time
            app.log(f"【Sleep完成】预期{t}秒，实际等待{elapsed:.2f}秒")
            return ("", False)
        else:
            return ("等待动作参数不足", True)

    elif action_type == "press":
        keys = [k.strip() for k in params if k.strip()]
        if not keys:
            return "press动作参数不足（需指定至少1个按键名称，如enter、ctrl,a）"

        try:
            if click_mode == "sendmessage" and window:
                hwnd = window._hWnd
                key_map = {
                    'enter': 0x0D, 'return': 0x0D,
                    'tab': 0x09, 'space': 0x20,
                    'backspace': 0x08, 'delete': 0x2E,
                    'ctrl': 0x11, 'shift': 0x10, 'alt': 0x12,
                    'esc': 0x1B, 'capslock': 0x14,
                    'f1': 0x70, 'f2': 0x71, 'f3': 0x72, 'f4': 0x73,
                    'f5': 0x74, 'f6': 0x75, 'f7': 0x76, 'f8': 0x77,
                    'f9': 0x78, 'f10': 0x79, 'f11': 0x7A, 'f12': 0x7B,
                    'up': 0x26, 'down': 0x28, 'left': 0x25, 'right': 0x27,
                    'home': 0x24, 'end': 0x23, 'pageup': 0x21, 'pagedown': 0x22,
                    'insert': 0x2D,
                }
                for c in 'abcdefghijklmnopqrstuvwxyz':
                    key_map[c] = ord(c.upper())
                for c in '0123456789':
                    key_map[c] = ord(c)

                for key in keys:
                    vk_code = key_map.get(key.lower())
                    if vk_code is None:
                        return f"SendMessage模式不支持的按键：{key}"
                    win32gui.SendMessage(hwnd, win32con.WM_KEYDOWN, vk_code, 0)
                    time.sleep(0.05)
                    win32gui.SendMessage(hwnd, win32con.WM_KEYUP, vk_code, 0)

                return f"执行键盘消息发送（SendMessage模式）：{'+' .join(keys)}"

            else:
                if len(keys) == 1:
                    key = keys[0]
                    pyautogui.press(key)
                    return f"执行键盘单键按下：{key}"
                else:
                    pyautogui.hotkey(*keys)
                    return f"执行键盘组合键按下：{'+' .join(keys)}"

        except ValueError as e:
            invalid_keys = []
            for k in keys:
                try:
                    pyautogui.press(k)
                except ValueError:
                    invalid_keys.append(k)
            if invalid_keys:
                return f"无效的键盘按键名称「{'、'.join(invalid_keys)}」：{str(e)}"
            else:
                return f"组合键格式错误：{str(e)}"
        except Exception as e:
            return f"键盘按键执行失败（按键：{'+' .join(keys)}）：{str(e)}"

    elif action_type == "goto_task":
        if len(params) >= 1:
            try:
                task_num = int(params[0])
                target_index = task_num - 1
                return f"跳转到任务{task_num}", False, target_index
            except ValueError:
                return "goto_task参数必须是整数"
        else:
            return "goto_task参数不足"

    elif action_type == "taskcall":
        notify_content = ",".join(params)
        send_windows_notification("TaskCall：", notify_content)
        return f"发送Windows通知: {notify_content}"

    elif action_type == "swipe":
        try:
            if len(params) < 5:
                return f"拖动参数不足（正确格式：x1,y1,x2,y2,t），当前仅传入{len(params)}个参数"
            x1 = int(float(params[0]))
            y1 = int(float(params[1]))
            x2 = int(float(params[2]))
            y2 = int(float(params[3]))
            t = float(params[4])

            hwnd = window._hWnd

            if click_mode == "pyautogui":
                client_left, client_top, _, _ = get_window_client_rect(hwnd)
                screen_x1 = client_left + x1
                screen_y1 = client_top + y1
                screen_x2 = client_left + x2
                screen_y2 = client_top + y2

                bring_window_to_front(window)
                time.sleep(0.1)

                pyautogui.moveTo(screen_x1, screen_y1, duration=0.05)
                pyautogui.mouseDown(screen_x1, screen_y1)
                pyautogui.moveTo(screen_x2, screen_y2, duration=t)
                pyautogui.mouseUp(screen_x2, screen_y2)

                return f"执行拖动（pyautogui模式）：窗口内({x1},{y1})→({x2},{y2}) | 屏幕({screen_x1},{screen_y1})→({screen_x2},{screen_y2}) | 耗时{t}秒"

            elif click_mode == "sendmessage":
                l_param_down = y1 << 16 | x1
                win32gui.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, l_param_down)
                time.sleep(0.05)

                steps = max(1, int(t / 0.05))
                dx = (x2 - x1) / steps
                dy = (y2 - y1) / steps
                current_x, current_y = x1, y1

                for _ in range(steps):
                    current_x += dx
                    current_y += dy
                    l_param_move = int(current_y) << 16 | int(current_x)
                    win32gui.SendMessage(hwnd, win32con.WM_MOUSEMOVE, win32con.MK_LBUTTON, l_param_move)
                    time.sleep(0.05)

                l_param_up = y2 << 16 | x2
                win32gui.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, l_param_up)

                return f"执行拖动（sendmessage模式）：窗口内({x1},{y1})→({x2},{y2}) | 耗时{t}秒"

        except ValueError as e:
            return f"拖动参数格式错误（正确格式：(x1,y1),(x2,y2),(t)）：{str(e)}"
        except Exception as e:
            return f"拖动执行失败：{str(e)}"

    elif action_type == "stop":
        return f"停止任务组（由stop指令触发）", True, "stop_action"

    elif action_type == "adbcall":
        from screenshot_capture import get_adb_executable_path
        adb_valid, adb_msg = get_adb_executable_path()
        if not adb_valid:
            return f"ADB执行前置检查失败：{adb_msg}"

        if not params or not params[0]:
            return "adbcall动作参数不足（格式：adbcall:完整ADB命令）"

        raw_adb_cmd = params[0].strip()
        if not raw_adb_cmd.startswith("adb "):
            return f"ADB命令格式错误：必须以'adb '开头（当前：{raw_adb_cmd}）"

        adb_usage_mode = action_config.get("adb_usage_mode", "1") if action_config else "1"
        _, adb_path = get_adb_executable_path()

        if adb_usage_mode == "3":
            final_adb_cmd = raw_adb_cmd
        else:
            final_adb_cmd = raw_adb_cmd.replace("adb ", f'"{adb_path}" ', 1)

        try:
            app.log(f"  - 执行ADB命令：{final_adb_cmd}")
            result = subprocess.run(
                final_adb_cmd, shell=True, check=True,
                stdout=subprocess.PIPE, stderr=subprocess.PIPE, encoding="utf-8", timeout=30,
            )
            stdout = result.stdout.strip() if result.stdout else ""
            stderr = result.stderr.strip() if result.stderr else ""
            return f"ADB命令执行成功：{final_adb_cmd}\n  输出：{stdout}\n  错误输出：{stderr}"
        except subprocess.CalledProcessError as e:
            return f"ADB命令执行失败（返回码：{e.returncode}）：{final_adb_cmd}\n  输出：{e.stdout}\n  错误：{e.stderr}"
        except subprocess.TimeoutExpired:
            return f"ADB命令执行超时（30秒）：{final_adb_cmd}"
        except Exception as e:
            return f"ADB命令执行异常：{final_adb_cmd} | 错误：{str(e)}"

    elif action_type == "serial_adbcall":
        from screenshot_capture import get_adb_executable_path
        adb_valid, adb_msg = get_adb_executable_path()
        if not adb_valid:
            return f"Serial ADB执行前置检查失败：{adb_msg}"

        adb_device_serial = action_config.get("adb_device_serial", "") if action_config else ""

        if not params or not params[0]:
            return "serial_adbcall动作参数不足（格式：serial_adbcall:完整ADB命令）"

        raw_adb_cmd = params[0].strip()
        if not raw_adb_cmd.startswith("adb "):
            return f"Serial ADB命令格式错误：必须以'adb '开头（当前：{raw_adb_cmd}）"

        processed_adb_cmd = raw_adb_cmd
        if adb_device_serial and "-s " not in raw_adb_cmd:
            adb_cmd_parts = raw_adb_cmd.split("adb ", 1)
            if len(adb_cmd_parts) == 2:
                processed_adb_cmd = f"adb -s {adb_device_serial} {adb_cmd_parts[1]}"

        adb_usage_mode = action_config.get("adb_usage_mode", "1") if action_config else "1"
        _, adb_path = get_adb_executable_path()

        if adb_usage_mode == "3":
            final_adb_cmd = processed_adb_cmd
        else:
            final_adb_cmd = processed_adb_cmd.replace("adb ", f'"{adb_path}" ', 1)

        try:
            app.log(f"  - 执行Serial ADB命令：{final_adb_cmd}（原始命令：{raw_adb_cmd}）")
            result = subprocess.run(
                final_adb_cmd, shell=True, check=True,
                stdout=subprocess.PIPE, stderr=subprocess.PIPE, encoding="utf-8", timeout=30,
            )
            stdout = result.stdout.strip() if result.stdout else ""
            stderr = result.stderr.strip() if result.stderr else ""
            return (
                f"Serial ADB命令执行成功：\n"
                f"  原始命令：{raw_adb_cmd}\n"
                f"  执行命令：{final_adb_cmd}\n"
                f"  输出：{stdout}\n"
                f"  错误输出：{stderr}"
            )
        except subprocess.CalledProcessError as e:
            return (
                f"Serial ADB命令执行失败（返回码：{e.returncode}）：\n"
                f"  原始命令：{raw_adb_cmd}\n"
                f"  执行命令：{final_adb_cmd}\n"
                f"  输出：{e.stdout}\n"
                f"  错误：{e.stderr}"
            )
        except subprocess.TimeoutExpired:
            return (
                f"Serial ADB命令执行超时（30秒）：\n"
                f"  原始命令：{raw_adb_cmd}\n"
                f"  执行命令：{final_adb_cmd}"
            )
        except Exception as e:
            return (
                f"Serial ADB命令执行异常：\n"
                f"  原始命令：{raw_adb_cmd}\n"
                f"  执行命令：{final_adb_cmd}\n"
                f"  错误：{str(e)}"
            )

    elif action_type == "cmd_call":
        if not params or not params[0]:
            return "cmd_call动作参数不足（格式：cmd_call:完整CMD命令）"

        cmd = params[0].strip()
        if not cmd:
            return "cmd_call命令不能为空"

        try:
            app.log(f"  - 执行CMD命令：{cmd}")
            result = subprocess.run(
                cmd, shell=True, check=True,
                stdout=subprocess.PIPE, stderr=subprocess.PIPE, encoding="utf-8", timeout=60,
            )
            stdout = result.stdout.strip() if result.stdout else ""
            stderr = result.stderr.strip() if result.stderr else ""
            return f"CMD命令执行成功：{cmd}\n  输出：{stdout}\n  错误输出：{stderr}"
        except subprocess.CalledProcessError as e:
            return f"CMD命令执行失败（返回码：{e.returncode}）：{cmd}\n  输出：{e.stdout}\n  错误：{e.stderr}"
        except subprocess.TimeoutExpired:
            return f"CMD命令执行超时（60秒）：{cmd}"
        except Exception as e:
            return f"CMD命令执行异常：{cmd} | 错误：{str(e)}"

    elif action_type == "shutdown":
        try:
            if params and params[0]:
                cmd = params[0]
            else:
                cmd = "shutdown /s /t 60"

            app.log(f"  - 执行关机命令：{cmd}")
            os.system(cmd)
            return f"关机命令执行成功：{cmd}"
        except Exception as e:
            return f"关机命令执行失败：{cmd} | 错误：{str(e)}"

    else:
        return f"未知动作: {action_type}"