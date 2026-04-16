import os
import time
import configparser
import threading
import datetime
import cv2
import json
import webbrowser
import numpy as np
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, simpledialog, filedialog
import subprocess
import sys
import pygetwindow as gw
import psutil
import win32gui
import win32con
import win32api
import win32process
import win32ui
import pyautogui
import ctypes
from plyer import notification
from PIL import ImageGrab, ImageTk, Image
from pynput import keyboard
from pynput.keyboard import Key
from pathlib import Path

# 解决高DPI显示模糊问题
try:
    ctypes.windll.user32.SetProcessDPIAware()
except:
    pass

# 基础路径与配置初始化 - 兼容.py与.exe
# 当作为.py文件运行时，使用脚本所在目录
# 当作为.exe文件运行时，使用当前工作目录（exe所在目录）
if getattr(sys, 'frozen', False):
    # 打包为exe的情况
    BASE_DIR = os.getcwd()
else:
    # 作为.py文件运行的情况
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

MAIN_CONFIG_PATH = os.path.join(BASE_DIR, "config.ini")
REFS_DIR = os.path.join(BASE_DIR, "refs")
TASKS_DIR = os.path.join(BASE_DIR, "tasks")
for dir_path in [REFS_DIR, TASKS_DIR]:
    if not os.path.exists(dir_path):
        os.makedirs(dir_path)
for dir_path in [REFS_DIR, TASKS_DIR, os.path.join(REFS_DIR, "subdir")]:
    if not os.path.exists(dir_path):
        os.makedirs(dir_path)

# 系统通知函数（改用plyer，兼容新版Windows/Python）
def send_windows_notification(title, message):
    try:
        from plyer import notification
        notification.notify(
            title=title,
            message=message,
            app_name="自动点击工具",
            timeout=15  # 通知显示时长（秒）
        )
    except Exception as e:
        print(f"Toast通知失败，使用备用弹窗：{e}")
        # 备用弹窗逻辑（同上）
        win32gui.MessageBox(
            0,
            message,
            title,
            win32con.MB_ICONINFORMATION | win32con.MB_OK | win32con.MB_TOPMOST
        )

# 配置初始化（增强默认窗口预设读取）
def init_main_config():
    config = configparser.ConfigParser()
    if os.path.exists(MAIN_CONFIG_PATH):
        config.read(MAIN_CONFIG_PATH, encoding="utf-8")

    # ========== 原有GENERAL节点保留 ==========
    if "GENERAL" not in config:
        config["GENERAL"] = {} 
    # ==========创建ADBConfig节点 ==========
    if "ADBConfig" not in config:
        config["ADBConfig"] = {}
    # ========== 创建WindowConfig节点 ==========
    if "WindowConfig" not in config:
        config["WindowConfig"] = {}
    
    
    # 配置项 + 窗口预设相关
    default_general_config = {
        "recognition_frequency": "1.0",
        "next_start_time": "00:00:00",
        "current_task_group": "default",
        "target_program_name": "",
        "target_window_title": "",
        "target_window_hwnd": "",
        "max_execution_errors": "5",
        "enable_schedule": "False",
        "schedule_mode": "once",
        "click_mode": "sendmessage",        # 点击模式配置，可选值 sendmessage/pyautogui
        "screenshot_mode": "win32gui" ,     # 截图模式（win32gui/adb）
        "template_match_step": "0.05"       # 模板匹配步长，默认0.05
    }
    
    # ========== 新增：ADBConfig节点默认配置 ==========
    default_adb_config = {
        "adb_enabled": "1",             # 1启用ADB检查，0忽略ADB配置
        "adb_usage_mode": "1",          # 1:脚本目录下adb\platform-tools\adb.exe（默认）；2:自定义路径；3:系统环境变量
        "adb_position": "",             # 自定义ADB路径（仅mode=2时生效）
        "adb_device_serial": ""         # 新增：设备Serial号
    }

    # ========== 新增：WindowConfig节点默认配置（核心修改2） ==========
    default_window_config = {
        "selfgeometry": "1200x600",  # 窗口长宽，格式axb（默认800x600）
        "selfxy_changeable": "1" ,    # 1=可改变脚本窗口尺寸，0=固定尺寸（默认可修改）
        "scrgeometry": "820x460",
        "scrxy_changeable": "1" ,
        "adbscrgeometry":"700x450", 
        "adbscrxy_changeable": "1" 
    }

    # 写入GENERAL默认配置
    for key, value in default_general_config.items():
        if key not in config["GENERAL"]:
            config["GENERAL"][key] = value
    
    # 写入ADBConfig默认配置
    for key, value in default_adb_config.items():
        if key not in config["ADBConfig"]:
            config["ADBConfig"][key] = value

    # 写入WindowConfig默认配置
    for key, value in default_window_config.items():
        if key not in config["WindowConfig"]:
            config["WindowConfig"][key] = value
    
    with open(MAIN_CONFIG_PATH, "w", encoding="utf-8") as f:
        config.write(f)
    return config

# 修改配置初始化函数
def init_task_config(task_group_name):
    # 后缀从.ini改为.json
    task_config_path = os.path.join(TASKS_DIR, f"{task_group_name}.json")
    # 定义用户指定的JSON格式配置，原生数据类型（无字符串包装）
    default_config = {
        "TASK1": {
            "ignore_occlusion": False, 
            "ref_images": [
                {
                    "image": "button1.png",
                    "similarity_threshold": 0.9,
                    "match_times": 1,
                    "actions": [
                        {"type": "click", "params": [100, 200]},
                        {"type": "sleep", "params": [1.0]}
                    ]
                },
                {
                    "image": "button2.png",
                    "similarity_threshold": 0.9,
                    "match_times": 2,
                    "actions": [
                        {
                            "type": "click",
                            "params": [
                                100,
                                200
                            ]
                        }
                    ]
                }
            ]
        },
        "TASK2": {
            "ignore_occlusion": False, 
            "ref_images": [
                {
                    "image": "close_button.png",
                    "similarity_threshold": 0.85,
                    "match_times": 1,
                    "actions": [
                        {"type": "click", "params": [100, 200]},
                        {"type": "sleep", "params": [1.0]},
                        {"type": "stop", "params": []}
                    ]
                }
            ]
        }
    }
    if not os.path.exists(task_config_path):
        # 写入JSON文件，带缩进保证可读性，兼容中文
        with open(task_config_path, "w", encoding="utf-8") as f:
            json.dump(default_config, f, ensure_ascii=False, indent=4)
    # 读取JSON并返回（替代原configparser对象）
    with open(task_config_path, "r", encoding="utf-8") as f:
        task_config = json.load(f)
    return task_config

def get_window_client_rect(hwnd):
    """
    获取窗口客户区的「屏幕绝对坐标」和「尺寸」
    返回：client_left, client_top, client_width, client_height
    """
    # 1. 获取客户区相对窗口的矩形（left=0, top=0, right=宽, bottom=高）
    client_rect = win32gui.GetClientRect(hwnd)
    client_width = client_rect[2] - client_rect[0]
    client_height = client_rect[3] - client_rect[1]
    
    # 2. 转换客户区左上角为屏幕绝对坐标（修正外框→客户区的偏移）
    client_left, client_top = win32gui.ClientToScreen(hwnd, (client_rect[0], client_rect[1]))
    
    return client_left, client_top, client_width, client_height

def get_adb_executable_path():
    """
    根据配置获取ADB可执行文件路径，并验证有效性
    返回：(是否有效, adb路径/错误信息)
    """
    # 读取配置
    main_config = configparser.ConfigParser()
    main_config.read(MAIN_CONFIG_PATH, encoding="utf-8")
    adb_enabled = main_config["ADBConfig"].get("adb_enabled", "1")
    adb_usage_mode = main_config["ADBConfig"].get("adb_usage_mode", "1")
    adb_position = main_config["ADBConfig"].get("adb_position", "").strip()

    # 若禁用ADB，直接返回无效
    if adb_enabled != "1":
        return False, "ADB功能已禁用（adb_enabled=0）"

    # 模式1：脚本目录下的adb\platform-tools\adb.exe
    if adb_usage_mode == "1":
        adb_path = os.path.join(BASE_DIR, "adb", "platform-tools", "adb.exe")
        if os.path.exists(adb_path):
            return True, adb_path
        else:
            return False, f"模式1 - ADB文件不存在：{adb_path}（请检查脚本目录下的adb子文件夹）"
    
    # 模式2：自定义路径
    elif adb_usage_mode == "2":
        if not adb_position:
            return False, "模式2 - 未配置自定义ADB路径（adb_position为空）"
        if os.path.exists(adb_position) and adb_position.endswith("adb.exe"):
            return True, adb_position
        else:
            return False, f"模式2 - 自定义ADB路径无效：{adb_position}"
    
    # 模式3：系统环境变量
    elif adb_usage_mode == "3":
        try:
            # 检查系统环境变量中是否有adb
            result = subprocess.run(
                "adb version",
                shell=True,
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                encoding="utf-8"
            )
            return True, "adb（系统环境变量）"
        except subprocess.CalledProcessError as e:
            return False, f"模式3 - 系统环境变量中ADB执行失败：{e.stderr}"
        except FileNotFoundError:
            return False, "模式3 - 系统环境变量中未找到ADB"
    
    # 无效模式
    else:
        return False, f"无效的ADB使用模式：{adb_usage_mode}（仅支持1/2/3）"

def validate_adb_environment():
    """验证ADB环境是否可用，返回：(是否可用, 提示信息)"""
    # ---------------------- 新增：读取adb_usage_mode配置 ----------------------
    main_config = configparser.ConfigParser()
    main_config.read(MAIN_CONFIG_PATH, encoding="utf-8")
    adb_usage_mode = main_config["ADBConfig"].get("adb_usage_mode", "1")  # 与get_adb_executable_path逻辑一致
    # --------------------------------------------------------------------------
    
    is_valid, msg = get_adb_executable_path()
    if not is_valid:
        return False, msg
    
    # 额外验证ADB版本（确保可执行）
    adb_path = msg if is_valid else ""
    try:
        # 现在使用读取到的adb_usage_mode变量，而非未定义变量
        cmd = f'"{adb_path}" version' if adb_usage_mode != "3" else "adb version"
        result = subprocess.run(
            cmd,
            shell=True,
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            encoding="utf-8",
            timeout=10
        )
        version_info = result.stdout.strip().split("\n")[0] if result.stdout else "未知版本"
        # 同样使用读取到的adb_usage_mode变量
        return True, f"ADB环境验证通过：{adb_path if adb_usage_mode !=3 else '系统环境变量'} | {version_info}"
    except Exception as e:
        return False, f"ADB可执行但版本检查失败：{str(e)}"

# 添加执行动作的函数
def execute_action(window, action_type, params, stop_flag=False):  
    """执行具体动作"""

    # 先检查停止标志，任何动作执行前都优先终止
    # 实时检测停止标志（而非依赖传入的静态值）
    if app and app.stop_flag:
        return "已触发停止指令，终止动作执行", True  # 复用原有stop的返回格式
    
    if action_type == "click":
        if len(params) >= 2:
            x, y = int(params[0]), int(params[1])
            # ---------------------- 新增：读取当前点击模式配置 ----------------------
            main_config = configparser.ConfigParser()
            main_config.read(MAIN_CONFIG_PATH, encoding="utf-8")
            click_mode = main_config["GENERAL"]["click_mode"]
            # ----------------------------------------------------------------------
            
            # 传递点击模式到 auto_click
            auto_click(window, x, y, times=1, interval=0.1, click_mode=click_mode)
            return f"执行点击: ({x},{y}) | 模式: {click_mode}"
        else:
            return "点击动作参数不足"
    
    elif action_type == "sleep":
        if len(params) >= 1:
            t = float(params[0])
            # 第一步：优先输出「执行等待」的前置日志（用户要求的顺序第一）
            app.log(f"  - 执行等待: {t}秒")
            start_time = time.monotonic()
            elapsed = 0.0
            # 第二步：输出Sleep开始日志（顺序第二）
            app.log(f"【Sleep开始】预期等待{t}秒 | 开始时间: {datetime.datetime.now()}")
            while elapsed < t:
                if app and app.stop_flag:
                    app.log(f"【Sleep中断】预期{t}秒，实际等待{elapsed:.2f}秒")
                    return ("等待被中断（用户停止）", True, True)
                sleep_chunk = min(0.1, t - elapsed)
                time.sleep(sleep_chunk)
                elapsed = time.monotonic() - start_time
            # 第三步：输出Sleep完成日志（顺序第三）
            app.log(f"【Sleep完成】预期{t}秒，实际等待{elapsed:.2f}秒")
            # 关键：返回("", False)，让外层不重复输出日志
            return ("", False)
        else:
            return ("等待动作参数不足", True)
        
    # ---------------------- press 键盘按键动作 ----------------------
    # -------------------- press 单键/组合键动作 ---------------------
    elif action_type == "press":
        # 清理参数：去除每个按键名称的空格，过滤空值（避免多余逗号导致的空参数）
        keys = [k.strip() for k in params if k.strip()]
        if not keys:
            return "press动作参数不足（需指定至少1个按键名称，如enter、ctrl,a）"
        
        try:
            if len(keys) == 1:
                # 单键：模拟单次按下
                key = keys[0]
                pyautogui.press(key)
                return f"执行键盘单键按下：{key}"
            else:
                # 多键：模拟组合键（如ctrl+a、shift+alt+d）
                pyautogui.hotkey(*keys)
                return f"执行键盘组合键按下：{'+' .join(keys)}"
        except ValueError as e:
            # 捕获无效按键名称的异常（明确指出哪个按键无效）
            invalid_keys = []
            for k in keys:
                try:
                    # 验证单个按键是否有效（pyautogui无直接验证方法，模拟press触发校验）
                    pyautogui.press(k)
                except ValueError:
                    invalid_keys.append(k)
            if invalid_keys:
                return f"无效的键盘按键名称「{'、'.join(invalid_keys)}」：{str(e)}"
            else:
                return f"组合键格式错误：{str(e)}"
        except Exception as e:
            # 捕获其他执行异常（如权限/环境问题）
            return f"键盘按键执行失败（按键：{'+' .join(keys)}）：{str(e)}"
        
    elif action_type == "goto_task":
        if len(params) >= 1:
            try:
                task_num = int(params[0])
                target_index = task_num - 1  # 转换为0基索引
                return f"跳转到任务{task_num}", False, target_index  # 第三个值为目标索引
            except ValueError:
                return "goto_task参数必须是整数"
        else:
            return "goto_task参数不足"
        
    elif action_type == "taskcall":
        # 拼接参数（兼容内容含逗号的场景）
        notify_content = ",".join(params)
        # 调用现有Windows通知函数发送消息
        send_windows_notification("TaskCall：", notify_content)
        return f"发送Windows通知: {notify_content}"
    
     # ---------------------- swipe 拖动动作 ----------------------
    elif action_type == "swipe":
        try:
            # 1. 修正参数解析：坐标转整数，时长转浮点数（补全t的读取）
            if len(params) < 5:
                return f"拖动参数不足（正确格式：x1,y1,x2,y2,t），当前仅传入{len(params)}个参数"
            x1 = int(float(params[0]))  # 兼容字符串/浮点数输入，最终转整数
            y1 = int(float(params[1]))
            x2 = int(float(params[2]))
            y2 = int(float(params[3]))
            t = float(params[4])        # 拖动时长（秒）
            
            # 读取当前点击模式配置
            main_config = configparser.ConfigParser()
            main_config.read(MAIN_CONFIG_PATH, encoding="utf-8")
            click_mode = main_config["GENERAL"]["click_mode"]
            hwnd = window._hWnd

            if click_mode == "pyautogui":
                # PyAutoGUI模式：硬件级拖动（屏幕绝对坐标）
                client_left, client_top, _, _ = get_window_client_rect(hwnd)
                screen_x1 = client_left + x1
                screen_y1 = client_top + y1
                screen_x2 = client_left + x2
                screen_y2 = client_top + y2
                
                # 确保窗口前置
                bring_window_to_front(window)
                time.sleep(0.1)
                
                # 模拟物理拖动（按下→移动→松开）
                pyautogui.moveTo(screen_x1, screen_y1, duration=0.05)
                pyautogui.mouseDown(screen_x1, screen_y1)
                pyautogui.moveTo(screen_x2, screen_y2, duration=t)
                pyautogui.mouseUp(screen_x2, screen_y2)
                
                return f"执行拖动（pyautogui模式）：窗口内({x1},{y1})→({x2},{y2}) | 屏幕({screen_x1},{screen_y1})→({screen_x2},{screen_y2}) | 耗时{t}秒"
            
            elif click_mode == "sendmessage":
                # SendMessage模式：消息级拖动（窗口相对坐标）
                # 左键按下（起始点）
                l_param_down = y1 << 16 | x1
                win32gui.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, l_param_down)
                time.sleep(0.05)
                
                # 分步骤移动（避免一次性移动失效）
                steps = max(1, int(t / 0.05))  # 每50ms一步
                dx = (x2 - x1) / steps
                dy = (y2 - y1) / steps
                current_x, current_y = x1, y1
                
                for _ in range(steps):
                    current_x += dx
                    current_y += dy
                    l_param_move = int(current_y) << 16 | int(current_x)
                    win32gui.SendMessage(hwnd, win32con.WM_MOUSEMOVE, win32con.MK_LBUTTON, l_param_move)
                    time.sleep(0.05)
                
                # 左键松开（结束点）
                l_param_up = y2 << 16 | x2
                win32gui.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, l_param_up)
                
                return f"执行拖动（sendmessage模式）：窗口内({x1},{y1})→({x2},{y2}) | 耗时{t}秒"
        
        except ValueError as e:
            return f"拖动参数格式错误（正确格式：(x1,y1),(x2,y2),(t)）：{str(e)}"
        except Exception as e:
            return f"拖动执行失败：{str(e)}"

    elif action_type == "stop":
        return f"停止任务组（由stop指令触发）", True, "stop_action" # 第二个返回值表示需要停止，第三个返回值标记stop来源
    
    elif action_type == "adbcall":
        # 第一步：检查ADB是否启用并验证环境
        adb_valid, adb_msg = get_adb_executable_path()
        if not adb_valid:
            return f"ADB执行前置检查失败：{adb_msg}"
        
        # 第二步：处理ADB命令（替换adb路径）
        if not params or not params[0]:
            return "adbcall动作参数不足（格式：adbcall:完整ADB命令）"
        
        # 获取原始命令
        raw_adb_cmd = params[0].strip()
        if not raw_adb_cmd.startswith("adb "):
            return f"ADB命令格式错误：必须以'adb '开头（当前：{raw_adb_cmd}）"
        
        # 读取配置，替换ADB执行路径
        main_config = configparser.ConfigParser()
        main_config.read(MAIN_CONFIG_PATH, encoding="utf-8")
        adb_usage_mode = main_config["ADBConfig"].get("adb_usage_mode", "1")
        _, adb_path = get_adb_executable_path()

        # 构建最终执行命令
        if adb_usage_mode == "3":
            # 模式3：直接使用系统adb
            final_adb_cmd = raw_adb_cmd
        else:
            # 模式1/2：替换为实际adb路径
            final_adb_cmd = raw_adb_cmd.replace("adb ", f'"{adb_path}" ', 1)
        
        try:
            app.log(f"  - 执行ADB命令：{final_adb_cmd}")
            # 执行完整ADB命令，捕获输出/错误（超时30秒）
            result = subprocess.run(
                final_adb_cmd,
                shell=True,
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                encoding="utf-8",
                timeout=30
            )
            # 输出执行结果
            stdout = result.stdout.strip()
            stderr = result.stderr.strip()
            return f"ADB命令执行成功：{final_adb_cmd}\n  输出：{stdout}\n  错误输出：{stderr}"
        except subprocess.CalledProcessError as e:
            return f"ADB命令执行失败（返回码：{e.returncode}）：{final_adb_cmd}\n  输出：{e.stdout}\n  错误：{e.stderr}"
        except subprocess.TimeoutExpired:
            return f"ADB命令执行超时（30秒）：{final_adb_cmd}"
        except Exception as e:
            return f"ADB命令执行异常：{final_adb_cmd} | 错误：{str(e)}"
        
    elif action_type == "serial_adbcall":
        # 第一步：检查ADB是否启用并验证环境（复用原有逻辑）
        adb_valid, adb_msg = get_adb_executable_path()
        if not adb_valid:
            return f"Serial ADB执行前置检查失败：{adb_msg}"
        
        # 第二步：读取配置中的adb_device_serial
        main_config = configparser.ConfigParser()
        main_config.read(MAIN_CONFIG_PATH, encoding="utf-8")
        adb_device_serial = main_config["ADBConfig"].get("adb_device_serial", "").strip()
        
        # 第三步：处理ADB命令参数（核心：自动添加serial）
        if not params or not params[0]:
            return "serial_adbcall动作参数不足（格式：serial_adbcall:完整ADB命令）"
        
        # 获取原始命令
        raw_adb_cmd = params[0].strip()
        if not raw_adb_cmd.startswith("adb "):
            return f"Serial ADB命令格式错误：必须以'adb '开头（当前：{raw_adb_cmd}）"
        
        # 核心逻辑：判断命令是否已包含 -s 参数，未包含则自动添加配置的serial
        processed_adb_cmd = raw_adb_cmd
        # 仅当配置有serial 且 命令中无 -s 参数时，才自动添加
        if adb_device_serial and "-s " not in raw_adb_cmd:
            # 拆分adb头和后续参数，插入 -s {serial}
            adb_cmd_parts = raw_adb_cmd.split("adb ", 1)
            if len(adb_cmd_parts) == 2:
                processed_adb_cmd = f"adb -s {adb_device_serial} {adb_cmd_parts[1]}"
        
        # 第四步：替换ADB执行路径（复用原有adbcall逻辑）
        adb_usage_mode = main_config["ADBConfig"].get("adb_usage_mode", "1")
        _, adb_path = get_adb_executable_path()

        # 构建最终执行命令
        if adb_usage_mode == "3":
            # 模式3：直接使用系统adb
            final_adb_cmd = processed_adb_cmd
        else:
            # 模式1/2：替换为实际adb路径
            final_adb_cmd = processed_adb_cmd.replace("adb ", f'"{adb_path}" ', 1)
        
        try:
            app.log(f"  - 执行Serial ADB命令：{final_adb_cmd}（原始命令：{raw_adb_cmd}）")
            # 执行完整ADB命令，捕获输出/错误（超时30秒）
            result = subprocess.run(
                final_adb_cmd,
                shell=True,
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                encoding="utf-8",
                timeout=30
            )
            # 输出执行结果（包含原始命令和处理后命令，方便排查）
            stdout = result.stdout.strip()
            stderr = result.stderr.strip()
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
    
    elif action_type == "shutdown":
        try:
            # 取冒号后完整命令，无参数则用默认值
            if params and params[0]:
                cmd = params[0]  # params[0] 是冒号后完整命令（如 "shutdown /s"）
            else:
                cmd = "shutdown /s /t 60"  # 默认60秒关机
            
            app.log(f"  - 执行关机命令：{cmd}")
            os.system(cmd)
            return f"关机命令执行成功：{cmd}"
        except Exception as e:
            return f"关机命令执行失败：{cmd} | 错误：{str(e)}"
    
    
    else:
        return f"未知动作: {action_type}"

# 窗口匹配逻辑（模糊识别，配置读取）
def get_all_visible_windows_simple():
    """精简版：获取窗口列表（仅标题+进程名+句柄，用于下拉列表）"""
    windows = []
    def callback(hwnd, extra):
        if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowTextLength(hwnd) > 0:
            title = win32gui.GetWindowText(hwnd)
            try:
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                process = psutil.Process(pid)
                process_name = process.name()
            except:
                process_name = "未知进程"
            # 格式：标题 (进程名) - 句柄
            display_text = f"{title[:50]} ({process_name})"  # 截断长标题，避免下拉框过长
            windows.append({
                "display": display_text,
                "hwnd": hwnd,
                "title": title,
                "process_name": process_name
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
    
    # 1. 优先精确匹配句柄
    if target_hwnd and target_hwnd.isdigit():
        hwnd = int(target_hwnd)
        if win32gui.IsWindow(hwnd) and win32gui.IsWindowVisible(hwnd):
            try:
                return gw.Window(hwnd)
            except:
                pass
    
    # 2. 模糊匹配
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
                    except:
                        pass
            return True
        
        win32gui.EnumWindows(enum_windows_callback, {"pid_list": pid_list})
        if win_list:
            return win_list[0]
    
    return None

# 核心函数（截图、点击、模板匹配等）
def is_window_visible(hwnd):
    try:
        return (win32gui.IsWindowVisible(hwnd) 
                and win32gui.GetWindowText(hwnd) != "")
    except:
        return False
    
def capture_adb_screenshot(device_serial=""):
    """
    通过ADB获取指定设备的截图，转为OpenCV BGR格式（兼容原截图处理逻辑）
    :param device_serial: 设备Serial号（为空则使用默认设备）
    :return: OpenCV BGR格式截图 / None（失败）
    """
    # 1. 前置检查：ADB环境有效性
    adb_valid, adb_path = get_adb_executable_path()
    if not adb_valid:
        app.log(f"ADB截图失败：{adb_path}")
        return None
    
    # 2. 构建ADB命令前缀（带设备Serial号）
    adb_cmd_prefix = f'"{adb_path}"'
    if device_serial.strip():
        adb_cmd_prefix += f' -s {device_serial.strip()}'
    
    # 3. 临时截图文件（避免中文/空格路径问题）
    temp_png = os.path.join(BASE_DIR, "adb_screenshot_temp.png")
    if os.path.exists(temp_png):
        os.remove(temp_png)
    
    # 4. 执行ADB截图命令（exec-out 避免权限问题）
    adb_screenshot_cmd = f'{adb_cmd_prefix} exec-out screencap -p > "{temp_png}"'
    try:
        result = subprocess.run(
            adb_screenshot_cmd,
            shell=True,
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            encoding="utf-8",
            timeout=10
        )
        # 校验临时文件有效性
        if not os.path.exists(temp_png) or os.path.getsize(temp_png) == 0:
            app.log(f"ADB截图命令执行成功，但临时文件为空：{temp_png}")
            return None
        
        # 5. 转换为OpenCV BGR格式（匹配原win32gui截图输出）
        screenshot_pil = Image.open(temp_png).convert('RGB')
        img = cv2.cvtColor(np.array(screenshot_pil), cv2.COLOR_RGB2BGR)
        
        # 清理临时文件
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
    """
    从窗口内存直接读取像素（Win32 Memory模式），返回OpenCV BGR格式图像
    修复：仅读取窗口客户区（排除标题栏/边框），解决非客户区截取问题
    :param hwnd: 窗口句柄
    :return: OpenCV BGR图像 / None（失败）
    """
    try:
        # 1. 获取窗口整体的屏幕矩形（用于计算客户区在窗口内的偏移）
        window_rect = win32gui.GetWindowRect(hwnd)
        window_left, window_top = window_rect[0], window_rect[1]
        
        # 2. 获取客户区的屏幕坐标+尺寸（复用现有函数）
        client_left, client_top, client_width, client_height = get_window_client_rect(hwnd)
        if client_width <= 0 or client_height <= 0:
            app.log(f"窗口客户区尺寸无效：{client_width}x{client_height}")
            return None
        
        # 3. 核心修复：计算客户区在「窗口自身坐标系」内的偏移量
        # （窗口自身坐标系的(0,0) = 窗口左上角的屏幕坐标）
        offset_x = client_left - window_left  # 客户区x轴在窗口内的偏移
        offset_y = client_top - window_top  # 客户区y轴在窗口内的偏移
        if offset_x < 0 or offset_y < 0:
            app.log(f"客户区偏移异常：offset_x={offset_x}, offset_y={offset_y}")
            return None

        # 4. 创建GDI设备上下文（DC）
        hdc_window = win32gui.GetDC(hwnd)  # 获取窗口DC（窗口自身坐标系）
        hdc_mem = win32ui.CreateDCFromHandle(hdc_window)  # 从窗口DC创建内存DC
        mem_dc = hdc_mem.CreateCompatibleDC()  # 创建兼容DC

        # 5. 创建位图对象，绑定到兼容DC（仅匹配客户区尺寸）
        bmp = win32ui.CreateBitmap()
        bmp.CreateCompatibleBitmap(hdc_mem, client_width, client_height)
        mem_dc.SelectObject(bmp)

        # 6. 修复BitBlt源坐标：从客户区在窗口内的偏移位置开始拷贝
        mem_dc.BitBlt(
            (0, 0),  # 目标DC（内存位图）的起始坐标
            (client_width, client_height),  # 拷贝尺寸（客户区宽高）
            hdc_mem,  # 源DC（窗口DC，基于窗口自身坐标系）
            (offset_x, offset_y),  # 源DC的起始坐标（客户区在窗口内的偏移）
            win32con.SRCCOPY
        )

        # 7. 转换位图为OpenCV格式（BGR）
        signed_ints_array = bmp.GetBitmapBits(True)  # 获取位图像素数据（BGRA）
        img = np.frombuffer(signed_ints_array, dtype='uint8')
        img.shape = (client_height, client_width, 4)  # 调整为BGRA维度
        img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)  # 转为BGR（兼容原有逻辑）

        app.log(f"win32memory截图成功：{client_width}x{client_height} | 窗口内偏移：({offset_x},{offset_y})")
        return img

    except Exception as e:
        app.log(f"win32memory截图失败：{str(e)}")
        return None

    finally:
        # 强制释放GDI资源（避免内存泄漏）
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

def capture_window_win32gui(hwnd):
    """
    Win32GUI模式截图（抽离独立函数）：基于pyautogui截取窗口客户区，返回OpenCV BGR格式图像
    :param hwnd: 窗口句柄
    :return: OpenCV BGR图像 / None（失败）
    """
    try:
        if not win32gui.IsWindow(hwnd) or not is_window_visible(hwnd):
            app.log("win32gui截图失败：窗口无效/不可见/无标题")
            return None

        # 获取窗口客户区尺寸与绝对坐标（复用现有函数）
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        width = right - left
        height = bottom - top
        client_left, client_top, client_width, client_height = get_window_client_rect(hwnd)

        if client_width <= 0 or client_height <= 0:
            app_log = f"窗口客户区尺寸无效：{client_width}x{client_height}"
            app.log(f"截图失败：{app_log}")
            return None

        # 核心截图逻辑（原win32gui分支内的代码）
        screenshot_pil = pyautogui.screenshot(region=(client_left, client_top, client_width, client_height))
        img = cv2.cvtColor(np.array(screenshot_pil), cv2.COLOR_RGB2BGR)

        app.log(f"win32gui截图成功（客户区）：{img.shape[1]}x{img.shape[0]}")
        return img

    except Exception as e:
        app.log(f"win32gui截图失败：{str(e)}")
        return None

def capture_window(window):
    hdc_window = None
    hdc_memdc = None
    hbitmap = None
    hwnd = None
    
    try:
        # ========== 读取截图模式配置 ==========
        main_config = configparser.ConfigParser()
        main_config.read(MAIN_CONFIG_PATH, encoding="utf-8")
        screenshot_mode = main_config["GENERAL"]["screenshot_mode"]
        adb_device_serial = main_config["ADBConfig"]["adb_device_serial"]
        # =====================================
        
        # ========== ADB截图模式（无需窗口句柄） ==========
        if screenshot_mode == "ADB":
            img = capture_adb_screenshot(device_serial=adb_device_serial)
            if img is None:
                app.log("ADB截图模式：截图失败")
                return None
            return img
        
        # ========== win32gui截图模式（解耦合后调用独立函数） ==========
        elif screenshot_mode == "Win32GUI":
            if not window or not win32gui.IsWindow(window._hWnd):
                app.log("win32gui截图失败：窗口句柄无效/已关闭")
                return None
            hwnd = window._hWnd
            img = capture_window_win32gui(hwnd)  # 调用独立函数
            return img
        
        # ========== 新增：win32memory截图模式 ==========
        elif screenshot_mode == "Win32Memory":
            if not window or not win32gui.IsWindow(window._hWnd):
                app.log("win32memory截图失败：窗口句柄无效/已关闭")
                return None
            hwnd = window._hWnd
            
            if not win32gui.IsWindowVisible(hwnd) or win32gui.GetWindowText(hwnd) == "":
                app.log("win32memory截图失败：窗口不可见/无标题")
                return None
            
            # 调用新增的内存读取函数
            img = capture_window_memory(hwnd)
            return img

        # ========== 无效截图模式处理 ==========
        else:
            app.log(f"无效的截图模式：{screenshot_mode}（仅支持win32gui/adb/win32memory）")
            return None
    except Exception as e:
        app.log(f"截图失败：{str(e)}")
        return None
    finally:
        # 原有资源释放逻辑不变
        if hbitmap and hdc_memdc:
            try:
                hdc_memdc.SelectObject(None)
            except:
                pass
        if hbitmap:
            try:
                win32gui.DeleteObject(hbitmap.GetHandle())
            except:
                app.log("警告：位图对象删除失败")
        if hdc_memdc:
            try:
                hdc_memdc.DeleteDC()
            except:
                pass
        if hdc_window and hwnd:
            try:
                win32gui.ReleaseDC(hwnd, hdc_window)
            except:
                pass

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
        print(f"窗口置顶失败：{e}")
        raise

def template_match(screenshot, ref_image_path, threshold=0.9):
    """
    支持同比例拉伸的多尺度模板匹配（兼容原有逻辑）
    :param screenshot: 窗口截图（OpenCV BGR格式）
    :param ref_image_path: 参考图路径
    :param threshold: 匹配阈值
    :return: (是否匹配成功, 最高相似度)
    """
    # 1. 读取参考图（转为灰度图，减少计算量，支持中文路径）
    try:
        # 用PIL读取（兼容中文路径），再转为OpenCV格式的灰度图
        ref_pil = Image.open(ref_image_path).convert('L')
        ref_img = np.array(ref_pil)
    except FileNotFoundError:
        print(f"错误：参考图不存在 → {ref_image_path}")
        return False, 0.0
    except Exception as e:
        print(f"错误：读取参考图失败 → {e}")
        return False, 0.0

    # 2. 处理截图（转为灰度图）
    if screenshot is None or screenshot.size == 0:
        return False, 0.0
    screenshot_gray = cv2.cvtColor(screenshot, cv2.COLOR_BGR2GRAY)
    ref_h, ref_w = ref_img.shape[:2]  # 原始参考图尺寸
    max_similarity = 0.0  # 存储所有尺度下的最高相似度

    # ========== 新增：读取配置中的匹配步长 ==========
    main_config = configparser.ConfigParser()
    main_config.read(MAIN_CONFIG_PATH, encoding="utf-8")
    # 读取步长，兼容配置缺失/格式错误，默认0.05
    try:
        match_step = float(main_config["GENERAL"].get("template_match_step", "0.05"))
        # 步长合法性校验，限制0.001~0.5（避免过细/过粗）
        match_step = max(0.001, min(0.5, match_step))
    except ValueError:
        match_step = 0.05
    # =============================================

    # 3. 多尺度匹配核心：生成0.1~4.0倍的缩放比例（覆盖常见拉伸范围）
    # 步长越小越精准，但耗时略增；桌面应用建议0.1~0.2即可
    scales = np.arange(0.1, 4.1, match_step)  
    for scale in scales:
        # 按当前比例缩放参考图
        scaled_w = int(ref_w * scale)
        scaled_h = int(ref_h * scale)
        
        # 跳过缩放后尺寸为0，或超过截图尺寸的情况
        if scaled_w <= 0 or scaled_h <= 0:
            continue
        if scaled_w > screenshot_gray.shape[1] or scaled_h > screenshot_gray.shape[0]:
            continue
        
        # 缩放参考图：缩小用INTER_AREA（更清晰），放大用INTER_CUBIC（抗锯齿）
        if scale < 1.0:
            scaled_ref = cv2.resize(ref_img, (scaled_w, scaled_h), interpolation=cv2.INTER_AREA)
        else:
            scaled_ref = cv2.resize(ref_img, (scaled_w, scaled_h), interpolation=cv2.INTER_CUBIC)
        
        # 4. 模板匹配
        result = cv2.matchTemplate(screenshot_gray, scaled_ref, cv2.TM_CCOEFF_NORMED)
        _, current_sim, _, _ = cv2.minMaxLoc(result)  # 取当前尺度的最高相似度
        
        # 5. 更新全局最高相似度
        if current_sim > max_similarity:
            max_similarity = current_sim
        
        # 提前终止：如果已达到阈值，无需继续匹配（提升效率）
        if max_similarity >= threshold:
            break

    # 6. 按最高相似度判断是否匹配
    is_match = max_similarity >= threshold
    return is_match, max_similarity


def auto_click(window, x, y, times=1, interval=0.5, click_mode="sendmessage"):
    """
    支持双模式的自动点击函数
    :param click_mode: 点击模式（sendmessage/pyautogui）
    """
    if not window or not is_window_visible(window._hWnd):
        raise RuntimeError("目标窗口不可见或已关闭")
    
    hwnd = window._hWnd

    # ---------------------- 模式1：win32gui.SendMessage 消息点击（原有逻辑）----------------------
    if click_mode == "sendmessage":
        def send_mouse_down():
            l_param = y << 16 | x  # 窗口内相对坐标
            win32gui.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, l_param)
        
        def send_mouse_up():
            l_param = y << 16 | x
            win32gui.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, l_param)
        
        for _ in range(times):
            send_mouse_down()
            time.sleep(0.05)
            send_mouse_up()
            time.sleep(interval)

    # ---------------------- 模式2：PyAutoGUI 硬件级点击（新增逻辑）----------------------
    elif click_mode == "pyautogui":

        client_left, client_top, _, _ = get_window_client_rect(hwnd)
        screen_x = client_left + x
        screen_y = client_top + y
        
        # 3. 确保窗口前置（避免点击到其他窗口）
        bring_window_to_front(window)
        time.sleep(0.1)  # 等待窗口置顶
        
        # 4. 模拟物理级鼠标点击（替代SendMessage）
        for _ in range(times):
            # 移动鼠标到目标坐标（可选，若需要可见鼠标移动）
            pyautogui.moveTo(screen_x, screen_y, duration=0.05)
            # 模拟左键点击（按下+松开）
            pyautogui.click(screen_x, screen_y)
            # 间隔时间
            time.sleep(interval)


# 工作线程（保留原有逻辑，适配新的窗口配置）
# 修改worker函数中的任务执行部分
current_task_index = 0 

def worker(app):
    stop_execution = False
    global current_task_index  # 声明使用全局变量
    current_task_index = 0  # 每次启动任务时重置任务索引为0
    consecutive_error_count = 0
    scheduled_once_triggered = False
    stop_source = ""  # 记录停止来源（manual/stop_action）
    # 新增：每个子任务的连续匹配成功计数器，key=任务名，value=连续成功次数，初始化为0
    task_continuous_match = {}     
    
    while not app.stop_flag:
        is_normal_execution = True
        error_msg = ""
        
        try:
            main_config = configparser.ConfigParser()
            main_config.read(MAIN_CONFIG_PATH, encoding="utf-8")
            freq = float(main_config["GENERAL"]["recognition_frequency"])
            start_time_str = main_config["GENERAL"]["next_start_time"]
            task_group = main_config["GENERAL"]["current_task_group"]
            max_consecutive_errors = int(main_config["GENERAL"]["max_execution_errors"])
            
            enable_schedule = main_config["GENERAL"].getboolean("enable_schedule")
            schedule_mode = main_config["GENERAL"]["schedule_mode"]
            
            now = datetime.datetime.now()
            need_wait = False
            
            if enable_schedule and not (schedule_mode == "once" and scheduled_once_triggered):
                try:
                    target_time = datetime.datetime.strptime(start_time_str, "%H:%M:%S").replace(
                        year=now.year, month=now.month, day=now.day
                    )
                    
                    if now < target_time:
                        wait_sec = (target_time - now).total_seconds()
                        app.log(f"【定时等待】{schedule_mode}模式，距离{start_time_str}还有 {wait_sec:.1f} 秒")
                        time.sleep(min(wait_sec, freq))
                        need_wait = True
                    else:
                        if schedule_mode == "once":
                            app.log(f"【定时触发】仅一次模式已触发（{start_time_str}）")
                            scheduled_once_triggered = True
                        else:
                            target_time += datetime.timedelta(days=1)
                            app.log(f"【定时触发】始终模式已触发，下次：{target_time.strftime('%H:%M:%S')}")
                except ValueError:
                    is_normal_execution = False
                    error_msg = f"定时时间格式错误（{start_time_str}）"
                    app.log(f"⚠️ {error_msg}")
            
            if not need_wait:
                # 从配置获取目标窗口
                target_window = get_target_window_from_config()
                if not target_window:
                    is_normal_execution = False
                    error_msg = "未找到目标程序窗口（请检查配置的进程/标题关键词）"
                    app.log(f"执行错误：{error_msg}（连续错误：{consecutive_error_count + 1}/{max_consecutive_errors}）")
                else:
                    try:
                        window_text = win32gui.GetWindowText(target_window._hWnd)
                        window_handle = target_window._hWnd
                        client_left, client_top, client_width, client_height = get_window_client_rect(target_window._hWnd)
                        app.log(f"✅ 找到目标窗口 - 标题：{window_text} | 句柄：{window_handle} | 客户区位置/尺寸：({client_left},{client_top}) {client_width}x{client_height}")
                            
                        # 后缀从.ini改为.json
                        task_path = os.path.join(TASKS_DIR, f"{task_group}.json")
                        if not os.path.exists(task_path):
                            init_task_config(task_group)
                        # 用json模块读取配置，返回原生字典
                        with open(task_path, "r", encoding="utf-8") as f:
                            task_config = json.load(f)
                        
                        task_error = False
                        stop_execution = False
                        
                        # 获取任务列表并转换为可迭代对象
                        # JSON字典的key即为任务名（TASK1/TASK2）
                        task_list = list(task_config.keys())

                        while current_task_index < len(task_list) and not stop_execution:

                            # 每个Task执行前重新截图（保留原有失败逻辑）
                            screenshot = capture_window(target_window)
                            if screenshot is None:
                                # 完全复用原有截图失败逻辑，不做任何修改
                                is_normal_execution = False
                                error_msg = "窗口截图失败"
                                app.log(f"执行错误：{error_msg}（连续错误：{consecutive_error_count + 1}/{max_consecutive_errors}）")
                                # 触发任务循环终止，回到外层错误处理
                                task_error = True
                                break
                            else:
                                app.log(f"✅ 任务[{task_list[current_task_index]}]截图成功 - 尺寸：{screenshot.shape[1]}x{screenshot.shape[0]}")

                            # ：进入任务前先检查停止，立即终止循环
                            if app.stop_flag:
                                stop_execution = True
                                break  # 直接跳出任务循环，不再执行任何任务

                            task_name = task_list[current_task_index]

                            try:
                                task = task_config[task_name]

                                # 新增：读取任务级 ignore_occlusion 配置
                                ignore_occlusion = task.get("ignore_occlusion", False)
                                
                                # 新增：遍历ref_images数组，查找第一个匹配的分支
                                ref_images_list = task.get("ref_images", [])
                                if not ref_images_list:
                                    task_error = True
                                    error_msg = f"任务[{task_name}]：ref_images数组为空"
                                    app.log(f"⚠️ {error_msg}")
                                    break

                                matched_branch = None  # 存储匹配到的分支（包含image/threshold/match_times/actions）
                                matched_branch_idx = -1  # 匹配分支的索引

                                 # 新增：如果配置了 ignore_occlusion，直接执行第一个分支
                                if ignore_occlusion:
                                    matched_branch = ref_images_list[0]
                                    matched_branch_idx = 0
                                    app.log(f"✅ 任务[{task_name}]：配置了忽略遮挡，直接执行第一个分支")

                                else:

                                    # 核心：遍历所有ref_images，按顺序查找第一个匹配的
                                    for branch_idx, ref_branch in enumerate(ref_images_list):
                                        # 提取当前分支的配置
                                        ref_path = os.path.join(REFS_DIR, ref_branch["image"])
                                        threshold = ref_branch.get("similarity_threshold", 0.9)
                                        match_times = ref_branch.get("match_times", 1)
                                        match_times = match_times if match_times >= 1 else 1

                                        # 初始化当前分支的连续匹配计数器
                                        branch_key = f"{task_name}_branch{branch_idx}"
                                        if branch_key not in task_continuous_match:
                                            task_continuous_match[branch_key] = 0

                                        # 执行模板匹配
                                        match_result = template_match(screenshot, ref_path, threshold)
                                        if len(match_result) == 2:
                                            is_match, similarity = match_result
                                        elif len(match_result) == 3:
                                            is_match, similarity, match_loc = match_result
                                        else:
                                            is_match = False
                                            similarity = 0.0
                                            app.log(f"⚠️ 任务[{task_name}]分支{branch_idx}：模板匹配返回值异常")

                                        # 更新连续匹配计数器
                                        if is_match:
                                            task_continuous_match[branch_key] += 1
                                        else:
                                            task_continuous_match[branch_key] = 0

                                        current_continuous = task_continuous_match[branch_key]
                                        app.log(f"任务[{task_name}]分支{branch_idx}({ref_branch['image']})：相似度 {similarity:.3f} | 单次匹配: {is_match} | 连续匹配成功次数: {current_continuous}/{match_times}")

                                        # 【核心逻辑】第一个达到match_times阈值的分支被选中，立即跳出循环
                                        if current_continuous >= match_times:
                                            matched_branch = ref_branch
                                            matched_branch_idx = branch_idx
                                            app.log(f"✅ 任务[{task_name}]：分支{branch_idx}({ref_branch['image']})连续匹配成功，已选中此分支")
                                            break

                                    # 处理分支匹配结果
                                    if matched_branch is None:
                                        # 无分支达到匹配条件：保持当前任务索引，下一周期继续检查
                                        app.log(f"任务[{task_name}]：所有分支均未达到匹配条件，下一周期将继续检查")
                                        break  # 跳出当前循环，等待下一个识别周期

                                # 有分支达到匹配条件：执行该分支的actions
                                app.log(f"任务[{task_name}]分支{matched_branch_idx}：开始执行{len(matched_branch.get('actions', []))}个动作...")
                                
                                # 提前初始化jump_triggered，确保所有分支都能访问
                                jump_triggered = False
                                actions = matched_branch.get("actions", [])

                                # 执行动作列表：遍历JSON的action对象，提取type和params
                                for action in actions:
                                    # 防护：跳过格式错误的action（无type/params）
                                    if not isinstance(action, dict) or "type" not in action or "params" not in action:
                                        app.log(f"⚠️ 任务[{task_name}]动作格式错误，跳过：{action}")
                                        continue
                                    action_type = action["type"]
                                    params = action["params"]
                                    #  动作执行前检查停止（原有逻辑完全保留）
                                    if app.stop_flag:
                                        stop_execution = True
                                        app.log("  - 检测到停止指令，终止当前动作执行")
                                        break
                                    
                                    # 执行动作（原有逻辑完全保留）
                                    result = execute_action(target_window, action_type, params, stop_flag=app.stop_flag)

                                    if isinstance(result, tuple):
                                        log_msg = result[0]
                                    else:
                                        log_msg = result
                                    app.log(f"  - {log_msg}")
                                    
                                    if isinstance(result, tuple):
                                        if len(result) >= 2 and result[1]:
                                            stop_execution = True
                                            if len(result) >=3 and result[2] == "stop_action":
                                                stop_source = "stop_action"
                                            app.root.after(0, lambda: app._stop(is_manual=False))
                                            app.stop_flag = True
                                            break
                                        
                                        # 处理goto_task跳转（原有逻辑保留）
                                        if len(result) == 3 and not result[1]:
                                            current_task_index = result[2]
                                            current_task_index = max(0, min(current_task_index, len(task_list)-1))
                                            # 标记触发跳转
                                            jump_triggered = True
                                            break
                                
                                # 动作执行完成后：重置所有分支的连续匹配次数为0（避免重复执行）
                                for branch_idx in range(len(ref_images_list)):
                                    branch_key = f"{task_name}_branch{branch_idx}"
                                    task_continuous_match[branch_key] = 0

                                # 仅在未跳转+未停止时推进索引
                                if not jump_triggered and not stop_execution:
                                    current_task_index += 1
                                
                            except Exception as e:
                                task_error = True
                                error_msg = f"任务[{task_name}]执行失败：{str(e)}"
                                app.log(f"⚠️ 任务执行错误：{error_msg}")
                                break
                        
                        if task_error:
                            is_normal_execution = False
                            app.log(f"执行错误：{error_msg}（连续错误：{consecutive_error_count + 1}/{max_consecutive_errors}）")
                        
                        if task_error:
                            is_normal_execution = False
                            app.log(f"执行错误：{error_msg}（连续错误：{consecutive_error_count + 1}/{max_consecutive_errors}）")
                    except Exception as e:
                        is_normal_execution = False
                        error_msg = f"窗口操作失败：{str(e)}"
                        app.log(f"执行错误：{error_msg}（连续错误：{consecutive_error_count + 1}/{max_consecutive_errors}）")
            
            if is_normal_execution:
                if consecutive_error_count > 0:
                    app.log(f"✅ 执行正常，连续错误计数器已重置为0（之前：{consecutive_error_count}）")
                consecutive_error_count = 0
            else:
                consecutive_error_count += 1
            
            if consecutive_error_count >= max_consecutive_errors:
                final_msg = f"连续执行错误达到{max_consecutive_errors}次，任务已停止"
                app.log(final_msg)
                send_windows_notification("自动点击工具 - 任务停止", final_msg)
                app.stop_flag = True
                app.root.after(0, lambda: app._stop(is_manual=False))
                break
            
            if not need_wait:
                time.sleep(freq)
        
        except Exception as e:
            is_normal_execution = False
            error_msg = str(e)
            consecutive_error_count += 1
            app.log(f"❌ 未预期的执行错误：{error_msg}（连续错误：{consecutive_error_count}/{max_consecutive_errors}）")
            
            main_config = configparser.ConfigParser()
            main_config.read(MAIN_CONFIG_PATH, encoding="utf-8")
            max_consecutive_errors = int(main_config["GENERAL"]["max_execution_errors"])
            if consecutive_error_count >= max_consecutive_errors:
                final_msg = f"连续执行错误达到{max_consecutive_errors}次，任务已停止"
                app.log(final_msg)
                send_windows_notification("自动点击工具 - 任务停止", final_msg)
                app.stop_flag = True
                app.root.after(0, app._stop)
                break
            if stop_source == "stop_action":
                app.log("🛑 【指令停止】由stop动作指令触发，程序已停止")
            elif app.stop_flag and stop_source == "":
                app.log("🛑 【手动停止】用户手动触发，程序已停止")
            else:
                app.log("🛑 程序已停止")
            
            time.sleep(1.0)

# GUI核心类
class AutoClickGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("自动识别点击工具")

        config = init_main_config()  # 确保配置已初始化
        selfgeometry = config["WindowConfig"]["selfgeometry"]
        selfxy_changeable = int(config["WindowConfig"]["selfxy_changeable"])
        self.root.geometry(selfgeometry)#主窗口长宽
        self.root.resizable(selfxy_changeable, selfxy_changeable)

        self.thread = None
        self.auto_scroll = True
        self.window_list = []  # 存储窗口信息（display: 显示文本, hwnd: 句柄, title: 标题, process_name: 进程名）
        self.SCREENXY_PATH = os.path.join(BASE_DIR, "screenxy.pyw") # 坐标拾取工具路径
        self.adb_screenshot_path = os.path.join(BASE_DIR, "adb_screenshot.pyw")

        self.stop_event = threading.Event()  # 线程安全的Event
        # 兼容原有stop_flag属性的读写逻辑
        @property
        def stop_flag(self):
            return self.stop_event.is_set()
        
        @stop_flag.setter
        def stop_flag(self, value):
            if value:
                self.stop_event.set()
            else:
                self.stop_event.clear()

        # 新增↓ 热键相关初始化
        self.hotkey_keys = {Key.f9, Key.f10}  # F9+F10组合键
        self.pressed_keys = set()  # 记录当前按下的键
        self.key_listener = None
        self._init_hotkey_listener()  # 初始化热键监听
        
        init_main_config()
        self._create_widgets()
        self._load_task_groups()
        self._load_schedule_config()
        self._load_window_combobox()  # 加载窗口下拉列表
        self._load_default_window_from_config()  # 读取配置中的默认窗口
        # 新增↓ 绑定窗口关闭事件（清理热键监听）
        self.root.protocol("WM_DELETE_WINDOW", self._on_window_close)

    def _init_hotkey_listener(self):
        """初始化F9+F10组合键监听"""
        def on_key_press(key):
            """按键按下回调"""
            try:
                self.pressed_keys.add(key)
                # 检测是否同时按下F9+F10
                if self.hotkey_keys.issubset(self.pressed_keys):
                    self._handle_hotkey_trigger()
            except Exception as e:
                self.log(f"热键按下异常：{str(e)}")

        def on_key_release(key):
            """按键释放回调"""
            try:
                if key in self.pressed_keys:
                    self.pressed_keys.remove(key)
            except Exception as e:
                self.log(f"热键释放异常：{str(e)}")

        # 创建非阻塞监听器（守护线程）
        self.key_listener = keyboard.Listener(on_press=on_key_press, on_release=on_key_release)
        self.key_listener.daemon = True
        self.key_listener.start()
    # 新增↑

    # 新增↓ 组合键触发处理（核心逻辑：复用原有启动/停止方法）
    def _handle_hotkey_trigger(self):
        """F9+F10触发时切换任务状态"""
        # 避免重复触发（短时间内只响应一次）
        if hasattr(self, "_hotkey_locked") and self._hotkey_locked:
            return
        self._hotkey_locked = True
        self.root.after(500, lambda: setattr(self, "_hotkey_locked", False))  # 500ms解锁

        if self.thread and self.thread.is_alive():
            # 任务运行中 → 停止（复用原有_stop方法）
            self._stop(is_manual=True)
            self.log("🛑 热键(F9+F10)触发：停止任务组")
        else:
            # 任务未运行 → 启动（复用原有_start方法）
            self._start()
            self.log("▶️ 热键(F9+F10)触发：启动任务组")
    # 新增↑

    # 新增↓ 窗口关闭清理
    def _on_window_close(self):
        """窗口关闭时清理资源"""
        # 停止热键监听
        if self.key_listener and self.key_listener.is_alive():
            self.key_listener.stop()
        # 停止任务线程
        if self.thread and self.thread.is_alive():
            self.stop_flag = True
            self.thread.join(timeout=2)
        self.root.destroy()
    # 新增↑

    def on_screenshot_mode_change(self, event=None):
        """截图模式变更回调，保存到配置（匹配点击模式回调风格）"""
        try:
            config = configparser.ConfigParser()
            config.read(MAIN_CONFIG_PATH, encoding="utf-8")
            config["GENERAL"]["screenshot_mode"] = self.screenshot_mode_var.get()
            with open(MAIN_CONFIG_PATH, "w", encoding="utf-8") as f:
                config.write(f)
            self.log(f"✅ 截图模式已切换为：{self.screenshot_mode_var.get()}")
        except Exception as e:
            self.log(f"⚠️ 保存截图模式配置失败：{str(e)}")
            messagebox.showerror("错误", f"保存截图模式失败：{e}")

    def on_adb_device_serial_change(self, event=None):
        """ADB设备Serial号变更回调，保存到配置"""
        try:
            config = configparser.ConfigParser()
            config.read(MAIN_CONFIG_PATH, encoding="utf-8")
            config["ADBConfig"]["adb_device_serial"] = self.adb_device_serial_var.get().strip()
            with open(MAIN_CONFIG_PATH, "w", encoding="utf-8") as f:
                config.write(f)
            self.log(f"✅ ADB设备Serial已更新为：{self.adb_device_serial_var.get().strip()}")
        except Exception as e:
            self.log(f"⚠️ 保存ADB设备Serial失败：{str(e)}")
            messagebox.showerror("错误", f"保存ADB设备Serial失败：{e}")

    def _on_click_mode_change(self, event):
        """点击模式选择变化时，实时保存配置"""
        # 读取当前配置
        main_config = configparser.ConfigParser()
        main_config.read(MAIN_CONFIG_PATH, encoding="utf-8")

        # 转换下拉选项为配置值
        selected_text = self.click_mode_var.get()
        click_mode = "sendmessage" if "SendMessage" in selected_text else "pyautogui"

        # 实时更新配置
        main_config["GENERAL"]["click_mode"] = click_mode
        with open(MAIN_CONFIG_PATH, "w", encoding="utf-8") as f:
            main_config.write(f)

        # 日志提示用户
        self.log(f"✅ 点击模式已切换为：{selected_text}")

    def _on_task_group_change(self, event):
        """任务组下拉框选中变化时，实时保存current_task_group到配置"""
        try:
            # 读取当前选中的任务组名称
            selected_task_group = self.task_var.get().strip()
            if not selected_task_group:
                self.log("⚠️ 任务组名称不能为空，跳过配置更新")
                return
            
            # 读取配置并更新current_task_group
            main_config = configparser.ConfigParser()
            main_config.read(MAIN_CONFIG_PATH, encoding="utf-8")
            main_config["GENERAL"]["current_task_group"] = selected_task_group
            
            # 保存到配置文件
            with open(MAIN_CONFIG_PATH, "w", encoding="utf-8") as f:
                main_config.write(f)
            
            # 打印日志提示
            self.log(f"✅ 任务组已切换为：{selected_task_group}")
        except Exception as e:
            self.log(f"⚠️ 保存任务组配置失败：{str(e)}")
            messagebox.showerror("错误", f"保存任务组配置失败：{e}")

    def _on_window_change(self, event):
        """窗口切换时，实时更新系统依赖的3个核心配置字段"""
        try:
            # 1. 通过索引获取完整的窗口对象（正确方式）
            idx = self.window_combobox.current()
            if idx < 0 or idx >= len(self.window_list):
                self.log("⚠️ 未选中有效窗口，跳过配置更新")
                return
            win = self.window_list[idx]  # 拿到完整窗口信息（hwnd/title/process_name）
            
            # 2. 写入系统依赖的3个核心字段（与get_target_window_from_config对应）
            config = configparser.ConfigParser()
            config.read(MAIN_CONFIG_PATH, encoding="utf-8")
            config["GENERAL"]["target_window_hwnd"] = str(win["hwnd"])
            config["GENERAL"]["target_window_title"] = win["title"]
            config["GENERAL"]["target_program_name"] = win["process_name"]
            
            # 3. 保存配置
            with open(MAIN_CONFIG_PATH, "w", encoding="utf-8") as f:
                config.write(f)
            
            # 4. 同时更新窗口信息显示（合并_on_window_select的逻辑，避免冗余绑定）
            self._update_window_info(idx)
            self.log(f"✅ 窗口已切换为：{win['display']}，配置已实时更新")
        except Exception as e:
            self.log(f"⚠️ 保存窗口配置失败：{str(e)}")
            messagebox.showerror("错误", f"保存窗口配置失败：{e}")

    def _run_adb_command(self):
        """执行用户输入的完整ADB命令（需带adb前缀），彻底剥离前缀避免重复"""
        # 1. 获取用户输入的完整命令
        full_cmd_input = self.adb_cmd_var.get().strip()
        if not full_cmd_input:
            self.log("⚠️ ADB命令不能为空！请输入完整命令（如'adb devices'、'adb shell'）")
            return

        # 2. 严格验证命令格式：必须以'adb'开头，且仅允许两种合法格式：
        #    - 格式1：仅输入'adb'（无参数）
        #    - 格式2：'adb ' + 命令参数（adb后必须跟空格）
        if not full_cmd_input.startswith("adb"):
            self.log(f"❌ 命令格式错误！必须以'adb'开头（当前输入：{full_cmd_input}）")
            return
        # 检查是否是非法格式（如'adbdevices'、'adb-xxx'，无空格且长度>3）
        if len(full_cmd_input) > 3 and not full_cmd_input[3].isspace():
            self.log(f"❌ 命令格式错误！'adb'后必须跟空格（当前输入：{full_cmd_input}，正确示例：'adb devices'）")
            return

        # 3. 验证ADB环境是否可用
        adb_valid, adb_path_or_msg = get_adb_executable_path()
        if not adb_valid:
            self.log(f"❌ ADB环境不可用：{adb_path_or_msg}")
            return

        # 4. 核心修复：彻底剥离开头的'adb'前缀，提取纯参数（关键步骤）
        main_config = configparser.ConfigParser()
        main_config.read(MAIN_CONFIG_PATH, encoding="utf-8")
        adb_usage_mode = main_config["ADBConfig"].get("adb_usage_mode", "1")
        
        # 剥离'adb'和后面的所有空格，剩下的就是纯参数（支持多空格场景）
        cmd_params = full_cmd_input[3:].lstrip()  # 从第3个字符（'adb'结束）开始，去掉左侧所有空格

        # 构建最终执行命令（无参数时直接执行ADB）
        if adb_usage_mode == "3":
            # 系统环境变量模式：直接使用用户输入的完整命令（避免二次处理）
            final_cmd = full_cmd_input
        else:
            # 模式1/2：用配置的ADB路径 + 纯参数（无参数时仅执行ADB）
            if cmd_params:
                final_cmd = f'"{adb_path_or_msg}" {cmd_params}'
            else:
                final_cmd = f'"{adb_path_or_msg}"'

        # 5. 执行命令并捕获结果
        self.log(f"▶️ 正在执行ADB命令：{final_cmd}")
        try:
            result = subprocess.run(
                final_cmd,
                shell=True,
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                encoding="utf-8",
                timeout=30  # 30秒超时
            )
            # 6. 输出执行结果到日志
            if not cmd_params:
                self.log(f"✅ ADB无参数命令执行成功（显示帮助信息）！")
            else:
                self.log(f"✅ ADB命令执行成功！")
            
            if result.stdout:
                self.log(f"📤 输出：\n{result.stdout.strip()}")
            if result.stderr:
                self.log(f"⚠️ 警告输出：\n{result.stderr.strip()}")
        except subprocess.TimeoutExpired:
            self.log(f"❌ ADB命令执行超时（30秒）：{final_cmd}")
        except subprocess.CalledProcessError as e:
            # 特殊处理：ADB无参数执行时返回码1（正常输出帮助信息）
            if not cmd_params and e.returncode == 1 and e.stdout:
                self.log(f"✅ ADB无参数命令执行成功（返回码1为正常现象）！")
                self.log(f"📤 输出（ADB帮助信息）：\n{e.stdout.strip()}")
            else:
                self.log(f"❌ ADB命令执行失败（返回码：{e.returncode}）")
                self.log(f"命令：{final_cmd}")
                if e.stdout:
                    self.log(f"📤 输出：\n{e.stdout.strip()}")
                if e.stderr:
                    self.log(f"❌ 错误输出：\n{e.stderr.strip()}")
        except Exception as e:
            self.log(f"❌ ADB命令执行异常：{str(e)}")


    
    def _create_widgets(self):
        # 主布局：移除坐标拾取标签页，保留原有结构
        main_notebook = ttk.Notebook(self.root)
        main_notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 1. 目标窗口标签页（精简为下拉列表）
        window_frame = ttk.Frame(main_notebook)
        main_notebook.add(window_frame, text="目标窗口")
        
        # 1.1 紧凑的窗口选择区域（下拉列表+刷新按钮）
        select_frame = ttk.LabelFrame(window_frame, text="窗口选择", padding="10")
        select_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # 下拉列表：显示窗口列表（标题+进程名）
        ttk.Label(select_frame, text="选择目标窗口：").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.window_var = tk.StringVar()
        self.window_combobox = ttk.Combobox(
            select_frame, 
            textvariable=self.window_var, 
            state="readonly",
            width=60  # 紧凑宽度
        )
        self.window_combobox.bind("<<ComboboxSelected>>", self._on_window_change)
        self.window_combobox.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        # 操作按钮：刷新+保存+模糊匹配配置
        ttk.Button(select_frame, text="刷新列表", command=self._load_window_combobox).grid(row=0, column=2, padx=5, pady=5)
        ttk.Button(select_frame, text="选中", command=self._save_window_to_config).grid(row=0, column=3, padx=5, pady=5)
        ttk.Button(select_frame, text="模糊匹配配置", command=self._edit_window_match_config).grid(row=0, column=4, padx=5, pady=5)
        
        # 1.2 极简窗口信息（仅一行显示关键信息）
        info_frame = ttk.LabelFrame(window_frame, text="窗口信息", padding="10")
        info_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.window_info_var = tk.StringVar(value="未选择窗口 | 标题：- | 进程：- | 句柄：-")
        ttk.Label(info_frame, textvariable=self.window_info_var).pack(side=tk.LEFT, padx=5, pady=5)
        
        # =============================================
        # 新增：模板匹配步长配置区域（复用原有样式）
        # =============================================
        # 1. 创建和原有区域样式一致的LabelFrame
        match_step_frame = ttk.LabelFrame(window_frame, text="模板匹配配置", padding="10")
        match_step_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # 2. 初始化配置读取（复用原有config逻辑）
        self.match_step_var = tk.StringVar()
        # 读取配置中的步长，兼容配置缺失/格式错误
        self._load_match_step_from_config()
        
        # 3. 步长输入区域（布局和"窗口选择"区域对齐，保持视觉统一）
        ttk.Label(match_step_frame, text="匹配缩略图步长：").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        # 输入框：宽度适配，仅允许输入数字
        step_entry = ttk.Entry(
            match_step_frame, 
            textvariable=self.match_step_var,
            width=20  # 紧凑宽度，和原有组件风格一致
        )
        step_entry.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        # 4. 保存按钮：复用原有配置写入逻辑风格
        ttk.Button(
            match_step_frame, 
            text="保存步长", 
            command=self._save_match_step_to_config
        ).grid(row=0, column=2, padx=5, pady=5)
        
        # 5. 提示标签：告知合法范围，提升用户体验
        ttk.Label(
            match_step_frame, 
            text="（步长范围：0.001~0.2，默认0.05）",
            foreground="#666666"  # 浅灰色，不干扰主视觉
        ).grid(row=0, column=3, padx=5, pady=5, sticky=tk.W)

        # 2. 任务配置标签页
        task_frame = ttk.Frame(main_notebook)
        main_notebook.add(task_frame, text="任务配置")
        
        # 任务组选择
        ttk.Label(task_frame, text="任务组：").grid(row=0, column=0, padx=5, pady=5)
        self.task_var = tk.StringVar()
        self.task_combobox = ttk.Combobox(task_frame, textvariable=self.task_var, state="readonly")
        self.task_combobox.grid(row=0, column=1, padx=5, pady=5)
        self.task_combobox.bind("<<ComboboxSelected>>", self._on_task_group_change)
        ttk.Button(task_frame, text="刷新", command=self._load_task_groups).grid(row=0, column=2, padx=5)
        ttk.Button(task_frame, text="新建", command=self._new_task_group).grid(row=0, column=3, padx=5)
        ttk.Button(task_frame, text="编辑", command=self._edit_task_config).grid(row=0, column=4, padx=5)
        ttk.Button(task_frame, text="坐标拾取", command=self._open_screenxy).grid(row=0, column=5, padx=5)
        ttk.Button(task_frame, text="ADB截图工具", command=self._open_adb_screenshot).grid(row=0, column=6, padx=5)
        
        # 定时配置
        schedule_frame = ttk.LabelFrame(task_frame, text="定时运行配置", padding="10")
        schedule_frame.grid(row=1, column=0, columnspan=5, sticky=tk.W+tk.E, padx=10, pady=5)
        
        self.enable_schedule_var = tk.BooleanVar()
        schedule_check = ttk.Checkbutton(
            schedule_frame, 
            text="启用定时运行", 
            variable=self.enable_schedule_var,
            command=self._toggle_schedule_widgets
        )
        schedule_check.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        
        ttk.Label(schedule_frame, text="定时时间：").grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        self.schedule_time_var = tk.StringVar(value="00:00:00")
        self.schedule_time_entry = ttk.Entry(
            schedule_frame, 
            textvariable=self.schedule_time_var, 
            width=12
        )
        self.schedule_time_entry.grid(row=0, column=2, padx=5, pady=5)
        ttk.Label(schedule_frame, text="（HH:MM:SS）").grid(row=0, column=3, padx=5, pady=5)
        
        ttk.Label(schedule_frame, text="运行模式：").grid(row=0, column=4, padx=5, pady=5, sticky=tk.W)
        self.schedule_mode_var = tk.StringVar()
        self.schedule_mode_combobox = ttk.Combobox(
            schedule_frame, 
            textvariable=self.schedule_mode_var,
            values=["仅一次", "始终"],
            state="readonly",
            width=8
        )
        self.schedule_mode_combobox.grid(row=0, column=5, padx=5, pady=5)
        self.schedule_mode_combobox.current(0)
        
        ttk.Button(
            schedule_frame, 
            text="保存配置", 
            command=self._save_schedule_config
        ).grid(row=0, column=6, padx=10, pady=5)

          # ---------------------- 新增：点击模式切换下拉列表 ---------------------
        # 放在定时运行配置框的下方（task_frame的row=2）
        ttk.Label(task_frame, text="点击模式：").grid(row=2, column=0, padx=5, pady=8, sticky=tk.W)
        self.click_mode_var = tk.StringVar()
        self.click_mode_combobox = ttk.Combobox(
            task_frame,
            textvariable=self.click_mode_var,
            values=["SendMessage消息点击", "PyAutoGUI硬件点击"],
            state="readonly",
            width=35  # 加长下拉栏宽度（原15→20）
        )
        self.click_mode_combobox.grid(row=2, column=1, padx=5, pady=8, sticky=tk.W)
        # 绑定选中事件，实现「选择即更新」
        self.click_mode_combobox.bind("<<ComboboxSelected>>", self._on_click_mode_change)

        main_config = configparser.ConfigParser()
        main_config.read(MAIN_CONFIG_PATH, encoding="utf-8")

        # ========== 新增：截图模式选择（点击模式下一行） ==========
        ttk.Label(task_frame, text="截图模式：").grid(row=6, column=0, padx=5, pady=3, sticky="w")
        self.screenshot_mode_var = tk.StringVar(value=main_config["GENERAL"]["screenshot_mode"])
        screenshot_mode_combo = ttk.Combobox(task_frame, textvariable=self.screenshot_mode_var, values=["Win32GUI", "ADB","Win32Memory"], state="readonly", width=10)
        screenshot_mode_combo.grid(row=6, column=1, padx=5, pady=3, sticky="we")
        screenshot_mode_combo.bind("<<ComboboxSelected>>", self.on_screenshot_mode_change)
        # =======================================================

        # ========== 新增：ADB设备Serial号配置（截图模式为adb时生效） ==========
        ttk.Label(task_frame, text="ADB设备Serial：").grid(row=7, column=0, padx=5, pady=3, sticky="w")
        self.adb_device_serial_var = tk.StringVar(value=main_config["ADBConfig"]["adb_device_serial"])
        adb_device_serial_entry = ttk.Entry(task_frame, textvariable=self.adb_device_serial_var, width=15)
        adb_device_serial_entry.grid(row=7, column=1, padx=5, pady=3, sticky="we")
        adb_device_serial_entry.bind("<FocusOut>", self.on_adb_device_serial_change)
        # ====================================================================
        def comfort_confirm():
            """仅作安慰作用，无实际功能"""
            self.log("✅ 设备Serial已更新")

        # 紧贴输入框右侧添加按钮（column=2，padx=0紧贴）
        comfort_btn = ttk.Button(task_frame, text="确认", command=comfort_confirm, width=6)
        comfort_btn.grid(row=7, column=2, padx=(0, 5), pady=3, sticky="w")
        # ====================================================================



        # ========== 2. ADB配置标签页 ==========
        # 创建ADB配置标签页（在任务配置右侧）
        adb_frame = ttk.Frame(main_notebook)
        main_notebook.add(adb_frame, text="ADB配置")

        # 读取当前ADB配置
        main_config = configparser.ConfigParser()
        main_config.read(MAIN_CONFIG_PATH, encoding="utf-8")
        if "ADBConfig" not in main_config:
            main_config["ADBConfig"] = {}

        # 初始化配置变量（原数字值）
        adb_enabled = main_config["ADBConfig"].get("adb_enabled", "1")
        adb_usage_mode = main_config["ADBConfig"].get("adb_usage_mode", "1")
        adb_position = main_config["ADBConfig"].get("adb_position", "")

        # ---------------------- 下拉框显示文字与实际值的映射 ----------------------
        enabled_map = {"是": "1", "否": "0"}
        mode_map = {"内置": "1", "自定义": "2", "系统环境变量": "3"}
        enabled_rev_map = {v: k for k, v in enabled_map.items()}
        mode_rev_map = {v: k for k, v in mode_map.items()}

        # ---------------------- 新增：灰显样式定义 ----------------------
        style = ttk.Style()
        style.configure("Gray.TEntry", foreground="#888888")  # 灰显输入框
        style.configure("Normal.TEntry", foreground="#000000") # 正常输入框
        style.configure("Gray.TLabel", foreground="#888888")   # 灰显标签
        style.configure("Normal.TLabel", foreground="#000000") # 正常标签

        # ---------------------- 控件布局 ----------------------
        # 1. ADB功能启用
        current_enabled_text = "是" if adb_enabled == "1" else "否"
        adb_enabled_var = tk.StringVar(value=current_enabled_text)

        ttk.Label(adb_frame, text="ADB功能启用：").grid(
            row=0, column=0, padx=10, pady=8, sticky="w"
        )
        adb_enabled_combobox = ttk.Combobox(
            adb_frame, textvariable=adb_enabled_var, values=["是", "否"], width=12, state="readonly"
        )
        adb_enabled_combobox.grid(row=0, column=1, padx=2, pady=8, sticky="w")
        ttk.Label(adb_frame, text="是=启用ADB检查 | 否=忽略配置").grid(
            row=0, column=2, padx=5, pady=8, sticky="w"
        )

        # 2. ADB使用模式
        current_mode_text = {v: k for k, v in mode_map.items()}.get(adb_usage_mode, "内置")
        adb_mode_var = tk.StringVar(value=current_mode_text)

        ttk.Label(adb_frame, text="ADB使用模式：").grid(
            row=1, column=0, padx=10, pady=8, sticky="w"
        )
        adb_mode_combobox = ttk.Combobox(
            adb_frame, textvariable=adb_mode_var, values=["内置", "自定义", "系统环境变量"], width=12, state="readonly"
        )
        adb_mode_combobox.grid(row=1, column=1, padx=2, pady=8, sticky="w")
        mode_note = ttk.Label(
            adb_frame,
            text="内置=脚本内置 | 自定义=手动路径 | 系统环境变量=全局ADB",
        )
        mode_note.grid(row=1, column=2, padx=5, pady=8, sticky="w")

        # 3. 自定义ADB路径（新增控件命名，用于样式切换）
        adb_path_label = ttk.Label(adb_frame, text="自定义ADB路径：")
        adb_path_label.grid(row=2, column=0, padx=10, pady=8, sticky="w")

        adb_position_var = tk.StringVar(value=adb_position)
        adb_position_entry = ttk.Entry(adb_frame, textvariable=adb_position_var, width=35)
        adb_position_entry.grid(row=2, column=1, padx=2, pady=8, sticky="w")

        def select_adb_path():
            adb_file = filedialog.askopenfilename(
                title="选择ADB可执行文件",
                initialdir=os.path.dirname(adb_position) if adb_position else BASE_DIR,
                filetypes=[("ADB执行文件", "adb.exe"), ("所有文件", "*.*")]
            )
            if adb_file:
                adb_position_var.set(adb_file)
        adb_browse_btn = ttk.Button(adb_frame, text="浏览", command=select_adb_path)
        adb_browse_btn.grid(row=2, column=2, padx=5, pady=8, sticky="w")


        # 4. 验证ADB环境按钮
        def validate_adb_env():
            #from your_script_name import validate_adb_environment
            is_valid, msg = validate_adb_environment()
            log_time = time.strftime("%Y-%m-%d %H:%M:%S")
            if is_valid:
                app.log(f"[{log_time}] ADB验证成功：{msg}")
            else:
                app.log(f"[{log_time}] ADB验证失败：{msg}")

        ttk.Button(
            adb_frame, text="验证ADB环境", command=validate_adb_env, width=15
        ).grid(row=3, column=1, padx=2, pady=10, sticky="w")

        # ========== 新增：ADB命令行执行区域 ==========
        # 命令输入框（row=4）
        ttk.Label(adb_frame, text="ADB命令执行：").grid(
            row=4, column=0, padx=10, pady=8, sticky="w"
        )
        self.adb_cmd_var = tk.StringVar(value="")  # 存储完整ADB命令（含adb前缀）
        adb_cmd_entry = ttk.Entry(
            adb_frame, textvariable=self.adb_cmd_var, width=50
        )
        adb_cmd_entry.grid(row=4, column=2, padx=2, pady=8, sticky="w")

        # 执行按钮（row=4，与输入框同排）
        def execute_adb_cmd():
            self._run_adb_command()
        ttk.Button(
            adb_frame, text="执行命令", command=execute_adb_cmd, width=10
        ).grid(row=4, column=3, padx=5, pady=8, sticky="w")

        # 绑定Enter键触发执行
        adb_cmd_entry.bind("<Return>", lambda event: self._run_adb_command())

        # 提示文本（修改为要求带adb前缀）
        cmd_tip_label = ttk.Label(
            adb_frame,
            text="提示：输入完整ADB命令（需带'adb '前缀，如'adb devices'、'adb shell getprop'）",
        )
        cmd_tip_label.grid(row=3, column=2, columnspan=3, padx=10, pady=5, sticky="w")

        # 5. 实时更新配置（精准记录修改内容，无变化不输出日志）
        def update_adb_config(*args):
            # ========== 1. 读取修改前的旧配置（原始值） ==========
            old_config = configparser.ConfigParser()
            old_config.read(MAIN_CONFIG_PATH, encoding="utf-8")
            if "ADBConfig" not in old_config:
                old_config["ADBConfig"] = {}
            
            # 旧值（无配置则用默认值）
            old_enabled = old_config["ADBConfig"].get("adb_enabled", "1")
            old_mode = old_config["ADBConfig"].get("adb_usage_mode", "1")
            old_position = old_config["ADBConfig"].get("adb_position", "")

            # ========== 2. 获取修改后的新值（转换后） ==========
            new_enabled_text = adb_enabled_var.get()  # 是/否
            new_mode_text = adb_mode_var.get()        # 内置/自定义/系统环境变量
            new_position = adb_position_var.get()     # 路径字符串

            # 转换为配置文件的原始值
            new_enabled = enabled_map[new_enabled_text]
            new_mode = mode_map[new_mode_text]

            # ========== 3. 对比新旧值，收集变化项 ==========
            changed_items = []
            # 对比ADB启用状态
            if old_enabled != new_enabled:
                old_enabled_text = "是" if old_enabled == "1" else "否"
                changed_items.append(f"ADB启用状态：{old_enabled_text} → {new_enabled_text}")
            # 对比ADB使用模式
            if old_mode != new_mode:
                old_mode_text = {v: k for k, v in mode_map.items()}.get(old_mode, "内置")
                changed_items.append(f"ADB使用模式：{old_mode_text} → {new_mode_text}")
            # 对比自定义ADB路径
            if old_position != new_position:
                # 路径为空时显示“空”，更易读
                old_pos_show = old_position if old_position else "空"
                new_pos_show = new_position if new_position else "空"
                changed_items.append(f"自定义ADB路径：{old_pos_show} → {new_pos_show}")

            # ========== 4. 无变化则直接返回，不输出日志 ==========
            if not changed_items:
                return

            # ========== 5. 有变化则写入配置 + 输出精准日志 ==========
            try:
                # 写入新配置
                main_config = configparser.ConfigParser()
                main_config.read(MAIN_CONFIG_PATH, encoding="utf-8")
                if "ADBConfig" not in main_config:
                    main_config["ADBConfig"] = {}
                
                main_config["ADBConfig"]["adb_enabled"] = new_enabled
                main_config["ADBConfig"]["adb_usage_mode"] = new_mode
                main_config["ADBConfig"]["adb_position"] = new_position
                
                with open(MAIN_CONFIG_PATH, "w", encoding="utf-8") as f:
                    main_config.write(f)
                
                # 输出精准日志：拼接所有变化项
                log_time = time.strftime("%Y-%m-%d %H:%M:%S")
                changed_detail = "；".join(changed_items)
                app.log(f"[{log_time}] ADB配置更新：{changed_detail}")
            
            except Exception as e:
                # 错误日志保留详细信息
                log_time = time.strftime("%Y-%m-%d %H:%M:%S")
                app.log(f"[{log_time}] ADB配置更新失败：{str(e)}")

        # 绑定更新事件
        adb_enabled_var.trace_add("write", update_adb_config)
        adb_mode_var.trace_add("write", update_adb_config)
        adb_position_var.trace_add("write", update_adb_config)

        # 6. 提示信息
        tip_label = ttk.Label(
            adb_frame,
            text="提示：自定义路径仅在“是+自定义”模式下生效",
        )
        tip_label.grid(row=4, column=0, columnspan=3, padx=10, pady=5, sticky="w")



        # 3. 控制和日志区域
        control_frame = ttk.Frame(self.root)
        control_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.start_btn = ttk.Button(control_frame, text="启动", command=self._start)
        self.start_btn.pack(side=tk.LEFT, padx=5)
        self.stop_btn = ttk.Button(control_frame, text="停止", command=self._stop, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=5)
        ttk.Button(control_frame, text="编辑总配置", command=self._edit_main_config).pack(side=tk.LEFT, padx=5)
        ttk.Button(control_frame, text="清除日志", command=self._clear_log).pack(side=tk.LEFT, padx=5)
        
        self.scroll_btn = ttk.Button(control_frame, text="日志自动滚动：开", command=self._toggle_auto_scroll)
        self.scroll_btn.pack(side=tk.RIGHT, padx=5)

        self.scroll_btn = ttk.Button(control_frame, text="关于", command=self._show_about_dialog)
        self.scroll_btn.pack(side=tk.RIGHT, padx=5)
        
        # 日志区
        log_frame = ttk.LabelFrame(self.root, text="运行日志", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, state=tk.DISABLED) 
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
    
    # 新增打开坐标拾取脚本的方法
    def _open_screenxy(self):
            try:
                if os.path.exists(self.SCREENXY_PATH):
                    if sys.platform == "win32":
                        os.startfile(self.SCREENXY_PATH)
                    elif sys.platform == "darwin":
                        subprocess.run(["open", self.SCREENXY_PATH])
                    else:
                        subprocess.run(["xdg-open", self.SCREENXY_PATH])
                    self.log(f"✅ 已打开坐标拾取工具：{self.SCREENXY_PATH}")
                else:
                    self.log(f"❌ 未找到坐标拾取工具：{self.SCREENXY_PATH}")
                    messagebox.showerror("错误", f"未找到脚本文件：{self.SCREENXY_PATH}")
            except Exception as e:
                self.log(f"❌ 打开坐标拾取工具失败：{str(e)}")
                messagebox.showerror("错误", f"打开脚本失败：{str(e)}")

    def _open_adb_screenshot(self):
            try:
                if os.path.exists(self.adb_screenshot_path):
                    if sys.platform == "win32":
                        os.startfile(self.adb_screenshot_path)
                    elif sys.platform == "darwin":
                        subprocess.run(["open", self.adb_screenshot_path])
                    else:
                        subprocess.run(["xdg-open", self.adb_screenshot_path])
                    self.log(f"✅ 已打开ADB截图工具：{self.adb_screenshot_path}")
                else:
                    self.log(f"❌ 未找到ADB截图工具：{self.adb_screenshot_path}")
                    messagebox.showerror("错误", f"未找到脚本文件：{self.adb_screenshot_path}")
            except Exception as e:
                self.log(f"❌ 打开ADB截图工具失败：{str(e)}")
                messagebox.showerror("错误", f"打开脚本失败：{str(e)}")

    # 新增：配置读写方法（复用原有config逻辑）
    # =============================================
    def _load_match_step_from_config(self):
        """从配置文件读取匹配步长，初始化输入框"""
        config = configparser.ConfigParser()
        if os.path.exists(MAIN_CONFIG_PATH):
            config.read(MAIN_CONFIG_PATH, encoding="utf-8")
        # 兼容GENERAL节点缺失的情况
        if "GENERAL" not in config:
            config["GENERAL"] = {}
        # 读取步长，默认值0.01，兼容格式错误
        try:
            step = float(config["GENERAL"].get("template_match_step", "0.01"))
            # 合法性校验，限制范围
            step = max(0.001, min(0.2, step))
            self.match_step_var.set(f"{step:.3f}")  # 保留3位小数，提升可读性
        except (ValueError, TypeError):
            self.match_step_var.set("0.01")  # 格式错误时用默认值

    def _save_match_step_to_config(self):
        """将输入的步长保存到配置文件，带合法性校验"""
        try:
            # 1. 读取输入值并校验
            input_step = self.match_step_var.get().strip()
            step_val = float(input_step)
            # 2. 范围校验
            if not (0.001 <= step_val <= 0.2):
                self.log(f"⚠️ 步长值{step_val}超出范围（0.001~0.2）")
                # 恢复合法值
                self._load_match_step_from_config()
                return
            # 3. 写入配置文件（复用原有config写入逻辑）
            config = configparser.ConfigParser()
            if os.path.exists(MAIN_CONFIG_PATH):
                config.read(MAIN_CONFIG_PATH, encoding="utf-8")
            if "GENERAL" not in config:
                config["GENERAL"] = {}
            # 保存为字符串，保留3位小数
            config["GENERAL"]["template_match_step"] = f"{step_val:.3f}"
            # 写入文件
            with open(MAIN_CONFIG_PATH, "w", encoding="utf-8") as f:
                config.write(f)
            self.log(f"✅ 匹配步长已保存为：{step_val:.3f}")
        except ValueError:
            self.log(f"❌ 请输入有效的数字（如0.01）！")
            # 恢复合法值
            self._load_match_step_from_config()

    # 窗口下拉列表相关
    def _load_window_combobox(self):
        """加载窗口列表到下拉框，并自动匹配默认窗口"""
        self.window_list = get_all_visible_windows_simple()
        display_list = [win["display"] for win in self.window_list]
        self.window_combobox["values"] = display_list
        
        # 正确获取下拉框实际选项数量，不手动修改其值
        actual_count = len(self.window_combobox["values"])
        if actual_count > 0:  # 仅当有有效选项时执行清空
            self.window_combobox.current(0)
        else:
            self.window_combobox.set("")  # 无选项时清空显示
            
        if display_list:
            # 尝试从配置加载默认窗口
            config = configparser.ConfigParser()
            config.read(MAIN_CONFIG_PATH, encoding="utf-8")
            target_hwnd = config["GENERAL"]["target_window_hwnd"]
            target_title = config["GENERAL"]["target_window_title"]
            target_program = config["GENERAL"]["target_program_name"]
            
            matched_idx = -1
            
            # 1. 优先匹配窗口句柄
            if target_hwnd and target_hwnd.isdigit():
                hwnd = int(target_hwnd)
                for idx, win in enumerate(self.window_list):
                    if win["hwnd"] == hwnd:
                        matched_idx = idx
                        break
            
            # 2. 句柄匹配失败时，尝试匹配窗口标题
            if matched_idx == -1 and target_title:
                for idx, win in enumerate(self.window_list):
                    if target_title in win["title"]:
                        matched_idx = idx
                        break
            
            # 3. 标题匹配失败时，尝试匹配进程名
            if matched_idx == -1 and target_program:
                for idx, win in enumerate(self.window_list):
                    if target_program in win["process_name"]:
                        matched_idx = idx
                        break
            
            # 设置选中项并更新信息
            if matched_idx != -1:
                self.window_combobox.current(matched_idx)
                self._update_window_info(matched_idx)
                self.log(f"✅ 已刷新窗口列表并匹配到默认窗口：{self.window_list[matched_idx]['display']}")
            else:
                # 无匹配时选中第一个窗口
                self.window_combobox.current(0)
                self._update_window_info(0)
                self.log(f"ℹ️ 已刷新窗口列表（共 {len(self.window_list)} 个可见窗口），未找到默认窗口，选中第一个窗口")
        else:
            self.log("ℹ️ 已刷新窗口列表，未发现可见窗口")
    
    #def _on_window_select(self, event):
    #    """选择下拉列表中的窗口"""
    #    idx = self.window_combobox.current()
    #    if idx >= 0 and idx < len(self.window_list):
    #        self._update_window_info(idx)
    
    def _update_window_info(self, idx):
        """更新窗口信息（极简显示）"""
        win = self.window_list[idx]
        info_text = f"标题：{win['title']} | 进程：{win['process_name']} | 句柄：{win['hwnd']}"
        self.window_info_var.set(info_text)
    
    def _save_window_to_config(self):
        """将选中的窗口保存为默认（写入总配置）"""
        idx = self.window_combobox.current()
        if idx < 0 or idx >= len(self.window_list):
            messagebox.showwarning("提示", "请先选择窗口")
            return
        
        win = self.window_list[idx]
        config = configparser.ConfigParser()
        config.read(MAIN_CONFIG_PATH, encoding="utf-8")
        config["GENERAL"]["target_window_hwnd"] = str(win["hwnd"])
        config["GENERAL"]["target_window_title"] = win["title"]
        config["GENERAL"]["target_program_name"] = win["process_name"]
        
        with open(MAIN_CONFIG_PATH, "w", encoding="utf-8") as f:
            config.write(f)
        
        self.log(f"✅ 已将窗口「{win['display']}」写入配置文件")

    
    
    def _load_default_window_from_config(self):
        """从配置读取默认窗口并选中"""
        config = configparser.ConfigParser()
        config.read(MAIN_CONFIG_PATH, encoding="utf-8")
        target_hwnd = config["GENERAL"]["target_window_hwnd"]
        
        if target_hwnd and target_hwnd.isdigit():
            hwnd = int(target_hwnd)
            # 查找对应窗口在下拉列表中的索引
            for idx, win in enumerate(self.window_list):
                if win["hwnd"] == hwnd:
                    self.window_combobox.current(idx)
                    self._update_window_info(idx)
                    self.log(f"✅ 已加载配置中的默认窗口：{win['display']}")
                    return
        
        # 无精确句柄时，尝试模糊匹配标题/进程名
        target_title = config["GENERAL"]["target_window_title"]
        target_program = config["GENERAL"]["target_program_name"]
        if target_title or target_program:
            for idx, win in enumerate(self.window_list):
                if (target_title and target_title in win["title"]) or (target_program and target_program in win["process_name"]):
                    self.window_combobox.current(idx)
                    self._update_window_info(idx)
                    self.log(f"✅ 模糊匹配到默认窗口：{win['display']}")
                    return
        
        self.log("ℹ️ 未找到配置中的默认窗口，使用第一个窗口")
    
    def _edit_window_match_config(self):
        """编辑模糊匹配的关键词（进程名/窗口标题）"""
        config = configparser.ConfigParser()
        config.read(MAIN_CONFIG_PATH, encoding="utf-8")
        
        # 新建临时窗口，输入模糊匹配关键词
        top = tk.Toplevel(self.root)
        top.title("模糊匹配配置")
        top.geometry("300x270")
        top.transient(self.root)
        top.grab_set()
        
        # 进程名关键词
        ttk.Label(top, text="进程名关键词（模糊匹配）：").grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)
        prog_var = tk.StringVar(value=config["GENERAL"]["target_program_name"])
        prog_entry = ttk.Entry(top, textvariable=prog_var, width=30)
        prog_entry.grid(row=1, column=0, padx=10, pady=10)
        
        # 窗口标题关键词
        ttk.Label(top, text="窗口标题关键词（模糊匹配）：").grid(row=2, column=0, padx=10, pady=10, sticky=tk.W)
        title_var = tk.StringVar(value=config["GENERAL"]["target_window_title"])
        title_entry = ttk.Entry(top, textvariable=title_var, width=30)
        title_entry.grid(row=3, column=0, padx=10, pady=10)
        
        # 保存按钮
        def save():
            config["GENERAL"]["target_program_name"] = prog_var.get().strip()
            config["GENERAL"]["target_window_title"] = title_var.get().strip()
            config["GENERAL"]["target_window_hwnd"] = ""  # 清空精确句柄，启用模糊匹配
            with open(MAIN_CONFIG_PATH, "w", encoding="utf-8") as f:
                config.write(f)
            self.log(f"✅ 保存模糊匹配配置：进程名={prog_var.get()}, 标题={title_var.get()}")
            top.destroy()
        
        ttk.Button(top, text="保存", command=save).grid(row=4, column=0, padx=10, pady=10)
    
    # 日志自动滚动功能
    def _toggle_auto_scroll(self):
        self.auto_scroll = not self.auto_scroll
        if self.auto_scroll:
            self.scroll_btn.config(text="日志自动滚动：开")
            self.log("📝 日志自动滚动已开启")
            self.log_text.see(tk.END)
        else:
            self.scroll_btn.config(text="日志自动滚动：关")
            self.log("📝 日志自动滚动已关闭")
    
    def _toggle_schedule_widgets(self):
        state = tk.NORMAL if self.enable_schedule_var.get() else tk.DISABLED
        self.schedule_time_entry.config(state=state)
        self.schedule_mode_combobox.config(state=state if self.enable_schedule_var.get() else "readonly")
    
    def _load_schedule_config(self):
        main_config = configparser.ConfigParser()
        main_config.read(MAIN_CONFIG_PATH, encoding="utf-8")
        
        self.enable_schedule_var.set(main_config["GENERAL"].getboolean("enable_schedule"))
        self.schedule_time_var.set(main_config["GENERAL"]["next_start_time"])
        
        mode = main_config["GENERAL"]["schedule_mode"]
        self.schedule_mode_var.set("仅一次" if mode == "once" else "始终")

        click_mode = main_config["GENERAL"]["click_mode"]

        # 根据配置值设置下拉列表选中项（sendmessage → 消息点击，pyautogui → 硬件点击）
        if click_mode == "pyautogui":
            self.click_mode_combobox.current(1)
        else:
            self.click_mode_combobox.current(0)  # 默认选中 SendMessage
            
            self._toggle_schedule_widgets()
    
    def _save_schedule_config(self):
        time_str = self.schedule_time_var.get().strip()
        try:
            datetime.datetime.strptime(time_str, "%H:%M:%S")
        except ValueError:
            messagebox.showerror("错误", "定时时间格式错误！请输入 HH:MM:SS 格式")
            return
        
        main_config = configparser.ConfigParser()
        main_config.read(MAIN_CONFIG_PATH, encoding="utf-8")
        
        main_config["GENERAL"]["enable_schedule"] = str(self.enable_schedule_var.get())
        main_config["GENERAL"]["next_start_time"] = time_str
        mode_text = self.schedule_mode_var.get()
        main_config["GENERAL"]["schedule_mode"] = "once" if mode_text == "仅一次" else "always"
        
        with open(MAIN_CONFIG_PATH, "w", encoding="utf-8") as f:
            main_config.write(f)
        
        self.log("✅ 定时配置已保存！")
    
    def _load_task_groups(self):
        tasks = [f[:-5] for f in os.listdir(TASKS_DIR) if f.endswith(".json")]
        if not tasks:
            init_task_config("default")
            tasks = ["default"]
        
        self.task_combobox["values"] = tasks
        main_config = configparser.ConfigParser()
        main_config.read(MAIN_CONFIG_PATH, encoding="utf-8")
        current = main_config["GENERAL"]["current_task_group"]
        self.task_var.set(current if current in tasks else tasks[0])
    
    def _new_task_group(self):
        """新建任务组 - 自定义大小窗口，命名校验逻辑 + 元素居中显示"""
        # 创建自定义顶级窗口
        new_win = tk.Toplevel(self.root)
        new_win.title("新建任务组")
        
        # ========== 核心：设置窗口大小（可按需调整） ==========
        new_win.geometry("350x180")  # 宽度350px，高度180px
        new_win.minsize(300, 150)    # 可选：设置最小尺寸，防止拖得太小
        new_win.resizable(True, True)  # 修正注释：允许拉伸窗口（原注释错误）
        
        # 可选：让窗口居中显示（优化体验）
        new_win.update_idletasks()
        screen_width = new_win.winfo_screenwidth()
        screen_height = new_win.winfo_screenheight()
        x = (screen_width - 350) // 2
        y = (screen_height - 180) // 2
        new_win.geometry(f"350x180+{x}+{y}")

        # ========== 设置列权重，实现元素居中 ==========
        new_win.columnconfigure(0, weight=1)  # 让第0列自适应窗口宽度，支撑居中

        # ========== 窗口内控件布局（修改为居中显示） ==========
        # 标签（使用sticky="center"实现居中）
        ttk.Label(new_win, text="请输入任务组名称：", font=("Arial", 10)).grid(
            row=0, column=0, 
            padx=20, pady=20, 
            sticky="ew"  # 控件在单元格内水平+垂直居中
        )
        # 输入框
        name_var = tk.StringVar()
        name_entry = ttk.Entry(new_win, textvariable=name_var, width=25, font=("Arial", 10))
        name_entry.grid(row=1, column=0,  # 单列布局
                        padx=20, pady=5, 
                        sticky="ew")  # 输入框居中
        name_entry.focus()  # 自动聚焦输入框

        # ========== 保存逻辑 ==========
        def save_group():
            # 1. 获取并清洗输入内容
            name = name_var.get().strip()
            if not name:  # 检查命名为空
                messagebox.showwarning("提示", "任务组名称不能为空！")
                return
            
            # 2. 检查任务组是否已存在
            config_path = os.path.join(TASKS_DIR, f"{name}.json")
            if os.path.exists(config_path):
                messagebox.showerror("错误", "任务组已存在！")
                return
            
            # 3. 执行新建逻辑
            init_task_config(name)
            self._load_task_groups()
            self.task_var.set(name)
            self.log(f"✅ 新建任务组：{name}")
            
            # 4. 关闭新建窗口
            new_win.destroy()

        # 保存按钮
        save_btn = ttk.Button(new_win, text="保存", command=save_group)
        save_btn.grid(row=2, column=0,  # 单列布局
                    padx=20, pady=15, 
                    sticky="ew")  # 按钮居中

        # 禁止点击父窗口
        new_win.grab_set()
    
    def _edit_main_config(self):
        self._open_file(MAIN_CONFIG_PATH)
    
    def _edit_task_config(self):
        task_name = self.task_var.get()
        self._open_file(os.path.join(TASKS_DIR, f"{task_name}.json"))
    
    def _open_file(self, path):
        try:
            if sys.platform == "win32":
                os.startfile(path)
            elif sys.platform == "darwin":
                subprocess.run(["open", path])
            else:
                subprocess.run(["xdg-open", path])
        except Exception as e:
            messagebox.showerror("错误", f"无法打开文件：{str(e)}")
    
    def _start(self):
        main_config = configparser.ConfigParser()
        main_config.read(MAIN_CONFIG_PATH, encoding="utf-8")
        main_config["GENERAL"]["current_task_group"] = self.task_var.get()
        with open(MAIN_CONFIG_PATH, "w", encoding="utf-8") as f:
            main_config.write(f)
        
        self.stop_flag = False
        self.thread = threading.Thread(target=worker, args=(self,), daemon=True)
        self.thread.start()
        
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.log("🚀 程序已启动！")
    
    # 在AutoClickGUI类中补充/修改_stop方法
    def _stop(self, is_manual=True):
        """停止任务并重置UI状态"""
        self.stop_flag = True
        # 确保线程结束后重置按钮状态
        if self.thread and self.thread.is_alive():
            self.thread.join(timeout=1.0)  # 等待线程结束
        # 强制重置按钮状态
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        # 记录停止日志
        if is_manual:
            self.log("🛑 用户手动停止任务")
        else:
            self.log("🛑 任务因条件触发自动停止")
        
    def _clear_log(self):
        # 临时启用编辑状态以清除日志
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)  # 恢复不可编辑状态
        self.log("📝 日志已清除！")  # 记录清除操作
    
    def _show_about_dialog(self):
        """显示关于对话框"""
        about_window = tk.Toplevel(self.root)
        about_window.title("关于")
        about_window.geometry("500x400")
        about_window.resizable(False, False)
        
        # 居中显示
        window_width = 480
        window_height = 390
        screen_width = about_window.winfo_screenwidth()
        screen_height = about_window.winfo_screenheight()
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        about_window.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # 创建内容框架
        content_frame = ttk.Frame(about_window, padding=20)
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        ttk.Label(content_frame, text="OCR自动点击工具", font=("微软雅黑", 16, "bold")).pack(pady=10)
        
        # 版本信息
        #ttk.Label(content_frame, text="版本: 1.4.1").pack(pady=5)
        
        # 功能描述
        # 创建描述文本
        description = "OCR自动点击工具 - 基于图像识别的自动化操作工具\n\n"
        description += " Copyright (C) 2026 XDista\n\n"
        description += "本程序绝对无任何保证；详情见GNU GPL V3 许可证。\n\n"
        
        # 创建普通文本标签，设置为居中对齐
        ttk.Label(content_frame, text=description, justify=tk.CENTER).pack(pady=10)
        
        # 创建包含GitHub链接的Frame，设置为居中
        github_frame = ttk.Frame(content_frame)
        github_frame.pack(pady=(0, 10), anchor='center')
        
        # 创建"项目GitHub主页："标签，设置为居中
        ttk.Label(github_frame, text="项目GitHub主页：").pack(side='left', anchor='center')
        
        # 创建超链接标签，添加下划线，设置为居中
        link_label = ttk.Label(github_frame, text="https://github.com/XDista/OCR_AutoClick", 
                             foreground="blue", cursor="hand2", font=('TkDefaultFont', 9, 'underline'))
        link_label.pack(side='left', anchor='center')
        
        def open_link(event):
            webbrowser.open_new("https://github.com/XDista/OCR_AutoClick")
        
        link_label.bind("<Button-1>", open_link)
        
        # 创建剩余描述文本，设置为居中
        ttk.Label(content_frame, text="本软件为开源免费项目，如果你以任何付费方式获得此软件，请立即尝试退款！", 
                 justify=tk.CENTER).pack(pady=(5, 10))
        #其实好像不是OCR，而是IR（Image Recognition）
        #算了，问起来就说“OCR是工具的设计目标与追求，实际功能请以成品为准”

        # 关闭按钮
        ttk.Button(content_frame, text="关闭", command=about_window.destroy).pack(pady=10)

    # 修改log方法
    def log(self, msg):
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        # 临时启用编辑状态以写入日志
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"[{ts}] {msg}\n")
        
        if self.auto_scroll:
            self.log_text.see(tk.END)
        
        self.root.update_idletasks()
        # 恢复不可编辑状态
        self.log_text.config(state=tk.DISABLED)

if __name__ == "__main__":
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except:
        pass
    
    root = tk.Tk()
    app = AutoClickGUI(root)
    
    def on_close():
        app._stop()
        root.destroy()
    root.protocol("WM_DELETE_WINDOW", on_close)
    
    root.mainloop()