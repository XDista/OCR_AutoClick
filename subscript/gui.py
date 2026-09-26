import os
import time
import configparser
import threading
import datetime
import webbrowser
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import subprocess
import sys
from pynput import keyboard
from pynput.keyboard import Key
from theme.theme_manager import (
    apply_theme,
    apply_theme_to_log,
    get_log_colors,
    get_theme_names,
    get_theme_name,
    get_theme_id_by_name,
    load_theme_config,
    save_theme_config,
)
from ocr_engine import OCREngine
from ocr_test import OCRTest
from project_paths import BASE_DIR, MAIN_CONFIG_PATH, TASKS_DIR
from config_manager import init_main_config, init_task_config
from window_utils import (
    get_all_visible_windows_simple,
    get_target_window_from_config,
)
from screenshot_capture import (
    capture_window,
    get_adb_executable_path,
    validate_adb_environment,
)
from task_worker import worker

class AutoClickGUI:
    def __init__(self, root, version="1.0.0"):
        self.root = root
        self.root.title("自动识别点击工具")
        self.version = version

        config = init_main_config()
        selfgeometry = config["WindowConfig"]["selfgeometry"]
        selfxy_resizable = int(config["WindowConfig"]["selfxy_resizable"])
        self.root.geometry(selfgeometry)
        self.root.resizable(selfxy_resizable, selfxy_resizable)

        self.ocr_lang = config["OCRConfig"].get("ocr_language", "ch_sim, en")
        self.ocr_device = config["OCRConfig"].get("ocr_device", "cpu")
        self.ocr_conf = float(config["OCRConfig"].get("ocr_confidence", "0.7"))
        self.ocr_gray = config["OCRConfig"].get("ocr_grayscale", "True") == "True"
        self.ocr_thresh = config["OCRConfig"].get("ocr_threshold", "False") == "True"

        saved_theme = load_theme_config(MAIN_CONFIG_PATH)
        if saved_theme and saved_theme in ("light", "dark"):
            apply_theme(root, saved_theme)
        else:
            apply_theme(root, "light")

        self.thread = None
        self.auto_scroll = True
        self.window_list = []
        self.SCREENXY_PATH = os.path.join(BASE_DIR, "subscript", "screenxy.pyw")
        self.adb_screenshot_path = os.path.join(BASE_DIR, "subscript", "adb_screenshot.pyw")
        self.SYNC_FILE_PATH = os.path.join(BASE_DIR, "subscript", "sync_file.py")

        self.stop_event = threading.Event()
        self.worker_generation = 0
        self.stop_flag = False

        self.ocr_engine = OCREngine()
        self.ocr_engine.configure(device=self.ocr_device, language=self.ocr_lang)
        self.ocr_task_dir = os.path.join(BASE_DIR, "task_ocr")
        self._ocr_test_instance = None

        self.hotkey_keys = {Key.f9, Key.f10}
        self.pressed_keys = set()
        self.key_listener = None
        self._init_hotkey_listener()

        self._create_widgets()
        self._refresh_ocr_device_list()
        self._load_task_groups()
        self._load_schedule_config()
        self._load_window_combobox()
        self._load_default_window_from_config()

    def _init_hotkey_listener(self):
        def on_key_press(key):
            try:
                self.pressed_keys.add(key)
                if self.hotkey_keys.issubset(self.pressed_keys):
                    self._handle_hotkey_trigger()
            except Exception as e:
                self.log(f"热键按下异常：{str(e)}")

        def on_key_release(key):
            try:
                if key in self.pressed_keys:
                    self.pressed_keys.remove(key)
            except Exception as e:
                self.log(f"热键释放异常：{str(e)}")

        self.key_listener = keyboard.Listener(on_press=on_key_press, on_release=on_key_release)
        self.key_listener.daemon = True
        self.key_listener.start()

    def _handle_hotkey_trigger(self):
        if hasattr(self, "_hotkey_locked") and self._hotkey_locked:
            return
        self._hotkey_locked = True
        self.root.after(500, lambda: setattr(self, "_hotkey_locked", False))

        if self.thread and self.thread.is_alive():
            self._stop(is_manual=True)
            self.log("🛑 热键(F9+F10)触发：停止任务组")
        else:
            self._start()
            self.log("▶️ 热键(F9+F10)触发：启动任务组")

    def _on_window_close(self):
        if self.key_listener and self.key_listener.is_alive():
            self.key_listener.stop()
        if self.thread and self.thread.is_alive():
            self.stop_flag = True
            self.thread.join(timeout=2)
        self.root.destroy()

    def on_screenshot_mode_change(self, event=None):
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

    def _on_tab_changed(self, event):
        for combobox in [
            getattr(self, 'window_combobox', None),
            getattr(self, 'task_combobox', None),
            getattr(self, 'click_mode_combobox', None),
            getattr(self, 'schedule_mode_combobox', None),
            getattr(self, 'theme_combobox', None),
        ]:
            if combobox:
                combobox.selection_clear()
        self.root.focus_set()

    def on_adb_device_serial_change(self, event=None):
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
        main_config = configparser.ConfigParser()
        main_config.read(MAIN_CONFIG_PATH, encoding="utf-8")

        selected_text = self.click_mode_var.get()
        click_mode = "sendmessage" if "SendMessage" in selected_text else "pyautogui"

        main_config["GENERAL"]["click_mode"] = click_mode
        with open(MAIN_CONFIG_PATH, "w", encoding="utf-8") as f:
            main_config.write(f)

        self.log(f"✅ 点击模式已切换为：{selected_text}")

        self.click_mode_combobox.selection_clear()
        self.root.focus_set()

    def _on_task_group_change(self, event):
        try:
            selected_task_group = self.task_var.get().strip()
            if not selected_task_group:
                self.log("⚠️ 任务组名称不能为空，跳过配置更新")
                return

            main_config = configparser.ConfigParser()
            main_config.read(MAIN_CONFIG_PATH, encoding="utf-8")
            main_config["GENERAL"]["current_task_group"] = selected_task_group

            with open(MAIN_CONFIG_PATH, "w", encoding="utf-8") as f:
                main_config.write(f)

            self.log(f"✅ 任务组已切换为：{selected_task_group}")
        except Exception as e:
            self.log(f"⚠️ 保存任务组配置失败：{str(e)}")
            messagebox.showerror("错误", f"保存任务组配置失败：{e}")
        finally:
            self.task_combobox.selection_clear()
            self.root.focus_set()

    def _on_window_change(self, event):
        try:
            idx = self.window_combobox.current()
            if idx < 0 or idx >= len(self.window_list):
                self.log("⚠️ 未选中有效窗口，跳过配置更新")
                return
            win = self.window_list[idx]

            config = configparser.ConfigParser()
            config.read(MAIN_CONFIG_PATH, encoding="utf-8")
            config["GENERAL"]["target_window_hwnd"] = str(win["hwnd"])
            config["GENERAL"]["target_window_title"] = win["title"]
            config["GENERAL"]["target_program_name"] = win["process_name"]

            with open(MAIN_CONFIG_PATH, "w", encoding="utf-8") as f:
                config.write(f)

            self._update_window_info(idx)
            self.log(f"✅ 窗口已切换为：{win['display']}，配置已实时更新")
        except Exception as e:
            self.log(f"⚠️ 保存窗口配置失败：{str(e)}")
            messagebox.showerror("错误", f"保存窗口配置失败：{e}")
        finally:
            self.window_combobox.selection_clear()
            self.root.focus_set()

    def _run_adb_command(self):
        full_cmd_input = self.adb_cmd_var.get().strip()
        if not full_cmd_input:
            self.log("⚠️ ADB命令不能为空！请输入完整命令（如'adb devices'、'adb shell'）")
            return

        if not full_cmd_input.startswith("adb"):
            self.log(f"❌ 命令格式错误！必须以'adb'开头（当前输入：{full_cmd_input}）")
            return
        if len(full_cmd_input) > 3 and not full_cmd_input[3].isspace():
            self.log(f"❌ 命令格式错误！'adb'后必须跟空格（当前输入：{full_cmd_input}，正确示例：'adb devices'）")
            return

        adb_valid, adb_path_or_msg = get_adb_executable_path()
        if not adb_valid:
            self.log(f"❌ ADB环境不可用：{adb_path_or_msg}")
            return

        main_config = configparser.ConfigParser()
        main_config.read(MAIN_CONFIG_PATH, encoding="utf-8")
        adb_usage_mode = main_config["ADBConfig"].get("adb_usage_mode", "1")

        cmd_params = full_cmd_input[3:].lstrip()

        if adb_usage_mode == "3":
            final_cmd = full_cmd_input
        else:
            if cmd_params:
                final_cmd = f'"{adb_path_or_msg}" {cmd_params}'
            else:
                final_cmd = f'"{adb_path_or_msg}"'

        self.log(f"▶️ 正在执行ADB命令：{final_cmd}")
        try:
            result = subprocess.run(
                final_cmd,
                shell=True,
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                encoding="utf-8",
                timeout=30
            )
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
        main_notebook = ttk.Notebook(self.root)

        main_notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        main_notebook.enable_traversal()
        main_notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        window_frame = ttk.Frame(main_notebook)
        main_notebook.add(window_frame, text="目标窗口")

        select_frame = ttk.LabelFrame(window_frame, text="窗口选择", padding="5")
        select_frame.pack(fill=tk.X, padx=10, pady=2)

        ttk.Label(select_frame, text="选择目标窗口：").grid(row=0, column=0, padx=5, pady=3, sticky=tk.W)

        self.window_var = tk.StringVar()
        self.window_combobox = ttk.Combobox(
            select_frame,
            textvariable=self.window_var,
            state="readonly",
            width=60
        )
        self.window_combobox.bind("<<ComboboxSelected>>", self._on_window_change)
        self.window_combobox.bind("<FocusOut>", lambda e: self.window_combobox.selection_clear())
        self.window_combobox.grid(row=0, column=1, padx=5, pady=3, sticky=tk.W)

        ttk.Button(select_frame, text="刷新列表", command=self._load_window_combobox).grid(row=0, column=2, padx=5, pady=3)
        ttk.Button(select_frame, text="模糊匹配配置", command=self._edit_window_match_config).grid(row=0, column=3, padx=5, pady=3)

        info_frame = ttk.LabelFrame(window_frame, text="窗口信息", padding="5")
        info_frame.pack(fill=tk.X, padx=10, pady=2)

        self.window_info_var = tk.StringVar(value="未选择窗口 | 标题：- | 进程：- | 句柄：-")
        ttk.Label(info_frame, textvariable=self.window_info_var).pack(side=tk.LEFT, padx=5, pady=3)

        match_step_frame = ttk.LabelFrame(window_frame, text="模板匹配配置", padding="5")
        match_step_frame.pack(fill=tk.X, padx=10, pady=2)

        self.match_step_var = tk.StringVar()
        self._load_match_step_from_config()

        ttk.Label(match_step_frame, text="匹配缩略图步长：").grid(row=0, column=0, padx=5, pady=3, sticky=tk.W)
        step_entry = ttk.Entry(
            match_step_frame,
            textvariable=self.match_step_var,
            width=20
        )
        step_entry.grid(row=0, column=1, padx=5, pady=3, sticky=tk.W)

        ttk.Button(
            match_step_frame,
            text="保存步长",
            command=self._save_match_step_to_config
        ).grid(row=0, column=2, padx=5, pady=3)

        ttk.Label(
            match_step_frame,
            text="（步长范围：0.001~0.2，默认0.05）",
            foreground="#666666"
        ).grid(row=0, column=3, padx=5, pady=3, sticky=tk.W)

        placeholder_frame = ttk.LabelFrame(window_frame, text="功能", padding="5")
        placeholder_frame.pack(fill=tk.X, padx=10, pady=2)

        restart_btn = ttk.Button(
            placeholder_frame,
            text="重启脚本",
            command=self._restart_script
        )
        restart_btn.pack(side=tk.LEFT, padx=5, pady=3)

        about_btn = ttk.Button(
            placeholder_frame,
            text="关于",
            command=self._show_about_dialog
        )
        about_btn.pack(side=tk.RIGHT, padx=5, pady=3)

        task_frame = ttk.Frame(main_notebook)
        main_notebook.add(task_frame, text="任务配置")

        task_group_frame = ttk.LabelFrame(task_frame, text="任务组", padding="5")
        task_group_frame.pack(fill=tk.X, padx=10, pady=2)

        ttk.Label(task_group_frame, text="选择任务组：").pack(side=tk.LEFT, padx=5, pady=3)
        self.task_var = tk.StringVar()
        self.task_combobox = ttk.Combobox(task_group_frame, textvariable=self.task_var, state="readonly")
        self.task_combobox.pack(side=tk.LEFT, padx=5, pady=3)
        self.task_combobox.bind("<<ComboboxSelected>>", self._on_task_group_change)
        self.task_combobox.bind("<FocusOut>", lambda e: self.task_combobox.selection_clear())
        ttk.Button(task_group_frame, text="刷新", command=self._load_task_groups).pack(side=tk.LEFT, padx=5, pady=3)
        ttk.Button(task_group_frame, text="新建", command=self._new_task_group).pack(side=tk.LEFT, padx=5, pady=3)
        ttk.Button(task_group_frame, text="编辑", command=self._edit_task_config).pack(side=tk.LEFT, padx=5, pady=3)
        ttk.Button(task_group_frame, text="坐标拾取", command=self._open_screenxy).pack(side=tk.LEFT, padx=5, pady=3)
        ttk.Button(task_group_frame, text="ADB截图工具", command=self._open_adb_screenshot).pack(side=tk.LEFT, padx=5, pady=3)

        schedule_frame = ttk.LabelFrame(task_frame, text="定时运行配置", padding="5")
        schedule_frame.pack(fill=tk.X, padx=10, pady=2)

        schedule_inner = ttk.Frame(schedule_frame)
        schedule_inner.pack(fill=tk.X, padx=5, pady=3)

        self.enable_schedule_var = tk.BooleanVar()
        schedule_check = ttk.Checkbutton(
            schedule_inner,
            text="启用定时运行",
            variable=self.enable_schedule_var,
            command=self._toggle_schedule_widgets
        )
        schedule_check.pack(side=tk.LEFT, padx=5, pady=3)

        ttk.Label(schedule_inner, text="定时时间：").pack(side=tk.LEFT, padx=5, pady=3)
        self.schedule_time_var = tk.StringVar(value="00:00:00")
        self.schedule_time_entry = ttk.Entry(
            schedule_inner,
            textvariable=self.schedule_time_var,
            width=12
        )
        self.schedule_time_entry.pack(side=tk.LEFT, padx=5, pady=3)
        ttk.Label(schedule_inner, text="（HH:MM:SS）").pack(side=tk.LEFT, padx=5, pady=3)

        ttk.Label(schedule_inner, text="运行模式：").pack(side=tk.LEFT, padx=5, pady=3)
        self.schedule_mode_var = tk.StringVar()
        self.schedule_mode_combobox = ttk.Combobox(
            schedule_inner,
            textvariable=self.schedule_mode_var,
            values=["仅一次", "始终"],
            state="readonly",
            width=8
        )
        self.schedule_mode_combobox.pack(side=tk.LEFT, padx=5, pady=3)
        self.schedule_mode_combobox.bind("<FocusOut>", lambda e: self.schedule_mode_combobox.selection_clear())
        self.schedule_mode_combobox.current(0)

        ttk.Button(
            schedule_inner,
            text="保存配置",
            command=self._save_schedule_config
        ).pack(side=tk.LEFT, padx=10, pady=3)

        config_frame = ttk.LabelFrame(task_frame, text="点击配置", padding="5")
        config_frame.pack(fill=tk.X, padx=10, pady=2)

        ttk.Label(config_frame, text="点击模式：").grid(row=0, column=0, padx=5, pady=3, sticky="w")
        self.click_mode_var = tk.StringVar()
        self.click_mode_combobox = ttk.Combobox(
            config_frame,
            textvariable=self.click_mode_var,
            values=["SendMessage消息点击", "PyAutoGUI硬件点击"],
            state="readonly",
            width=35
        )
        self.click_mode_combobox.grid(row=0, column=1, padx=5, pady=3, sticky=tk.W)
        self.click_mode_combobox.bind("<<ComboboxSelected>>", self._on_click_mode_change)
        self.click_mode_combobox.bind("<FocusOut>", lambda e: self.click_mode_combobox.selection_clear())

        main_config = configparser.ConfigParser()
        main_config.read(MAIN_CONFIG_PATH, encoding="utf-8")

        ttk.Label(config_frame, text="截图模式：").grid(row=1, column=0, padx=5, pady=2, sticky="w")
        self.screenshot_mode_var = tk.StringVar(value=main_config["GENERAL"]["screenshot_mode"])
        screenshot_mode_combo = ttk.Combobox(config_frame, textvariable=self.screenshot_mode_var,
            values=["Win32GUI", "PrintWindow", "Win32Memory", "ADB"], state="readonly", width=35)
        screenshot_mode_combo.grid(row=1, column=1, padx=5, pady=3, sticky="w")
        screenshot_mode_combo.bind("<<ComboboxSelected>>", self.on_screenshot_mode_change)
        screenshot_mode_combo.bind("<FocusOut>", lambda e: screenshot_mode_combo.selection_clear())

        ttk.Label(config_frame, text="ADB设备Serial：").grid(row=2, column=0, padx=5, pady=3, sticky="w")
        self.adb_device_serial_var = tk.StringVar(value=main_config["ADBConfig"]["adb_device_serial"])
        adb_device_serial_entry = ttk.Entry(config_frame, textvariable=self.adb_device_serial_var, width=38)
        adb_device_serial_entry.grid(row=2, column=1, padx=5, pady=3, sticky="w")
        adb_device_serial_entry.bind("<FocusOut>", self.on_adb_device_serial_change)

        def comfort_confirm():
            self.log("✅ 设备Serial已更新")

        comfort_btn = ttk.Button(config_frame, text="确认", command=comfort_confirm, width=6)
        comfort_btn.grid(row=2, column=2, padx=(0, 5), pady=3, sticky="w")

        adb_frame = ttk.Frame(main_notebook)
        main_notebook.add(adb_frame, text="ADB配置")

        main_config = configparser.ConfigParser()
        main_config.read(MAIN_CONFIG_PATH, encoding="utf-8")
        if "ADBConfig" not in main_config:
            main_config["ADBConfig"] = {}

        adb_enabled = main_config["ADBConfig"].get("adb_enabled", "1")
        adb_usage_mode = main_config["ADBConfig"].get("adb_usage_mode", "1")
        adb_position = main_config["ADBConfig"].get("adb_position", "")

        enabled_map = {"是": "1", "否": "0"}
        mode_map = {"内置": "1", "自定义": "2", "系统环境变量": "3", "Alas附带": "4"}
        enabled_rev_map = {v: k for k, v in enabled_map.items()}
        mode_rev_map = {v: k for k, v in mode_map.items()}

        style = ttk.Style()
        style.configure("Gray.TEntry", foreground="#888888")
        style.configure("Normal.TEntry", foreground="#000000")
        style.configure("Gray.TLabel", foreground="#888888")
        style.configure("Normal.TLabel", foreground="#000000")

        adb_basic_frame = ttk.LabelFrame(adb_frame, text="ADB基本设置", padding="5")
        adb_basic_frame.pack(fill=tk.X, padx=10, pady=2)

        current_enabled_text = "是" if adb_enabled == "1" else "否"
        adb_enabled_var = tk.StringVar(value=current_enabled_text)

        ttk.Label(adb_basic_frame, text="ADB启用控制：").grid(
            row=0, column=0, padx=5, pady=3, sticky="w"
        )
        adb_enabled_combobox = ttk.Combobox(
            adb_basic_frame, textvariable=adb_enabled_var, values=["是", "否"], width=12, state="readonly"
        )
        adb_enabled_combobox.grid(row=0, column=1, padx=2, pady=3, sticky="w")
        adb_enabled_combobox.bind("<FocusOut>", lambda e: adb_enabled_combobox.selection_clear())

        current_mode_text = {v: k for k, v in mode_map.items()}.get(adb_usage_mode, "内置")
        adb_mode_var = tk.StringVar(value=current_mode_text)

        ttk.Label(adb_basic_frame, text="ADB使用模式：").grid(
            row=0, column=2, padx=10, pady=3, sticky="w"
        )
        adb_mode_combobox = ttk.Combobox(
            adb_basic_frame, textvariable=adb_mode_var, values=["内置", "自定义", "系统环境变量", "Alas附带"], width=12, state="readonly"
        )
        adb_mode_combobox.grid(row=0, column=3, padx=2, pady=3, sticky="w")
        adb_mode_combobox.bind("<FocusOut>", lambda e: adb_mode_combobox.selection_clear())

        adb_path_frame = ttk.LabelFrame(adb_frame, text="自定义ADB路径", padding="5")
        adb_path_frame.pack(fill=tk.X, padx=10, pady=2)

        adb_path_label = ttk.Label(adb_path_frame, text="路径：")
        adb_path_label.grid(row=0, column=0, padx=5, pady=3, sticky="w")

        adb_position_var = tk.StringVar(value=adb_position)
        adb_position_entry = ttk.Entry(adb_path_frame, textvariable=adb_position_var, width=35)
        adb_position_entry.grid(row=0, column=1, padx=2, pady=3, sticky="w")

        def select_adb_path():
            adb_file = filedialog.askopenfilename(
                title="选择ADB可执行文件",
                initialdir=os.path.dirname(adb_position) if adb_position else BASE_DIR,
                filetypes=[("ADB执行文件", "adb.exe"), ("所有文件", "*.*")]
            )
            if adb_file:
                adb_position_var.set(adb_file)
        adb_browse_btn = ttk.Button(adb_path_frame, text="浏览", command=select_adb_path, width=6)
        adb_browse_btn.grid(row=0, column=2, padx=5, pady=3, sticky="w")

        def validate_adb_env():
            is_valid, msg = validate_adb_environment()
            if is_valid:
                self.log(f"ADB验证成功：{msg}")
            else:
                self.log(f"ADB验证失败：{msg}")

        ttk.Button(
            adb_path_frame, text="验证ADB环境", command=validate_adb_env, width=12
        ).grid(row=0, column=3, padx=5, pady=3, sticky="w")

        tip_label = ttk.Label(
            adb_path_frame,
            text="仅在\"是+自定义\"模式下生效",
            foreground="#666666",
            font=(None, 10)
        )
        tip_label.grid(row=1, column=0, columnspan=4, padx=5, pady=2, sticky="w")

        adb_cmd_frame = ttk.LabelFrame(adb_frame, text="ADB命令执行", padding="5")
        adb_cmd_frame.pack(fill=tk.X, padx=10, pady=2)

        ttk.Label(adb_cmd_frame, text="命令：").grid(
            row=0, column=0, padx=5, pady=3, sticky="w"
        )
        self.adb_cmd_var = tk.StringVar(value="")
        adb_cmd_entry = ttk.Entry(
            adb_cmd_frame, textvariable=self.adb_cmd_var, width=45
        )
        adb_cmd_entry.grid(row=0, column=1, padx=2, pady=3, sticky="w")

        def execute_adb_cmd():
            self._run_adb_command()
        ttk.Button(
            adb_cmd_frame, text="执行命令", command=execute_adb_cmd, width=10
        ).grid(row=0, column=2, padx=5, pady=8, sticky="w")

        adb_cmd_entry.bind("<Return>", lambda event: self._run_adb_command())

        cmd_tip_label = ttk.Label(
            adb_cmd_frame,
            text="输入完整ADB命令（需带'adb '前缀，如'adb devices'、'adb shell getprop'）",
            foreground="#666666"
        )
        cmd_tip_label.grid(row=1, column=0, columnspan=3, padx=10, pady=5, sticky="w")

        def update_adb_config(*args):
            old_config = configparser.ConfigParser()
            old_config.read(MAIN_CONFIG_PATH, encoding="utf-8")
            if "ADBConfig" not in old_config:
                old_config["ADBConfig"] = {}

            old_enabled = old_config["ADBConfig"].get("adb_enabled", "1")
            old_mode = old_config["ADBConfig"].get("adb_usage_mode", "1")
            old_position = old_config["ADBConfig"].get("adb_position", "")

            new_enabled_text = adb_enabled_var.get()
            new_mode_text = adb_mode_var.get()
            new_position = adb_position_var.get()

            new_enabled = enabled_map[new_enabled_text]
            new_mode = mode_map[new_mode_text]

            changed_items = []
            if old_enabled != new_enabled:
                old_enabled_text = "是" if old_enabled == "1" else "否"
                changed_items.append(f"ADB启用状态：{old_enabled_text} → {new_enabled_text}")
            if old_mode != new_mode:
                old_mode_text = {v: k for k, v in mode_map.items()}.get(old_mode, "内置")
                changed_items.append(f"ADB使用模式：{old_mode_text} → {new_mode_text}")
            if old_position != new_position:
                old_pos_show = old_position if old_position else "空"
                new_pos_show = new_position if new_position else "空"
                changed_items.append(f"自定义ADB路径：{old_pos_show} → {new_pos_show}")

            if not changed_items:
                return

            try:
                main_config = configparser.ConfigParser()
                main_config.read(MAIN_CONFIG_PATH, encoding="utf-8")
                if "ADBConfig" not in main_config:
                    main_config["ADBConfig"] = {}

                main_config["ADBConfig"]["adb_enabled"] = new_enabled
                main_config["ADBConfig"]["adb_usage_mode"] = new_mode
                main_config["ADBConfig"]["adb_position"] = new_position

                with open(MAIN_CONFIG_PATH, "w", encoding="utf-8") as f:
                    main_config.write(f)

                changed_detail = "；".join(changed_items)
                self.log(f"ADB配置更新：{changed_detail}")

            except Exception as e:
                self.log(f"ADB配置更新失败：{str(e)}")

        adb_enabled_var.trace_add("write", update_adb_config)
        adb_mode_var.trace_add("write", update_adb_config)
        adb_position_var.trace_add("write", update_adb_config)

        ocr_frame = ttk.Frame(main_notebook)
        main_notebook.add(ocr_frame, text="OCR")

        ocr_task_frame = ttk.LabelFrame(ocr_frame, text="OCR 任务组", padding="10")
        ocr_task_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(ocr_task_frame, text="选择任务组：").grid(row=0, column=0, padx=5, pady=3, sticky="w")
        self.ocr_task_var = tk.StringVar()
        self.ocr_task_combo = ttk.Combobox(
            ocr_task_frame,
            textvariable=self.ocr_task_var,
            state="readonly",
            width=25
        )
        self.ocr_task_combo.grid(row=0, column=1, padx=2, pady=3, sticky="w")
        self.ocr_task_combo.bind("<FocusOut>", lambda e: self.ocr_task_combo.selection_clear())

        def _refresh_ocr_tasks():
            tasks = self.ocr_engine.list_ocr_tasks(self.ocr_task_dir)
            self.ocr_task_combo["values"] = tasks
            if tasks:
                main_config = configparser.ConfigParser()
                main_config.read(MAIN_CONFIG_PATH, encoding="utf-8")
                current = main_config["OCRConfig"].get("current_ocr_task", "")
                if current in tasks:
                    self.ocr_task_var.set(current)
                elif not self.ocr_task_var.get() or self.ocr_task_var.get() not in tasks:
                    self.ocr_task_var.set(tasks[0])
            else:
                self.ocr_task_var.set("")

        def _on_ocr_task_change(*args):
            selected = self.ocr_task_var.get()
            if not selected:
                return
            try:
                main_config = configparser.ConfigParser()
                main_config.read(MAIN_CONFIG_PATH, encoding="utf-8")
                if "OCRConfig" not in main_config:
                    main_config["OCRConfig"] = {}
                if main_config["OCRConfig"].get("current_ocr_task") != selected:
                    main_config["OCRConfig"]["current_ocr_task"] = selected
                    with open(MAIN_CONFIG_PATH, "w", encoding="utf-8") as f:
                        main_config.write(f)
                    self.log(f"✅ OCR任务组切换为：{selected}")
                    try:
                        task_config = self.ocr_engine.load_ocr_task(selected, self.ocr_task_dir)
                        self.ocr_engine.apply_ocr_settings(task_config)
                        self.log(f"  已应用任务中的 OCR 设置")
                    except Exception as e:
                        self.log(f"  ⚠️ 应用 OCR 设置失败：{e}")
            except Exception as e:
                self.log(f"OCR任务组保存失败：{e}")

        self.ocr_task_var.trace_add("write", _on_ocr_task_change)
        ttk.Button(ocr_task_frame, text="刷新", command=_refresh_ocr_tasks, width=5).grid(
            row=0, column=2, padx=2, pady=3, sticky="w")
        ttk.Button(ocr_task_frame, text="编辑", command=self._edit_ocr_task, width=5).grid(
            row=0, column=3, padx=2, pady=3, sticky="w")
        ttk.Button(ocr_task_frame, text="检测GPU", command=self._detect_gpu, width=7).grid(
            row=0, column=4, padx=2, pady=3, sticky="w")

        ocr_param_frame = ttk.LabelFrame(ocr_frame, text="识别参数", padding="10")
        ocr_param_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(ocr_param_frame, text="识别语言：").grid(row=0, column=0, padx=5, pady=3, sticky="w")
        self.ocr_lang_var = tk.StringVar(value=self.ocr_lang)
        ocr_lang_combo = ttk.Combobox(
            ocr_param_frame,
            textvariable=self.ocr_lang_var,
            values=["ch_sim, en", "ch_tra, en", "en", "ja, en", "ko, en",
                    "ch_sim", "ch_tra", "ja", "ko", "fr", "de", "es", "th", "vi"],
            state="readonly",
            width=35
        )
        ocr_lang_combo.grid(row=0, column=1, padx=2, pady=3, sticky="w")
        ocr_lang_combo.bind("<FocusOut>", lambda e: ocr_lang_combo.selection_clear())

        def _on_ocr_lang_change(*args):
            lang = self.ocr_lang_var.get()
            self.ocr_engine.configure(language=lang)
            self._save_ocr_config_key("ocr_language", lang)
            self.log(f"✅ 识别语言已切换为：{lang}")
        self.ocr_lang_var.trace_add("write", _on_ocr_lang_change)

        ttk.Label(ocr_param_frame, text="置信度阈值：").grid(row=0, column=2, padx=(15, 2), pady=3, sticky="w")
        self.ocr_conf_var = tk.DoubleVar(value=self.ocr_conf)
        ocr_conf_scale = ttk.Scale(
            ocr_param_frame,
            from_=0.3, to=1.0,
            variable=self.ocr_conf_var,
            orient=tk.HORIZONTAL,
            length=100
        )
        ocr_conf_scale.grid(row=0, column=3, padx=2, pady=3, sticky="w")
        self.ocr_conf_label = ttk.Label(ocr_param_frame, text=f"{self.ocr_conf:.2f}", width=4)
        self.ocr_conf_label.grid(row=0, column=4, padx=2, pady=3, sticky="w")

        def _on_ocr_conf_change(*args):
            val = round(self.ocr_conf_var.get(), 2)
            self.ocr_conf_label.config(text=f"{val:.2f}")
            self.ocr_engine.configure(confidence_threshold=val)
            self._save_ocr_config_key("ocr_confidence", str(val))
            self.log(f"✅ 置信度阈值已更新为：{val:.2f}")
        self.ocr_conf_var.trace_add("write", _on_ocr_conf_change)

        ttk.Label(ocr_param_frame, text="加速设备：").grid(row=1, column=0, padx=5, pady=3, sticky="w")
        self.ocr_device_var = tk.StringVar(value=self.ocr_device)
        self.ocr_device_combo = ttk.Combobox(
            ocr_param_frame,
            textvariable=self.ocr_device_var,
            values=[self.ocr_device],
            state="readonly",
            width=35
        )
        self.ocr_device_combo.grid(row=1, column=1, padx=2, pady=3, sticky="w")
        self.ocr_device_combo.bind("<FocusOut>", lambda e: self.ocr_device_combo.selection_clear())

        def _on_ocr_device_change(*args):
            display_name = self.ocr_device_var.get()
            device = getattr(self, '_device_name_map', {}).get(display_name, display_name)
            self.ocr_engine.configure(device=device)
            self._save_ocr_config_key("ocr_device", device)
            self.log(f"✅ 加速设备已切换为：{display_name}")
        self.ocr_device_var.trace_add("write", _on_ocr_device_change)

        ocr_pre_frame = ttk.LabelFrame(ocr_frame, text="图像预处理", padding="10")
        ocr_pre_frame.pack(fill=tk.X, padx=10, pady=5)

        self.ocr_gray_var = tk.BooleanVar(value=self.ocr_gray)
        def _on_gray_change():
            enabled = self.ocr_gray_var.get()
            self.ocr_engine.configure(preprocess={"grayscale": enabled})
            self._save_ocr_config_key("ocr_grayscale", str(enabled))
            self.log(f"✅ 灰度化已{'开启' if enabled else '关闭'}")
        ttk.Checkbutton(ocr_pre_frame, text="灰度化", variable=self.ocr_gray_var,
                        command=_on_gray_change
                        ).grid(row=0, column=0, padx=5, pady=3, sticky="w")

        self.ocr_thresh_var = tk.BooleanVar(value=self.ocr_thresh)
        def _on_thresh_change():
            enabled = self.ocr_thresh_var.get()
            self.ocr_engine.configure(preprocess={"threshold_enabled": enabled})
            self._save_ocr_config_key("ocr_threshold", str(enabled))
            self.log(f"✅ 二值化已{'开启' if enabled else '关闭'}")
        ttk.Checkbutton(ocr_pre_frame, text="二值化", variable=self.ocr_thresh_var,
                        command=_on_thresh_change
                        ).grid(row=0, column=1, padx=5, pady=3, sticky="w")

        self.ocr_test_btn = ttk.Button(ocr_pre_frame, text="OCR 测试",
                                       command=self._ocr_test, width=8)
        self.ocr_test_btn.grid(row=0, column=2, padx=(20, 2), pady=3, sticky="w")

        self.ocr_stop_btn = ttk.Button(ocr_pre_frame, text="停止",
                                       command=self._stop_ocr_test, width=5,
                                       state=tk.DISABLED)
        self.ocr_stop_btn.grid(row=0, column=3, padx=2, pady=3, sticky="w")

        _refresh_ocr_tasks()

        other_frame = ttk.Frame(main_notebook)
        main_notebook.add(other_frame, text="其他")

        theme_frame = ttk.LabelFrame(other_frame, text="主题配置", padding="10")
        theme_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(theme_frame, text="选择主题：").grid(row=0, column=0, padx=5, pady=3, sticky="w")
        self.theme_var = tk.StringVar()
        theme_display_names = get_theme_names()
        self.theme_combobox = ttk.Combobox(
            theme_frame,
            textvariable=self.theme_var,
            values=theme_display_names,
            state="readonly",
            width=25
        )
        self.theme_combobox.grid(row=0, column=1, padx=5, pady=3, sticky="w")
        self.theme_combobox.bind("<FocusOut>", lambda e: self.theme_combobox.selection_clear())

        saved_theme_id = load_theme_config(MAIN_CONFIG_PATH)
        if saved_theme_id:
            self.theme_var.set(get_theme_name(saved_theme_id))
        else:
            self.theme_var.set(get_theme_name("light"))

        def apply_selected_theme():
            selected_display = self.theme_var.get()
            selected_id = get_theme_id_by_name(selected_display)
            if selected_id:
                apply_theme(self.root, selected_id)
                save_theme_config(selected_id, MAIN_CONFIG_PATH)
                apply_theme_to_log(self.log_text, selected_id)
                self.log(f"✅ 主题已切换为：{selected_display}")
            else:
                self.log(f"⚠️ 主题切换失败：{selected_display}")

        ttk.Button(theme_frame, text="应用主题", command=apply_selected_theme).grid(row=0, column=2, padx=5, pady=3, sticky="w")
        ttk.Label(theme_frame, text="（切换后自动保存，重启后生效）", foreground="#666666").grid(row=0, column=3, padx=5, pady=3, sticky="w")

        empty_frame = ttk.LabelFrame(other_frame, text="Alas配置", padding="10")
        empty_frame.pack(fill=tk.X, padx=10, pady=5)

        if "GENERAL" not in main_config:
            main_config["GENERAL"] = {}
        where_alas = main_config["GENERAL"].get("where_alas", "")
        where_alas_var = tk.StringVar(value=where_alas)

        ttk.Label(empty_frame, text="Alas路径：").grid(row=0, column=0, padx=5, pady=3, sticky="w")
        where_alas_entry = ttk.Entry(empty_frame, textvariable=where_alas_var, width=35)
        where_alas_entry.grid(row=0, column=1, padx=2, pady=3, sticky="w")

        def select_alas_path():
            alas_file = filedialog.askopenfilename(
                title="选择Alas可执行文件",
                initialdir=os.path.dirname(where_alas) if where_alas else BASE_DIR,
                filetypes=[("可执行文件", "*.exe"), ("所有文件", "*.*")]
            )
            if alas_file:
                where_alas_var.set(alas_file)
        ttk.Button(empty_frame, text="浏览", command=select_alas_path, width=6).grid(row=0, column=2, padx=5, pady=3, sticky="w")
        ttk.Button(empty_frame, text="同步配置", command=self._open_sync_file, width=8).grid(row=0, column=3, padx=5, pady=3, sticky="w")

        def update_alas_config(*args):
            new_alas_path = where_alas_var.get()
            if new_alas_path != where_alas:
                try:
                    main_config = configparser.ConfigParser()
                    main_config.read(MAIN_CONFIG_PATH, encoding="utf-8")
                    if "GENERAL" not in main_config:
                        main_config["GENERAL"] = {}
                    main_config["GENERAL"]["where_alas"] = new_alas_path
                    with open(MAIN_CONFIG_PATH, "w", encoding="utf-8") as f:
                        main_config.write(f)
                    self.log(f"Alas路径更新：{where_alas if where_alas else '空'} → {new_alas_path if new_alas_path else '空'}")
                except Exception as e:
                    self.log(f"Alas路径更新失败：{str(e)}")

        where_alas_var.trace_add("write", update_alas_config)

        placeholder_frame2 = ttk.LabelFrame(other_frame, text="预留区域", padding="10")
        placeholder_frame2.pack(fill=tk.X, padx=10, pady=5)

        placeholder_label = ttk.Label(
            placeholder_frame2,
            text="前面的区域，以后再来探索吧！",
            foreground="#666666"
        )
        placeholder_label.pack(padx=5, pady=5)

        control_frame = ttk.Frame(self.root)
        control_frame.pack(fill=tk.X, padx=10, pady=5)

        self.start_btn = ttk.Button(control_frame, text="启动", command=self._start)
        self.start_btn.pack(side=tk.LEFT, padx=5)
        self.stop_btn = ttk.Button(control_frame, text="停止", command=self._stop, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=5)
        ttk.Button(control_frame, text="编辑总配置", command=self._edit_main_config).pack(side=tk.LEFT, padx=5)
        self.scroll_btn = ttk.Button(control_frame, text="日志自动滚动：开", command=self._toggle_auto_scroll)
        self.scroll_btn.pack(side=tk.RIGHT, padx=5)

        ttk.Button(control_frame, text="清除日志", command=self._clear_log).pack(side=tk.RIGHT, padx=5)

        log_frame = ttk.LabelFrame(self.root, text="运行日志", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        log_bg, log_fg = get_log_colors()
        log_text_frame = ttk.Frame(log_frame)
        log_text_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = tk.Text(log_text_frame, wrap=tk.WORD, state=tk.DISABLED, bg=log_bg, fg=log_fg)
        log_scrollbar = ttk.Scrollbar(log_text_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        apply_theme_to_log(self.log_text)

        self.root.after(200, lambda: apply_theme(
            self.root,
            get_theme_id_by_name(self.theme_var.get()) if self.theme_var.get() else "light"
        ))
        apply_theme(
            self.root,
            get_theme_id_by_name(self.theme_var.get()) if self.theme_var.get() else "light"
        )
        self.root.after(150, lambda: apply_theme(
            self.root,
            get_theme_id_by_name(self.theme_var.get()) if self.theme_var.get() else "light"
        ))

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

    def _open_sync_file(self):
        try:
            if not os.path.exists(self.SYNC_FILE_PATH):
                self.log(f"❌ 未找到配置同步工具：{self.SYNC_FILE_PATH}")
                messagebox.showerror("错误", f"未找到脚本文件：{self.SYNC_FILE_PATH}")
                return

            main_config = configparser.ConfigParser()
            main_config.read(MAIN_CONFIG_PATH, encoding="utf-8")
            where_alas = main_config["GENERAL"].get("where_alas", "").strip()
            if not where_alas or not os.path.exists(where_alas):
                self.log("❌ 未配置有效的Alas路径，无法确定同步目标目录")
                messagebox.showerror("错误", "请先在「Alas配置」中配置有效的Alas路径，\n同步工具将打开Alas同目录下的 config 文件夹。")
                return

            alas_dir = os.path.dirname(where_alas)
            target_config_dir = os.path.join(alas_dir, "config")
            if not os.path.isdir(target_config_dir):
                self.log(f"❌ 目标配置目录不存在：{target_config_dir}")
                messagebox.showerror("错误", f"目标配置目录不存在：\n{target_config_dir}")
                return

            subprocess.Popen(
                [sys.executable, self.SYNC_FILE_PATH, target_config_dir],
                creationflags=subprocess.CREATE_NEW_CONSOLE if sys.platform == "win32" else 0
            )
            self.log(f"✅ 已打开配置同步工具，目标目录：{target_config_dir}")
        except Exception as e:
            self.log(f"❌ 打开配置同步工具失败：{str(e)}")
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

    def _load_match_step_from_config(self):
        config = configparser.ConfigParser()
        if os.path.exists(MAIN_CONFIG_PATH):
            config.read(MAIN_CONFIG_PATH, encoding="utf-8")
        if "GENERAL" not in config:
            config["GENERAL"] = {}
        try:
            step = float(config["GENERAL"].get("template_match_step", "0.01"))
            step = max(0.001, min(0.2, step))
            self.match_step_var.set(f"{step:.3f}")
        except (ValueError, TypeError):
            self.match_step_var.set("0.01")

    def _save_match_step_to_config(self):
        try:
            input_step = self.match_step_var.get().strip()
            step_val = float(input_step)
            if not (0.001 <= step_val <= 0.2):
                self.log(f"⚠️ 步长值{step_val}超出范围（0.001~0.2）")
                self._load_match_step_from_config()
                return
            config = configparser.ConfigParser()
            if os.path.exists(MAIN_CONFIG_PATH):
                config.read(MAIN_CONFIG_PATH, encoding="utf-8")
            if "GENERAL" not in config:
                config["GENERAL"] = {}
            config["GENERAL"]["template_match_step"] = f"{step_val:.3f}"
            with open(MAIN_CONFIG_PATH, "w", encoding="utf-8") as f:
                config.write(f)
            self.log(f"✅ 匹配步长已保存为：{step_val:.3f}")
        except ValueError:
            self.log(f"❌ 请输入有效的数字（如0.01）！")
            self._load_match_step_from_config()

    def _save_ocr_config_key(self, key, value):
        try:
            config = configparser.ConfigParser()
            if os.path.exists(MAIN_CONFIG_PATH):
                config.read(MAIN_CONFIG_PATH, encoding="utf-8")
            if "OCRConfig" not in config:
                config["OCRConfig"] = {}
            config["OCRConfig"][key] = value
            with open(MAIN_CONFIG_PATH, "w", encoding="utf-8") as f:
                config.write(f)
        except Exception as e:
            self.log(f"⚠️ 保存 OCR 配置失败：{e}")

    def _load_window_combobox(self):
        self.window_list = get_all_visible_windows_simple()
        display_list = [win["display"] for win in self.window_list]
        self.window_combobox["values"] = display_list

        actual_count = len(self.window_combobox["values"])
        if actual_count > 0:
            self.window_combobox.current(0)
        else:
            self.window_combobox.set("")

        if display_list:
            config = configparser.ConfigParser()
            config.read(MAIN_CONFIG_PATH, encoding="utf-8")
            target_hwnd = config["GENERAL"]["target_window_hwnd"]
            target_title = config["GENERAL"]["target_window_title"]
            target_program = config["GENERAL"]["target_program_name"]

            matched_idx = -1

            if target_hwnd and target_hwnd.isdigit():
                hwnd = int(target_hwnd)
                for idx, win in enumerate(self.window_list):
                    if win["hwnd"] == hwnd:
                        matched_idx = idx
                        break

            if matched_idx == -1 and target_title:
                for idx, win in enumerate(self.window_list):
                    if target_title in win["title"]:
                        matched_idx = idx
                        break

            if matched_idx == -1 and target_program:
                for idx, win in enumerate(self.window_list):
                    if target_program in win["process_name"]:
                        matched_idx = idx
                        break

            if matched_idx != -1:
                self.window_combobox.current(matched_idx)
                self._update_window_info(matched_idx)
                self.log(f"✅ 已刷新窗口列表并匹配到默认窗口：{self.window_list[matched_idx]['display']}")
            else:
                self.window_combobox.current(0)
                self._update_window_info(0)
                self.log(f"ℹ️ 已刷新窗口列表（共 {len(self.window_list)} 个可见窗口），未找到默认窗口，选中第一个窗口")
        else:
            self.log("ℹ️ 已刷新窗口列表，未发现可见窗口")

    def _update_window_info(self, idx):
        win = self.window_list[idx]
        info_text = f"标题：{win['title']} | 进程：{win['process_name']} | 句柄：{win['hwnd']}"
        self.window_info_var.set(info_text)

    def _load_default_window_from_config(self):
        config = configparser.ConfigParser()
        config.read(MAIN_CONFIG_PATH, encoding="utf-8")
        target_hwnd = config["GENERAL"]["target_window_hwnd"]

        if target_hwnd and target_hwnd.isdigit():
            hwnd = int(target_hwnd)
            for idx, win in enumerate(self.window_list):
                if win["hwnd"] == hwnd:
                    self.window_combobox.current(idx)
                    self._update_window_info(idx)
                    self.log(f"✅ 已加载配置中的默认窗口：{win['display']}")
                    return

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
        config = configparser.ConfigParser()
        config.read(MAIN_CONFIG_PATH, encoding="utf-8")

        top = tk.Toplevel(self.root)
        top.title("模糊匹配配置")
        top.geometry("300x270")
        top.transient(self.root)
        top.grab_set()

        ttk.Label(top, text="进程名关键词（模糊匹配）：").grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)
        prog_var = tk.StringVar(value=config["GENERAL"]["target_program_name"])
        prog_entry = ttk.Entry(top, textvariable=prog_var, width=30)
        prog_entry.grid(row=1, column=0, padx=10, pady=10)

        ttk.Label(top, text="窗口标题关键词（模糊匹配）：").grid(row=2, column=0, padx=10, pady=10, sticky=tk.W)
        title_var = tk.StringVar(value=config["GENERAL"]["target_window_title"])
        title_entry = ttk.Entry(top, textvariable=title_var, width=30)
        title_entry.grid(row=3, column=0, padx=10, pady=10)

        def save():
            config["GENERAL"]["target_program_name"] = prog_var.get().strip()
            config["GENERAL"]["target_window_title"] = title_var.get().strip()
            config["GENERAL"]["target_window_hwnd"] = ""
            with open(MAIN_CONFIG_PATH, "w", encoding="utf-8") as f:
                config.write(f)
            self.log(f"✅ 保存模糊匹配配置：进程名={prog_var.get()}, 标题={title_var.get()}")
            top.destroy()

        ttk.Button(top, text="保存", command=save).grid(row=4, column=0, padx=10, pady=10)

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

        if click_mode == "pyautogui":
            self.click_mode_combobox.current(1)
        else:
            self.click_mode_combobox.current(0)

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
        new_win = tk.Toplevel(self.root)
        new_win.title("新建任务组")

        new_win.geometry("350x180")
        new_win.minsize(300, 150)
        new_win.resizable(True, True)

        new_win.update_idletasks()
        screen_width = new_win.winfo_screenwidth()
        screen_height = new_win.winfo_screenheight()
        x = (screen_width - 350) // 2
        y = (screen_height - 180) // 2
        new_win.geometry(f"350x180+{x}+{y}")

        new_win.columnconfigure(0, weight=1)

        ttk.Label(new_win, text="请输入任务组名称：", font=("Arial", 10)).grid(
            row=0, column=0,
            padx=20, pady=20,
            sticky="ew"
        )
        name_var = tk.StringVar()
        name_entry = ttk.Entry(new_win, textvariable=name_var, width=25, font=("Arial", 10))
        name_entry.grid(row=1, column=0,
                        padx=20, pady=5,
                        sticky="ew")
        name_entry.focus()

        def save_group():
            name = name_var.get().strip()
            if not name:
                messagebox.showwarning("提示", "任务组名称不能为空！")
                return

            config_path = os.path.join(TASKS_DIR, f"{name}.json")
            if os.path.exists(config_path):
                messagebox.showerror("错误", "任务组已存在！")
                return

            init_task_config(name)
            self._load_task_groups()
            self.task_var.set(name)
            self.log(f"✅ 新建任务组：{name}")

            new_win.destroy()

        save_btn = ttk.Button(new_win, text="保存", command=save_group)
        save_btn.grid(row=2, column=0,
                    padx=20, pady=15,
                    sticky="ew")

        new_win.grab_set()

    def _edit_main_config(self):
        self._open_file(MAIN_CONFIG_PATH)

    def _edit_task_config(self):
        task_name = self.task_var.get()
        self._open_file(os.path.join(TASKS_DIR, f"{task_name}.json"))

    def _edit_ocr_task(self):
        task_name = self.ocr_task_var.get()
        if not task_name:
            messagebox.showwarning("提示", "请先选择一个 OCR 任务组")
            return
        self._open_file(os.path.join(self.ocr_task_dir, f"{task_name}.json"))

    def _refresh_ocr_device_list(self, force_refresh=False):
        """检测 GPU 并更新加速设备下拉框

        下拉框显示 gpu.json 中的 "name" 字段（便于用户识别），
        内部仍使用 "device" 字段（如 dml:0）进行配置。

        :param force_refresh: 是否强制实时检测（忽略 gpu.json 缓存）
        """
        from gpu_accelerator import GPUAccelerator
        devices = GPUAccelerator.detect_cuda_devices(force_refresh=force_refresh)

        self._device_name_map = {}  # {display_name: device_key}
        display_names = []
        for d in devices:
            clean_name = d["name"].replace("\u0000", "")
            self._device_name_map[clean_name] = d["device"]
            display_names.append(clean_name)

        self.ocr_device_combo.config(values=display_names)

        current_key = self.ocr_device_var.get()
        current_display = None
        for name, key in self._device_name_map.items():
            if key == current_key:
                current_display = name
                break

        if current_display is not None:
            self.ocr_device_var.set(current_display)
        elif display_names:
            self.ocr_device_var.set(display_names[0])

    def _detect_gpu(self):
        """按需检测 GPU 硬件加速状态"""
        self.log("正在检测 GPU ...")
        self.root.update_idletasks()
        gpu_info = self.ocr_engine.gpu_summary()
        for line in gpu_info.split("\n"):
            self.log(f"🖥️ {line}")
        self._refresh_ocr_device_list(force_refresh=True)

    def _ocr_test(self):
        target_win = get_target_window_from_config()
        if not target_win:
            messagebox.showwarning("提示", "请先在主界面选择一个有效的目标窗口")
            return

        main_config = configparser.ConfigParser()
        main_config.read(MAIN_CONFIG_PATH, encoding="utf-8")
        screenshot_mode = main_config["GENERAL"].get("screenshot_mode", "Win32GUI")

        ocr_cfg = self.ocr_engine.get_config()
        device = ocr_cfg["device"]
        language = ",".join(ocr_cfg["language"]) if isinstance(ocr_cfg["language"], list) else ocr_cfg["language"]

        self.ocr_test_btn.config(state=tk.DISABLED)
        self.ocr_stop_btn.config(state=tk.NORMAL)

        self._ocr_test_instance = OCRTest(
            hwnd=target_win._hWnd,
            device=device,
            language=language,
            confidence_threshold=ocr_cfg["confidence_threshold"],
            screenshot_mode=screenshot_mode,
            grayscale=self.ocr_gray_var.get(),
            threshold_enabled=self.ocr_thresh_var.get(),
            on_progress=lambda msg: self.root.after(0, lambda m=msg: self.log(m)),
            on_complete=lambda result: self.root.after(
                0, lambda r=result: self._on_ocr_test_complete(r)
            ),
        )
        self._ocr_test_instance.start()

    def _stop_ocr_test(self):
        """停止 OCR 测试"""
        if self._ocr_test_instance is not None:
            self._ocr_test_instance.cancel()
            self.log("⏹ OCR 测试停止请求已发出...")

    def _on_ocr_test_complete(self, result):
        """OCR 测试完成回调（主线程）"""
        self.ocr_test_btn.config(state=tk.NORMAL)
        self.ocr_stop_btn.config(state=tk.DISABLED)
        self._ocr_test_instance = None

        if result["status"] == "cancelled":
            self.log("OCR 测试已取消")
            return

        if result["status"] == "error":
            self.log("OCR 测试出错")
            return

        self.log(f"  共识别到 {result['text_count']} 条文字：")
        for r in result["texts"]:
            self.log(
                f"    [{r['confidence']}%] {r['text']}  "
                f"@({r['bbox'][0]},{r['bbox'][1]})"
            )

        if not result["texts"]:
            self.log("  未识别到任何文字")

        self.log(f"✅ OCR 测试结束  |  耗时: {result['duration_ms']} ms")

    def _log_threadsafe(self, msg):
        """线程安全的日志输出"""
        self.root.after(0, lambda: self.log(msg))

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
        if self.thread and self.thread.is_alive():
            self.stop_flag = True
            self.worker_generation += 1
            self.log("⏳ 等待旧线程停止...")

            start_wait = time.time()
            self.thread.join(timeout=5.0)
            wait_time = time.time() - start_wait

            if self.thread.is_alive():
                self.log(f"⚠️ 旧线程停止超时（等待{wait_time:.2f}秒），但世代号已递增，旧 worker 将在下次检测时自行退出")
            else:
                self.log(f"✅ 旧线程已停止（耗时{wait_time:.2f}秒）")

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

    def _stop(self, is_manual=True):
        if self._ocr_test_instance is not None:
            self._ocr_test_instance.cancel()
            self._ocr_test_instance = None
        self.stop_flag = True
        if self.thread and self.thread.is_alive():
            self.thread.join(timeout=1.0)
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.ocr_test_btn.config(state=tk.NORMAL)
        self.ocr_stop_btn.config(state=tk.DISABLED)
        if is_manual:
            self.log("🛑 用户手动停止任务")
        else:
            self.log("🛑 任务因条件触发自动停止")

    def _restart_script(self):
        self.log("🔄 正在重启脚本...")

        self._stop(is_manual=False)

        if getattr(sys, 'frozen', False):
            script_path = sys.executable
        else:
            script_path = os.path.abspath(sys.argv[0])

        subprocess.Popen([sys.executable, script_path] if not getattr(sys, 'frozen', False) else [script_path])

        self.root.destroy()
        sys.exit(0)

    def _clear_log(self):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.log("📝 日志已清除！")

    def _show_about_dialog(self):
        about_window = tk.Toplevel(self.root)
        about_window.title("关于")
        about_window.geometry("500x400")
        about_window.resizable(False, False)

        window_width = 480
        window_height = 390
        screen_width = about_window.winfo_screenwidth()
        screen_height = about_window.winfo_screenheight()
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        about_window.geometry(f"{window_width}x{window_height}+{x}+{y}")

        content_frame = ttk.Frame(about_window, padding=20)
        content_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(content_frame, text="OCR自动点击工具", font=("微软雅黑", 16, "bold")).pack(pady=10)

        ttk.Label(content_frame, text=f"版本: {self.version}").pack(pady=5)

        description = "OCR自动点击工具 - 基于图像识别的自动化操作工具\n\n"
        description += " Copyright (C) 2026 XDista\n\n"
        description += "本程序绝对无任何保证；详情见GNU GPL V3 许可证。\n\n"

        ttk.Label(content_frame, text=description, justify=tk.CENTER).pack(pady=10)

        github_frame = ttk.Frame(content_frame)
        github_frame.pack(pady=(0, 10), anchor='center')

        ttk.Label(github_frame, text="项目GitHub主页：").pack(side='left', anchor='center')

        link_label = ttk.Label(github_frame, text="https://github.com/XDista/OCR_AutoClick",
                             foreground="blue", cursor="hand2", font=('TkDefaultFont', 9, 'underline'))
        link_label.pack(side='left', anchor='center')

        def open_link(event):
            webbrowser.open_new("https://github.com/XDista/OCR_AutoClick")

        link_label.bind("<Button-1>", open_link)

        ttk.Label(content_frame, text="本软件为开源免费项目，如果你以任何付费方式获得此软件，请立即尝试退款！",
                 justify=tk.CENTER).pack(pady=(5, 10))

        ttk.Button(content_frame, text="关闭", command=about_window.destroy).pack(pady=10)

    def log(self, msg):
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"[{ts}] {msg}\n")

        if self.auto_scroll:
            self.log_text.see(tk.END)

        self.root.update_idletasks()
        self.log_text.config(state=tk.DISABLED)