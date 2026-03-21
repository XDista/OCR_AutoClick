import os
import subprocess
import time
import sys
import threading
import configparser
import tkinter as tk
import ctypes
from tkinter import ttk, messagebox, filedialog

# 解决高DPI显示模糊问题
try:
    ctypes.windll.user32.SetProcessDPIAware()
except:
    pass

# -------------------------- 全局配置 --------------------------
SCRIPT_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))
CONFIG_FILE = os.path.join(SCRIPT_DIR, "config.ini")  # 配置文件路径
SCREENSHOT_DIR = os.path.join(SCRIPT_DIR, "adb_temp_screenshot")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MAIN_CONFIG_PATH = os.path.join(BASE_DIR, "config.ini")

# 初始化配置文件
def init_main_config():
    config = configparser.ConfigParser()
    if os.path.exists(MAIN_CONFIG_PATH):
        config.read(MAIN_CONFIG_PATH, encoding="utf-8")
    
    with open(MAIN_CONFIG_PATH, "w", encoding="utf-8") as f:
        config.write(f)
    return config

class ADBScreenshotGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("ADB 设备截图工具")

        config = init_main_config()  # 确保配置已初始化
        adbscrgeometry = config["WindowConfig"]["adbscrgeometry"]
        adbscrxy_changeable = int(config["WindowConfig"]["adbscrxy_changeable"])
        self.root.geometry(adbscrgeometry)#主窗口长宽
        self.root.resizable(adbscrxy_changeable, adbscrxy_changeable)

        # 1. 读取配置文件（从config.ini读取配置）
        self.adb_path, self.device_serial = self._read_config()

        # 2. 初始化界面变量（用配置文件的值填充）
        self.adb_path_var = tk.StringVar(value=self.adb_path)
        self.serial_var = tk.StringVar(value=self.device_serial)

        # 3. 创建界面组件
        self._create_widgets()
        # 4. 初始化截图目录
        self._init_screenshot_dir()

    def _read_config(self):
        """读取config.ini配置（只读，不写入），处理配置文件缺失/字段缺失"""
        config = configparser.ConfigParser()
        adb_path = ""
        device_serial = ""

        # 检查配置文件是否存在
        if not os.path.exists(CONFIG_FILE):
            self._show_config_warning(f"未找到配置文件：{CONFIG_FILE}\n请在脚本目录下创建config.ini并配置相关项。")
            return adb_path, device_serial

        # 读取配置文件
        try:
            config.read(CONFIG_FILE, encoding="utf-8")
            # 读取adb_position（ADB路径）
            if "ADBConfig" in config and "adb_position" in config["ADBConfig"]:
                adb_path = config["ADBConfig"]["adb_position"].strip()
            else:
                self._show_config_warning("config.ini中缺失 [ADBConfig] 节或 adb_position 配置项！")

            # 读取adb_device_serial（设备序列号）
            if "ADBConfig" in config and "adb_device_serial" in config["ADBConfig"]:
                device_serial = config["ADBConfig"]["adb_device_serial"].strip()
            else:
                self._show_config_warning("config.ini中缺失 [ADBConfig] 节或 adb_device_serial 配置项！")

        except Exception as e:
            self._show_config_warning(f"读取配置文件失败：{str(e)}")

        return adb_path, device_serial

    def _show_config_warning(self, msg):
        """显示配置文件相关警告"""
        messagebox.showwarning("配置文件警告", msg)

    def _create_widgets(self):
        """创建GUI界面组件"""
        # 1. ADB路径配置区域
        frame_adb = ttk.LabelFrame(self.root, text="ADB 路径")
        frame_adb.pack(padx=10, pady=5, fill=tk.X)

        ttk.Label(frame_adb, text="ADB路径：").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        entry_adb = ttk.Entry(frame_adb, textvariable=self.adb_path_var, width=50)
        entry_adb.grid(row=0, column=1, padx=5, pady=5)
        btn_select_adb = ttk.Button(frame_adb, text="选择文件", command=self._select_adb_path)
        btn_select_adb.grid(row=0, column=2, padx=5, pady=5)

        # 2. 设备序列号区域
        frame_serial = ttk.LabelFrame(self.root, text="设备序列号")
        frame_serial.pack(padx=10, pady=5, fill=tk.X)

        ttk.Label(frame_serial, text="设备序列号：").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        entry_serial = ttk.Entry(frame_serial, textvariable=self.serial_var, width=50)
        entry_serial.grid(row=0, column=1, padx=5, pady=5)

        # 3. 操作按钮区域
        frame_btn = ttk.Frame(self.root)
        frame_btn.pack(padx=10, pady=10)

        self.btn_screenshot = ttk.Button(
            frame_btn, text="截取屏幕", command=self._run_screenshot_thread
        )
        self.btn_screenshot.grid(row=0, column=0, padx=5)

        btn_open_dir = ttk.Button(
            frame_btn, text="打开截图文件夹", command=self._open_screenshot_dir
        )
        btn_open_dir.grid(row=0, column=1, padx=5)

        # 4. 日志输出区域
        frame_log = ttk.LabelFrame(self.root, text="操作日志")
        frame_log.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(frame_log)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.text_log = tk.Text(frame_log, wrap=tk.WORD, yscrollcommand=scrollbar.set, height=15)
        self.text_log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        scrollbar.config(command=self.text_log.yview)
        self.text_log.config(state=tk.DISABLED)

        # 初始化日志
        self._log(f"📄 配置文件路径：{CONFIG_FILE}")
        self._log(f"📌 从配置文件读取的ADB路径：{self.adb_path if self.adb_path else '未配置'}")
        self._log(f"📌 从配置文件读取的设备序列号：{self.device_serial if self.device_serial else '未配置'}")

    def _init_screenshot_dir(self):
        """初始化截图目录"""
        try:
            os.makedirs(SCREENSHOT_DIR, exist_ok=True)
            self._log(f"✅ 截图保存目录已准备好：{SCREENSHOT_DIR}")
        except Exception as e:
            messagebox.showerror("错误", f"创建截图目录失败：{e}")
            sys.exit(1)

    def _select_adb_path(self):
        """选择ADB.exe文件路径（手动修改）"""
        file_path = filedialog.askopenfilename(
            title="选择adb.exe",
            filetypes=[("可执行文件", "*.exe"), ("所有文件", "*.*")],
            initialdir=os.path.dirname(self.adb_path_var.get()) if self.adb_path_var.get() else SCRIPT_DIR
        )
        if file_path:
            self.adb_path_var.set(file_path)
            self._log(f"📝 ADB路径已手动更新：{file_path}")

    def _log(self, msg):
        """日志输出到文本框"""
        self.text_log.config(state=tk.NORMAL)
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
        self.text_log.insert(tk.END, f"[{timestamp}] {msg}\n")
        self.text_log.see(tk.END)
        self.text_log.config(state=tk.DISABLED)

    def _open_screenshot_dir(self):
        """打开截图保存文件夹"""
        if os.path.exists(SCREENSHOT_DIR):
            os.startfile(SCREENSHOT_DIR)
        else:
            messagebox.warning("提示", "截图文件夹不存在！")

    def _adb_screenshot(self):
        """核心截图逻辑（后台执行）"""
        adb_path = self.adb_path_var.get().strip()
        serial = self.serial_var.get().strip()

        # 前置检查
        if not adb_path or not os.path.exists(adb_path):
            self._log("❌ ADB路径无效，请检查config.ini或手动选择！")
            return
        if not serial:
            self._log("❌ 设备序列号不能为空，请检查config.ini或手动输入！")
            return

        # 定义路径
        device_temp_path = "/sdcard/temp_screenshot.png"
        local_filename = f"screenshot_{int(time.time())}.png"
        local_save_path = os.path.join(SCREENSHOT_DIR, local_filename)

        try:
            # 步骤1：设备端截图
            self._log(f"\n📱 开始截取设备 [{serial}] 的屏幕...")
            screencap_cmd = [adb_path, "-s", serial, "shell", "screencap", "-p", device_temp_path]
            result = subprocess.run(
                screencap_cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, encoding="utf-8"
            )
            if result.returncode != 0:
                raise RuntimeError(f"截图命令执行失败：{result.stderr.strip()}")

            # 步骤2：拉取截图到本地
            self._log("📤 正在拉取截图到本地...")
            pull_cmd = [adb_path, "-s", serial, "pull", device_temp_path, local_save_path]
            result = subprocess.run(
                pull_cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, encoding="utf-8"
            )
            if result.returncode != 0:
                raise RuntimeError(f"拉取截图失败：{result.stderr.strip()}")

            # 步骤3：清理设备临时文件
            subprocess.run(
                [adb_path, "-s", serial, "shell", "rm", "-f", device_temp_path],
                stdout=subprocess.PIPE, stderr=subprocess.PIPE
            )

            self._log(f"🎉 截图成功！保存路径：{local_save_path}")

        except Exception as e:
            self._log(f"❌ 截图失败：{str(e)}")

    def _run_screenshot_thread(self):
        """启动线程执行截图（避免GUI卡死）"""
        screenshot_thread = threading.Thread(target=self._adb_screenshot)
        screenshot_thread.daemon = True
        screenshot_thread.start()

def main():
    """启动GUI"""
    root = tk.Tk()
    app = ADBScreenshotGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()