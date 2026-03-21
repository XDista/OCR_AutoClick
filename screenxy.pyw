import tkinter as tk
from tkinter import ttk, messagebox
import win32gui
import win32con
import threading
import time
import pyautogui
import ctypes
import win32api
import configparser
import os
import psutil
import win32process
import pygetwindow as gw

# 解决高DPI显示模糊问题
try:
    ctypes.windll.user32.SetProcessDPIAware()
except:
    pass

# 配置路径设置
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MAIN_CONFIG_PATH = os.path.join(BASE_DIR, "config.ini")

# 初始化配置文件
def init_main_config():
    config = configparser.ConfigParser()
    if os.path.exists(MAIN_CONFIG_PATH):
        config.read(MAIN_CONFIG_PATH, encoding="utf-8")
    
    if "GENERAL" not in config:
        config["GENERAL"] = {}
    
    default_config = {
        "target_program_name": "",  # 模糊匹配：进程名关键词
        "target_window_title": "",  # 模糊匹配：窗口标题关键词
        "target_window_hwnd": ""    # 精确匹配：窗口句柄
    }
    
    for key, value in default_config.items():
        if key not in config["GENERAL"]:
            config["GENERAL"][key] = value
    
    with open(MAIN_CONFIG_PATH, "w", encoding="utf-8") as f:
        config.write(f)
    return config

def get_client_rect_screen(hwnd):
    """
    获取窗口客户区的屏幕坐标和尺寸
    返回：(client_left, client_top, client_width, client_height)
    - client_left/client_top：客户区左上角的屏幕绝对坐标
    - client_width/client_height：客户区的宽/高（纯可交互区域）
    """
    try:
        # 获取客户区在窗口内的矩形（相对于窗口左上角，通常left/top=0,0）
        client_rect = win32gui.GetClientRect(hwnd)
        client_l_win, client_t_win, client_r_win, client_b_win = client_rect
        client_width = client_r_win - client_l_win
        client_height = client_b_win - client_t_win

        # 将客户区左上角（窗口内的0,0）转换为屏幕绝对坐标
        client_left_screen, client_top_screen = win32gui.ClientToScreen(hwnd, (client_l_win, client_t_win))
        return client_left_screen, client_top_screen, client_width, client_height
    except:
        return 0, 0, 0, 0

# 获取可见窗口列表
def get_all_visible_windows_simple():
    """获取窗口列表（仅标题+进程名+句柄，用于下拉列表）"""
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
            # 格式：标题 (进程名)
            display_text = f"{title[:50]} ({process_name})" 
            windows.append({
                "display": display_text,
                "hwnd": hwnd,
                "title": title,
                "process_name": process_name
            })
    win32gui.EnumWindows(callback, None)
    return windows

class AdvancedCoordinatePicker:
    def __init__(self, root):
        self.root = root
        self.root.title("窗口坐标拾取工具")

        config = init_main_config()  # 确保配置已初始化
        scrgeometry = config["WindowConfig"]["scrgeometry"]
        scr_changeable = int(config["WindowConfig"]["scrxy_changeable"])
        self.root.geometry(scrgeometry)#主窗口长宽
        self.root.resizable(scr_changeable, scr_changeable)
        
        self.picking = False
        self.selected_hwnd = None
        self.window_list = []  # 存储窗口信息
        
        # 初始化配置
        init_main_config()
        
        # 初始化变量
        self.relative_coord_var = tk.StringVar(value="X: --, Y: --")
        self.screen_coord_var = tk.StringVar(value="X: --, Y: --")
        self.window_size_var = tk.StringVar(value="宽: --, 高: --")
        self.client_size_var = tk.StringVar(value="宽: --, 高: --")  # 新增：客户区尺寸
        self.status_var = tk.StringVar(value="就绪：选择窗口后开始拾取坐标")
        self.window_info_var = tk.StringVar(value="未选择窗口 | 标题：- | 进程：- | 句柄：-")
            
        self._create_gui()
        self._load_window_combobox()
        self._load_default_window_from_config()
        self.pick_event = threading.Event()

    def _create_gui(self):
        # 整体容器（分三部分）
        main_container = ttk.Frame(self.root, padding="10")
        main_container.pack(fill=tk.BOTH, expand=True)

        # 1. 窗口选择区（使用与main一致的逻辑）
        window_frame = ttk.LabelFrame(main_container, text="目标窗口选择", padding="10")
        window_frame.pack(fill=tk.X, pady=5)

        # 窗口选择下拉列表和按钮
        ttk.Label(window_frame, text="选择目标窗口：").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.window_var = tk.StringVar()
        self.window_combobox = ttk.Combobox(
            window_frame, textvariable=self.window_var, state="readonly", width=30
        )
        self.window_combobox.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)

        # 操作按钮：刷新+保存为默认+模糊匹配配置
        ttk.Button(window_frame, text="刷新列表", command=self._load_window_combobox).grid(row=0, column=2, padx=5, pady=5)
        ttk.Button(window_frame, text="设为默认", command=self._save_window_to_config).grid(row=0, column=3, padx=5, pady=5)
        ttk.Button(window_frame, text="模糊匹配配置", command=self._edit_window_match_config).grid(row=0, column=4, padx=5, pady=5)
        
        # 窗口信息显示
        info_frame = ttk.LabelFrame(window_frame, text="窗口信息", padding="10")
        info_frame.grid(row=1, column=0, columnspan=5, sticky=tk.W+tk.E, padx=5, pady=5)
        ttk.Label(info_frame, textvariable=self.window_info_var).pack(side=tk.LEFT, padx=5, pady=5)

        # 2. 信息显示区
        info_frame = ttk.LabelFrame(main_container, text="坐标&窗口信息", padding="10")
        info_frame.pack(fill=tk.X, pady=5)

        # 信息行0：窗口内相对坐标（以窗口左上角为原点）
        ttk.Label(info_frame, text="客户区相对坐标").grid(row=0, column=0, padx=5, pady=3, sticky=tk.W)
        ttk.Label(
            info_frame, textvariable=self.relative_coord_var, font=("Consolas", 10), width=20
        ).grid(row=0, column=1, padx=5, pady=3, sticky=tk.W)

        # 信息行2：屏幕绝对坐标
        ttk.Label(info_frame, text="屏幕绝对坐标：").grid(row=1, column=0, padx=5, pady=3, sticky=tk.W)
        ttk.Label(
            info_frame, textvariable=self.screen_coord_var, font=("Consolas", 10), width=20
        ).grid(row=1, column=1, padx=5, pady=3, sticky=tk.W)

        # 信息行3：窗口尺寸
        ttk.Label(info_frame, text="目标窗口尺寸：").grid(row=2, column=0, padx=5, pady=3, sticky=tk.W)
        ttk.Label(
            info_frame, textvariable=self.window_size_var, font=("Consolas", 10), width=20
        ).grid(row=2, column=1, padx=5, pady=3, sticky=tk.W)

        # 信息行4：新增客户区尺寸
        ttk.Label(info_frame, text="窗口客户区尺寸：").grid(row=3, column=0, padx=5, pady=3, sticky=tk.W)
        ttk.Label(
            info_frame, textvariable=self.client_size_var, font=("Consolas", 10), width=20
        ).grid(row=3, column=1, padx=5, pady=3, sticky=tk.W)

        # 3. 操作按钮区
        btn_frame = ttk.Frame(main_container, padding="10")
        btn_frame.pack(fill=tk.X, pady=5)

        self.pick_btn = ttk.Button(
            btn_frame, text="开始拾取坐标", command=self._start_picking, width=20
        )
        self.pick_btn.pack(side=tk.LEFT, padx=5)

        # 4. 状态提示栏
        status_bar = ttk.Label(
            self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W
        )
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 绑定下拉列表选择事件
        self.window_combobox.bind("<<ComboboxSelected>>", self._on_window_select)

    def _load_window_combobox(self):
        """加载窗口列表到下拉框"""
        self.window_list = get_all_visible_windows_simple()
        display_list = [win["display"] for win in self.window_list]
        self.window_combobox["values"] = display_list
        if display_list:
            self.window_combobox.current(0)
            self._update_window_info(0)  # 显示第一个窗口的信息
        self.status_var.set(f"已刷新窗口列表，共 {len(self.window_list)} 个可见窗口")

    def _on_window_select(self, event):
        """选择下拉列表中的窗口"""
        idx = self.window_combobox.current()
        if idx >= 0 and idx < len(self.window_list):
            self._update_window_info(idx)

    def _update_window_info(self, idx):
        """更新窗口信息显示"""
        win = self.window_list[idx]
        info_text = f"标题：{win['title']} | 进程：{win['process_name']} | 句柄：{win['hwnd']}"
        self.window_info_var.set(info_text)
        
        # 更新窗口尺寸信息
        try:
            left, top, right, bottom = win32gui.GetWindowRect(win['hwnd'])
            self.window_size_var.set(f"宽: {right-left}, 高: {bottom-top}")
            # 新增：更新客户区尺寸
            client_left, client_top, client_w, client_h = get_client_rect_screen(win['hwnd'])
            self.client_size_var.set(f"宽: {client_w}, 高: {client_h}")
        except:
            self.window_size_var.set("宽: --, 高: --")
            self.client_size_var.set("宽: --, 高: --")  # 新增异常处理

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
        
        self.status_var.set(f"已将窗口「{win['display']}」设为默认，写入配置文件")

    def _load_default_window_from_config(self):
        """从配置读取默认默认窗口并选中"""
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
                    self.status_var.set(f"已加载配置中的默认窗口：{win['display']}")
                    return
        
        # 无精确句柄时，尝试模糊匹配标题/进程名
        target_title = config["GENERAL"]["target_window_title"]
        target_program = config["GENERAL"]["target_program_name"]
        if target_title or target_program:
            for idx, win in enumerate(self.window_list):
                if (target_title and target_title in win["title"]) or (target_program and target_program in win["process_name"]):
                    self.window_combobox.current(idx)
                    self._update_window_info(idx)
                    self.status_var.set(f"模糊匹配到默认窗口：{win['display']}")
                    return
        
        self.status_var.set("未找到配置中的默认窗口，使用第一个窗口")

    def _edit_window_match_config(self):
        """编辑模糊匹配的关键词（进程名/窗口标题）"""
        config = configparser.ConfigParser()
        config.read(MAIN_CONFIG_PATH, encoding="utf-8")
        
        # 新建临时窗口，输入模糊匹配关键词
        top = tk.Toplevel(self.root)
        top.title("模糊匹配配置")
        top.geometry("400x200")
        top.transient(self.root)
        top.grab_set()
        
        # 进程名关键词
        ttk.Label(top, text="进程名关键词（模糊匹配）：").grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)
        prog_var = tk.StringVar(value=config["GENERAL"]["target_program_name"])
        prog_entry = ttk.Entry(top, textvariable=prog_var, width=30)
        prog_entry.grid(row=0, column=1, padx=10, pady=10)
        
        # 窗口标题关键词
        ttk.Label(top, text="窗口标题关键词（模糊匹配）：").grid(row=1, column=0, padx=10, pady=10, sticky=tk.W)
        title_var = tk.StringVar(value=config["GENERAL"]["target_window_title"])
        title_entry = ttk.Entry(top, textvariable=title_var, width=30)
        title_entry.grid(row=1, column=1, padx=10, pady=10)
        
        # 保存按钮
        def save():
            config["GENERAL"]["target_program_name"] = prog_var.get().strip()
            config["GENERAL"]["target_window_title"] = title_var.get().strip()
            config["GENERAL"]["target_window_hwnd"] = ""  # 清空精确句柄，启用模糊匹配
            with open(MAIN_CONFIG_PATH, "w", encoding="utf-8") as f:
                config.write(f)
            self.status_var.set(f"保存模糊匹配配置：进程名={prog_var.get()}, 标题={title_var.get()}")
            top.destroy()
        
        ttk.Button(top, text="保存", command=save).grid(row=2, column=1, padx=10, pady=10)

    def _get_selected_window(self):
        """获取选中的窗口信息"""
        if not self.window_list:
            return None
        idx = self.window_combobox.current()
        return self.window_list[idx] if 0 <= idx < len(self.window_list) else None

    def _start_picking(self):
        """启动坐标拾取流程"""
        selected = self._get_selected_window()
        if not selected:
            messagebox.showerror("错误", "请先选择一个目标窗口")
            return
        
        title, hwnd = selected["title"], selected["hwnd"]
        self.selected_hwnd = hwnd

        # 激活目标窗口（只在最小化时恢复，保留最大化状态）
        try:
            # 获取窗口当前状态
            placement = win32gui.GetWindowPlacement(hwnd)

            # 如果窗口处于最小化状态，才进行恢复
            if placement[1] == win32con.SW_SHOWMINIMIZED:
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)

            # 无论何种状态都将窗口置于前台
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.1)  # 缩短激活窗口后的延迟，提升响应

        except Exception as e:
            messagebox.showerror("失败", f"无法激活窗口：{str(e)}")
            return
        # 更新窗口尺寸（含客户区）
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        self.window_size_var.set(f"宽: {right-left}, 高: {bottom-top}")
        # 新增：更新客户区尺寸
        client_left, client_top, client_w, client_h = get_client_rect_screen(hwnd)
        self.client_size_var.set(f"宽: {client_w}, 高: {client_h}")

        # 切换拾取状态
        self.picking = True
        self.pick_event.set()
        self.pick_btn.config(text="取消拾取", command=self._stop_picking)
        self.status_var.set(f"拾取中：在「{title}」窗口客户区内点击目标位置（ESC键取消）")

        # 启动拾取线程（传入客户区左上角坐标）
        threading.Thread(target=self._pick_loop, args=(client_left, client_top), daemon=True).start()
    

        # 更新窗口尺寸信息
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        self.window_size_var.set(f"宽: {right-left}, 高: {bottom-top}")

        # 切换拾取状态
        self.picking = True
        self.pick_event.set()  # 设置事件标志
        self.pick_btn.config(text="取消拾取", command=self._stop_picking)
        self.status_var.set(f"拾取中：在「{title}」窗口内点击目标位置（ESC键取消）")

        # 启动拾取线程
        threading.Thread(target=self._pick_loop, daemon=True).start()

    def _stop_picking(self):
        """停止坐标拾取"""
        self.picking = False
        self.pick_event.clear()  # 清除事件标志
        self.pick_btn.config(text="开始拾取坐标", command=self._start_picking)
        self.status_var.set("就绪：选择窗口后开始拾取坐标")

    def _pick_loop(self, client_left, client_top):  # 新增参数：客户区左上角屏幕坐标
        """坐标拾取核心循环"""
        hwnd = self.selected_hwnd
        screen_width, screen_height = pyautogui.size()
        edge_threshold = 1  # 边缘检测

        # 标记浮窗是否已创建
        self.hint_win_created = False
        self.hint_win = None
        self.hint_label = None

        # 定义创建浮窗的函数（必须在主线程执行）
        def create_hint_window():
            self.hint_win = tk.Toplevel(self.root)
            self.hint_win.overrideredirect(True)
            self.hint_win.attributes("-alpha", 0.9, "-topmost", True)
            self.hint_label = tk.Label(
                self.hint_win, text="相对坐标：X-- Y--",
                bg="#FFEE00", font=("微软雅黑", 9),
                padx=3, pady=1
            )
            self.hint_label.pack()
            self.hint_win_created = True

        # 请求主线程创建浮窗
        self.root.after(0, create_hint_window)
        
        # 等待浮窗创建完成
        while not self.hint_win_created:
            time.sleep(0.01)

        try:
           # 使用事件标志和picking状态共同控制循环
           while self.picking and self.pick_event.is_set():
                mouse_x, mouse_y = win32api.GetCursorPos()
                
                is_near_edge = (
                    mouse_x <= edge_threshold
                    or mouse_x >= screen_width - edge_threshold
                    or mouse_y <= edge_threshold
                    or mouse_y >= screen_height - edge_threshold
                )
                
                if is_near_edge:
                    self.status_var.set("⚠️ 提示：鼠标接近屏幕边缘，注意故障安全机制")
                    time.sleep(0.01)

                # 核心修改：计算客户区相对坐标（替代原有整个窗口的相对坐标）
                rel_x = mouse_x - client_left
                rel_y = mouse_y - client_top
                
                # 校验坐标是否在客户区内（可选：负数/超过客户区尺寸则标红提示）
                client_w, client_h = get_client_rect_screen(hwnd)[2:]
                is_in_client = 0 <= rel_x <= client_w and 0 <= rel_y <= client_h
                
                # 请求主线程更新浮窗（线程安全方式）
                def update_hint(rel_x, rel_y, mouse_x, mouse_y):
                    if self.hint_win and self.hint_label:
                        self.hint_win.geometry(f"+{mouse_x+10}+{mouse_y+10}")
                        self.hint_label.config(text=f"相对坐标：X：{rel_x}   Y：{rel_y}")
                
                # 注意：将rel_x/rel_y作为参数传入update_hint，避免闭包变量延迟绑定问题
                self.root.after(0, update_hint, rel_x, rel_y, mouse_x, mouse_y)

                # 检测左键点击
                if win32api.GetKeyState(0x01) < 0:
                    if win32gui.PtInRect(win32gui.GetWindowRect(hwnd), (mouse_x, mouse_y)):
                        self.relative_coord_var.set(f"X: {rel_x}, Y: {rel_y}")
                        self.screen_coord_var.set(f"X: {mouse_x}, Y: {mouse_y}")
                        self.status_var.set(f"已拾取：相对({rel_x},{rel_y}) | 绝对({mouse_x},{mouse_y})")
                    else:
                        self.status_var.set("提示：点击位置不在目标窗口内")
                    self.root.after(0, self._stop_picking)
                    break

                 # 修复ESC键检测：使用GetAsyncKeyState并检查最高位（按下状态）
                if win32api.GetAsyncKeyState(0x1B) & 0x8000:
                    self.root.after(0, self._stop_picking)
                    break

                time.sleep(0.005)
        finally:
            # 主线程销毁浮窗
            def destroy_hint():
                if self.hint_win:
                    self.hint_win.destroy()
                    self.hint_win = None
                    self.hint_label = None
            
            self.root.after(0, destroy_hint)
            
if __name__ == "__main__":
    root = tk.Tk()
    app = AdvancedCoordinatePicker(root)
    root.mainloop()