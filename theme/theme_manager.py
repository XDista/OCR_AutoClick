"""
主题管理模块 - 提供浅色/深色两种主题
基于 Windows vista 原生样式，完全独立于 GUI 代码
"""

import os
import configparser
import tkinter as tk
from tkinter import ttk

# ============================================================
# Windows 原生暗色模式 API（模块级工具函数）
# ============================================================

import ctypes

_WINDOWS_DARK_MODE_AVAILABLE = False
_RefreshImmersiveColorPolicyState = None
_SetWindowTheme = None
_AllowDarkModeForWindow = None
_SetPreferredAppMode = None

try:
    _uxtheme = ctypes.windll.uxtheme
    _SetPreferredAppMode = _uxtheme[135]
    _SetPreferredAppMode.argtypes = [ctypes.c_int]
    _SetPreferredAppMode.restype = ctypes.c_int

    _AllowDarkModeForWindow = _uxtheme[133]
    _AllowDarkModeForWindow.argtypes = [ctypes.c_int, ctypes.c_int]
    _AllowDarkModeForWindow.restype = ctypes.c_int

    try:
        _RefreshImmersiveColorPolicyState = _uxtheme[104]
        _RefreshImmersiveColorPolicyState.argtypes = []
        _RefreshImmersiveColorPolicyState.restype = None
    except Exception:
        pass

    try:
        _SetWindowTheme = ctypes.windll.uxtheme.SetWindowTheme
        _SetWindowTheme.argtypes = [ctypes.c_int, ctypes.c_wchar_p, ctypes.c_wchar_p]
        _SetWindowTheme.restype = ctypes.c_int
    except Exception:
        pass

    _SetPreferredAppMode(1)
    _WINDOWS_DARK_MODE_AVAILABLE = True
except Exception:
    pass

# ============================================================
# 主题颜色定义
# ============================================================

LIGHT_COLORS = {
    "window_bg": "#F0F0F0",
    "frame_bg": "#F0F0F0",
    "label_fg": "#000000",
    "button_bg": "#E8E8E8",
    "button_fg": "#000000",
    "button_active_bg": "#D0D0D0",
    "button_pressed_bg": "#C0C0C0",
    "button_border": "#ADADAD",
    "entry_bg": "#FFFFFF",
    "entry_fg": "#000000",
    "entry_border": "#8A8A8A",
    "combobox_bg": "#FFFFFF",
    "combobox_fg": "#000000",
    "combobox_border": "#8A8A8A",
    "highlight": "#0078D4",
    "highlight_fg": "#FFFFFF",
    "separator": "#D0D0D0",
    "notebook_tab_bg": "#EBEBEB",
    "notebook_tab_selected_bg": "#FFFFFF",
    "notebook_tab_fg": "#000000",
    "notebook_border": "#AAAAAA",
    "scrollbar_thumb": "#C0C0C0",
    "scrollbar_bg": "#F0F0F0",
    "log_bg": "#FFFFFF",
    "log_fg": "#000000",
}

DARK_COLORS = {
    "window_bg": "#1E1E1E",
    "frame_bg": "#2D2D2D",
    "label_fg": "#FFFFFF",
    "button_bg": "#3C3C3C",
    "button_fg": "#FFFFFF",
    "button_active_bg": "#4A4A4A",
    "button_pressed_bg": "#5A5A5A",
    "button_border": "#555555",
    "entry_bg": "#2D2D2D",
    "entry_fg": "#FFFFFF",
    "entry_border": "#555555",
    "combobox_bg": "#2D2D2D",
    "combobox_fg": "#FFFFFF",
    "combobox_border": "#555555",
    "highlight": "#0078D4",
    "highlight_fg": "#FFFFFF",
    "separator": "#555555",
    "notebook_tab_bg": "#333333",
    "notebook_tab_selected_bg": "#1E1E1E",
    "notebook_tab_fg": "#FFFFFF",
    "notebook_border": "#6A6A6A",
    "scrollbar_thumb": "#555555",
    "scrollbar_bg": "#2D2D2D",
    "log_bg": "#1E1E1E",
    "log_fg": "#FFFFFF",
}

FONT_NAME = "Microsoft YaHei UI"
FONT_SIZE = 9
FONT_TITLE_SIZE = 10

THEMES = {
    "light": {"display_name": "浅色", "theme_type": "light", "colors": LIGHT_COLORS},
    "dark": {"display_name": "深色", "theme_type": "dark", "colors": DARK_COLORS},
}

DEFAULT_THEME = "light"

# 模块级当前主题 ID，供 get_log_colors / apply_theme_to_log 在未传参时回退使用
_current_theme_id = DEFAULT_THEME


# ============================================================
# 公共工具函数
# ============================================================

def _refresh_color_policy():
    if _RefreshImmersiveColorPolicyState is not None:
        try:
            _RefreshImmersiveColorPolicyState()
        except Exception:
            pass


def _redraw_window(hwnd):
    if hwnd and hwnd != 0:
        try:
            ctypes.windll.user32.RedrawWindow(
                hwnd, None, 0,
                0x0401  # RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN
            )
        except Exception:
            pass


def _force_titlebar_repaint(hwnd):
    """强制标题栏重绘，确保 DwmSetWindowAttribute 生效"""
    if not hwnd or hwnd == 0:
        return
    try:
        WM_NCACTIVATE = 0x0086
        ctypes.windll.user32.SendMessageW(hwnd, WM_NCACTIVATE, 0, 0)
        ctypes.windll.user32.SendMessageW(hwnd, WM_NCACTIVATE, 1, 0)
    except Exception:
        pass


def get_theme_names():
    """返回可用主题的显示名称列表"""
    return [THEMES[t]["display_name"] for t in THEMES]


def get_theme_ids():
    """返回可用主题的 ID 列表"""
    return list(THEMES.keys())


def get_theme_id_by_name(display_name):
    """通过显示名称获取主题 ID"""
    for tid, data in THEMES.items():
        if data["display_name"] == display_name:
            return tid
    return DEFAULT_THEME


def get_theme_name(theme_id):
    """通过主题 ID 获取显示名称"""
    if theme_id in THEMES:
        return THEMES[theme_id]["display_name"]
    return THEMES[DEFAULT_THEME]["display_name"]


# ============================================================
# 主题应用核心
# ============================================================

def _configure_ttk_styles(style, colors):
    """配置所有 ttk 控件样式（基于 vista 主题）"""
    fn = FONT_NAME
    fs = FONT_SIZE

    style.configure(".", background=colors["window_bg"], foreground=colors["label_fg"],
                    font=(fn, fs))
    style.configure("TFrame", background=colors["frame_bg"])
    style.configure("TLabelframe", background=colors["frame_bg"], foreground=colors["label_fg"])
    style.configure("TLabelframe.Label", foreground=colors["label_fg"],
                    font=(fn, FONT_TITLE_SIZE))

    border_c = colors["button_border"]
    style.configure("TButton", background=colors["button_bg"], foreground=colors["button_fg"],
                    font=(fn, fs), borderwidth=1, focuscolor="none", relief="solid",
                    bordercolor=border_c)
    style.map("TButton",
              background=[("active", colors["button_active_bg"]),
                          ("pressed", colors["button_pressed_bg"])],
              bordercolor=[("active", border_c), ("pressed", border_c)])

    style.configure("TEntry", fieldbackground=colors["entry_bg"],
                    foreground=colors["entry_fg"], bordercolor=colors["entry_border"],
                    font=(fn, fs), borderwidth=1, relief="solid")

    style.configure("TCombobox", fieldbackground=colors["combobox_bg"],
                    foreground=colors["combobox_fg"], bordercolor=colors["combobox_border"],
                    font=(fn, fs), arrowcolor=colors["combobox_fg"],
                    borderwidth=1, relief="solid", background=colors["combobox_bg"])
    style.map("TCombobox",
              fieldbackground=[("readonly", colors["combobox_bg"])],
              foreground=[("readonly", colors["combobox_fg"])])

    style.configure("TNotebook", background=colors["frame_bg"], borderwidth=1,
                    bordercolor=colors["separator"], tabmargins=[2, 2, 2, 0])

    style.configure("TNotebook.Tab", background=colors["notebook_tab_bg"],
                    foreground=colors["notebook_tab_fg"], font=(fn, fs),
                    padding=[20, 5, 20, 5], borderwidth=1, relief="solid",
                    bordercolor=colors["separator"])
    style.map("TNotebook.Tab",
              background=[("selected", colors["notebook_tab_selected_bg"]),
                          ("active", colors["button_active_bg"])],
              bordercolor=[("selected", colors["notebook_border"])])

    style.configure("TCheckbutton", background=colors["frame_bg"],
                    foreground=colors["label_fg"], font=(fn, fs))
    style.configure("TLabel", background=colors["frame_bg"],
                    foreground=colors["label_fg"], font=(fn, fs))

    style.configure("TScrollbar", background=colors["scrollbar_thumb"],
                    troughcolor=colors["scrollbar_bg"], gripcount=0, borderwidth=0)
    style.configure("Vertical.TScrollbar", background=colors["scrollbar_thumb"],
                    troughcolor=colors["scrollbar_bg"], gripcount=0, borderwidth=0)
    style.configure("Horizontal.TScrollbar", background=colors["scrollbar_thumb"],
                    troughcolor=colors["scrollbar_bg"], gripcount=0, borderwidth=0)


def _set_windows_dark_mode(hwnd, enable):
    """对单个窗口 HWND 应用 Windows 原生暗色/亮色模式"""
    if not _WINDOWS_DARK_MODE_AVAILABLE or not hwnd or hwnd == 0:
        return
    try:
        _SetPreferredAppMode(2 if enable else 1)
        _AllowDarkModeForWindow(hwnd, 1 if enable else 0)

        dark_val = ctypes.c_int(1 if enable else 0)
        ctypes.windll.dwmapi.DwmSetWindowAttribute(
            hwnd, 20, ctypes.byref(dark_val), ctypes.sizeof(ctypes.c_int))

        _refresh_color_policy()

        SWP_FRAMECHANGED = 0x0020
        ctypes.windll.user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0,
                                           0x0027 | SWP_FRAMECHANGED)

        if _SetWindowTheme is not None:
            theme_name = "DarkMode_Explorer" if enable else "Explorer"
            _SetWindowTheme(hwnd, theme_name, None)

        _force_titlebar_repaint(hwnd)
    except Exception:
        pass


def _recursive_set_window_theme(parent, enable):
    """递归对 parent 下所有控件设置 SetWindowTheme

    Entry / Combobox / Scrollbar 在 DarkMode_Explorer 下会强制白色背景，
    因此对它们清除原生主题，完全交由 ttk 样式控制。
    """
    if _SetWindowTheme is None:
        return
    theme_name = "DarkMode_Explorer" if enable else "Explorer"
    _text_input_types = (ttk.Entry, ttk.Combobox, ttk.Scrollbar)
    try:
        for child in parent.winfo_children():
            try:
                ch = child.winfo_id()
                if ch and ch != 0:
                    if enable and isinstance(child, _text_input_types):
                        _SetWindowTheme(ch, "", "")
                    else:
                        _SetWindowTheme(ch, theme_name, None)
            except Exception:
                pass
            _recursive_set_window_theme(child, enable)
    except Exception:
        pass


def _apply_combobox_dark_mode(parent, enable):
    """对 Combobox 控件清除原生主题，使其完全由 ttk 样式控制"""
    if _SetWindowTheme is None:
        return
    try:
        for child in parent.winfo_children():
            if isinstance(child, ttk.Combobox):
                hwnd = child.winfo_id()
                if hwnd and hwnd != 0:
                    if enable:
                        _SetWindowTheme(hwnd, "", "")
                    else:
                        _SetWindowTheme(hwnd, "Explorer", None)
            _apply_combobox_dark_mode(child, enable)
    except Exception:
        pass


def _apply_notebook_dark_mode(parent, enable):
    """递归处理 Notebook 控件的暗色/亮色模式"""
    if _SetWindowTheme is None:
        return
    try:
        for child in parent.winfo_children():
            if isinstance(child, ttk.Notebook):
                hwnd = child.winfo_id()
                if hwnd and hwnd != 0:
                    if enable:
                        _SetWindowTheme(hwnd, "", "")
                        _AllowDarkModeForWindow(hwnd, 1)
                    else:
                        _SetWindowTheme(hwnd, "Explorer", None)
                        _AllowDarkModeForWindow(hwnd, 0)
                    for tab_child in child.winfo_children():
                        th = tab_child.winfo_id()
                        if th and th != 0:
                            _SetWindowTheme(th, "" if enable else "Explorer", None)
                    _redraw_window(hwnd)
            _apply_notebook_dark_mode(child, enable)
        _refresh_color_policy()
    except Exception:
        pass


def _configure_combobox_popdown(combobox, is_dark, colors):
    """配置单个 Combobox 的下拉弹窗颜色"""
    if _SetWindowTheme is None or not _WINDOWS_DARK_MODE_AVAILABLE:
        return
    try:
        cb_path = str(combobox)
        bg = colors["combobox_bg"]
        fg = colors["combobox_fg"]
        sel_bg = colors["highlight"]
        sel_fg = colors["highlight_fg"]

        pop_hw = combobox.tk.eval(
            f'catch {{ winfo id [ttk::combobox::PopdownWindow {cb_path}] }}')
        if pop_hw and pop_hw != '' and pop_hw.startswith('0x'):
            pop_hwnd = int(pop_hw, 16)
            _AllowDarkModeForWindow(pop_hwnd, 1 if is_dark else 0)
            _SetWindowTheme(pop_hwnd, "" if is_dark else "Explorer", None)

        lb_hw = combobox.tk.eval(
            f'catch {{ winfo id [ttk::combobox::PopdownWindow {cb_path}].f.l }}')
        if lb_hw and lb_hw != '' and lb_hw.startswith('0x'):
            lb_hwnd = int(lb_hw, 16)
            if lb_hwnd and lb_hwnd != 0:
                _SetWindowTheme(lb_hwnd, "" if is_dark else "Explorer", None)

        sb_hw = combobox.tk.eval(
            f'catch {{ winfo id [ttk::combobox::PopdownWindow {cb_path}].f.sb }}')
        if not sb_hw or sb_hw == '' or not sb_hw.startswith('0x'):
            sb_hw = combobox.tk.eval(
                f'catch {{ winfo id [ttk::combobox::PopdownWindow {cb_path}].f.scrollbar }}')
        if sb_hw and sb_hw != '' and sb_hw.startswith('0x'):
            sb_hwnd = int(sb_hw, 16)
            if sb_hwnd and sb_hwnd != 0:
                _SetWindowTheme(sb_hwnd, "" if is_dark else "Explorer", None)

        combobox.tk.eval(f'''
            catch {{
                set pop [ttk::combobox::PopdownWindow {cb_path}]
                $pop configure -background {bg}
                $pop.f configure -background {bg}
                $pop.f.l configure \
                    -background {bg} \
                    -foreground {fg} \
                    -selectbackground {sel_bg} \
                    -selectforeground {sel_fg}
                $pop.f.l selection clear 0 end
            }}
        ''')
    except Exception:
        pass


def _stretch_notebook_tabs(widget):
    """让 Notebook 标签页铺满宽度"""
    if isinstance(widget, ttk.Notebook):
        num_tabs = widget.index("end")
        if num_tabs > 0:
            widget_width = widget.winfo_width()
            if widget_width > 0:
                tab_width = max(40, (widget_width - 20) // num_tabs)
                for tab_id in range(num_tabs):
                    try:
                        widget.tab(tab_id, sticky="ew")
                    except Exception:
                        pass
                style = ttk.Style()
                style.configure("TNotebook.Tab", width=tab_width, anchor=tk.CENTER)
    for child in widget.winfo_children():
        try:
            _stretch_notebook_tabs(child)
        except Exception:
            pass


# ============================================================
# 主题应用入口
# ============================================================

def apply_theme(root, theme_id=None):
    """应用主题到指定根窗口，返回实际应用的 theme_id"""
    global _current_theme_id

    if theme_id is None or theme_id not in THEMES:
        theme_id = DEFAULT_THEME

    _current_theme_id = theme_id

    theme_data = THEMES[theme_id]
    colors = theme_data["colors"]
    is_dark = theme_data["theme_type"] == "dark"

    style = ttk.Style()
    style.theme_use("vista")

    root.configure(bg=colors["window_bg"])

    _configure_ttk_styles(style, colors)

    # Windows 原生暗色模式
    try:
        hwnd = root.winfo_id()
        if hwnd and hwnd != 0:
            _set_windows_dark_mode(hwnd, is_dark)
            _recursive_set_window_theme(root, is_dark)
            _redraw_window(hwnd)
    except Exception:
        pass

    _apply_notebook_dark_mode(root, is_dark)
    _apply_combobox_dark_mode(root, is_dark)

    # Notebook 标签页铺满
    try:
        _stretch_notebook_tabs(root)
    except Exception:
        pass

    # Combobox 下拉列表颜色
    try:
        _refresh_combobox_popdowns(root, is_dark, colors)
    except Exception:
        pass

    root.update_idletasks()
    return theme_id


def _refresh_combobox_popdowns(parent, is_dark, colors):
    """遍历所有 Combobox，配置弹窗颜色并绑定事件"""
    for child in parent.winfo_children():
        if isinstance(child, ttk.Combobox):
            if not hasattr(child, '_theme_popdown_bound'):
                child._theme_popdown_bound = True

                def make_postcmd(w, dark_on, cols):
                    def fn():
                        w.after(10, lambda: _configure_combobox_popdown(w, dark_on, cols))
                    return fn

                orig_post = child.cget("postcommand")
                child._orig_postcmd = orig_post if orig_post else None
                child.configure(postcommand=make_postcmd(child, is_dark, colors))

                def make_clearer(w, path):
                    def on_selected(event=None):
                        try:
                            w.tk.eval(
                                'catch { [ttk::combobox::PopdownWindow '
                                + path + '].f.l selection clear 0 end }')
                            w.tk.eval(f'catch {{ after idle {{{path} selection clear}} }}')
                        except Exception:
                            pass
                    return on_selected
                child.bind("<<ComboboxSelected>>", make_clearer(child, str(child)), add="+")
        _refresh_combobox_popdowns(child, is_dark, colors)


def apply_theme_to_log(log_text_widget, theme_id=None):
    """对日志 Text 控件应用主题颜色"""
    if theme_id is None or theme_id not in THEMES:
        theme_id = _current_theme_id
    colors = THEMES[theme_id]["colors"]
    log_text_widget.config(bg=colors["log_bg"], fg=colors["log_fg"])


def get_log_colors(theme_id=None):
    """获取日志区域的前景色和背景色"""
    if theme_id is None or theme_id not in THEMES:
        theme_id = _current_theme_id
    colors = THEMES[theme_id]["colors"]
    return colors["log_bg"], colors["log_fg"]


# ============================================================
# 配置读写
# ============================================================

def save_theme_config(theme_id, config_path):
    """保存主题到配置文件"""
    c = configparser.ConfigParser()
    if os.path.exists(config_path):
        c.read(config_path, encoding="utf-8")
    if "Theme" not in c:
        c["Theme"] = {}
    c["Theme"]["current_theme"] = theme_id
    with open(config_path, "w", encoding="utf-8") as f:
        c.write(f)


def load_theme_config(config_path):
    """从配置文件读取主题，返回 theme_id 或 None"""
    c = configparser.ConfigParser()
    if not os.path.exists(config_path):
        return None
    c.read(config_path, encoding="utf-8")
    return c.get("Theme", "current_theme", fallback=None)