"""DPI 感知与缩放工具，供主程序与独立小工具统一使用。

背景：
    tkinter 中字体以「磅」为单位，会随系统 DPI 自动缩放；而 padx/pady、
    geometry 宽高等固定像素值不会自动缩放，导致高缩放比例下 UI 错位。

方案：
    1. enable_dpi_awareness() 开启每显示器 DPI 感知 v2；
    2. get_dpi_scale() 计算相对 96 DPI 的缩放因子；
    3. scale_pixels() / scale_geometry() 显式换算固定像素值；
    4. patch_geometry_managers() 让 pack/grid 的 padx/pady/ipadx/ipady
       自动按缩放因子换算，避免逐处改写大量硬编码像素。
"""
import ctypes


def enable_dpi_awareness():
    """开启每显示器 DPI 感知 v2（Windows 10 1703+），旧版系统逐级降级。"""
    try:
        ctypes.windll.user32.SetProcessDpiAwarenessContext(-4)
        return
    except Exception:
        pass
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
        return
    except Exception:
        pass
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


def get_dpi_scale(root):
    """计算相对 96 DPI 基准的缩放因子（100%→1.0，125%→1.25，150%→1.5）。"""
    scale = root.winfo_fpixels("1i") / 96.0
    if scale <= 0:
        scale = 1.0
    return scale


def scale_pixels(value, scale):
    """按缩放因子换算固定像素值（支持单个数值或元组/列表）。"""
    if isinstance(value, (tuple, list)):
        return tuple(scale_pixels(v, scale) for v in value)
    return int(round(value * scale))


def scale_geometry(geometry, scale):
    """缩放 'WxH' 或 'WxH+X+Y' 形式 geometry 字符串中的宽高（位置部分保持不变）。"""
    import re
    m = re.match(r"^(\d+)x(\d+)([+-]\d+[+-]\d+)?$", geometry)
    if not m:
        return geometry
    w = scale_pixels(int(m.group(1)), scale)
    h = scale_pixels(int(m.group(2)), scale)
    pos = m.group(3) or ""
    return f"{w}x{h}{pos}"


def patch_geometry_managers(scale):
    """让 pack/grid 的 padx/pady/ipadx/ipady 按 DPI 等比缩放（全局、幂等、仅一次）。

    返回是否实际应用补丁。100% 缩放（scale==1.0）时无需处理，直接跳过。
    """
    if abs(scale - 1.0) < 1e-6:
        return False
    if getattr(patch_geometry_managers, "_applied", False):
        return False
    patch_geometry_managers._applied = True

    import tkinter as _tk

    def _sc(v):
        return scale_pixels(v, scale)

    _orig_pack = _tk.Pack.pack
    _orig_grid = _tk.Grid.grid

    def _pack(self, *args, **kw):
        for _k in ("padx", "pady", "ipadx", "ipady"):
            if _k in kw:
                kw[_k] = _sc(kw[_k])
        return _orig_pack(self, *args, **kw)

    def _grid(self, *args, **kw):
        for _k in ("padx", "pady", "ipadx", "ipady"):
            if _k in kw:
                kw[_k] = _sc(kw[_k])
        return _orig_grid(self, *args, **kw)

    _tk.Pack.pack = _pack
    _tk.Grid.grid = _grid
    _tk.Pack.pack_configure = _pack
    _tk.Grid.grid_configure = _grid
    return True