import ctypes
import os
import sys
import tkinter as tk

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "subscript"))

from subscript.gui import AutoClickGUI
from app_registry import set_app

APP_VERSION = "1.4.5"
UPDATE_TIME = "2026-09-27 00:00:00"

try:
    ctypes.windll.user32.SetProcessDPIAware()
except:
    pass

try:
    ctypes.windll.uxtheme[135](1)
except:
    pass

if __name__ == "__main__":
    root = tk.Tk()
    app = AutoClickGUI(root, version=APP_VERSION)
    set_app(app)

    def on_close():
        if app.key_listener and app.key_listener.is_alive():
            app.key_listener.stop()
        app._stop()
        root.destroy()
    root.protocol("WM_DELETE_WINDOW", on_close)

    root.mainloop()