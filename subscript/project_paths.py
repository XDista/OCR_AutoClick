import os
import sys

# 基础路径
if getattr(sys, 'frozen', False):
    BASE_DIR = os.getcwd()
else:
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

MAIN_CONFIG_PATH = os.path.join(BASE_DIR, "config.ini")
REFS_DIR = os.path.join(BASE_DIR, "refs")
TASKS_DIR = os.path.join(BASE_DIR, "tasks")

# 确保必要目录存在
for _dir in [REFS_DIR, TASKS_DIR, os.path.join(REFS_DIR, "subdir")]:
    if not os.path.exists(_dir):
        os.makedirs(_dir)