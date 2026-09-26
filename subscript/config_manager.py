import os
import json
import configparser
from project_paths import MAIN_CONFIG_PATH, TASKS_DIR


def init_main_config():
    config = configparser.ConfigParser()
    if os.path.exists(MAIN_CONFIG_PATH):
        config.read(MAIN_CONFIG_PATH, encoding="utf-8")

    if "GENERAL" not in config:
        config["GENERAL"] = {}
    if "ADBConfig" not in config:
        config["ADBConfig"] = {}
    if "WindowConfig" not in config:
        config["WindowConfig"] = {}
    if "OCRConfig" not in config:
        config["OCRConfig"] = {}

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
        "click_mode": "sendmessage",
        "screenshot_mode": "win32gui",
        "template_match_step": "0.05",
    }

    default_adb_config = {
        "adb_enabled": "1",
        "adb_usage_mode": "1",
        "adb_position": "",
        "adb_device_serial": "",
    }

    default_window_config = {
        "selfgeometry": "1200x600",
        "selfxy_resizable": "1",
        "scrgeometry": "820x460",
        "scrxy_resizable": "1",
        "adbscrgeometry": "700x450",
        "adbscrxy_resizable": "1",
    }

    default_theme_config = {
        "current_theme": "light",
    }

    for key, value in default_general_config.items():
        if key not in config["GENERAL"]:
            config["GENERAL"][key] = value

    for key, value in default_adb_config.items():
        if key not in config["ADBConfig"]:
            config["ADBConfig"][key] = value

    for key, value in default_window_config.items():
        if key not in config["WindowConfig"]:
            config["WindowConfig"][key] = value

    if "Theme" not in config:
        config["Theme"] = {}
    for key, value in default_theme_config.items():
        if key not in config["Theme"]:
            config["Theme"][key] = value

    default_ocr_config = {
        "ocr_language": "ch_sim, en",
        "ocr_device": "cpu",
        "current_ocr_task": "",
    }
    for key, value in default_ocr_config.items():
        if key not in config["OCRConfig"]:
            config["OCRConfig"][key] = value

    with open(MAIN_CONFIG_PATH, "w", encoding="utf-8") as f:
        config.write(f)
    return config


def init_task_config(task_group_name):
    task_config_path = os.path.join(TASKS_DIR, f"{task_group_name}.json")
    default_config = {
        "TASK1": {
            "ignore_occlusion": False,
            "ref_images": [
                {
                    "image": "button1.png",
                    "similarity_threshold": 0.9,
                    "match_times": 1,
                    "run_on_match": True,
                    "actions": [
                        {"type": "click", "params": [100, 200]},
                        {"type": "sleep", "params": [1.0]},
                    ],
                },
                {
                    "image": "button2.png",
                    "similarity_threshold": 0.9,
                    "match_times": 2,
                    "run_on_match": True,
                    "actions": [
                        {"type": "click", "params": [100, 200]},
                    ],
                },
            ],
        },
        "TASK2": {
            "ignore_occlusion": False,
            "ref_images": [
                {
                    "image": "close_button.png",
                    "similarity_threshold": 0.85,
                    "match_times": 1,
                    "run_on_match": True,
                    "actions": [
                        {"type": "click", "params": [100, 200]},
                        {"type": "sleep", "params": [1.0]},
                        {"type": "stop", "params": []},
                    ],
                },
            ],
        },
    }
    if not os.path.exists(task_config_path):
        with open(task_config_path, "w", encoding="utf-8") as f:
            json.dump(default_config, f, ensure_ascii=False, indent=4)
    with open(task_config_path, "r", encoding="utf-8") as f:
        task_config = json.load(f)
    return task_config