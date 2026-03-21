import configparser
import json
import os
import sys

def get_next_path_or_exit():
    """
    等待用户输入下一个INI文件路径或退出。
    输入为空或输入'exit'/'quit'则退出。
    返回路径字符串或None（表示退出）
    """
    try:
        import msvcrt
        # 只在Windows上提示ESC键的用法
        path_input = input("\n请输入下一个INI文件路径（或按Ctrl+C退出）：").strip('"').strip()
    except KeyboardInterrupt:
        # 捕捉Ctrl+C
        print("\n已退出")
        return None
    except:
        # 其他环境降级处理
        path_input = input("\n请输入下一个INI文件路径（输入'exit'退出）：").strip('"').strip()
    
    # 检查退出命令
    if path_input.lower() in ('exit', 'quit', 'q'):
        print("已退出")
        return None
    
    return path_input if path_input else ""

def parse_ini_actions(actions_str):
    actions = []
    param_split_actions = ["click", "sleep", "press", "swipe", "goto_task", "taskcall"]
    full_param_actions = ["shutdown", "adbcall", "serial_adbcall"]
    all_known_actions = param_split_actions + full_param_actions + ["stop"]

    cleaned = actions_str.replace("\r", "").replace("\n", "").replace("\t", " ").strip()
    items = [item.strip() for item in cleaned.split(";") if item.strip()]

    for item in items:
        if item.startswith("#"):
            continue
        if ":" not in item:
            action_type = item.lower()
            if action_type in all_known_actions:
                actions.append({"type": action_type, "params": []})
            continue

        action_type, param_str = item.split(":", 1)
        action_type = action_type.strip().lower()
        param_str = param_str.strip()

        params = []
        if action_type in param_split_actions:
            parts = [x.strip() for x in param_str.split(",") if x.strip()]
            for p in parts:
                if p.startswith("(") and p.endswith(")"):
                    p = p[1:-1]
                try:
                    if "." in p:
                        params.append(float(p))
                    else:
                        params.append(int(p))
                except ValueError:
                    params.append(p)
        elif action_type in full_param_actions:
            params = [param_str] if param_str else []

        if action_type in all_known_actions:
            actions.append({"type": action_type, "params": params})
    return actions

def ini_to_new_json(ini_file_path):
    if not os.path.exists(ini_file_path):
        print("❌ 文件不存在：", ini_file_path)
        return False
    if not ini_file_path.lower().endswith(".ini"):
        print("❌ 请输入INI文件路径")
        return False

    config = configparser.ConfigParser()
    try:
        config.read(ini_file_path, encoding="utf-8")
    except Exception as e:
        print("❌ 读取INI失败：", str(e))
        return False

    new_config = {}
    for section in config.sections():
        task = config[section]
        image = task.get("ref_image", "").strip()
        similarity = float(task.get("similarity_threshold", 0.9))
        match_times = int(task.get("match_times", 1))
        if match_times < 1:
            match_times = 1
        actions_str = task.get("actions", "")
        actions = parse_ini_actions(actions_str)
        
        # 读取并转换 ignore_occlusion（支持字符串 "True"/"False" 或 "1"/"0"）
        ignore_occlusion_str = task.get("ignore_occlusion", "False").strip().lower()
        ignore_occlusion = ignore_occlusion_str in ("true", "1", "yes")

        new_config[section] = {
            "ignore_occlusion": ignore_occlusion,
            "ref_images": [
                {
                    "image": image,
                    "similarity_threshold": similarity,
                    "match_times": match_times,
                    "actions": actions
                }
            ]
        }

    out_path = os.path.splitext(ini_file_path)[0] + ".json"
    try:
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(new_config, f, ensure_ascii=False, indent=4)
    except Exception as e:
        print("❌ 写入JSON失败：", str(e))
        return False

    print("✅ 转换完成:", out_path)
    return True

def main():
    # 首次获取路径
    if len(sys.argv) >= 2:
        ini_path = sys.argv[1].strip('"')
    else:
        ini_path = input("请输入INI文件路径（如 D:\\tasks\\test_all.ini）：").strip('"').strip()

    # 循环处理文件转换
    while ini_path:
        success = ini_to_new_json(ini_path)
        
        # 获取下一个路径或退出
        ini_path = get_next_path_or_exit()
        if ini_path is None:
            break

if __name__ == "__main__":
    main()