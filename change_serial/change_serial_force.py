import json
import os
import time
import configparser
import argparse

# ============ 从配置文件读取配置 ============
def load_config():
    """从 serial_config.ini 加载配置"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(script_dir, 'serial_config.ini')
    
    config = configparser.ConfigParser()
    
    if not os.path.exists(config_path):
        print(f"错误：配置文件不存在: {config_path}")
        print("请创建 serial_config.ini 配置文件")
        input("按回车键退出...")
        exit(1)
    
    config.read(config_path, encoding='utf-8')
    
    # 读取 FORCE_MODE
    FORCE_MODE = config.getboolean('Settings', 'FORCE_MODE', fallback=True)
    
    # 读取 FORCE_MODE_SERIAL
    FORCE_MODE_SERIAL = config.get('Settings', 'FORCE_MODE_SERIAL', fallback='R5CW7017LDM')
    
    # 读取批量处理文件列表
    batch_files_str = config.get('BatchFiles', 'Files', fallback='')
    BATCH_JSON_FILES = [line.strip() for line in batch_files_str.split('\n') if line.strip()]
    
    # 读取预设 Serial 值列表
    preset_serials_str = config.get('PresetSerials', 'Serials', fallback='')
    PRESET_SERIALS = [line.strip() for line in preset_serials_str.split('\n') if line.strip()]
    
    return BATCH_JSON_FILES, PRESET_SERIALS, FORCE_MODE, FORCE_MODE_SERIAL

# 加载配置
BATCH_JSON_FILES, PRESET_SERIALS, FORCE_MODE, FORCE_MODE_SERIAL = load_config()

# ================ 命令行参数解析 =================
def parse_args():
    """解析命令行参数"""
    parser = argparse.ArgumentParser(description='修改 JSON 文件中的 Serial 字段')
    
    # 模式选择组
    mode_group = parser.add_mutually_exclusive_group(required=False)
    mode_group.add_argument('-f', '--file', metavar='PATH', help='指定单个 JSON 文件路径')
    mode_group.add_argument('-b', '--batch', action='store_true', help='使用批量模式（处理配置文件中的文件列表）')
    
    # Serial 值参数
    serial_group = parser.add_mutually_exclusive_group(required=False)
    serial_group.add_argument('-s', '--serial', metavar='SERIAL', help='指定新的 Serial 值')
    serial_group.add_argument('-p', '--preset', type=int, metavar='INDEX', help=f'使用预设的 Serial 值（索引 1-{len(PRESET_SERIALS)}）')
    
    # 其他选项
    parser.add_argument('-y', '--yes', action='store_true', help='跳过确认提示（直接执行）')
    parser.add_argument('-c', '--config', metavar='PATH', help='指定配置文件路径（默认：serial_config.ini）')
    
    return parser.parse_args()

# =======================================================

def modify_serial(json_path, new_serial, batch_mode=False):
    """修改 JSON 文件中的 Serial 字段
    
    Args:
        json_path: JSON 文件路径
        new_serial: 新的 Serial 值
        batch_mode: 是否为批量模式（批量模式下错误不终止程序）
    
    Returns:
        bool: 修改是否成功
    """
    try:
        # 读取 JSON 文件
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # 检查 Alas > Emulator > Serial 路径是否存在
        if 'Alas' not in data:
            print(f"\n[{json_path}] 错误：未找到 'Alas' 字段")
            if batch_mode:
                return False
            print("按任意键退出...")
            os.system('pause')
            exit(0)
        
        if 'Emulator' not in data['Alas']:
            print(f"\n[{json_path}] 错误：未找到 'Alas' -> 'Emulator' 字段")
            if batch_mode:
                return False
            print("按任意键退出...")
            os.system('pause')
            exit(0)
        
        if 'Serial' not in data['Alas']['Emulator']:
            print(f"\n[{json_path}] 错误：未找到 'Alas' -> 'Emulator' -> 'Serial' 字段")
            if batch_mode:
                return False
            print("按任意键退出...")
            os.system('pause')
            exit(0)
        
        # 获取当前 Serial 值
        current_serial = data['Alas']['Emulator']['Serial']
        
        print(f"\n[{json_path}]")
        print(f"  当前 Serial 值: {current_serial}")
        print(f"  新的 Serial 值: {new_serial}")
        
        # 批量模式下不需要二次确认
        if not batch_mode:
            # 二次确认
            confirm = input("  确认要修改吗？(y/n): ").strip().lower()
            if confirm != 'y':
                print("  操作已取消")
                return False
        
        # 修改 Serial 值
        data['Alas']['Emulator']['Serial'] = new_serial
        
        # 保存修改
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        
        print(f"  修改成功！Serial 已更新为: {new_serial}")
        return True
        
    except FileNotFoundError:
        print(f"\n[{json_path}] 错误：文件不存在")
        return False
    except json.JSONDecodeError:
        print(f"\n[{json_path}] 错误：不是有效的 JSON 格式")
        return False
    except Exception as e:
        print(f"\n[{json_path}] 发生错误: {str(e)}")
        return False

def batch_modify_serial(new_serial, skip_confirm=False):
    """批量修改多个 JSON 文件的 Serial 字段
    
    Args:
        new_serial: 新的 Serial 值
        skip_confirm: 是否跳过确认提示（默认 False）
    """
    if not BATCH_JSON_FILES:
        print("错误：批量处理列表为空！")
        print("请在配置文件的 BatchFiles 列表中添加要处理的文件路径")
        print("按任意键退出...")
        os.system('pause')
        return
    
    script_dir = os.path.dirname(os.path.abspath(__file__))
    success_count = 0
    fail_count = 0
    
    print(f"\n=== 开始批量处理 ===")
    print(f"共 {len(BATCH_JSON_FILES)} 个文件待处理")
    
    # 强制模式或 skip_confirm=True 时跳过确认，否则需要二次确认
    if not FORCE_MODE and not skip_confirm:
        confirm = input(f"确认要将所有文件的 Serial 修改为 '{new_serial}' 吗？(y/n): ").strip().lower()
        if confirm != 'y':
            print("操作已取消")
            return
    else:
        print(f"【强制模式】将所有文件的 Serial 修改为: {new_serial}")
    
    for json_file in BATCH_JSON_FILES:
        # 处理相对路径
        if not os.path.isabs(json_file):
            json_path = os.path.join(script_dir, json_file)
        else:
            json_path = json_file
        
        if modify_serial(json_path, new_serial, batch_mode=True):
            success_count += 1
        else:
            fail_count += 1
    
    print(f"\n=== 批量处理完成 ===")
    print(f"成功: {success_count} 个文件")
    print(f"失败: {fail_count} 个文件")

def get_serial_input():
    """获取 Serial 值输入，支持预设值选择
    
    Returns:
        str: Serial 值，如果输入无效返回 None
    """
    print("\n请选择 Serial 值：")
    print("0. 手动输入（默认，直接按 Enter）")
    
    # 显示预设值列表
    for i, preset in enumerate(PRESET_SERIALS, 1):
        print(f"{i}. {preset}")
    
    choice = input("请选择 (0-{}，默认手动输入): ".format(len(PRESET_SERIALS))).strip()
    
    # 直接按 Enter 或输入 0，手动输入
    if choice == '' or choice == '0':
        new_serial = input("请输入新的 Serial 值: ").strip()
        if not new_serial:
            print("错误：Serial 值不能为空")
            input("按回车键退出...")
            return None
        return new_serial
    
    # 选择预设值
    try:
        index = int(choice)
        if 1 <= index <= len(PRESET_SERIALS):
            return PRESET_SERIALS[index - 1]
        else:
            print(f"错误：无效的选择，请输入 0-{len(PRESET_SERIALS)} 之间的数字")
            input("按回车键退出...")
            return None
    except ValueError:
        print("错误：请输入有效的数字")
        input("按回车键退出...")
        return None

if __name__ == "__main__":
    script_dir = os.path.dirname(os.path.abspath(__file__))
    args = parse_args()
    
    # 确定操作模式和目标文件
    mode = 'batch' if args.batch else 'single' if args.file else None
    
    # 确定 Serial 值
    new_serial = None
    if args.serial:
        new_serial = args.serial
    elif args.preset:
        index = args.preset - 1
        if 0 <= index < len(PRESET_SERIALS):
            new_serial = PRESET_SERIALS[index]
        else:
            print(f"错误：预设索引 {args.preset} 无效（有效范围 1-{len(PRESET_SERIALS)}）")
            exit(1)
    
    # 强制模式：直接执行批量修改，跳过所有交互
    if FORCE_MODE:
        print("=== 修改 JSON Serial 工具（强制模式） ===")
        print(f"强制模式已启用，将直接修改以下文件:")
        for i, json_file in enumerate(BATCH_JSON_FILES, 1):
            print(f"  {i}. {json_file}")
        
        # 使用强制模式专用的 Serial 值
        new_serial = FORCE_MODE_SERIAL
        print(f"\n使用强制模式 Serial 值: {new_serial}")
        
        batch_modify_serial(new_serial)
        
        # 强制模式下等待5秒后自动退出
        print("\n=== 操作完成 ===")
        for i in range(5, 0, -1):
            print(f"将在 {i} 秒后自动退出...", end='\r')
            time.sleep(1)
        print("\n退出程序")
        exit(0)
    
    # 命令行参数模式
    if mode is not None or new_serial is not None:
        print("=== 修改 JSON Serial 工具（命令行模式） ===")
        
        # 如果模式未指定，默认批量模式
        if mode is None:
            mode = 'batch'
        
        # 如果 Serial 值未指定，提示输入
        if new_serial is None:
            new_serial = get_serial_input()
            if new_serial is None:
                exit(0)
        
        if mode == 'batch':
            # 批量修改模式
            batch_modify_serial(new_serial, skip_confirm=args.yes)
        else:
            # 单文件修改模式
            json_path = args.file
            if not os.path.isabs(json_path):
                json_path = os.path.join(script_dir, json_path)
            
            if not args.yes:
                confirm = input(f"确认要修改 '{json_path}' 的 Serial 为 '{new_serial}' 吗？(y/n): ").strip().lower()
                if confirm != 'y':
                    print("操作已取消")
                    exit(0)
            # 传递 batch_mode=args.yes，避免函数内部再次询问确认
            modify_serial(json_path, new_serial, batch_mode=args.yes)
        
        exit(0)
    
    # 正常交互模式
    print("=== 修改 JSON Serial 工具 ===")
    print("1. 单文件修改")
    print("2. 批量修改（默认，直接按 Enter）")
    
    choice = input("请选择操作模式 (1/2，默认批量修改): ").strip()
    
    # 默认进入批量修改模式
    if choice == '' or choice == '2':
        # 批量修改模式
        new_serial = get_serial_input()
        if new_serial is None:
            exit(0)
        batch_modify_serial(new_serial)
        
    elif choice == '1':
        # 单文件修改模式
        default_json_path = os.path.join(script_dir, 'example.json')
        
        # 获取 JSON 文件路径
        json_path = input(f"请输入 JSON 文件路径（默认: {default_json_path}）: ").strip()
        if not json_path:
            json_path = default_json_path
        
        # 获取新的 Serial 值
        new_serial = get_serial_input()
        if new_serial is None:
            exit(0)
        
        # 执行修改
        modify_serial(json_path, new_serial)
    
    else:
        print("错误：无效的选择")
    
    # 保持窗口打开
    input("\n按回车键退出...")

# ================ 命令行参数用法说明 ================
# 命令行模式：提供参数时进入命令行模式，不显示交互式菜单
#
# 参数说明：
#   -f, --file PATH      指定单个 JSON 文件路径
#   -b, --batch          使用批量模式（处理配置文件中的文件列表）
#   -s, --serial SERIAL  指定新的 Serial 值
#   -p, --preset INDEX   使用预设的 Serial 值（索引 1-N，预设值在配置文件中定义）
#   -y, --yes            跳过确认提示，直接执行修改
#   -c, --config PATH    指定配置文件路径（默认：serial_config.ini）
#
# 使用示例：
#
# 1. 单文件修改（指定Serial值）
#    python change_serial_force.py -f example.json -s 123456
#
# 2. 单文件修改（使用预设值，跳过确认）
#    python change_serial_force.py -f example.json -p 1 -y
#
# 3. 批量修改（指定Serial值，跳过确认）
#    python change_serial_force.py -b -s 123R5CW7017LDM -y
#
# 4. 批量修改（使用预设值第2个）
#    python change_serial_force.py -b -p 2
#
# 5. 仅指定Serial值，使用默认批量模式
#    python change_serial_force.py -s 127.0.0.1:16384
#
# 6. 使用自定义配置文件
#    python change_serial_force.py -c my_config.ini -b -s 123456 -y
#
# 注意：
#   - 模式参数（-f/-b）和 Serial 参数（-s/-p）都是可选的
#   - 如果不提供任何参数，脚本进入交互式菜单模式
#   - FORCE_MODE=True 时，忽略命令行参数，直接执行批量修改
# ========================================================