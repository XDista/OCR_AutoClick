import os
import json
import sys

TARGET_DIR = "."

def list_json_files():
    """列出目标目录下所有JSON文件并排序"""
    json_files = sorted([f for f in os.listdir(TARGET_DIR) if f.endswith('.json')])
    if not json_files:
        print("当前目录没有JSON文件")
        return None
    print("\n当前目录的JSON文件列表：")
    for i, file in enumerate(json_files, 1):
        print(f"  {i}. {file}")
    return json_files

def get_all_paths(data, prefix="", max_depth=0, current_depth=1):
    """递归获取所有节点的点分隔路径（包含 dict 节点和叶子节点）

    Args:
        data: JSON 对象（dict）
        prefix: 当前路径前缀（递归用）
        max_depth: 最大展开层级（0=无限），超过此深度的节点不继续递归
        current_depth: 当前深度（根节点的子节点深度为1）

    Returns:
        list[str]: 所有节点的点分隔路径列表
    """
    paths = []
    for key, value in data.items():
        full_key = f"{prefix}.{key}" if prefix else key
        paths.append(full_key)  # 所有节点都列出（含叶子）
        if isinstance(value, dict):
            if max_depth == 0 or current_depth < max_depth:
                sub_paths = get_all_paths(value, full_key, max_depth, current_depth + 1)
                paths.extend(sub_paths)
    return paths

def get_nested_value(data, path):
    """根据点分隔路径（如 "A.B.C"）从嵌套 dict 中获取值"""
    keys = path.split('.')
    d = data
    for key in keys:
        d = d[key]
    return d

def set_nested_value(data, path, value):
    """根据点分隔路径设置嵌套 dict 的值，自动创建缺失的中间节点"""
    keys = path.split('.')
    d = data
    for key in keys[:-1]:
        if key not in d:
            d[key] = {}
        d = d[key]
    d[keys[-1]] = value

def format_hierarchical_numbers(paths):
    """为路径列表生成层级编号（如 1, 1.1, 1.2, 2, 2.1, 2.2.1 ...）

    Args:
        paths: 点分隔路径列表（DFS顺序）

    Returns:
        list[str]: 与 paths 等长的层级编号字符串列表
    """
    counters = []  # 各层级的当前计数，索引0对应第1层
    numbers = []
    for path in paths:
        depth = path.count('.') + 1
        # 调整计数器长度
        while len(counters) < depth:
            counters.append(0)
        # 剥离多余层级
        counters = counters[:depth]
        # 当前层计数器+1
        counters[-1] += 1
        numbers.append('.'.join(str(c) for c in counters))
    return numbers

def select_nodes(prompt, content_keys):
    """选择要同步的节点：支持序号或完整路径名，可混用

    Args:
        prompt: 提示文字
        content_keys: 所有可选的节点路径列表

    Returns:
        list[str]: 选中的节点路径列表
    """
    max_num = len(content_keys)
    while True:
        user_input = input(f"\n{prompt}")
        raw_items = user_input.replace(',', ' ').split()
        items = [t.strip() for t in raw_items if t.strip()]
        if not items:
            print("输入不能为空，请重新输入")
            continue

        selected = []
        invalid = []
        for item in items:
            # 尝试作为序号解析
            try:
                idx = int(item) - 1
                if 0 <= idx < max_num:
                    key = content_keys[idx]
                    if key not in selected:
                        selected.append(key)
                    continue
                else:
                    invalid.append(f"序号{idx + 1}（超出范围1-{max_num}）")
                    continue
            except ValueError:
                pass

            # 尝试作为路径名精确匹配
            if item in content_keys:
                if item not in selected:
                    selected.append(item)
            else:
                invalid.append(f"'{item}'（路径名不存在）")

        if invalid:
            print(f"无效输入: {'; '.join(invalid)}")
            continue

        return selected

def select_files(prompt, json_files, allow_multiple=False):
    """选择文件：支持序号或完整文件名，可混用

    Args:
        prompt: 提示文字
        json_files: 文件名列表
        allow_multiple: 是否允许多选

    Returns:
        list[str]: 选中的文件名列表
    """
    max_num = len(json_files)
    while True:
        user_input = input(f"\n{prompt}")
        raw_items = user_input.replace(',', ' ').split()
        items = [t.strip() for t in raw_items if t.strip()]
        if not items:
            print("输入不能为空，请重新输入")
            continue

        if not allow_multiple and len(items) > 1:
            print("只能输入一个文件，请重新输入")
            continue

        selected = []
        invalid = []
        for item in items:
            # 尝试作为序号解析
            try:
                idx = int(item) - 1
                if 0 <= idx < max_num:
                    fname = json_files[idx]
                    if fname not in selected:
                        selected.append(fname)
                    continue
                else:
                    invalid.append(f"序号{idx + 1}（超出范围1-{max_num}）")
                    continue
            except ValueError:
                pass

            # 尝试作为文件名精确匹配
            if item in json_files:
                if item not in selected:
                    selected.append(item)
            else:
                invalid.append(f"'{item}'（文件不存在）")

        if invalid:
            print(f"无效输入: {'; '.join(invalid)}")
            continue

        return selected

def sync_json_content():
    """主同步流程（支持任意层级）"""
    while True:
        # 步骤1: 列出所有JSON文件，选择模板文件
        json_files = list_json_files()
        if not json_files:
            return

        template_file = select_files(
            "请输入作为模板的JSON文件（支持序号或文件名）: ",
            json_files
        )[0]

        # 读取模板文件
        try:
            template_path = os.path.join(TARGET_DIR, template_file)
            with open(template_path, 'r', encoding='utf-8') as f:
                template_data = json.load(f)
        except Exception as e:
            print(f"读取模板文件失败: {e}")
            continue

        # 步骤2: 询问展开层级数
        while True:
            depth_input = input("\n请输入本次展开层级数（0=展开全部，默认0）: ").strip()
            if depth_input == "":
                max_depth = 0
                break
            try:
                max_depth = int(depth_input)
                if max_depth >= 0:
                    break
                print("请输入 >=0 的整数")
            except ValueError:
                print("请输入有效整数")

        # 步骤3: 列出所有可同步的节点路径
        content_keys = get_all_paths(template_data, max_depth=max_depth)
        if not content_keys:
            print("模板JSON没有可同步的节点")
            continue

        depth_label = f"全部（{max(content_keys, key=lambda k: k.count('.')).count('.') + 1}层）" if max_depth == 0 else f"{max_depth}层"
        print(f"\n模板文件 '{template_file}' 可同步节点列表（展开{depth_label}）：")
        hierarchical_nums = format_hierarchical_numbers(content_keys)
        for i, (key, num) in enumerate(zip(content_keys, hierarchical_nums), 1):
            depth = key.count('.') + 1
            indent = "  " * (depth - 1)
            print(f"  {i:>3}. [{num}] {indent}{key}")

        # 步骤4: 用户选择要同步的内容（支持序号或路径名混输）
        selected_keys = select_nodes(
            "请输入要同步的节点（支持序号、完整路径名，空格/逗号分隔）: ",
            content_keys
        )

        # 步骤5: 再次列出所有JSON文件，选择目标文件
        print("\n再次列出JSON文件列表：")
        for i, file in enumerate(json_files, 1):
            print(f"  {i}. {file}")

        target_files = select_files(
            "请输入待修改的JSON文件（支持序号或文件名，空格/逗号分隔）: ",
            json_files,
            allow_multiple=True
        )

        # 步骤6: 二次确认并列出修改项
        print("\n" + "=" * 50)
        print("修改项确认")
        print("=" * 50)
        print(f"模板文件: {template_file}")
        print(f"待同步节点 ({len(selected_keys)} 项):")
        for key in selected_keys:
            idx = content_keys.index(key)
            num = hierarchical_nums[idx]
            print(f"  - [{num}] {key}")
        print(f"同步到 ({len(target_files)} 个):")
        for target in target_files:
            if target == template_file:
                print(f"  - {target} (将被跳过)")
            else:
                print(f"  - {target}")
        print("=" * 50)

        confirm = input("确认执行以上修改？(y/n): ").strip().lower()
        if confirm != 'y' and confirm != 'yes':
            print("取消操作，返回初始状态")
            continue

        # 步骤7: 同步内容
        print("\n开始同步内容...")
        for target_file in target_files:
            if target_file == template_file:
                print(f"  跳过模板文件本身: {target_file}")
                continue

            try:
                target_path = os.path.join(TARGET_DIR, target_file)
                with open(target_path, 'r', encoding='utf-8') as f:
                    target_data = json.load(f)

                # 通用同步：根据路径深度设置嵌套值
                for key in selected_keys:
                    try:
                        template_value = get_nested_value(template_data, key)
                        set_nested_value(target_data, key, template_value)
                        print(f"    已同步: {key}")
                    except KeyError as ke:
                        print(f"    ⚠ 跳过: {key}（模板中路径不存在: {ke}）")
                    except Exception as ex:
                        print(f"    ⚠ 跳过: {key}（错误: {ex}）")

                # 保存修改后的文件
                with open(target_path, 'w', encoding='utf-8') as f:
                    json.dump(target_data, f, ensure_ascii=False, indent=2)

                print(f"  ✓ {target_file} 同步完成")
            except Exception as e:
                print(f"  ✗ {target_file} 同步失败: {e}")

        # 步骤8: 询问是否回到初始状态
        print("\n同步完成！")
        choice = input("是否回到初始状态继续操作？(y/n): ").strip().lower()
        if choice != 'y' and choice != 'yes':
            print("退出脚本")
            break

if __name__ == "__main__":
    if len(sys.argv) > 1:
        TARGET_DIR = sys.argv[1]
    if not os.path.isdir(TARGET_DIR):
        print(f"错误：目标目录不存在或无效：{TARGET_DIR}")
        input("按回车键退出...")
        sys.exit(1)
    print("=== JSON内容同步工具（支持任意层级） ===")
    print(f"目标目录：{os.path.abspath(TARGET_DIR)}")
    print("流程：选择模板文件 -> 设置展开层级 -> 选择需要同步的节点 -> 选择目标文件 -> 确认修改项 -> 同步")
    print("=" * 40)
    sync_json_content()