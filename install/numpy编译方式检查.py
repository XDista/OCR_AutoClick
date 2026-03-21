# -*- coding: utf-8 -*-
"""
查看NumPy版本和编译配置信息的脚本
特点：输出长久显示，窗口不会自动关闭
"""
import numpy as np  # 导入NumPy库（别名np是行业通用写法）
import warnings

# 可选：消除pyyaml警告（不影响功能，仅让输出更整洁）
warnings.filterwarnings("ignore", message="Install `pyyaml` for better output")

def print_numpy_info():
    """打印NumPy核心信息"""
    print("=" * 60)
    print("NumPy 版本与编译配置信息")
    print("=" * 60)
    
    # 1. 打印NumPy版本
    print(f"\n1. NumPy版本：{np.__version__}")
    
    # 2. 打印编译配置信息
    print("\n2. 编译配置详情：")
    config_info = np.__config__.show()  # 输出编译配置
    if config_info is None:  # 兼容show()返回None的情况
        pass
    
    print("\n" + "=" * 60)
    print("信息输出完成！")
    print("=" * 60)

if __name__ == "__main__":
    # 主逻辑执行
    try:
        print_numpy_info()
    except ImportError as e:
        print(f"错误：未安装NumPy库！\n报错信息：{e}")
        print("请先执行：pip install numpy")
    except Exception as e:
        print(f"运行出错：{e}")
    
    # 核心：暂停窗口，防止闪退（按任意键才会关闭）
    input("\n\n按【任意键】关闭窗口...")