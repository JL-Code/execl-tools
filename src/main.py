#!/usr/bin/env python3
"""
Main module for execl-tools
Excel文件拆分工具的主入口
"""

import sys
import os

# 检测是否在PyInstaller打包环境中运行
def is_frozen():
    return getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS')

# 获取资源路径（适配PyInstaller）
def get_resource_path(relative_path):
    if is_frozen():
        # PyInstaller打包后的临时目录
        base_path = sys._MEIPASS
    else:
        # 开发环境
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

# 设置模块搜索路径
if is_frozen():
    # PyInstaller环境：模块已经打包在exe中，直接导入
    pass
else:
    # 开发环境：添加 src 目录到 Python 路径
    current_dir = os.path.dirname(os.path.abspath(__file__))
    if current_dir not in sys.path:
        sys.path.insert(0, current_dir)

from excel_splitter_gui import main as gui_main


def main():
    """Main entry point for the application."""
    print("启动Excel文件拆分工具...")
    gui_main()


if __name__ == "__main__":
    main()

