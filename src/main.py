#!/usr/bin/env python3
"""
Main module for execl-tools
Excel文件拆分工具的主入口
"""

from excel_splitter_gui import main as gui_main


def main():
    """Main entry point for the application."""
    print("启动Excel文件拆分工具...")
    gui_main()


if __name__ == "__main__":
    main()

