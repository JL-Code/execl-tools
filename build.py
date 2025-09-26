#!/usr/bin/env python3
"""
Excel文件拆分工具构建脚本
使用PyInstaller将应用打包成可执行文件
"""

import os
import sys
import shutil
import subprocess
from pathlib import Path

def clean_build():
    """清理之前的构建文件"""
    print("🧹 清理之前的构建文件...")
    
    # 要清理的目录和文件
    clean_paths = [
        'build',
        'dist', 
        '__pycache__',
        'src/__pycache__',
        '*.spec'
    ]
    
    for path in clean_paths:
        if path.endswith('*.spec'):
            # 删除所有.spec文件
            for spec_file in Path('.').glob('*.spec'):
                print(f"  删除: {spec_file}")
                spec_file.unlink()
        else:
            if os.path.exists(path):
                if os.path.isdir(path):
                    print(f"  删除目录: {path}")
                    shutil.rmtree(path)
                else:
                    print(f"  删除文件: {path}")
                    os.remove(path)

def build_exe():
    """构建可执行文件"""
    print("🔨 开始构建可执行文件...")
    
    # PyInstaller命令参数
    cmd = [
        sys.executable, '-m', 'PyInstaller',
        '--onefile',                    # 打包成单个文件
        '--windowed',                   # Windows下隐藏控制台窗口
        '--name=Excel文件拆分工具',        # 可执行文件名称
        '--icon=assets/icon.ico',       # 图标文件（如果存在）
        '--add-data=src:src',           # 添加源代码目录
        '--hidden-import=pandas',       # 确保pandas被包含
        '--hidden-import=openpyxl',     # 确保openpyxl被包含
        '--hidden-import=xlrd',         # 确保xlrd被包含
        '--hidden-import=xlsxwriter',   # 确保xlsxwriter被包含
        '--hidden-import=tkinter',      # 确保tkinter被包含
        '--clean',                      # 清理临时文件
        'src/main.py'                   # 主入口文件
    ]
    
    # 如果没有图标文件，移除图标参数
    if not os.path.exists('assets/icon.ico'):
        cmd = [arg for arg in cmd if not arg.startswith('--icon')]
    
    print(f"执行命令: {' '.join(cmd)}")
    
    try:
        # 执行PyInstaller命令
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        print("✅ 构建成功！")
        print(result.stdout)
        
        # 检查生成的文件
        exe_path = Path('dist/Excel文件拆分工具.exe')
        if exe_path.exists():
            file_size = exe_path.stat().st_size / (1024 * 1024)  # MB
            print(f"📦 生成的可执行文件: {exe_path}")
            print(f"📏 文件大小: {file_size:.1f} MB")
        else:
            print("⚠️  未找到生成的可执行文件")
            
    except subprocess.CalledProcessError as e:
        print("❌ 构建失败！")
        print(f"错误信息: {e}")
        print(f"标准输出: {e.stdout}")
        print(f"错误输出: {e.stderr}")
        return False
    
    return True

def create_installer_script():
    """创建安装脚本"""
    print("📝 创建安装脚本...")
    
    installer_script = """@echo off
echo 正在安装Excel文件拆分工具...

REM 创建程序目录
if not exist "%PROGRAMFILES%\\Excel文件拆分工具" (
    mkdir "%PROGRAMFILES%\\Excel文件拆分工具"
)

REM 复制可执行文件
copy "Excel文件拆分工具.exe" "%PROGRAMFILES%\\Excel文件拆分工具\\"

REM 创建桌面快捷方式
echo 创建桌面快捷方式...
powershell "$WshShell = New-Object -comObject WScript.Shell; $Shortcut = $WshShell.CreateShortcut('%USERPROFILE%\\Desktop\\Excel文件拆分工具.lnk'); $Shortcut.TargetPath = '%PROGRAMFILES%\\Excel文件拆分工具\\Excel文件拆分工具.exe'; $Shortcut.Save()"

echo 安装完成！
pause
"""
    
    with open('dist/install.bat', 'w', encoding='gbk') as f:
        f.write(installer_script)
    
    print("✅ 安装脚本创建完成: dist/install.bat")

def main():
    """主函数"""
    print("🚀 Excel文件拆分工具构建脚本")
    print("=" * 50)
    
    # 检查是否在正确的目录
    if not os.path.exists('src/main.py'):
        print("❌ 错误: 请在项目根目录运行此脚本")
        sys.exit(1)
    
    # 清理构建文件
    clean_build()
    
    # 构建可执行文件
    if build_exe():
        # 创建安装脚本
        create_installer_script()
        
        print("\n🎉 构建完成！")
        print("📁 可执行文件位置: dist/Excel文件拆分工具.exe")
        print("📁 安装脚本位置: dist/install.bat")
        print("\n使用说明:")
        print("1. 直接运行 dist/Excel文件拆分工具.exe")
        print("2. 或者运行 dist/install.bat 进行系统安装")
    else:
        print("❌ 构建失败，请检查错误信息")
        sys.exit(1)

if __name__ == '__main__':
    main()