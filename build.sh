#!/bin/bash
# Excel文件拆分工具构建脚本 (macOS/Linux版本)

set -e  # 遇到错误立即退出

echo "🚀 Excel文件拆分工具构建脚本"
echo "=================================================="

# 检查是否在正确的目录
if [ ! -f "src/main.py" ]; then
    echo "❌ 错误: 请在项目根目录运行此脚本"
    exit 1
fi

# 激活虚拟环境
if [ -d "venv" ]; then
    echo "🔧 激活虚拟环境..."
    source venv/bin/activate
else
    echo "⚠️  警告: 未找到虚拟环境，使用系统Python"
fi

# 清理之前的构建文件
echo "🧹 清理之前的构建文件..."
rm -rf build dist src/__pycache__ __pycache__

# 构建可执行文件
echo "🔨 开始构建可执行文件..."

if [[ "$OSTYPE" == "darwin"* ]]; then
    # macOS
    echo "📱 检测到macOS系统，构建.app应用程序..."
    pyinstaller excel_splitter.spec
    
    if [ -f "dist/excel_splitter.app/Contents/MacOS/excel_splitter" ]; then
        echo "✅ macOS应用程序构建成功！"
        echo "📦 应用程序位置: dist/excel_splitter.app"
        
        # 创建DMG安装包
        echo "📦 创建DMG安装包..."
        hdiutil create -volname "Excel文件拆分工具" -srcfolder "dist/excel_splitter.app" -ov -format UDZO "dist/Excel文件拆分工具.dmg"
        echo "✅ DMG安装包创建完成: dist/Excel文件拆分工具.dmg"
    else
        echo "❌ 应用程序构建失败"
        exit 1
    fi
else
    # Linux
    echo "🐧 检测到Linux系统，构建可执行文件..."
    pyinstaller --onefile --windowed --name="Excel文件拆分工具" \
                --hidden-import=pandas --hidden-import=openpyxl \
                --hidden-import=xlrd --hidden-import=xlsxwriter \
                --hidden-import=tkinter --clean src/main.py
    
    if [ -f "dist/Excel文件拆分工具" ]; then
        echo "✅ Linux可执行文件构建成功！"
        echo "📦 可执行文件位置: dist/Excel文件拆分工具"
        
        # 设置执行权限
        chmod +x "dist/Excel文件拆分工具"
        
        # 创建桌面文件
        cat > "dist/excel-splitter.desktop" << EOF
[Desktop Entry]
Version=1.0
Type=Application
Name=Excel文件拆分工具
Comment=Excel文件拆分工具
Exec=$(pwd)/dist/Excel文件拆分工具
Icon=applications-office
Terminal=false
Categories=Office;
EOF
        echo "✅ 桌面文件创建完成: dist/excel-splitter.desktop"
    else
        echo "❌ 可执行文件构建失败"
        exit 1
    fi
fi

echo ""
echo "🎉 构建完成！"
echo "使用说明:"
if [[ "$OSTYPE" == "darwin"* ]]; then
    echo "1. 双击运行 dist/Excel文件拆分工具.app"
    echo "2. 或安装 dist/Excel文件拆分工具.dmg"
else
    echo "1. 运行 ./dist/Excel文件拆分工具"
    echo "2. 或复制 dist/excel-splitter.desktop 到 ~/.local/share/applications/"
fi