#!/bin/bash
# Excel文件拆分工具启动脚本

# 检查虚拟环境是否存在
if [ ! -d "venv" ]; then
    echo "创建虚拟环境..."
    python3 -m venv venv
fi

# 激活虚拟环境并安装依赖
echo "激活虚拟环境并检查依赖..."
source venv/bin/activate

# 检查是否需要安装依赖
if ! python -c "import pandas, openpyxl" 2>/dev/null; then
    echo "安装依赖包..."
    pip install -r requirements.txt
fi

# 启动GUI工具
echo "启动Excel文件拆分工具..."
python src/main.py