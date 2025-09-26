# Excel文件拆分工具

一个基于Python 3和Tkinter的GUI工具，用于将大Excel文件拆分为多个小Excel文件。

## 功能特性

- 🎯 **简单易用的GUI界面** - 直观的图形用户界面，无需命令行操作
- 📊 **支持多种Excel格式** - 支持.xlsx和.xls格式文件
- ⚡ **自定义拆分行数** - 可以自由设置每个小文件包含的行数
- 📁 **智能文件命名** - 自动按照"原文件名_序号"格式命名拆分后的文件
- 📈 **实时进度显示** - 显示拆分进度和详细信息
- 🔍 **文件信息预览** - 自动分析并显示Excel文件的行数、列数等信息
- ⚠️ **错误处理** - 完善的错误提示和异常处理机制

## 使用示例

假设你有一个包含1000行数据的Excel文件，设置每个小文件包含50行，工具将自动拆分为20个文件：
- 原文件：`数据表.xlsx`
- 拆分后：`数据表_001.xlsx`, `数据表_002.xlsx`, ..., `数据表_020.xlsx`

## 项目结构

```
execl-tools/
├── src/                           # 源代码目录
│   ├── __init__.py               # 包初始化文件
│   ├── main.py                   # 主程序入口
│   └── excel_splitter_gui.py     # GUI界面和拆分逻辑
├── tests/                         # 测试文件目录
├── docs/                          # 文档目录
├── venv/                          # 虚拟环境（运行后自动创建）
├── requirements.txt               # 项目依赖
├── run.sh                        # 启动脚本（推荐使用）
├── .gitignore                    # Git忽略文件
├── README.md                     # 项目说明文档
└── 副本样例数据.xls              # 示例数据文件
```

## 快速开始

### 方法一：使用启动脚本（推荐）

```bash
# 克隆或下载项目到本地
cd execl-tools

# 运行启动脚本（会自动处理环境和依赖）
./run.sh
```

### 方法二：手动安装

1. **克隆项目**
   ```bash
   git clone <repository-url>
   cd execl-tools
   ```

2. **创建虚拟环境**
   ```bash
   python3 -m venv venv
   source venv/bin/activate  # macOS/Linux
   # 或者在Windows上: venv\Scripts\activate
   ```

3. **安装依赖**
   ```bash
   pip install -r requirements.txt
   ```

4. **运行程序**
   ```bash
   python src/main.py
   ```

## 使用说明

1. **启动程序** - 运行后会打开GUI界面
2. **选择Excel文件** - 点击"浏览"按钮选择要拆分的Excel文件
3. **设置行数** - 在"每个小文件行数"输入框中输入期望的行数（默认50行）
4. **选择输出目录** - 选择拆分后文件的保存位置（默认为原文件所在目录）
5. **开始拆分** - 点击"开始拆分"按钮，程序会显示进度和详细信息
6. **完成** - 拆分完成后会显示成功提示和输出目录

## 依赖包

- `pandas>=1.5.0` - 数据处理和Excel文件读写
- `openpyxl>=3.0.0` - Excel文件格式支持
- `tkinter-tooltip>=2.0.0` - GUI工具提示功能

## 系统要求

- Python 3.7+
- macOS/Linux/Windows
- 支持GUI的环境（桌面环境）

## 注意事项

- 确保有足够的磁盘空间存储拆分后的文件
- 大文件拆分可能需要一些时间，请耐心等待
- 建议在拆分前备份原始文件
- 如果遇到权限问题，请确保对输出目录有写入权限

## 开发

如果你想参与开发或自定义功能：

```bash
# 安装开发依赖（可选）
pip install pytest black flake8 mypy

# 运行测试
python -m pytest tests/

# 代码格式化
black src/

# 代码检查
flake8 src/
```

## 许可证

本项目采用MIT许可证，详见LICENSE文件。

## 贡献

欢迎提交Issue和Pull Request来改进这个工具！

### Running Tests
```bash
# Install test dependencies
pip install pytest

# Run tests
pytest tests/
```

### Code Formatting
```bash
# Install formatting tools
pip install black flake8

# Format code
black src/ tests/

# Lint code
flake8 src/ tests/
```

## License

[Add your license here]

## Contributing

[Add contribution guidelines here]

