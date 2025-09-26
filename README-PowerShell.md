# PowerShell 脚本使用说明

本项目提供了Windows PowerShell版本的构建和运行脚本，方便在Windows 11系统上使用。

## 文件说明

### build.ps1 - 构建脚本
用于在Windows系统上构建可执行文件的PowerShell脚本。

**功能特性：**
- 自动检测和管理Python虚拟环境
- 智能依赖包安装
- 清理旧的构建文件
- 支持使用现有spec文件或直接构建
- 自动创建桌面快捷方式（可选）
- 构建完成后可选择立即测试

**使用方法：**
```powershell
# 基本构建
.\build.ps1

# 跳过虚拟环境使用系统Python
.\build.ps1 -NoVenv

# 清理构建（强制重新构建）
.\build.ps1 -Clean
```

### run.ps1 - 运行脚本
用于在Windows系统上运行Excel文件拆分工具的PowerShell脚本。

**功能特性：**
- 自动创建和激活Python虚拟环境
- 智能检测和安装依赖包
- 支持开发模式和生产模式
- 优先使用已构建的可执行文件
- 详细的错误诊断和故障排除建议

**使用方法：**
```powershell
# 基本运行
.\run.ps1

# 跳过虚拟环境使用系统Python
.\run.ps1 -NoVenv

# 开发模式（显示详细信息）
.\run.ps1 -Dev

# 强制重新安装依赖包
.\run.ps1 -Install
```

## 系统要求

### 必需组件
- **Windows 11** 或 Windows 10
- **PowerShell 5.1+** （Windows内置）
- **Python 3.8+** （从 [python.org](https://python.org) 下载）

### 推荐配置
- **PowerShell 7+** （PowerShell Core）
- **Python 3.9+** 
- 至少 **2GB** 可用磁盘空间

## 首次使用设置

### 1. 设置PowerShell执行策略
如果遇到执行策略错误，请运行：
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### 2. 验证Python安装
```powershell
python --version
```

### 3. 运行应用程序
```powershell
.\run.ps1
```

## 构建可执行文件

### 快速构建
```powershell
.\build.ps1
```

### 构建选项
```powershell
# 使用系统Python构建（不推荐）
.\build.ps1 -NoVenv

# 清理后重新构建
.\build.ps1 -Clean
```

### 构建输出
构建成功后，将在 `dist\` 目录下生成：
- `Excel文件拆分工具.exe` - 主可执行文件
- 可选的桌面快捷方式

## 故障排除

### 常见问题

**1. 执行策略错误**
```
无法加载文件 *.ps1，因为在此系统上禁止运行脚本
```
**解决方案：**
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

**2. Python未找到**
```
❌ 错误: 未找到Python，请先安装Python
```
**解决方案：**
- 从 [python.org](https://python.org) 下载并安装Python
- 确保安装时勾选"Add Python to PATH"

**3. 虚拟环境激活失败**
```
⚠️ 警告: 虚拟环境激活失败，使用系统Python
```
**解决方案：**
```powershell
# 删除现有虚拟环境重新创建
Remove-Item -Recurse -Force venv
.\run.ps1
```

**4. 依赖包安装失败**
```
❌ 依赖包安装失败
```
**解决方案：**
```powershell
# 升级pip
python -m pip install --upgrade pip
# 重新安装依赖
.\run.ps1 -Install
```

### 开发模式调试
使用开发模式获取详细信息：
```powershell
.\run.ps1 -Dev
```

## 脚本参数说明

### build.ps1 参数
- `-Clean`: 强制清理所有构建文件后重新构建
- `-NoVenv`: 跳过虚拟环境，使用系统Python

### run.ps1 参数
- `-NoVenv`: 跳过虚拟环境，使用系统Python
- `-Dev`: 开发模式，显示详细的系统信息
- `-Install`: 强制重新安装所有依赖包

## 与Linux/macOS版本的差异

| 功能 | Linux/macOS (bash) | Windows (PowerShell) |
|------|-------------------|---------------------|
| 虚拟环境激活 | `source venv/bin/activate` | `venv\Scripts\Activate.ps1` |
| 路径分隔符 | `/` | `\` |
| 可执行文件扩展名 | 无 | `.exe` |
| 桌面集成 | `.desktop` 文件 | `.lnk` 快捷方式 |
| 包管理器 | 系统包管理器 | Windows包管理器 |

## 性能优化建议

1. **使用虚拟环境**：避免与系统Python包冲突
2. **定期清理**：使用 `-Clean` 参数清理构建缓存
3. **SSD存储**：将项目放在SSD上以提高构建速度
4. **关闭杀毒软件实时保护**：构建时临时关闭以提高速度

## 技术支持

如果遇到问题，请：
1. 首先尝试开发模式：`.\run.ps1 -Dev`
2. 检查系统要求是否满足
3. 查看故障排除部分
4. 提供详细的错误信息和系统环境

---

**注意：** 这些PowerShell脚本专为Windows系统设计，在Linux/macOS系统上请使用对应的bash脚本（`build.sh` 和 `run.sh`）。