# Excel文件拆分工具构建脚本 (Windows PowerShell版本)
# 使用方法: .\build.ps1

param(
    [switch]$Clean,
    [switch]$NoVenv
)

# 设置错误处理
$ErrorActionPreference = "Stop"

Write-Host "🚀 Excel文件拆分工具构建脚本" -ForegroundColor Green
Write-Host "==================================================" -ForegroundColor Green

# 检查是否在正确的目录
if (-not (Test-Path "src\main.py")) {
    Write-Host "❌ 错误: 请在项目根目录运行此脚本" -ForegroundColor Red
    exit 1
}

# 检查Python是否安装
try {
    $pythonVersion = python --version 2>&1
    Write-Host "🐍 检测到Python: $pythonVersion" -ForegroundColor Cyan
} catch {
    Write-Host "❌ 错误: 未找到Python，请先安装Python" -ForegroundColor Red
    exit 1
}

# 虚拟环境管理
if (-not $NoVenv) {
    if (Test-Path "venv") {
        Write-Host "🔧 激活虚拟环境..." -ForegroundColor Yellow
        & "venv\Scripts\Activate.ps1"
        if ($LASTEXITCODE -ne 0) {
            Write-Host "⚠️  警告: 虚拟环境激活失败，使用系统Python" -ForegroundColor Yellow
        }
    } else {
        Write-Host "📦 创建虚拟环境..." -ForegroundColor Yellow
        python -m venv venv
        & "venv\Scripts\Activate.ps1"
        Write-Host "📥 安装依赖包..." -ForegroundColor Yellow
        pip install -r requirements.txt
    }
} else {
    Write-Host "⚠️  跳过虚拟环境，使用系统Python" -ForegroundColor Yellow
}

# 检查PyInstaller
try {
    pyinstaller --version | Out-Null
} catch {
    Write-Host "📥 安装PyInstaller..." -ForegroundColor Yellow
    pip install pyinstaller
}

# 清理之前的构建文件
Write-Host "🧹 清理之前的构建文件..." -ForegroundColor Yellow
if (Test-Path "build") { Remove-Item -Recurse -Force "build" }
if (Test-Path "dist") { Remove-Item -Recurse -Force "dist" }
if (Test-Path "src\__pycache__") { Remove-Item -Recurse -Force "src\__pycache__" }
if (Test-Path "__pycache__") { Remove-Item -Recurse -Force "__pycache__" }

# 构建可执行文件
Write-Host "🔨 开始构建Windows可执行文件..." -ForegroundColor Yellow

# 检查是否存在spec文件
if (Test-Path "excel_splitter.spec") {
    Write-Host "📋 使用现有的spec文件构建..." -ForegroundColor Cyan
    pyinstaller excel_splitter.spec --clean
} else {
    Write-Host "📋 使用PyInstaller直接构建..." -ForegroundColor Cyan
    pyinstaller --onefile --windowed --name="Excel文件拆分工具" `
                --hidden-import=pandas --hidden-import=openpyxl `
                --hidden-import=xlrd --hidden-import=xlsxwriter `
                --hidden-import=tkinter --clean src\main.py
}

# 检查构建结果
$exeName = "Excel文件拆分工具.exe"
$exePath = "dist\$exeName"

if (Test-Path $exePath) {
    Write-Host "✅ Windows可执行文件构建成功！" -ForegroundColor Green
    Write-Host "📦 可执行文件位置: $exePath" -ForegroundColor Green
    
    # 获取文件大小
    $fileSize = [math]::Round((Get-Item $exePath).Length / 1MB, 2)
    Write-Host "📊 文件大小: ${fileSize}MB" -ForegroundColor Cyan
    
    # 创建快捷方式到桌面（可选）
    $createShortcut = Read-Host "是否创建桌面快捷方式？(y/N)"
    if ($createShortcut -eq "y" -or $createShortcut -eq "Y") {
        $desktopPath = [Environment]::GetFolderPath("Desktop")
        $shortcutPath = "$desktopPath\Excel文件拆分工具.lnk"
        $targetPath = (Resolve-Path $exePath).Path
        
        $WshShell = New-Object -comObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut($shortcutPath)
        $Shortcut.TargetPath = $targetPath
        $Shortcut.WorkingDirectory = (Get-Location).Path
        $Shortcut.Description = "Excel文件拆分工具"
        $Shortcut.Save()
        
        Write-Host "✅ 桌面快捷方式创建完成: $shortcutPath" -ForegroundColor Green
    }
    
} else {
    Write-Host "❌ 可执行文件构建失败" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "🎉 构建完成！" -ForegroundColor Green
Write-Host "使用说明:" -ForegroundColor Cyan
Write-Host "1. 双击运行 $exePath" -ForegroundColor White
Write-Host "2. 或在PowerShell中运行: & '$exePath'" -ForegroundColor White

# 询问是否立即测试
$testNow = Read-Host "是否立即测试运行？(y/N)"
if ($testNow -eq "y" -or $testNow -eq "Y") {
    Write-Host "🧪 启动应用程序测试..." -ForegroundColor Yellow
    Start-Process -FilePath $exePath
}

Write-Host "✨ 构建脚本执行完成！" -ForegroundColor Green