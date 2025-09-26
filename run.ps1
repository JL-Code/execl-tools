# Excel文件拆分工具启动脚本 (Windows PowerShell版本)
# 使用方法: .\run.ps1

param(
    [switch]$NoVenv,
    [switch]$Dev,
    [switch]$Install
)

# 设置错误处理
$ErrorActionPreference = "Stop"

Write-Host "🚀 Excel文件拆分工具启动脚本" -ForegroundColor Green
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
    Write-Host "💡 提示: 请从 https://python.org 下载并安装Python" -ForegroundColor Yellow
    exit 1
}

# 虚拟环境管理
if (-not $NoVenv) {
    if (-not (Test-Path "venv")) {
        Write-Host "📦 创建虚拟环境..." -ForegroundColor Yellow
        python -m venv venv
        if ($LASTEXITCODE -ne 0) {
            Write-Host "❌ 虚拟环境创建失败" -ForegroundColor Red
            exit 1
        }
    }

    Write-Host "🔧 激活虚拟环境..." -ForegroundColor Yellow
    try {
        & "venv\Scripts\Activate.ps1"
        if ($LASTEXITCODE -ne 0) {
            throw "虚拟环境激活失败"
        }
    } catch {
        Write-Host "⚠️  警告: 虚拟环境激活失败，使用系统Python" -ForegroundColor Yellow
        Write-Host "💡 提示: 可能需要执行 Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser" -ForegroundColor Yellow
    }
} else {
    Write-Host "⚠️  跳过虚拟环境，使用系统Python" -ForegroundColor Yellow
}

# 检查并安装依赖
Write-Host "📋 检查依赖包..." -ForegroundColor Yellow

$requiredPackages = @("pandas", "openpyxl", "xlrd", "xlsxwriter")
$missingPackages = @()

foreach ($package in $requiredPackages) {
    try {
        python -c "import $package" 2>$null
        if ($LASTEXITCODE -ne 0) {
            $missingPackages += $package
        }
    } catch {
        $missingPackages += $package
    }
}

if ($missingPackages.Count -gt 0 -or $Install) {
    Write-Host "📥 安装/更新依赖包..." -ForegroundColor Yellow
    
    if (Test-Path "requirements.txt") {
        pip install -r requirements.txt
    } else {
        # 如果没有requirements.txt，安装基本依赖
        pip install pandas openpyxl xlrd xlsxwriter
    }
    
    if ($LASTEXITCODE -ne 0) {
        Write-Host "❌ 依赖包安装失败" -ForegroundColor Red
        exit 1
    }
    Write-Host "✅ 依赖包安装完成" -ForegroundColor Green
} else {
    Write-Host "✅ 所有依赖包已安装" -ForegroundColor Green
}

# 开发模式信息
if ($Dev) {
    Write-Host ""
    Write-Host "🔧 开发模式信息:" -ForegroundColor Cyan
    Write-Host "Python版本: $(python --version)" -ForegroundColor White
    Write-Host "工作目录: $(Get-Location)" -ForegroundColor White
    Write-Host "虚拟环境: $(if (Test-Path 'venv') { '已创建' } else { '未创建' })" -ForegroundColor White
    
    Write-Host ""
    Write-Host "📦 已安装的包:" -ForegroundColor Cyan
    pip list | Select-String "pandas|openpyxl|xlrd|xlsxwriter|tkinter"
    Write-Host ""
}

# 启动应用程序
Write-Host "🎯 启动Excel文件拆分工具..." -ForegroundColor Green

try {
    # 检查是否有构建好的可执行文件
    $exePath = "dist\Excel文件拆分工具.exe"
    if ((Test-Path $exePath) -and (-not $Dev)) {
        $useExe = Read-Host "检测到已构建的可执行文件，是否使用？(Y/n)"
        if ($useExe -ne "n" -and $useExe -ne "N") {
            Write-Host "🚀 启动可执行文件..." -ForegroundColor Green
            Start-Process -FilePath $exePath
            return
        }
    }
    
    # 使用Python源码运行
    Write-Host "🐍 使用Python源码运行..." -ForegroundColor Green
    python src\main.py
    
} catch {
    Write-Host "❌ 启动失败: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "🔍 故障排除建议:" -ForegroundColor Yellow
    Write-Host "1. 确保所有依赖包已正确安装" -ForegroundColor White
    Write-Host "2. 检查Python版本是否兼容 (推荐3.8+)" -ForegroundColor White
    Write-Host "3. 尝试重新创建虚拟环境: Remove-Item -Recurse venv; .\run.ps1" -ForegroundColor White
    Write-Host "4. 使用开发模式查看详细信息: .\run.ps1 -Dev" -ForegroundColor White
    exit 1
}

Write-Host "✨ 应用程序已启动！" -ForegroundColor Green