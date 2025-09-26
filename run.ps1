# Excel File Split Tool Launcher (Windows PowerShell)
# Usage: .\run.ps1
param(
    [switch]$NoVenv,
    [switch]$Dev,
    [switch]$Install
)

# Set error handling
$ErrorActionPreference = "Stop"

Write-Host "Excel File Split Tool Launcher" -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Green

# Check if running in correct directory
if (-not (Test-Path "src\main.py")) {
    Write-Host "Error: Please run this script from project root directory" -ForegroundColor Red
    exit 1
}

# Check if Python is installed
try {
    $pythonVersion = python --version 2>&1
    Write-Host "Python detected: $pythonVersion" -ForegroundColor Cyan
}
catch {
    Write-Host "Error: Python not found, please install Python first" -ForegroundColor Red
    Write-Host "Tip: Download from https://python.org" -ForegroundColor Yellow
    exit 1
}

# Virtual environment management
if (-not $NoVenv) {
    if (-not (Test-Path "venv")) {
        Write-Host "Creating virtual environment..." -ForegroundColor Yellow
        python -m venv venv
        if ($LASTEXITCODE -ne 0) {
            Write-Host "Failed to create virtual environment" -ForegroundColor Red
            exit 1
        }
    }
    
    Write-Host "Activating virtual environment..." -ForegroundColor Yellow
    try {
        & "venv\Scripts\Activate.ps1"
        if ($LASTEXITCODE -ne 0) {
            throw "Virtual environment activation failed"
        }
    }
    catch {
        Write-Host "Warning: Virtual environment activation failed, using system Python" -ForegroundColor Yellow
        Write-Host "Tip: You may need to run: Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser" -ForegroundColor Yellow
    }
}
else {
    Write-Host "Skipping virtual environment, using system Python" -ForegroundColor Yellow
}

# Check and install dependencies
Write-Host "Checking dependencies..." -ForegroundColor Yellow
$requiredPackages = @("pandas", "openpyxl", "xlrd", "xlsxwriter")
$missingPackages = @()

foreach ($package in $requiredPackages) {
    try {
        python -c "import $package" 2>$null
        if ($LASTEXITCODE -ne 0) {
            $missingPackages += $package
        }
    }
    catch {
        $missingPackages += $package
    }
}

if ($missingPackages.Count -gt 0 -or $Install) {
    Write-Host "Installing/updating dependencies..." -ForegroundColor Yellow
    
    if (Test-Path "requirements.txt") {
        pip install -r requirements.txt
    }
    else {
        # Install basic dependencies if no requirements.txt
        pip install pandas openpyxl xlrd xlsxwriter
    }
    
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Failed to install dependencies" -ForegroundColor Red
        exit 1
    }
    Write-Host "Dependencies installed successfully" -ForegroundColor Green
}
else {
    Write-Host "All dependencies are installed" -ForegroundColor Green
}

# Development mode information
if ($Dev) {
    Write-Host ""
    Write-Host "Development Mode Info:" -ForegroundColor Cyan
    Write-Host "Python Version: $(python --version)" -ForegroundColor White
    Write-Host "Working Directory: $(Get-Location)" -ForegroundColor White
    $venvStatus = if (Test-Path 'venv') { 'Created' } else { 'Not Created' }
    Write-Host "Virtual Environment: $venvStatus" -ForegroundColor White
    
    Write-Host ""
    Write-Host "Installed Packages:" -ForegroundColor Cyan
    pip list | Select-String "pandas|openpyxl|xlrd|xlsxwriter|tkinter"
    Write-Host ""
}

# Launch application
Write-Host "Starting Excel File Split Tool..." -ForegroundColor Green
try {
    # Check for built executable
    $exePath = "dist\Excel文件拆分工具.exe"
    if ((Test-Path $exePath) -and (-not $Dev)) {
        $useExe = Read-Host "Built executable found, use it? (Y/n)"
        if ($useExe -ne "n" -and $useExe -ne "N") {
            Write-Host "Starting executable..." -ForegroundColor Green
            Start-Process -FilePath $exePath
            return
        }
    }
    
    # Run from Python source
    Write-Host "Running from Python source..." -ForegroundColor Green
    python src\main.py
}
catch {
    Write-Host "Launch failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "Troubleshooting suggestions:" -ForegroundColor Yellow
    Write-Host "1. Ensure all dependencies are correctly installed" -ForegroundColor White
    Write-Host "2. Check Python version compatibility (recommended 3.8+)" -ForegroundColor White
    Write-Host "3. Try recreating virtual environment: Remove-Item -Recurse venv; .\run.ps1" -ForegroundColor White
    Write-Host "4. Use development mode for details: .\run.ps1 -Dev" -ForegroundColor White
    exit 1
}

Write-Host "Application started successfully!" -ForegroundColor Green