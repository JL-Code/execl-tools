# Excel File Split Tool Portable Build Script
# This script creates a more portable executable with better compatibility
param(
    [switch]$Clean,
    [switch]$NoVenv
)

# Set console encoding to UTF-8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding = [System.Text.Encoding]::UTF8
$PSDefaultParameterValues['*:Encoding'] = 'utf8'
chcp 65001 | Out-Null
$ErrorActionPreference = "Stop"

function Write-ColorMessage {
    param([string]$Message, [string]$Color = "White")
    Write-Host $Message -ForegroundColor $Color
}

Write-ColorMessage "Excel File Split Tool - Portable Build" "Green"
Write-ColorMessage "=======================================" "Green"

# Check Python installation
try {
    $pythonVersion = python --version 2>&1
    Write-ColorMessage "Python detected: $pythonVersion" "Cyan"
}
catch {
    Write-ColorMessage "Error: Python not found" "Red"
    exit 1
}

# Virtual environment management
if (-not $NoVenv -and (Test-Path "venv")) {
    Write-ColorMessage "Activating virtual environment..." "Yellow"
    try {
        & "venv\Scripts\Activate.ps1"
    }
    catch {
        Write-ColorMessage "Warning: Using system Python" "Yellow"
    }
}

# Clean previous builds
if ($Clean -or (Test-Path "build") -or (Test-Path "dist")) {
    Write-ColorMessage "Cleaning previous build files..." "Yellow"
    @("build", "dist", "src\__pycache__", "__pycache__") | ForEach-Object {
        if (Test-Path $_) { Remove-Item -Recurse -Force $_ }
    }
    Get-ChildItem -Path "." -Filter "*.spec" | Remove-Item -Force
}

# Install/update dependencies
Write-ColorMessage "Ensuring dependencies are installed..." "Yellow"
pip install --upgrade pip
pip install --upgrade pyinstaller pandas openpyxl xlrd xlsxwriter

# Build with maximum compatibility
Write-ColorMessage "Building portable executable..." "Yellow"

$appName = "ExcelFileSplitTool_Portable"

# Set environment for maximum compatibility
$env:PYTHONIOENCODING = "utf-8"
$env:PYTHONUTF8 = "1"
$env:PYTHONHASHSEED = "1"
$env:PYTHONDONTWRITEBYTECODE = "1"

# Clear PyInstaller cache
if (Test-Path "$env:APPDATA\pyinstaller") {
    Remove-Item -Recurse -Force "$env:APPDATA\pyinstaller" -ErrorAction SilentlyContinue
}

# Comprehensive build arguments for maximum compatibility
$buildArgs = @(
    "--onefile"
    "--windowed"
    "--name=$appName"
    "--distpath=dist"
    "--workpath=build"
    "--specpath=."
    
    # All possible hidden imports
    "--hidden-import=pandas"
    "--hidden-import=numpy"
    "--hidden-import=openpyxl"
    "--hidden-import=openpyxl.workbook"
    "--hidden-import=openpyxl.worksheet"
    "--hidden-import=openpyxl.reader.excel"
    "--hidden-import=openpyxl.writer.excel"
    "--hidden-import=openpyxl.styles"
    "--hidden-import=xlrd"
    "--hidden-import=xlsxwriter"
    "--hidden-import=tkinter"
    "--hidden-import=tkinter.filedialog"
    "--hidden-import=tkinter.messagebox"
    "--hidden-import=tkinter.ttk"
    "--hidden-import=tkinter.font"
    
    # System modules
    "--hidden-import=encodings"
    "--hidden-import=encodings.utf_8"
    "--hidden-import=encodings.cp1252"
    "--hidden-import=encodings.latin1"
    "--hidden-import=encodings.ascii"
    "--hidden-import=datetime"
    "--hidden-import=dateutil"
    "--hidden-import=dateutil.parser"
    "--hidden-import=os"
    "--hidden-import=sys"
    "--hidden-import=subprocess"
    "--hidden-import=platform"
    "--hidden-import=pathlib"
    "--hidden-import=shutil"
    
    # Application modules
    "--hidden-import=excel_splitter_gui"
    
    # Data and paths
    "--add-data=src;src"
    "--paths=src"
    
    # Build options for compatibility
    "--clean"
    "--noconfirm"
    "--noupx"
    "--strip"
    "--console"  # Enable console for debugging
    
    # Runtime options
    "--runtime-tmpdir=."
    
    "src\main.py"
)

try {
    Write-ColorMessage "Executing PyInstaller with enhanced compatibility settings..." "Cyan"
    & python -m PyInstaller @buildArgs
    
    if ($LASTEXITCODE -ne 0) {
        throw "PyInstaller failed"
    }
    
    $exePath = "dist\$appName.exe"
    if (Test-Path $exePath) {
        $fileInfo = Get-Item $exePath
        $fileSize = [math]::Round($fileInfo.Length / 1MB, 2)
        
        Write-ColorMessage "Build successful!" "Green"
        Write-ColorMessage "Executable: $exePath" "Green"
        Write-ColorMessage "Size: ${fileSize}MB" "Cyan"
        Write-ColorMessage "This version includes console output for debugging" "Yellow"
        
        # Test the executable
        Write-ColorMessage "Testing executable..." "Yellow"
        try {
            $testProcess = Start-Process -FilePath $exePath -PassThru -WindowStyle Hidden
            Start-Sleep -Seconds 3
            if (-not $testProcess.HasExited) {
                $testProcess.Kill()
                Write-ColorMessage "Executable test passed!" "Green"
            }
        }
        catch {
            Write-ColorMessage "Warning: Could not test executable automatically" "Yellow"
        }
    }
    else {
        throw "Executable not found after build"
    }
}
catch {
    Write-ColorMessage "Build failed: $($_.Exception.Message)" "Red"
    Write-ColorMessage "Check the output above for detailed error information" "Yellow"
    exit 1
}

Write-ColorMessage "`nPortable build completed!" "Green"
Write-ColorMessage "The executable should work on other Windows 11 systems" "Cyan"
Write-ColorMessage "If issues persist, the console version will show error details" "Yellow"