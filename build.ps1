# Excel File Split Tool Build Script (Windows PowerShell)
# Usage: .\build.ps1
param(
    [switch]$Clean,
    [switch]$NoVenv,
    [switch]$CreateInstaller,
    [switch]$Dev
)

# Set console encoding to UTF-8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding = [System.Text.Encoding]::UTF8

# Set PowerShell output encoding
$PSDefaultParameterValues['*:Encoding'] = 'utf8'

# Change console code page to UTF-8
chcp 65001 | Out-Null

# Set error handling
$ErrorActionPreference = "Stop"

function Write-ColorMessage {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    Write-Host $Message -ForegroundColor $Color
}

function Test-PythonModule {
    param([string]$ModuleName)
    try {
        python -c "import $ModuleName" 2>$null
        return $LASTEXITCODE -eq 0
    }
    catch {
        return $false
    }
}

function New-DesktopShortcut {
    param(
        [string]$TargetPath,
        [string]$ShortcutName
    )
    
    $desktopPath = [Environment]::GetFolderPath("Desktop")
    $shortcutPath = Join-Path $desktopPath "$ShortcutName.lnk"
    
    try {
        $WshShell = New-Object -comObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut($shortcutPath)
        $Shortcut.TargetPath = (Resolve-Path $TargetPath).Path
        $Shortcut.WorkingDirectory = (Get-Location).Path
        $Shortcut.Description = $ShortcutName
        $Shortcut.Save()
        
        Write-ColorMessage "Desktop shortcut created: $shortcutPath" "Green"
        return $true
    }
    catch {
        Write-ColorMessage "Failed to create desktop shortcut: $($_.Exception.Message)" "Red"
        return $false
    }
}

function New-InstallerScript {
    Write-ColorMessage "Creating installer scripts..." "Yellow"
    
    # Create installer script with proper encoding
    $installerScript = @'
@echo off
chcp 65001 >nul
echo.
echo ==========================================
echo   Excel File Split Tool - Installer
echo ==========================================
echo.

REM Check admin privileges
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    echo Administrator privileges required...
    echo Please right-click and select "Run as administrator"
    pause
    exit /b 1
)

echo Installing Excel File Split Tool...

REM Create program directory
set INSTALL_DIR=%PROGRAMFILES%\ExcelFileSplitTool
if not exist "%INSTALL_DIR%" (
    echo Creating installation directory: %INSTALL_DIR%
    mkdir "%INSTALL_DIR%"
)

REM Copy executable
echo Copying program files...
copy "ExcelFileSplitTool.exe" "%INSTALL_DIR%\" >nul
if %errorlevel% neq 0 (
    echo Error: Unable to copy program files
    pause
    exit /b 1
)

REM Create desktop shortcut
echo Creating desktop shortcut...
powershell -Command "& {$WshShell = New-Object -comObject WScript.Shell; $Shortcut = $WshShell.CreateShortcut('%USERPROFILE%\Desktop\ExcelFileSplitTool.lnk'); $Shortcut.TargetPath = '%INSTALL_DIR%\ExcelFileSplitTool.exe'; $Shortcut.WorkingDirectory = '%INSTALL_DIR%'; $Shortcut.Description = 'Excel File Split Tool'; $Shortcut.Save()}"

REM Create start menu shortcut
echo Creating start menu shortcut...
if not exist "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Excel Tools" (
    mkdir "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Excel Tools"
)
powershell -Command "& {$WshShell = New-Object -comObject WScript.Shell; $Shortcut = $WshShell.CreateShortcut('%APPDATA%\Microsoft\Windows\Start Menu\Programs\Excel Tools\ExcelFileSplitTool.lnk'); $Shortcut.TargetPath = '%INSTALL_DIR%\ExcelFileSplitTool.exe'; $Shortcut.WorkingDirectory = '%INSTALL_DIR%'; $Shortcut.Description = 'Excel File Split Tool'; $Shortcut.Save()}"

echo.
echo ==========================================
echo   Installation Complete!
echo ==========================================
echo.
echo Installation path: %INSTALL_DIR%
echo Desktop shortcut: Created
echo Start menu: Excel Tools ^> Excel File Split Tool
echo.
pause
'@

    # Create uninstaller script
    $uninstallerScript = @'
@echo off
chcp 65001 >nul
echo.
echo ==========================================
echo   Excel File Split Tool - Uninstaller
echo ==========================================
echo.

REM Check admin privileges
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    echo Administrator privileges required...
    echo Please right-click and select "Run as administrator"
    pause
    exit /b 1
)

set INSTALL_DIR=%PROGRAMFILES%\ExcelFileSplitTool

echo Are you sure you want to uninstall Excel File Split Tool?
choice /c YN /m "Press Y to confirm, N to cancel"
if errorlevel 2 goto :cancel

echo Uninstalling...

REM Remove shortcuts
echo Removing shortcuts...
del "%USERPROFILE%\Desktop\ExcelFileSplitTool.lnk" >nul 2>&1
del "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Excel Tools\ExcelFileSplitTool.lnk" >nul 2>&1
rmdir "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Excel Tools" >nul 2>&1

REM Remove program directory
if exist "%INSTALL_DIR%" (
    echo Removing program files...
    rmdir /s /q "%INSTALL_DIR%"
)

echo.
echo Uninstallation complete!
pause
goto :end

:cancel
echo Uninstallation cancelled
pause

:end
'@

    try {
        # Use UTF-8 encoding for batch files
        [System.IO.File]::WriteAllText((Join-Path "dist" "install.bat"), $installerScript, [System.Text.Encoding]::UTF8)
        [System.IO.File]::WriteAllText((Join-Path "dist" "uninstall.bat"), $uninstallerScript, [System.Text.Encoding]::UTF8)
        
        Write-ColorMessage "Installer scripts created:" "Green"
        Write-ColorMessage "  - dist\install.bat" "White"
        Write-ColorMessage "  - dist\uninstall.bat" "White"
    }
    catch {
        Write-ColorMessage "Failed to create installer scripts: $($_.Exception.Message)" "Red"
    }
}

Write-ColorMessage "Excel File Split Tool - Build Script" "Green"
Write-ColorMessage "================================================" "Green"

# Check if running from correct directory
if (-not (Test-Path "src\main.py")) {
    Write-ColorMessage "Error: Please run this script from project root directory" "Red"
    Write-ColorMessage "Make sure src\main.py exists in current directory" "Yellow"
    exit 1
}

# Check Python installation
try {
    $pythonVersion = python --version 2>&1
    Write-ColorMessage "Python detected: $pythonVersion" "Cyan"
}
catch {
    Write-ColorMessage "Error: Python not found, please install Python first" "Red"
    Write-ColorMessage "Download from: https://python.org" "Yellow"
    exit 1
}

# Virtual environment management
if (-not $NoVenv) {
    if (Test-Path "venv") {
        Write-ColorMessage "Activating virtual environment..." "Yellow"
        try {
            & "venv\Scripts\Activate.ps1"
            if ($LASTEXITCODE -ne 0) {
                throw "Virtual environment activation failed"
            }
        }
        catch {
            Write-ColorMessage "Warning: Virtual environment activation failed, using system Python" "Yellow"
        }
    }
    else {
        Write-ColorMessage "Creating virtual environment..." "Yellow"
        python -m venv venv
        
        try {
            & "venv\Scripts\Activate.ps1"
            Write-ColorMessage "Installing dependencies..." "Yellow"
            
            if (Test-Path "requirements.txt") {
                pip install -r requirements.txt
            }
            else {
                pip install pandas openpyxl xlrd xlsxwriter pyinstaller
            }
        }
        catch {
            Write-ColorMessage "Warning: Using system Python environment" "Yellow"
        }
    }
}
else {
    Write-ColorMessage "Skipping virtual environment, using system Python" "Yellow"
}

# Check required modules
Write-ColorMessage "Checking required modules..." "Yellow"
$requiredModules = @("pandas", "openpyxl", "xlrd", "xlsxwriter", "tkinter")
$missingModules = @()

foreach ($module in $requiredModules) {
    if (-not (Test-PythonModule $module)) {
        $missingModules += $module
        Write-ColorMessage "Missing: $module" "Red"
    }
    else {
        Write-ColorMessage "Found: $module" "Green"
    }
}

if ($missingModules.Count -gt 0) {
    Write-ColorMessage "Installing missing modules..." "Yellow"
    foreach ($module in $missingModules) {
        if ($module -ne "tkinter") {  # tkinter comes with Python
            pip install $module
        }
    }
}

# Check PyInstaller
try {
    $pyinstallerVersion = python -m PyInstaller --version 2>&1
    Write-ColorMessage "PyInstaller detected" "Green"
}
catch {
    Write-ColorMessage "Installing PyInstaller..." "Yellow"
    pip install pyinstaller
}

# Clean previous build files
if ($Clean -or (Test-Path "build") -or (Test-Path "dist")) {
    Write-ColorMessage "Cleaning previous build files..." "Yellow"
    if (Test-Path "build") { Remove-Item -Recurse -Force "build" }
    if (Test-Path "dist") { Remove-Item -Recurse -Force "dist" }
    if (Test-Path "src\__pycache__") { Remove-Item -Recurse -Force "src\__pycache__" }
    if (Test-Path "__pycache__") { Remove-Item -Recurse -Force "__pycache__" }
    Get-ChildItem -Path "." -Filter "*.spec" | Remove-Item -Force
}

# Build executable
Write-ColorMessage "Starting build process..." "Yellow"

# Use English name to avoid encoding issues
$appName = "ExcelFileSplitTool"

# Build arguments with enhanced compatibility
$buildArgs = @(
    "--onefile"
    "--windowed"
    "--name=$appName"
    # Core dependencies
    "--hidden-import=pandas"
    "--hidden-import=numpy"
    "--hidden-import=openpyxl"
    "--hidden-import=openpyxl.workbook"
    "--hidden-import=openpyxl.worksheet"
    "--hidden-import=openpyxl.reader.excel"
    "--hidden-import=openpyxl.writer.excel"
    "--hidden-import=xlrd"
    "--hidden-import=xlsxwriter"
    # Tkinter and GUI
    "--hidden-import=tkinter"
    "--hidden-import=tkinter.filedialog"
    "--hidden-import=tkinter.messagebox"
    "--hidden-import=tkinter.ttk"
    # System and encoding
    "--hidden-import=encodings"
    "--hidden-import=encodings.utf_8"
    "--hidden-import=encodings.cp1252"
    "--hidden-import=encodings.latin1"
    # Date and time
    "--hidden-import=datetime"
    "--hidden-import=dateutil"
    "--hidden-import=dateutil.parser"
    # File operations
    "--hidden-import=os"
    "--hidden-import=sys"
    "--hidden-import=subprocess"
    "--hidden-import=platform"
    # Application modules
    "--hidden-import=excel_splitter_gui"
    # Data files
    "--add-data=src;src"
    "--paths=src"
    # Build options
    "--clean"
    "--noconfirm"
    "--noupx"
    # Runtime options for better compatibility
    "--runtime-tmpdir=."
)

# Check for icon files
$iconPaths = @("assets\icon.ico", "icon.ico", "src\icon.ico")
foreach ($iconPath in $iconPaths) {
    if (Test-Path $iconPath) {
        $buildArgs += "--icon=$iconPath"
        Write-ColorMessage "Using icon: $iconPath" "Cyan"
        break
    }
}

$buildArgs += "src\main.py"

# Execute build with proper encoding and enhanced error handling
try {
    Write-ColorMessage "Executing: python -m PyInstaller $($buildArgs -join ' ')" "Cyan"
    
    # Set environment variables for proper encoding and compatibility
    $env:PYTHONIOENCODING = "utf-8"
    $env:PYTHONUTF8 = "1"
    $env:PYTHONHASHSEED = "1"
    $env:PYTHONDONTWRITEBYTECODE = "1"
    
    # Clear any existing PyInstaller cache
    if (Test-Path "$env:APPDATA\pyinstaller") {
        Remove-Item -Recurse -Force "$env:APPDATA\pyinstaller" -ErrorAction SilentlyContinue
    }
    
    & python -m PyInstaller @buildArgs
    
    if ($LASTEXITCODE -ne 0) {
        throw "PyInstaller failed with exit code: $LASTEXITCODE"
    }
    
    # Additional verification
    $exeName = "$appName.exe"
    $exePath = "dist\$exeName"
    
    if (-not (Test-Path $exePath)) {
        throw "Executable was not created successfully"
    }
    
    # Check if the executable is valid
    try {
        $fileInfo = Get-Item $exePath
        if ($fileInfo.Length -lt 1MB) {
            Write-ColorMessage "Warning: Executable size is unusually small ($([math]::Round($fileInfo.Length / 1KB, 2))KB)" "Yellow"
        }
    }
    catch {
        Write-ColorMessage "Warning: Could not verify executable properties" "Yellow"
    }
}
catch {
    Write-ColorMessage "Build failed: $($_.Exception.Message)" "Red"
    Write-ColorMessage "Troubleshooting tips:" "Yellow"
    Write-ColorMessage "1. Try running with -Clean parameter" "White"
    Write-ColorMessage "2. Check if antivirus is blocking the build" "White"
    Write-ColorMessage "3. Ensure all dependencies are properly installed" "White"
    Write-ColorMessage "4. Try building without virtual environment (-NoVenv)" "White"
    exit 1
}

# Check build result
$exeName = "$appName.exe"
$exePath = "dist\$exeName"

if (Test-Path $exePath) {
    Write-ColorMessage "Build successful!" "Green"
    Write-ColorMessage "Executable location: $exePath" "Green"
    
    # Get file information
    $fileInfo = Get-Item $exePath
    $fileSize = [math]::Round($fileInfo.Length / 1MB, 2)
    $buildTime = $fileInfo.CreationTime
    
    Write-ColorMessage "File size: ${fileSize}MB" "Cyan"
    Write-ColorMessage "Build time: $buildTime" "Cyan"
    
    # Development mode information
    if ($Dev) {
        Write-ColorMessage "`nDevelopment Info:" "Cyan"
        Write-ColorMessage "Working Directory: $(Get-Location)" "White"
        Write-ColorMessage "Python Version: $(python --version)" "White"
        Write-ColorMessage "PyInstaller Version: $(pyinstaller --version)" "White"
        Write-ColorMessage "Console Encoding: $([Console]::OutputEncoding.EncodingName)" "White"
    }
    
    # Create installer scripts
    if ($CreateInstaller) {
        New-InstallerScript
    }
    
    # Create desktop shortcut
    $createShortcut = Read-Host "Create desktop shortcut? (y/N)"
    if ($createShortcut -eq "y" -or $createShortcut -eq "Y") {
        New-DesktopShortcut -TargetPath $exePath -ShortcutName "Excel File Split Tool"
    }
}
else {
    Write-ColorMessage "Build failed - executable not found" "Red"
    Write-ColorMessage "Check the build output above for errors" "Yellow"
    exit 1
}

Write-ColorMessage "`nBuild Complete!" "Green"
Write-ColorMessage "================================================" "Green"
Write-ColorMessage "Usage:" "Cyan"
Write-ColorMessage "1. Double-click: $exePath" "White"
Write-ColorMessage "2. PowerShell: & '$exePath'" "White"
if ($CreateInstaller) {
    Write-ColorMessage "3. Install system-wide: Right-click dist\install.bat -> Run as administrator" "White"
}

# Test run option
$testNow = Read-Host "`nTest run the application now? (y/N)"
if ($testNow -eq "y" -or $testNow -eq "Y") {
    Write-ColorMessage "Starting application for testing..." "Yellow"
    try {
        Start-Process -FilePath $exePath
        Write-ColorMessage "Application started successfully!" "Green"
    }
    catch {
        Write-ColorMessage "Failed to start application: $($_.Exception.Message)" "Red"
    }
}

Write-ColorMessage "`nBuild script completed successfully!" "Green"