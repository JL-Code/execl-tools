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

function Test-WindowsDefender {
    Write-ColorMessage "Checking Windows Defender exclusions..." "Yellow"
    
    $currentPath = Get-Location
    $distPath = Join-Path $currentPath "dist"
    
    try {
        # Check if current directory is in exclusions
        $exclusions = Get-MpPreference | Select-Object -ExpandProperty ExclusionPath -ErrorAction SilentlyContinue
        
        if ($exclusions -and ($exclusions -contains $currentPath -or $exclusions -contains $distPath)) {
            Write-ColorMessage "Directory already in Windows Defender exclusions" "Green"
        } else {
            Write-ColorMessage "Warning: Directory not in Windows Defender exclusions" "Yellow"
            Write-ColorMessage "This may cause the executable to be blocked or corrupted" "Yellow"
            
            $addExclusion = Read-Host "Add current directory to Windows Defender exclusions? (Requires Admin) (y/N)"
            if ($addExclusion -eq "y" -or $addExclusion -eq "Y") {
                try {
                    Add-MpPreference -ExclusionPath $currentPath
                    Add-MpPreference -ExclusionPath $distPath
                    Write-ColorMessage "Added to Windows Defender exclusions" "Green"
                } catch {
                    Write-ColorMessage "Failed to add exclusion (need admin rights): $($_.Exception.Message)" "Red"
                    Write-ColorMessage "Please manually add '$currentPath' to Windows Defender exclusions" "Yellow"
                }
            }
        }
    } catch {
        Write-ColorMessage "Could not check Windows Defender status" "Yellow"
    }
}

function Test-ExecutableIntegrity {
    param([string]$ExePath)
    
    Write-ColorMessage "Testing executable integrity..." "Yellow"
    
    if (-not (Test-Path $ExePath)) {
        Write-ColorMessage "Executable not found: $ExePath" "Red"
        return $false
    }
    
    try {
        # Test if file can be read
        $fileStream = [System.IO.File]::OpenRead($ExePath)
        $fileStream.Close()
        Write-ColorMessage "File is readable" "Green"
        
        # Check file size
        $fileInfo = Get-Item $ExePath
        if ($fileInfo.Length -lt 1KB) {
            Write-ColorMessage "Warning: File size too small ($($fileInfo.Length) bytes)" "Yellow"
            return $false
        }
        
        # Test execution (dry run)
        $processInfo = New-Object System.Diagnostics.ProcessStartInfo
        $processInfo.FileName = $ExePath
        $processInfo.Arguments = "--version"  # Try to get version info
        $processInfo.UseShellExecute = $false
        $processInfo.RedirectStandardOutput = $true
        $processInfo.RedirectStandardError = $true
        $processInfo.CreateNoWindow = $true
        
        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $processInfo
        
        # Quick test - start and immediately kill
        $started = $process.Start()
        if ($started) {
            Start-Sleep -Milliseconds 100
            if (-not $process.HasExited) {
                $process.Kill()
            }
            Write-ColorMessage "Executable appears to be valid" "Green"
            return $true
        } else {
            Write-ColorMessage "Failed to start executable" "Red"
            return $false
        }
    } catch {
        Write-ColorMessage "Integrity test failed: $($_.Exception.Message)" "Red"
        return $false
    }
}
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
    pyinstaller --version | Out-Null
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

# Check Windows Defender and antivirus issues
Test-WindowsDefender

# Build executable with proper encoding
Write-ColorMessage "Starting build process..." "Yellow"

# Use English name to avoid encoding issues
$appName = "ExcelFileSplitTool"

# Build arguments - optimized for Windows 11 compatibility
$buildArgs = @(
    "--onedir"                           # Use onedir instead of onefile for better compatibility
    "--windowed"
    "--name=$appName"
    "--paths=src"
    "--distpath=dist"
    "--workpath=build"
    "--specpath=."
    "--add-data=src;src"                 # Ensure all source files are included
    "--hidden-import=pandas"
    "--hidden-import=openpyxl" 
    "--hidden-import=xlrd"
    "--hidden-import=xlsxwriter"
    "--hidden-import=tkinter"
    "--hidden-import=tkinter.filedialog"
    "--hidden-import=tkinter.messagebox"
    "--hidden-import=tkinter.ttk"
    "--hidden-import=excel_splitter_gui"
    "--hidden-import=pkg_resources.py2_warn"
    "--exclude-module=matplotlib"        # Exclude unused modules to reduce size
    "--exclude-module=IPython"
    "--exclude-module=jupyter"
    "--noupx"                           # Disable UPX compression that can cause issues
    "--clean"
    "--noconfirm"
    "--console"                         # Temporarily enable console for debugging
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

# Execute build with proper encoding
try {
    Write-ColorMessage "Executing: pyinstaller $($buildArgs -join ' ')" "Cyan"
    
    # Set environment variables for proper encoding
    $env:PYTHONIOENCODING = "utf-8"
    $env:PYTHONUTF8 = "1"
    
    & pyinstaller @buildArgs
    
    if ($LASTEXITCODE -ne 0) {
        throw "PyInstaller failed with exit code: $LASTEXITCODE"
    }
}
catch {
    Write-ColorMessage "Build failed: $($_.Exception.Message)" "Red"
    exit 1
}

# Check build result with integrity testing
$exeName = "$appName.exe"
$exePath = "dist\$appName\$exeName"  # Note: onedir creates a folder

if (Test-Path $exePath) {
    Write-ColorMessage "Build successful!" "Green"
    Write-ColorMessage "Executable location: $exePath" "Green"
    
    # Test executable integrity
    if (Test-ExecutableIntegrity $exePath) {
        Write-ColorMessage "Executable integrity check passed" "Green"
    } else {
        Write-ColorMessage "Warning: Executable may have issues" "Yellow"
    }
    
    # Get file information
    $fileInfo = Get-Item $exePath
    $fileSize = [math]::Round($fileInfo.Length / 1MB, 2)
    $buildTime = $fileInfo.CreationTime
    
    Write-ColorMessage "File size: ${fileSize}MB" "Cyan"
    Write-ColorMessage "Build time: $buildTime" "Cyan"
    
    # Create a batch launcher for easier distribution
    $launcherScript = @"
@echo off
cd /d "%~dp0$appName"
start "" "$exeName"
"@
    
    $launcherScript | Out-File -FilePath "dist\Launch_$appName.bat" -Encoding ASCII
    Write-ColorMessage "Created launcher: dist\Launch_$appName.bat" "Green"
    
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