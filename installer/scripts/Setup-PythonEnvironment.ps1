# Setup-PythonEnvironment.ps1
# SuperBookTools Python Environment Setup Script
# Sets up RealEsrgan and YomiToku environments

param(
    [Parameter()]
    [ValidateSet('cu126', 'cu128', 'cu130')]
    [string]$CudaVersion = 'cu126',
    
    [Parameter()]
    [switch]$Force,
    
    [Parameter()]
    [switch]$Silent
)

$ErrorActionPreference = "Stop"

# Get app root from script directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$appRoot = Split-Path -Parent $scriptDir

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    if (-not $Silent) {
        switch ($Level) {
            "ERROR" { Write-Host $logMessage -ForegroundColor Red }
            "WARN"  { Write-Host $logMessage -ForegroundColor Yellow }
            "OK"    { Write-Host $logMessage -ForegroundColor Green }
            default { Write-Host $logMessage }
        }
    }
    
    # Write to log file
    $logFile = Join-Path $appRoot "setup.log"
    Add-Content -Path $logFile -Value $logMessage
}

function Test-PythonWorks {
    param([string]$PythonPath)
    
    try {
        if ($PythonPath -eq "py") {
            $result = & py --version 2>&1
        } else {
            $result = & $PythonPath --version 2>&1
        }
        
        # Check if output contains valid Python version (e.g., "Python 3.12.9")
        if ($result -match "Python\s+(\d+)\.(\d+)\.(\d+)") {
            $major = [int]$matches[1]
            $minor = [int]$matches[2]
            # Require Python 3.10 or higher
            if ($major -eq 3 -and $minor -ge 10) {
                return $result
            }
            Write-Log "Python version too old: $result (requires 3.10+)" "WARN"
            return $null
        }
        return $null
    }
    catch {
        return $null
    }
}

function Find-Python {
    Write-Log "Searching for Python..."
    
    # Search in PATH (excluding WindowsApps stub first)
    $pythonCmd = Get-Command python -ErrorAction SilentlyContinue
    if ($pythonCmd) {
        # Skip WindowsApps stub - it may not be a real Python
        if ($pythonCmd.Source -notlike "*WindowsApps*") {
            $version = Test-PythonWorks -PythonPath $pythonCmd.Source
            if ($version) {
                Write-Log "Found Python in PATH: $($pythonCmd.Source) ($version)"
                return $pythonCmd.Source
            }
        }
    }
    
    # Search using py launcher
    $pyCmd = Get-Command py -ErrorAction SilentlyContinue
    if ($pyCmd) {
        $version = Test-PythonWorks -PythonPath "py"
        if ($version) {
            Write-Log "Found py launcher: $version"
            return "py"
        }
    }
    
    # Search common installation paths
    $localAppData = $env:LOCALAPPDATA
    $commonPaths = @(
        # Standard installer - Program Files (prioritize)
        "C:\Program Files\Python312\python.exe",
        "C:\Program Files\Python311\python.exe",
        "C:\Program Files\Python310\python.exe",
        # Standard installer - User install
        "$env:USERPROFILE\AppData\Local\Programs\Python\Python312\python.exe",
        "$env:USERPROFILE\AppData\Local\Programs\Python\Python311\python.exe",
        "$env:USERPROFILE\AppData\Local\Programs\Python\Python310\python.exe",
        # Legacy paths
        "C:\Python312\python.exe",
        "C:\Python311\python.exe",
        "C:\Python310\python.exe",
        # Microsoft Store version (check last - may be stub)
        "$localAppData\Microsoft\WindowsApps\python3.12.exe",
        "$localAppData\Microsoft\WindowsApps\python3.11.exe",
        "$localAppData\Microsoft\WindowsApps\python3.exe",
        "$localAppData\Microsoft\WindowsApps\python.exe"
    )
    
    foreach ($path in $commonPaths) {
        if (Test-Path $path) {
            $version = Test-PythonWorks -PythonPath $path
            if ($version) {
                Write-Log "Found Python at: $path ($version)"
                return $path
            }
        }
    }
    
    # Also check PATH python again if it was WindowsApps (in case Store version is actually installed)
    if ($pythonCmd -and $pythonCmd.Source -like "*WindowsApps*") {
        $version = Test-PythonWorks -PythonPath $pythonCmd.Source
        if ($version) {
            Write-Log "Found Python (Store) in PATH: $($pythonCmd.Source) ($version)"
            return $pythonCmd.Source
        }
    }
    
    return $null
}

function Setup-RealEsrgan {
    param([string]$PythonPath, [string]$CudaVer)
    
    $toolsPath = Join-Path $appRoot "external_tools\image_tools\RealEsrgan\RealEsrgan_Repo"
    $venvPath = Join-Path $toolsPath "venv"
    
    if ((Test-Path $venvPath) -and (-not $Force)) {
        Write-Log "RealEsrgan environment already exists. Skipping." "WARN"
        return $true
    }
    
    Write-Log "Setting up RealEsrgan environment..."
    
    try {
        # Create directory
        if (-not (Test-Path $toolsPath)) {
            New-Item -ItemType Directory -Path $toolsPath -Force | Out-Null
        }
        
        # Remove existing venv if -Force
        if ((Test-Path $venvPath) -and $Force) {
            Write-Log "Removing existing venv..."
            Remove-Item -Recurse -Force $venvPath
        }
        
        # Create venv
        Write-Log "Creating virtual environment..."
        if ($PythonPath -eq "py") {
            & py -3 -m venv $venvPath
        } else {
            & $PythonPath -m venv $venvPath
        }
        
        # Upgrade pip
        Write-Log "Upgrading pip..."
        & "$venvPath\Scripts\python.exe" -m pip install --upgrade pip --quiet
        
        # Install PyTorch
        Write-Log "Installing PyTorch ($CudaVer)... This may take several minutes."
        & "$venvPath\Scripts\pip.exe" install torch torchvision torchaudio --index-url "https://download.pytorch.org/whl/$CudaVer" --quiet
        
        # Clone Real-ESRGAN repository
        $repoPath = Join-Path $toolsPath "Real-ESRGAN"
        if (-not (Test-Path $repoPath)) {
            Write-Log "Cloning Real-ESRGAN repository..."
            git clone https://github.com/xinntao/Real-ESRGAN.git $repoPath --quiet
            Push-Location $repoPath
            git checkout a4abfb2979a7bbff3f69f58f58ae324608821e27 --quiet
            Pop-Location
        }
        
        # Download model
        $weightsPath = Join-Path $repoPath "weights"
        if (-not (Test-Path $weightsPath)) {
            New-Item -ItemType Directory -Path $weightsPath -Force | Out-Null
        }
        $modelPath = Join-Path $weightsPath "RealESRGAN_x4plus.pth"
        if (-not (Test-Path $modelPath)) {
            Write-Log "Downloading RealESRGAN model..."
            Invoke-WebRequest -Uri "https://github.com/xinntao/Real-ESRGAN/releases/download/v0.1.0/RealESRGAN_x4plus.pth" -OutFile $modelPath
        }
        
        # Install requirements.txt
        $reqPath = Join-Path $repoPath "requirements.txt"
        if (Test-Path $reqPath) {
            Write-Log "Installing dependencies..."
            & "$venvPath\Scripts\pip.exe" install -r $reqPath --quiet
        }
        
        # Patch degradations.py
        $degradationsPath = Join-Path $venvPath "Lib\site-packages\basicsr\data\degradations.py"
        if (Test-Path $degradationsPath) {
            Write-Log "Patching degradations.py..."
            $content = Get-Content $degradationsPath -Raw
            $content = $content -replace 'from torchvision.transforms.functional_tensor import rgb_to_grayscale', 'from torchvision.transforms.functional import rgb_to_grayscale'
            Set-Content $degradationsPath $content
        }
        
        # Create version.py
        $versionPath = Join-Path $repoPath "realesrgan\version.py"
        if (-not (Test-Path $versionPath)) {
            $versionDir = Split-Path $versionPath -Parent
            if (-not (Test-Path $versionDir)) {
                New-Item -ItemType Directory -Path $versionDir -Force | Out-Null
            }
            New-Item -ItemType File -Path $versionPath -Force | Out-Null
        }
        
        Write-Log "RealEsrgan setup completed" "OK"
        return $true
    }
    catch {
        Write-Log "RealEsrgan setup failed: $_" "ERROR"
        return $false
    }
}

function Setup-YomiToku {
    param([string]$PythonPath, [string]$CudaVer)
    
    $toolsPath = Join-Path $appRoot "external_tools\image_tools\yomitoku"
    $venvPath = Join-Path $toolsPath "venv"
    
    if ((Test-Path $venvPath) -and (-not $Force)) {
        Write-Log "YomiToku environment already exists. Skipping." "WARN"
        return $true
    }
    
    Write-Log "Setting up YomiToku environment..."
    
    try {
        # Create directory
        if (-not (Test-Path $toolsPath)) {
            New-Item -ItemType Directory -Path $toolsPath -Force | Out-Null
        }
        
        # Remove existing venv if -Force
        if ((Test-Path $venvPath) -and $Force) {
            Write-Log "Removing existing venv..."
            Remove-Item -Recurse -Force $venvPath
        }
        
        # Create venv
        Write-Log "Creating virtual environment..."
        if ($PythonPath -eq "py") {
            & py -3 -m venv $venvPath
        } else {
            & $PythonPath -m venv $venvPath
        }
        
        # Upgrade pip
        Write-Log "Upgrading pip..."
        & "$venvPath\Scripts\python.exe" -m pip install --upgrade pip --quiet
        
        # Install PyTorch
        Write-Log "Installing PyTorch ($CudaVer)... This may take several minutes."
        & "$venvPath\Scripts\pip.exe" install torch torchvision torchaudio --index-url "https://download.pytorch.org/whl/$CudaVer" --quiet
        
        # Install YomiToku
        Write-Log "Installing YomiToku..."
        & "$venvPath\Scripts\pip.exe" install "yomitoku==0.10.3" --quiet
        
        Write-Log "YomiToku setup completed" "OK"
        return $true
    }
    catch {
        Write-Log "YomiToku setup failed: $_" "ERROR"
        return $false
    }
}

# Main
function Main {
    Write-Log "=========================================="
    Write-Log "SuperBookTools Python Environment Setup"
    Write-Log "=========================================="
    Write-Log "Application Root: $appRoot"
    Write-Log "CUDA Version: $CudaVersion"
    Write-Log ""
    
    # Check Git
    $gitCmd = Get-Command git -ErrorAction SilentlyContinue
    if (-not $gitCmd) {
        Write-Log "Git is not installed." "ERROR"
        Write-Log "Please install Git from https://git-scm.com/download/win"
        if (-not $Silent) {
            Read-Host "Press Enter to exit"
        }
        exit 1
    }
    
    # Check Python
    $pythonPath = Find-Python
    if (-not $pythonPath) {
        Write-Log "" 
        Write-Log "=========================================="
        Write-Log "  Python is NOT installed" "WARN"
        Write-Log "=========================================="
        Write-Log ""
        Write-Log "Python 3.10 or higher is required to set up AI features." "WARN"
        Write-Log "(RealEsrgan image upscaling & YomiToku OCR)"
        Write-Log ""
        Write-Log "To install Python:"
        Write-Log "  1. Visit: https://www.python.org/downloads/windows/"
        Write-Log "  2. Click 'Latest Python 3 Release' or download Python 3.12.x"
        Write-Log "  3. Scroll down and click 'Windows installer (64-bit)'"
        Write-Log "  4. Run the downloaded installer"
        Write-Log "  5. IMPORTANT: Check 'Add Python to PATH' at the bottom"
        Write-Log "  6. Click 'Install Now' (or Customize for all users)"
        Write-Log ""
        Write-Log "After installing Python, run this script again."
        Write-Log ""
        Write-Log "NOTE: SuperBookTools basic features will work without Python." "OK"
        Write-Log "      Only AI-powered features require Python environment."
        Write-Log ""
        
        if (-not $Silent) {
            Write-Host ""
            Write-Host "Would you like to open the Python download page? (Y/N): " -NoNewline -ForegroundColor Yellow
            $response = Read-Host
            if ($response -eq 'Y' -or $response -eq 'y') {
                Start-Process "https://www.python.org/downloads/windows/"
            }
            Write-Host ""
            Read-Host "Press Enter to exit"
        }
        exit 0  # Exit with 0 since this is not a fatal error
    }
    
    Write-Log ""
    Write-Log "--- RealEsrgan Setup ---"
    $realEsrganResult = Setup-RealEsrgan -PythonPath $pythonPath -CudaVer $CudaVersion
    
    Write-Log ""
    Write-Log "--- YomiToku Setup ---"
    $yomitokuResult = Setup-YomiToku -PythonPath $pythonPath -CudaVer $CudaVersion
    
    Write-Log ""
    Write-Log "=========================================="
    if ($realEsrganResult -and $yomitokuResult) {
        Write-Log "All setup completed successfully!" "OK"
        $exitCode = 0
    } else {
        Write-Log "Some setup failed." "ERROR"
        Write-Log "Please check setup.log for details."
        $exitCode = 1
    }
    Write-Log "=========================================="
    
    if (-not $Silent) {
        Read-Host "Press Enter to exit"
    }
    
    exit $exitCode
}

# Run
Main
