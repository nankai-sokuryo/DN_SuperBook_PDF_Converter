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

function Find-Python {
    Write-Log "Searching for Python..."
    
    # Search in PATH
    $pythonCmd = Get-Command python -ErrorAction SilentlyContinue
    if ($pythonCmd) {
        $version = & python --version 2>&1
        Write-Log "Found Python in PATH: $($pythonCmd.Source) ($version)"
        return $pythonCmd.Source
    }
    
    # Search using py launcher
    $pyCmd = Get-Command py -ErrorAction SilentlyContinue
    if ($pyCmd) {
        $version = & py --version 2>&1
        Write-Log "Found py launcher: $version"
        return "py"
    }
    
    # Search common installation paths (including Microsoft Store version)
    $localAppData = $env:LOCALAPPDATA
    $commonPaths = @(
        # Microsoft Store version
        "$localAppData\Microsoft\WindowsApps\python3.12.exe",
        "$localAppData\Microsoft\WindowsApps\python3.11.exe",
        "$localAppData\Microsoft\WindowsApps\python3.exe",
        "$localAppData\Microsoft\WindowsApps\python.exe",
        # Standard installer - Program Files
        "C:\Program Files\Python312\python.exe",
        "C:\Program Files\Python311\python.exe",
        "C:\Program Files\Python310\python.exe",
        # Standard installer - User install
        "$env:USERPROFILE\AppData\Local\Programs\Python\Python312\python.exe",
        "$env:USERPROFILE\AppData\Local\Programs\Python\Python311\python.exe",
        # Legacy paths
        "C:\Python312\python.exe",
        "C:\Python311\python.exe",
        "C:\Python310\python.exe"
    )
    
    foreach ($path in $commonPaths) {
        if (Test-Path $path) {
            $version = & $path --version 2>&1
            Write-Log "Found Python at: $path ($version)"
            return $path
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
        Write-Log "Python not found." "ERROR"
        Write-Log "Please install Python 3.11 or 3.12 from https://www.python.org/downloads/"
        Write-Log "Make sure to check 'Add Python to PATH' and 'Install for all users' during installation."
        if (-not $Silent) {
            Read-Host "Press Enter to exit"
        }
        exit 1
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
