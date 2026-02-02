# Setup-ExternalTools.ps1
# SuperBookTools External Tools Setup Script
# Downloads and sets up ImageMagick, QPDF, pdfcpu, exiftool, and Python environments

param(
    [Parameter()]
    [switch]$Force,
    
    [Parameter()]
    [switch]$Silent,
    
    [Parameter()]
    [string]$AppRoot = "",
    
    [Parameter()]
    [string]$ToolsPath = "",
    
    [Parameter()]
    [ValidateSet('cu126', 'cu128', 'cu130')]
    [string]$CudaVersion = 'cu126',
    
    [Parameter()]
    [switch]$SkipPython,
    
    [Parameter()]
    [switch]$SkipSymlink
)

$ErrorActionPreference = "Stop"

# Get app root - use parameter if provided, otherwise derive from script location
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if ($AppRoot -ne "") {
    $appRoot = $AppRoot
} else {
    # Default: repository root (parent of installer/ which is parent of scripts/)
    # scripts/ -> installer/ -> repo root
    $installerDir = Split-Path -Parent $scriptDir
    $appRoot = Split-Path -Parent $installerDir
}

# Tools base directory - use ToolsPath if provided, otherwise default to external_tools/image_tools
if ($ToolsPath -ne "") {
    $toolsBase = $ToolsPath
} else {
    $toolsBase = Join-Path $appRoot "external_tools\image_tools"
}

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
    $logFile = Join-Path $appRoot "setup-tools.log"
    Add-Content -Path $logFile -Value $logMessage
}

function Download-File {
    param([string]$Url, [string]$OutFile)
    
    Write-Log "Downloading: $Url"
    try {
        $ProgressPreference = 'SilentlyContinue'
        Invoke-WebRequest -Uri $Url -OutFile $OutFile -UseBasicParsing
        return $true
    }
    catch {
        Write-Log "Download failed: $_" "ERROR"
        return $false
    }
}

function Setup-ImageMagick {
    $toolDir = Join-Path $toolsBase "ImageMagick-portable-Q16-HDRI-x64"
    $marker = Join-Path $toolDir "magick.exe"
    
    # Ensure tool directory exists
    if (-not (Test-Path $toolDir)) {
        New-Item -ItemType Directory -Path $toolDir -Force | Out-Null
    }
    
    if ((Test-Path $marker) -and (-not $Force)) {
        Write-Log "ImageMagick already exists. Skipping." "WARN"
        return $true
    }
    
    Write-Log "Setting up ImageMagick..."
    Write-Log "  Tool directory: $toolDir"
    
    # ImageMagick portable Q16-HDRI x64
    # Dynamically find latest version from binaries page
    $binariesUrl = "https://imagemagick.org/archive/binaries/"
    
    try {
        # Try to find latest version dynamically
        $response = Invoke-WebRequest -Uri $binariesUrl -UseBasicParsing -ErrorAction Stop
        $pattern = 'ImageMagick-7\.1\.[\d\-]+-portable-Q16-HDRI-x64\.7z'
        $matches7z = [regex]::Matches($response.Content, $pattern)
        
        if ($matches7z.Count -gt 0) {
            $fileName = $matches7z[$matches7z.Count - 1].Value
        } else {
            # Fallback to known version
            $fileName = "ImageMagick-7.1.2-13-portable-Q16-HDRI-x64.7z"
        }
        
        $url = "$binariesUrl$fileName"
        $downloadPath = Join-Path $env:TEMP $fileName
        
        Write-Log "Downloading: $url"
        Invoke-WebRequest -Uri $url -OutFile $downloadPath -UseBasicParsing
        
        Write-Log "Extracting ImageMagick..."
        
        # Clear directory except .gitignore and _empty.txt
        Get-ChildItem $toolDir -Exclude ".gitignore", "_empty.txt" | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
        
        # Extract .7z using tar (Windows 10+ has built-in tar that handles 7z)
        tar -xf $downloadPath -C $toolDir
        
        # Cleanup
        Remove-Item $downloadPath -Force -ErrorAction SilentlyContinue
        
        Write-Log "ImageMagick setup complete." "OK"
        return $true
    }
    catch {
        Write-Log "ImageMagick setup failed: $_" "ERROR"
        return $false
    }
}

function Setup-QPDF {
    $toolDir = Join-Path $toolsBase "QPDF"
    $marker = Join-Path $toolDir "bin\qpdf.exe"
    
    # Ensure tool directory exists
    if (-not (Test-Path $toolDir)) {
        New-Item -ItemType Directory -Path $toolDir -Force | Out-Null
    }
    
    if ((Test-Path $marker) -and (-not $Force)) {
        Write-Log "QPDF already exists. Skipping." "WARN"
        return $true
    }
    
    Write-Log "Setting up QPDF..."
    Write-Log "  Tool directory: $toolDir"
    
    # QPDF releases: https://github.com/qpdf/qpdf/releases
    $version = "11.9.1"
    $url = "https://github.com/qpdf/qpdf/releases/download/v$version/qpdf-$version-msvc64.zip"
    $zipFile = Join-Path $env:TEMP "QPDF.zip"
    
    try {
        if (-not (Download-File -Url $url -OutFile $zipFile)) {
            return $false
        }
        
        Write-Log "Extracting QPDF..."
        
        # Clear directory except .gitignore and _empty.txt
        Get-ChildItem $toolDir -Exclude ".gitignore", "_empty.txt" | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
        
        # Extract to temp first
        $tempExtract = Join-Path $env:TEMP "QPDF_extract"
        if (Test-Path $tempExtract) { Remove-Item $tempExtract -Recurse -Force }
        
        Expand-Archive -Path $zipFile -DestinationPath $tempExtract -Force
        
        # QPDF extracts to qpdf-version subdirectory
        $subDir = Get-ChildItem $tempExtract -Directory | Select-Object -First 1
        if ($subDir) {
            Get-ChildItem $subDir.FullName | Move-Item -Destination $toolDir -Force
        }
        
        # Cleanup
        Remove-Item $zipFile -Force -ErrorAction SilentlyContinue
        Remove-Item $tempExtract -Recurse -Force -ErrorAction SilentlyContinue
        
        Write-Log "QPDF setup complete." "OK"
        return $true
    }
    catch {
        Write-Log "QPDF setup failed: $_" "ERROR"
        return $false
    }
}

function Setup-Pdfcpu {
    $toolDir = Join-Path $toolsBase "pdfcpu"
    $marker = Join-Path $toolDir "pdfcpu.exe"
    
    # Ensure tool directory exists
    if (-not (Test-Path $toolDir)) {
        New-Item -ItemType Directory -Path $toolDir -Force | Out-Null
    }
    
    if ((Test-Path $marker) -and (-not $Force)) {
        Write-Log "pdfcpu already exists. Skipping." "WARN"
        return $true
    }
    
    Write-Log "Setting up pdfcpu..."
    Write-Log "  Tool directory: $toolDir"
    
    # pdfcpu releases: https://github.com/pdfcpu/pdfcpu/releases
    $version = "0.11.0"
    $url = "https://github.com/pdfcpu/pdfcpu/releases/download/v$version/pdfcpu_$version`_Windows_x86_64.zip"
    $zipFile = Join-Path $env:TEMP "pdfcpu.zip"
    
    try {
        if (-not (Download-File -Url $url -OutFile $zipFile)) {
            return $false
        }
        
        Write-Log "Extracting pdfcpu..."
        
        # Clear directory except .gitignore and _empty.txt
        Get-ChildItem $toolDir -Exclude ".gitignore", "_empty.txt" | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
        
        Expand-Archive -Path $zipFile -DestinationPath $toolDir -Force
        
        # Cleanup
        Remove-Item $zipFile -Force -ErrorAction SilentlyContinue
        
        Write-Log "pdfcpu setup complete." "OK"
        return $true
    }
    catch {
        Write-Log "pdfcpu setup failed: $_" "ERROR"
        return $false
    }
}

function Setup-Exiftool {
    $toolDir = Join-Path $toolsBase "exiftool-13.30_64"
    $marker = Join-Path $toolDir "exiftool.exe"
    
    # Ensure tool directory exists
    if (-not (Test-Path $toolDir)) {
        New-Item -ItemType Directory -Path $toolDir -Force | Out-Null
    }
    
    if ((Test-Path $marker) -and (-not $Force)) {
        Write-Log "exiftool already exists. Skipping." "WARN"
        return $true
    }
    
    Write-Log "Setting up exiftool..."
    Write-Log "  Tool directory: $toolDir"
    
    # exiftool - download from SoftEther mirror (official site blocks automated downloads)
    $url = "https://filecenter.softether-upload.com/d/260118_003_73929/exiftool-13.30_64.zip"
    $zipFile = Join-Path $env:TEMP "exiftool.zip"
    
    try {
        if (-not (Download-File -Url $url -OutFile $zipFile)) {
            return $false
        }
        
        Write-Log "Extracting exiftool..."
        
        # Clear directory except .gitignore and _empty.txt
        Get-ChildItem $toolDir -Exclude ".gitignore", "_empty.txt" | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
        
        # Extract to temp first
        $tempExtract = Join-Path $env:TEMP "exiftool_extract"
        if (Test-Path $tempExtract) { Remove-Item $tempExtract -Recurse -Force }
        New-Item -ItemType Directory -Path $tempExtract -Force | Out-Null
        
        Expand-Archive -Path $zipFile -DestinationPath $tempExtract -Force
        
        # Check if there's a subdirectory
        $extractedSubDir = Join-Path $tempExtract "exiftool-13.30_64"
        if (Test-Path $extractedSubDir) {
            Copy-Item -Path "$extractedSubDir\*" -Destination $toolDir -Recurse -Force
        } else {
            Copy-Item -Path "$tempExtract\*" -Destination $toolDir -Recurse -Force
        }
        
        # Rename exiftool(-k).exe to exiftool.exe
        $srcExe = Join-Path $toolDir "exiftool(-k).exe"
        if (Test-Path $srcExe) {
            Rename-Item -Path $srcExe -NewName "exiftool.exe" -Force
        }
        
        # Cleanup
        Remove-Item $zipFile -Force -ErrorAction SilentlyContinue
        Remove-Item $tempExtract -Recurse -Force -ErrorAction SilentlyContinue
        
        Write-Log "exiftool setup complete." "OK"
        return $true
    }
    catch {
        Write-Log "exiftool setup failed: $_" "ERROR"
        return $false
    }
}

function Setup-TesseractOCR {
    $toolDir = Join-Path $toolsBase "TesseractOCR_Data"
    $marker = Join-Path $toolDir "eng.traineddata"
    
    # Ensure tool directory exists
    if (-not (Test-Path $toolDir)) {
        New-Item -ItemType Directory -Path $toolDir -Force | Out-Null
    }
    
    if ((Test-Path $marker) -and (-not $Force)) {
        Write-Log "TesseractOCR_Data already exists. Skipping." "WARN"
        return $true
    }
    
    Write-Log "Setting up Tesseract OCR Data..."
    Write-Log "  Tool directory: $toolDir"
    
    # Tesseract OCR trained data (best quality)
    $url = "https://github.com/tesseract-ocr/tessdata_best/archive/refs/tags/4.1.0.zip"
    $zipFile = Join-Path $env:TEMP "tessdata.zip"
    
    try {
        if (-not (Download-File -Url $url -OutFile $zipFile)) {
            return $false
        }
        
        Write-Log "Extracting Tesseract OCR Data..."
        
        # Clear directory except .gitignore and _empty.txt
        Get-ChildItem $toolDir -Exclude ".gitignore", "_empty.txt" | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
        
        # Extract to temp first
        $tempExtract = Join-Path $env:TEMP "tessdata_extract"
        if (Test-Path $tempExtract) { Remove-Item $tempExtract -Recurse -Force }
        
        Expand-Archive -Path $zipFile -DestinationPath $tempExtract -Force
        
        # Copy only needed language files (eng and jpn)
        $extractedDir = Join-Path $tempExtract "tessdata_best-4.1.0"
        Copy-Item (Join-Path $extractedDir "eng.traineddata") $toolDir -Force
        Copy-Item (Join-Path $extractedDir "jpn.traineddata") $toolDir -Force
        
        # Cleanup
        Remove-Item $zipFile -Force -ErrorAction SilentlyContinue
        Remove-Item $tempExtract -Recurse -Force -ErrorAction SilentlyContinue
        
        Write-Log "Tesseract OCR Data setup complete." "OK"
        return $true
    }
    catch {
        Write-Log "Tesseract OCR Data setup failed: $_" "ERROR"
        return $false
    }
}

# Main
Write-Log "=========================================="
Write-Log "SuperBookTools External Tools Setup"
Write-Log "=========================================="
Write-Log "Application Root: $appRoot"
Write-Log "Tools Base: $toolsBase"
Write-Log ""

# Ensure tools base directory exists
if (-not (Test-Path $toolsBase)) {
    New-Item -ItemType Directory -Path $toolsBase -Force | Out-Null
}

# Create symlinks for development (SuperBookToolsGui) FIRST
if (-not $SkipSymlink) {
    Write-Log "--- Development Symlinks ---"

    $externalToolsDir = Join-Path $appRoot "external_tools"
    $guiProjectDir = Join-Path $appRoot "SuperBookToolsGui"

    if (Test-Path $guiProjectDir) {
        $symlinkPath = Join-Path $guiProjectDir "external_tools"
        
        # Check if symlink already exists and points to correct target
        if (Test-Path $symlinkPath) {
            $item = Get-Item $symlinkPath -Force
            if ($item.LinkType -eq "SymbolicLink" -and $item.Target -eq $externalToolsDir) {
                Write-Log "Symlink already exists: $symlinkPath -> $externalToolsDir" "WARN"
            } else {
                # Remove existing item (file/folder/wrong symlink)
                Remove-Item $symlinkPath -Force -Recurse -ErrorAction SilentlyContinue
                try {
                    New-Item -ItemType SymbolicLink -Path $symlinkPath -Target $externalToolsDir -Force | Out-Null
                    Write-Log "Symlink created: $symlinkPath -> $externalToolsDir" "OK"
                } catch {
                    Write-Log "Failed to create symlink (run as Administrator): $_" "ERROR"
                }
            }
        } else {
            try {
                New-Item -ItemType SymbolicLink -Path $symlinkPath -Target $externalToolsDir -Force | Out-Null
                Write-Log "Symlink created: $symlinkPath -> $externalToolsDir" "OK"
            } catch {
                Write-Log "Failed to create symlink (run as Administrator): $_" "ERROR"
            }
        }
    } else {
        Write-Log "SuperBookToolsGui directory not found, skipping symlink creation" "WARN"
    }
} else {
    Write-Log "--- Skipping Symlink Creation (-SkipSymlink specified) ---"
}

Write-Log ""
Write-Log "--- Downloading External Tools ---"

$results = @{
    ImageMagick = Setup-ImageMagick
    QPDF = Setup-QPDF
    pdfcpu = Setup-Pdfcpu
    exiftool = Setup-Exiftool
    TesseractOCR = Setup-TesseractOCR
}

Write-Log ""
Write-Log "=========================================="
Write-Log "Setup Summary"
Write-Log "=========================================="

$allSuccess = $true
foreach ($tool in $results.Keys) {
    if ($results[$tool]) {
        Write-Log "$tool : OK" "OK"
    } else {
        Write-Log "$tool : FAILED" "ERROR"
        $allSuccess = $false
    }
}

# Run Python environment setup
if (-not $SkipPython) {
    Write-Log ""
    Write-Log "--- Python Environment Setup ---"
    
    $pythonSetupScript = Join-Path $scriptDir "Setup-PythonEnvironment.ps1"
    
    if (Test-Path $pythonSetupScript) {
        Write-Log "Running Python environment setup (CUDA: $CudaVersion)..."
        
        $pythonArgs = @("-AppRoot", $appRoot, "-CudaVersion", $CudaVersion)
        if ($Force) { $pythonArgs += "-Force" }
        if ($Silent) { $pythonArgs += "-Silent" }
        
        try {
            & $pythonSetupScript @pythonArgs
            if ($LASTEXITCODE -eq 0) {
                Write-Log "Python environment setup completed." "OK"
            } else {
                Write-Log "Python environment setup completed with warnings." "WARN"
            }
        } catch {
            Write-Log "Python environment setup failed: $_" "ERROR"
            $allSuccess = $false
        }
    } else {
        Write-Log "Python setup script not found: $pythonSetupScript" "ERROR"
        $allSuccess = $false
    }
} else {
    Write-Log ""
    Write-Log "Skipping Python environment setup (-SkipPython specified)" "WARN"
}

if ($allSuccess) {
    Write-Log ""
    Write-Log "All setup completed successfully!" "OK"
    exit 0
} else {
    Write-Log ""
    Write-Log "Some components failed to setup. Please check the log." "ERROR"
    exit 1
}
