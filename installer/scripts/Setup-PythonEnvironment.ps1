# Setup-PythonEnvironment.ps1
# SuperBookTools Python環境セットアップスクリプト
# RealEsrgan と YomiToku の環境を構築します

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

# スクリプトのディレクトリからアプリルートを取得
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
    
    # ログファイルに出力
    $logFile = Join-Path $appRoot "setup.log"
    Add-Content -Path $logFile -Value $logMessage
}

function Find-Python {
    Write-Log "Pythonを検索中..."
    
    # PATHから検索
    $pythonCmd = Get-Command python -ErrorAction SilentlyContinue
    if ($pythonCmd) {
        $version = & python --version 2>&1
        Write-Log "PATHでPythonを検出: $($pythonCmd.Source) ($version)"
        return $pythonCmd.Source
    }
    
    # py ランチャーから検索
    $pyCmd = Get-Command py -ErrorAction SilentlyContinue
    if ($pyCmd) {
        $version = & py --version 2>&1
        Write-Log "py ランチャーを検出: $version"
        return "py"
    }
    
    # 一般的なインストール場所を検索
    $commonPaths = @(
        "C:\Program Files\Python312\python.exe",
        "C:\Program Files\Python311\python.exe",
        "C:\Python312\python.exe",
        "C:\Python311\python.exe"
    )
    
    foreach ($path in $commonPaths) {
        if (Test-Path $path) {
            $version = & $path --version 2>&1
            Write-Log "Pythonを検出: $path ($version)"
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
        Write-Log "RealEsrgan環境は既に存在します。スキップします。" "WARN"
        return $true
    }
    
    Write-Log "RealEsrgan環境をセットアップ中..."
    
    try {
        # ディレクトリ作成
        if (-not (Test-Path $toolsPath)) {
            New-Item -ItemType Directory -Path $toolsPath -Force | Out-Null
        }
        
        # 既存のvenvを削除（-Force時）
        if ((Test-Path $venvPath) -and $Force) {
            Write-Log "既存のvenv環境を削除中..."
            Remove-Item -Recurse -Force $venvPath
        }
        
        # venv作成
        Write-Log "仮想環境を作成中..."
        if ($PythonPath -eq "py") {
            & py -3 -m venv $venvPath
        } else {
            & $PythonPath -m venv $venvPath
        }
        
        # pip アップグレード
        Write-Log "pipをアップグレード中..."
        & "$venvPath\Scripts\python.exe" -m pip install --upgrade pip --quiet
        
        # PyTorchインストール
        Write-Log "PyTorchをインストール中 ($CudaVer)... これには数分かかります。"
        & "$venvPath\Scripts\pip.exe" install torch torchvision torchaudio --index-url "https://download.pytorch.org/whl/$CudaVer" --quiet
        
        # Real-ESRGANリポジトリをクローン
        $repoPath = Join-Path $toolsPath "Real-ESRGAN"
        if (-not (Test-Path $repoPath)) {
            Write-Log "Real-ESRGANリポジトリをクローン中..."
            git clone https://github.com/xinntao/Real-ESRGAN.git $repoPath --quiet
            Push-Location $repoPath
            git checkout a4abfb2979a7bbff3f69f58f58ae324608821e27 --quiet
            Pop-Location
        }
        
        # モデルダウンロード
        $weightsPath = Join-Path $repoPath "weights"
        if (-not (Test-Path $weightsPath)) {
            New-Item -ItemType Directory -Path $weightsPath -Force | Out-Null
        }
        $modelPath = Join-Path $weightsPath "RealESRGAN_x4plus.pth"
        if (-not (Test-Path $modelPath)) {
            Write-Log "RealESRGANモデルをダウンロード中..."
            Invoke-WebRequest -Uri "https://github.com/xinntao/Real-ESRGAN/releases/download/v0.1.0/RealESRGAN_x4plus.pth" -OutFile $modelPath
        }
        
        # requirements.txtインストール
        $reqPath = Join-Path $repoPath "requirements.txt"
        if (Test-Path $reqPath) {
            Write-Log "依存関係をインストール中..."
            & "$venvPath\Scripts\pip.exe" install -r $reqPath --quiet
        }
        
        # degradations.pyのパッチ
        $degradationsPath = Join-Path $venvPath "Lib\site-packages\basicsr\data\degradations.py"
        if (Test-Path $degradationsPath) {
            Write-Log "degradations.pyにパッチを適用中..."
            $content = Get-Content $degradationsPath -Raw
            $content = $content -replace 'from torchvision.transforms.functional_tensor import rgb_to_grayscale', 'from torchvision.transforms.functional import rgb_to_grayscale'
            Set-Content $degradationsPath $content
        }
        
        # version.py作成
        $versionPath = Join-Path $repoPath "realesrgan\version.py"
        if (-not (Test-Path $versionPath)) {
            $versionDir = Split-Path $versionPath -Parent
            if (-not (Test-Path $versionDir)) {
                New-Item -ItemType Directory -Path $versionDir -Force | Out-Null
            }
            New-Item -ItemType File -Path $versionPath -Force | Out-Null
        }
        
        Write-Log "RealEsrganセットアップ完了" "OK"
        return $true
    }
    catch {
        Write-Log "RealEsrganセットアップ失敗: $_" "ERROR"
        return $false
    }
}

function Setup-YomiToku {
    param([string]$PythonPath, [string]$CudaVer)
    
    $toolsPath = Join-Path $appRoot "external_tools\image_tools\yomitoku"
    $venvPath = Join-Path $toolsPath "venv"
    
    if ((Test-Path $venvPath) -and (-not $Force)) {
        Write-Log "YomiToku環境は既に存在します。スキップします。" "WARN"
        return $true
    }
    
    Write-Log "YomiToku環境をセットアップ中..."
    
    try {
        # ディレクトリ作成
        if (-not (Test-Path $toolsPath)) {
            New-Item -ItemType Directory -Path $toolsPath -Force | Out-Null
        }
        
        # 既存のvenvを削除（-Force時）
        if ((Test-Path $venvPath) -and $Force) {
            Write-Log "既存のvenv環境を削除中..."
            Remove-Item -Recurse -Force $venvPath
        }
        
        # venv作成
        Write-Log "仮想環境を作成中..."
        if ($PythonPath -eq "py") {
            & py -3 -m venv $venvPath
        } else {
            & $PythonPath -m venv $venvPath
        }
        
        # pip アップグレード
        Write-Log "pipをアップグレード中..."
        & "$venvPath\Scripts\python.exe" -m pip install --upgrade pip --quiet
        
        # PyTorchインストール
        Write-Log "PyTorchをインストール中 ($CudaVer)... これには数分かかります。"
        & "$venvPath\Scripts\pip.exe" install torch torchvision torchaudio --index-url "https://download.pytorch.org/whl/$CudaVer" --quiet
        
        # YomiTokuインストール
        Write-Log "YomiTokuをインストール中..."
        & "$venvPath\Scripts\pip.exe" install "yomitoku==0.10.3" --quiet
        
        Write-Log "YomiTokuセットアップ完了" "OK"
        return $true
    }
    catch {
        Write-Log "YomiTokuセットアップ失敗: $_" "ERROR"
        return $false
    }
}

# メイン処理
function Main {
    Write-Log "=========================================="
    Write-Log "SuperBookTools Python環境セットアップ"
    Write-Log "=========================================="
    Write-Log "アプリケーションルート: $appRoot"
    Write-Log "CUDAバージョン: $CudaVersion"
    Write-Log ""
    
    # Gitの確認
    $gitCmd = Get-Command git -ErrorAction SilentlyContinue
    if (-not $gitCmd) {
        Write-Log "Gitがインストールされていません。" "ERROR"
        Write-Log "https://git-scm.com/download/win からGitをインストールしてください。"
        if (-not $Silent) {
            Read-Host "Enterキーを押して終了"
        }
        exit 1
    }
    
    # Pythonの確認
    $pythonPath = Find-Python
    if (-not $pythonPath) {
        Write-Log "Pythonが見つかりません。" "ERROR"
        Write-Log "https://www.python.org/downloads/ からPython 3.11または3.12をインストールしてください。"
        Write-Log "インストール時に「Add Python to PATH」と「Install for all users」を選択してください。"
        if (-not $Silent) {
            Read-Host "Enterキーを押して終了"
        }
        exit 1
    }
    
    Write-Log ""
    Write-Log "--- RealEsrgan セットアップ ---"
    $realEsrganResult = Setup-RealEsrgan -PythonPath $pythonPath -CudaVer $CudaVersion
    
    Write-Log ""
    Write-Log "--- YomiToku セットアップ ---"
    $yomitokuResult = Setup-YomiToku -PythonPath $pythonPath -CudaVer $CudaVersion
    
    Write-Log ""
    Write-Log "=========================================="
    if ($realEsrganResult -and $yomitokuResult) {
        Write-Log "すべてのセットアップが完了しました！" "OK"
        $exitCode = 0
    } else {
        Write-Log "一部のセットアップに失敗しました。" "ERROR"
        Write-Log "setup.log を確認してください。"
        $exitCode = 1
    }
    Write-Log "=========================================="
    
    if (-not $Silent) {
        Read-Host "Enterキーを押して終了"
    }
    
    exit $exitCode
}

# 実行
Main
