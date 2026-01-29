# SuperBookTools インストーラ

このフォルダには、SuperBookToolsのWindowsインストーラをビルドするためのファイルが含まれています。

## ファイル構成

```
installer/
├── SuperBookTools.iss          # Inno Setupスクリプト
├── scripts/
│   └── Setup-PythonEnvironment.ps1  # Python環境セットアップスクリプト
└── tools/                      # ビルド時に自動ダウンロードされる外部ツール
```

## インストーラの特徴

### オンラインインストーラ方式
- インストーラ自体は軽量（約200MB）
- Python環境（RealEsrgan、YomiToku）は初回起動時またはセットアップ時に構築
- CUDAバージョンをインストール時に選択可能

### 含まれるもの
- SuperBookToolsアプリケーション本体
- ExifTool
- ImageMagick
- pdfcpu
- QPDF
- Tesseract OCR データ

### 含まれないもの（初回セットアップで構築）
- RealEsrgan Python環境（〜2.5GB）
- YomiToku Python環境（〜2.5GB）

## ローカルでのビルド方法

### 前提条件
1. [Inno Setup 6](https://jrsoftware.org/isinfo.php) をインストール
2. .NET 6.0 SDK
3. PowerShell 5.1以上

### ビルド手順

```powershell
# 1. アプリケーションをビルド
dotnet publish SuperBookToolsApp/SuperBookToolsApp.csproj -c Release -r win-x64 --self-contained true -o ./publish/win-x64

# 2. 外部ツールをダウンロード（手動またはスクリプトで）
# ./installer/tools/ に配置

# 3. インストーラをビルド
& "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" ./installer/SuperBookTools.iss
```

出力先: `./installer_output/SuperBookTools_Setup_1.0.0.exe`

## GitHub Actionsでのビルド

main/masterブランチへのプッシュ時に自動的にインストーラがビルドされます。

ビルドされたインストーラは「Artifacts」からダウンロードできます。

## インストール後のPython環境セットアップ

インストール時にPython環境セットアップをスキップした場合、以下の方法で後からセットアップできます：

### 方法1: スタートメニューから
「SuperBookTools」→「Python環境セットアップ」を実行

### 方法2: コマンドラインから
```powershell
# デフォルト（cu126）でセットアップ
powershell -ExecutionPolicy Bypass -File "C:\SuperBookTools\scripts\Setup-PythonEnvironment.ps1"

# CUDAバージョンを指定
powershell -ExecutionPolicy Bypass -File "C:\SuperBookTools\scripts\Setup-PythonEnvironment.ps1" -CudaVersion cu130

# 強制再セットアップ
powershell -ExecutionPolicy Bypass -File "C:\SuperBookTools\scripts\Setup-PythonEnvironment.ps1" -Force
```

## バージョン更新方法

`SuperBookTools.iss` の以下の行を更新：

```iss
#define MyAppVersion "1.0.0"
```
