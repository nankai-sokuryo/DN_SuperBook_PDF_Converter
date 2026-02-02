# SuperBookTools インストーラ

このフォルダには、SuperBookToolsのWindowsインストーラをビルドするためのファイルが含まれています。

## ファイル構成

```
installer/
├── SuperBookTools.iss              # Inno Setupスクリプト
├── scripts/
│   ├── Setup-ExternalTools.ps1     # 外部ツール＋Python環境 統合セットアップスクリプト
│   └── Setup-PythonEnvironment.ps1 # Python環境セットアップスクリプト
└── tools/                          # ビルド時に自動ダウンロードされる外部ツール
```

## インストーラの特徴

### オンラインインストーラ方式
- インストーラ自体は軽量（約200MB）
- Python環境（RealEsrgan、YomiToku）は初回起動時またはセットアップ時に構築
- CUDAバージョンをインストール時に選択可能

### 含まれるもの
- SuperBookToolsアプリケーション本体（CLI版）
- SuperBookToolsアプリケーション本体（GUI版）
- ExifTool
- ImageMagick
- pdfcpu
- QPDF
- Tesseract OCR データ

### 含まれないもの（初回セットアップで構築）
- RealEsrgan Python環境（〜2.5GB）
- YomiToku Python環境（〜2.5GB）

## 開発環境セットアップ

開発時に`dotnet run`でGUIアプリを実行するには、外部ツールとPython環境のセットアップが必要です。

### 統合セットアップスクリプト（推奨）

```powershell
# リポジトリルートで実行
# 外部ツールのダウンロード + Python環境セットアップ + シンボリックリンク作成
& "installer\scripts\Setup-ExternalTools.ps1" -AppRoot "."

# CUDAバージョンを指定する場合
& "installer\scripts\Setup-ExternalTools.ps1" -AppRoot "." -CudaVersion cu130

# Python環境をスキップする場合（外部ツールのみ）
& "installer\scripts\Setup-ExternalTools.ps1" -AppRoot "." -SkipPython

# 強制再セットアップ
& "installer\scripts\Setup-ExternalTools.ps1" -AppRoot "." -Force
```

このスクリプトは以下を実行します：
1. **シンボリックリンク作成** - `SuperBookToolsGui\external_tools` → `external_tools`
2. **外部ツールダウンロード** - ImageMagick, QPDF, pdfcpu, exiftool
3. **Python環境セットアップ** - RealEsrgan, YomiToku（`-SkipPython`で省略可能）

### 個別セットアップ

#### Python環境のみセットアップ
```powershell
& "installer\scripts\Setup-PythonEnvironment.ps1" -AppRoot "." -CudaVersion cu126
```

## ローカルでのビルド方法

### 前提条件
1. [Inno Setup 6](https://jrsoftware.org/isinfo.php) をインストール
2. .NET 10.0 SDK
3. PowerShell 5.1以上

### ビルド手順

```powershell
# 1. CLI版アプリケーションをビルド
dotnet publish SuperBookToolsApp/SuperBookToolsApp.csproj -c Release -r win-x64 --self-contained true -o ./publish/win-x64

# 2. GUI版アプリケーションをビルド（同じディレクトリに出力してDLLを共有）
dotnet publish SuperBookToolsGui/SuperBookToolsGui.csproj -c Release -r win-x64 --self-contained true -o ./publish/win-x64

# 3. 外部ツールをダウンロード（手動またはスクリプトで）
# ./installer/tools/ に配置

# 4. インストーラをビルド
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
#define MyAppVersion "2.0.0"
```
