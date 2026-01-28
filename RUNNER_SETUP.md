# GitHub Actions Self-Hosted Runner セットアップガイド

このドキュメントでは、Windows 11 にGitHub Actions セルフホストランナーをインストールする手順を説明します。

## 目次

- [前提条件](#前提条件)
- [インストール手順](#インストール手順)
- [実行方法](#実行方法)
- [ワークフローでの使用](#ワークフローでの使用)
- [サービス管理コマンド](#サービス管理コマンド)
- [トラブルシューティング](#トラブルシューティング)
- [アンインストール](#アンインストール)

---

## 前提条件

### システム要件
- Windows 11 (64-bit)
- PowerShell 5.1 以上
- インターネット接続
- GitHubリポジトリへの管理者アクセス権

### 必須ソフトウェア（事前インストール）

以下のソフトウェアを **「すべてのユーザー向け」** にインストールしてください。  
※ セルフホストランナーは NETWORK SERVICE アカウントで動作するため、ユーザー単位のインストールでは認識されません。

#### 1. Python 3.11 または 3.12（推奨）

- **ダウンロード**: https://www.python.org/downloads/
- **推奨バージョン**: **Python 3.11.x** または **Python 3.12.x**
- インストール時に以下を選択：
  - ✅ **Add Python to PATH**
  - ✅ **Install for all users**（Customize installationから選択）
- インストール先: `C:\Program Files\Python3xx\`

> ⚠️ **重要**:
> - **Microsoft Store版のPython**は使用しないでください。NETWORK SERVICEアカウントからアクセスできません。
> - **Python 3.13/3.14は使用しないでください**。機械学習ライブラリ（basicsr等）との互換性問題があります。

#### 2. Ghostscript 10.x

- **ダウンロード**: https://ghostscript.com/releases/gsdnld.html
- 64-bit版をインストール
- インストール先: `C:\Program Files\gs\gs10.x.x\`

#### 3. Git for Windows

- **ダウンロード**: https://git-scm.com/download/win
- インストール時にデフォルト設定でOK
- インストール先: `C:\Program Files\Git\`

---

## インストール手順

### Step 1: GitHubでRunnerトークンを取得

1. リポジトリの **Settings** を開く
2. 左メニューから **Actions** → **Runners** を選択
3. **New self-hosted runner** ボタンをクリック
4. **Windows** を選択
5. 表示される**トークン**をコピーしておく（有効期限あり）

### Step 2: Runnerのダウンロードと解凍

PowerShellを **管理者として実行** し、以下のコマンドを順番に実行します：

```powershell
# Runnerをインストールするフォルダを作成
mkdir C:\actions-runner
cd C:\actions-runner

# 最新のRunnerパッケージをダウンロード
# ※バージョン番号は最新版に置き換えてください
Invoke-WebRequest -Uri https://github.com/actions/runner/releases/download/v2.331.0/actions-runner-win-x64-2.331.0.zip -OutFile actions-runner-win-x64.zip

# ZIPファイルを解凍
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::ExtractToDirectory("$PWD\actions-runner-win-x64.zip", "$PWD")

# ダウンロードしたZIPファイルを削除（任意）
Remove-Item actions-runner-win-x64.zip
```

> 💡 **最新バージョンの確認**: https://github.com/actions/runner/releases

### Step 3: Runnerの設定

```powershell
# 設定スクリプトを実行
# <OWNER>: GitHubユーザー名またはOrganization名
# <REPO>: リポジトリ名
# <TOKEN>: Step 1で取得したトークン

.\config.cmd --url https://github.com/<OWNER>/<REPO> --token <TOKEN>
```

**設定時の質問と推奨回答:**

| 質問 | 推奨回答 |
|------|----------|
| Enter the name of the runner group | `Default` (Enterキー) |
| Enter the name of runner | `windows-11-runner` (任意の名前) |
| Enter any additional labels | `windows,x64,win11` (任意) |
| Enter name of work folder | `_work` (Enterキー) |
| **Would you like to run the runner as service? (Y/N)** | **`Y` (推奨)** |
| User account to use for the service | `NT AUTHORITY\NETWORK SERVICE` (Enterキー) |

---

## ワークフローでの使用

### セルフホストランナーを使用する場合

`.github/workflows/build.yml` を以下のように編集：

```yaml
jobs:
  build:
    # GitHub提供のランナーの場合
    # runs-on: windows-latest
    
    # セルフホストランナーの場合
    runs-on: self-hosted
```

### ラベルを使用して特定のランナーを指定

```yaml
jobs:
  build:
    runs-on: [self-hosted, Windows, X64]
```

### GitHub提供とセルフホストの両方を使用

```yaml
jobs:
  build-github:
    runs-on: windows-latest
    steps:
      - name: Build on GitHub Runner
        run: echo "GitHub hosted runner"

  build-self:
    runs-on: self-hosted
    steps:
      - name: Build on Self-hosted Runner
        run: echo "Self-hosted runner"
```

---

## サービス管理コマンド

```powershell
cd C:\actions-runner
```

| コマンド | 説明 |
|----------|------|
| `.\svc.cmd install` | サービスをインストール |
| `.\svc.cmd start` | サービスを開始 |
| `.\svc.cmd stop` | サービスを停止 |
| `.\svc.cmd status` | サービス状態を確認 |
| `.\svc.cmd uninstall` | サービスをアンインストール |

### サービスの状態確認（PowerShell）

```powershell
Get-Service -Name "actions.runner.*"
```

---

## トラブルシューティング

### 実行ポリシーエラー

```powershell
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### ファイアウォールの設定

以下のドメインへの接続を許可してください：
- `github.com`
- `api.github.com`
- `*.actions.githubusercontent.com`
- `codeload.github.com`
- `*.blob.core.windows.net`

### ログの確認

```powershell
# Runnerのログを確認
Get-Content C:\actions-runner\_diag\Runner_*.log -Tail 100

# Workerのログを確認
Get-Content C:\actions-runner\_diag\Worker_*.log -Tail 100
```

### Runnerがオフライン状態の場合

1. サービスの状態を確認: `.\svc.cmd status`
2. サービスを再起動: `.\svc.cmd stop` → `.\svc.cmd start`
3. ネットワーク接続を確認
4. トークンの有効期限を確認（再設定が必要な場合あり）

### 認証エラー（トークン期限切れ）

```powershell
cd C:\actions-runner
.\svc.cmd stop
.\config.cmd remove --token <REMOVE_TOKEN>
.\config.cmd --url https://github.com/<OWNER>/<REPO> --token <NEW_TOKEN>
.\svc.cmd start
```

---

## アンインストール

### Step 1: サービスを停止・削除

```powershell
cd C:\actions-runner
.\svc.cmd stop
.\svc.cmd uninstall
```

### Step 2: GitHubからRunnerを削除

```powershell
# 削除用トークンはGitHubの Settings > Actions > Runners から取得
.\config.cmd remove --token <REMOVE_TOKEN>
```

### Step 3: フォルダを削除

```powershell
cd C:\
Remove-Item -Recurse -Force C:\actions-runner
```

---

## 参考リンク

- [GitHub Actions Runner 公式リポジトリ](https://github.com/actions/runner)
- [Self-hosted runners ドキュメント](https://docs.github.com/en/actions/hosting-your-own-runners)
- [Runner リリースページ](https://github.com/actions/runner/releases)

---

*最終更新: 2026年1月27日*
