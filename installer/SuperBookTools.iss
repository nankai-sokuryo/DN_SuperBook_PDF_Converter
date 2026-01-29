; SuperBookTools Inno Setup Script
; オンラインインストーラ - 軽量なインストーラでPython環境は初回起動時に構築

#define MyAppName "SuperBookTools"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "Daiyuu Nobori"
#define MyAppURL "https://github.com/IPA-CyberLab/DN_SuperBook_PDF_Converter"
#define MyAppExeName "SuperBookToolsApp.exe"

[Setup]
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}/releases
DefaultDirName=C:\SuperBookTools
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
; インストーラ出力先
OutputDir=..\installer_output
OutputBaseFilename=SuperBookTools_Setup_{#MyAppVersion}
; 圧縮設定
Compression=lzma2/ultra64
SolidCompression=yes
; UAC設定（管理者権限が必要）
PrivilegesRequired=admin
; Windows 10以降
MinVersion=10.0
; アーキテクチャ
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
; アンインストール時のログ
UninstallLogMode=overwrite
; ウィザードスタイル
WizardStyle=modern
; アイコン（オプション - 後で追加可能）
; SetupIconFile=..\assets\icon.ico

[Languages]
Name: "japanese"; MessagesFile: "compiler:Languages\Japanese.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; アプリケーション本体
Source: "..\publish\win-x64\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

; 外部ツール（静的ツールのみ - Python環境は含まない）
Source: "..\installer\tools\exiftool-13.30_64\*"; DestDir: "{app}\external_tools\image_tools\exiftool-13.30_64"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\installer\tools\ImageMagick-portable-Q16-HDRI-x64\*"; DestDir: "{app}\external_tools\image_tools\ImageMagick-portable-Q16-HDRI-x64"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\installer\tools\pdfcpu\*"; DestDir: "{app}\external_tools\image_tools\pdfcpu"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\installer\tools\QPDF\*"; DestDir: "{app}\external_tools\image_tools\QPDF"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\installer\tools\TesseractOCR_Data\*"; DestDir: "{app}\external_tools\image_tools\TesseractOCR_Data"; Flags: ignoreversion recursesubdirs createallsubdirs

; セットアップスクリプト
Source: "..\installer\scripts\Setup-PythonEnvironment.ps1"; DestDir: "{app}\scripts"; Flags: ignoreversion

[Dirs]
; Python環境用のディレクトリを事前作成
Name: "{app}\external_tools\image_tools\RealEsrgan"
Name: "{app}\external_tools\image_tools\yomitoku"

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Python環境セットアップ"; Filename: "powershell.exe"; Parameters: "-ExecutionPolicy Bypass -File ""{app}\scripts\Setup-PythonEnvironment.ps1"""; WorkingDir: "{app}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
; インストール後にPython環境セットアップを実行（オプション）
Filename: "powershell.exe"; Parameters: "-ExecutionPolicy Bypass -File ""{app}\scripts\Setup-PythonEnvironment.ps1"" -CudaVersion cu126"; \
  StatusMsg: "Python環境をセットアップしています（これには数分かかる場合があります）..."; \
  Description: "Python環境をセットアップする（推奨）"; \
  Flags: postinstall skipifsilent runhidden waituntilterminated; \
  Check: ShouldSetupPython

[UninstallDelete]
; アンインストール時にPython環境を削除
Type: filesandordirs; Name: "{app}\external_tools\image_tools\RealEsrgan"
Type: filesandordirs; Name: "{app}\external_tools\image_tools\yomitoku"

[Code]
var
  SetupPythonPage: TInputOptionWizardPage;
  CudaVersionPage: TInputOptionWizardPage;
  SelectedCudaVersion: String;

function ShouldSetupPython: Boolean;
begin
  Result := SetupPythonPage.Values[0];
end;

procedure InitializeWizard;
begin
  // Python環境セットアップの選択ページ
  SetupPythonPage := CreateInputOptionPage(wpSelectTasks,
    'Python環境のセットアップ',
    'AI機能（OCR、画像超解像）を使用するにはPython環境が必要です。',
    'Python環境をインストール後にセットアップしますか？' + #13#10 + #13#10 +
    '※ 約5GBのディスク容量と、インターネット接続が必要です。' + #13#10 +
    '※ セットアップには10〜30分程度かかります。' + #13#10 +
    '※ 後から「Python環境セットアップ」ショートカットで実行することもできます。',
    True, False);
  SetupPythonPage.Add('Python環境をセットアップする（推奨）');
  SetupPythonPage.Values[0] := True;
  
  // CUDAバージョン選択ページ
  CudaVersionPage := CreateInputOptionPage(SetupPythonPage.ID,
    'CUDA バージョンの選択',
    'お使いのNVIDIAドライバーに対応したCUDAバージョンを選択してください。',
    'nvidia-smi コマンドで「CUDA Version」を確認できます。' + #13#10 + #13#10 +
    '不明な場合は「cu126（推奨）」を選択してください。',
    True, False);
  CudaVersionPage.Add('cu126 - CUDA 12.6（推奨、ドライバー 525.60以上）');
  CudaVersionPage.Add('cu128 - CUDA 12.8（ドライバー 555.42以上）');
  CudaVersionPage.Add('cu130 - CUDA 13.0（最新、ドライバー 570以上）');
  CudaVersionPage.Values[0] := True;
  
  SelectedCudaVersion := 'cu126';
end;

function NextButtonClick(CurPageID: Integer): Boolean;
begin
  Result := True;
  
  if CurPageID = CudaVersionPage.ID then
  begin
    if CudaVersionPage.Values[0] then
      SelectedCudaVersion := 'cu126'
    else if CudaVersionPage.Values[1] then
      SelectedCudaVersion := 'cu128'
    else if CudaVersionPage.Values[2] then
      SelectedCudaVersion := 'cu130'
    else
      SelectedCudaVersion := 'cu126';
  end;
end;

function ShouldSkipPage(PageID: Integer): Boolean;
begin
  Result := False;
  // Python環境をセットアップしない場合はCUDAバージョン選択をスキップ
  if (PageID = CudaVersionPage.ID) and (not SetupPythonPage.Values[0]) then
    Result := True;
end;

function GetCudaVersion(Param: String): String;
begin
  Result := SelectedCudaVersion;
end;

// インストール後に実行するコマンドのパラメータを動的に設定
function UpdateReadyMemo(Space, NewLine, MemoUserInfoInfo, MemoDirInfo, MemoTypeInfo, MemoComponentsInfo, MemoGroupInfo, MemoTasksInfo: String): String;
begin
  Result := '';
  Result := Result + MemoDirInfo + NewLine + NewLine;
  
  if SetupPythonPage.Values[0] then
  begin
    Result := Result + 'Python環境:' + NewLine;
    Result := Result + Space + 'セットアップする' + NewLine;
    Result := Result + Space + 'CUDAバージョン: ' + SelectedCudaVersion + NewLine + NewLine;
  end
  else
  begin
    Result := Result + 'Python環境:' + NewLine;
    Result := Result + Space + 'スキップ（後で手動セットアップ可能）' + NewLine + NewLine;
  end;
end;
