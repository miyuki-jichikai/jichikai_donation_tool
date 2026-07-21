; ===================================================================
; 三幸町自治会 寄付依頼書生成ツール インストーラースクリプト
; Inno Setup 6.x 用
; ===================================================================

#define MyAppName "三幸町自治会 寄付依頼書生成ツール"
#define MyAppVersion "2.9.0"
#define MyAppPublisher "久野耕司"
#define MyAppExeName "jichikai_donation_tool.exe"

[Setup]
; AppId は一度生成したら変更しないでください（再インストール・更新の識別に使用）
AppId={{8F2C5A1E-4B7D-4E3A-9C6F-1A2B3C4D5E6F}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
; Program Files は書き込み権限がなく、本ツールは実行中に progress.json や
; tmp_ad_images フォルダをインストール先に作成するため、ユーザーが自由に
; 書き込める場所（ユーザーごとの AppData\Local）にインストールする
DefaultDirName={localappdata}\JichikaiDonationTool
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
; 出力されるインストーラーEXEのファイル名（拡張子なし）
OutputBaseFilename=jichikai_donation_tool_setup
; アイコン（任意・あれば指定）
SetupIconFile=app_icon_final.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
; ユーザーごとのフォルダにインストールするため管理者権限は不要
PrivilegesRequired=lowest
; 64bit専用にする場合
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "japanese"; MessagesFile: "compiler:Languages\Japanese.isl"

[Tasks]
Name: "desktopicon"; Description: "デスクトップにアイコンを作成する"; GroupDescription: "追加のアイコン:"; Flags: unchecked

[Files]
; 本体EXE
Source: "dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
; poppler本体一式（PDF→画像変換に必要。配布先PCにpopplerが入っていなくても
; 動作するように、EXEと同じフォルダに poppler_bin として同梱する）
Source: "poppler_bin\*"; DestDir: "{app}\poppler_bin"; Flags: recursesubdirs ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent
