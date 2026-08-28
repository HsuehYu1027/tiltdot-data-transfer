; Inno Setup 腳本：把 dist\tiltdot 打包成安裝檔
; 使用方式：先跑 build\build.bat，再用 Inno Setup Compiler 開啟本檔按 Build。
;
; 刻意安裝到 %LOCALAPPDATA%\Programs 而不是 Program Files：
; 程式預設把「初始檔案」「修改過的」建在 exe 旁邊，裝在 Program Files 會被權限擋住。

#define MyAppName "tiltDot 資料轉換工具"
#define MyAppShortName "tiltdot"
#define MyAppVersion "1.1.0"
#define MyAppExeName "tiltdot.exe"

[Setup]
AppId={{9C4A1F2E-7B3D-4E58-9A61-2F0D5C8E1B44}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
DefaultDirName={localappdata}\Programs\{#MyAppShortName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=..\dist
OutputBaseFilename=tiltdot-setup-{#MyAppVersion}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "chinesetrad"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "建立桌面捷徑"; GroupDescription: "附加捷徑:"

[Files]
Source: "..\dist\tiltdot\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "立即執行"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; 使用者的設定檔隨程式一起移除；產出的資料夾保留不動
Type: files; Name: "{app}\tiltdot_config.json"
