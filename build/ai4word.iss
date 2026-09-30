; AI4Word 安装包脚本（Inno Setup 6）
; 用法：ISCC build\ai4word.iss   （build\build.py 会自动调用）
#define MyAppName "AI4Word"
#define MyAppVersion "8.0"
#define MyAppPublisher "AI4Word"
#define MyAppExeName "AI4Word.exe"
#define MyAppDir "..\dist\AI4Word"

[Setup]
AppId={{A4B9C2D1-8F30-4E66-9B7C-3D5E1F2A8C04}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=..\dist
OutputBaseFilename=AI4Word-Setup-{#MyAppVersion}
SetupIconFile=..\assets\app_icon.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
Compression=lzma2/ultra
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64
PrivilegesRequired=lowest
CloseApplications=force

[Languages]
{LANGUAGES}

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "startup"; Description: "开机自动启动"; GroupDescription: "其他选项:"

[Files]
Source: "{#MyAppDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\卸载 {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon
Name: "{autostartup}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: startup

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#MyAppName}}"; Flags: nowait postinstall skipifsilent

[UninstallRun]
; 卸载时清理开机自启注册项（若由程序内自己管理则兜底）
Filename: "{cmd}"; Parameters: "/c reg delete HKCU\Software\Microsoft\Windows\CurrentVersion\Run /v AI4Word /f"; Flags: runhidden; RunOnceId: "RemoveAutostart"
