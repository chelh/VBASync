#define MyAppName "VBA Sync Tool"
#define MyAppVersion "2.2.0"
#define MyAppPublisher "Chelsea Hughes"
#define MyAppURL "https://github.com/chelh/VBASync"
#define MyAppExeName "VBASync.exe"

[Setup]
AppId={{FCE92422-DABC-447E-8DC4-504C206D2784}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
;AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={pf}\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
LicenseFile=C:\Users\Chel\Documents\VBA Sync Tool\LICENSE.rtf
OutputBaseFilename=Install {#MyAppName} {#MyAppVersion}
OutputDir=dist
Compression=lzma
SolidCompression=yes

[Languages]
Name: "en"; MessagesFile: "compiler:Default.isl"
Name: "fr"; MessagesFile: "compiler:Languages\French.isl"

[Files]
Source: "src\VBASync\bin\Release\VBASync.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "src\VBASync.WPF\bin\Release\VBASync.WPF.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "src\VBACompressionCodec.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "LICENSE.rtf"; DestDir: "{app}"; Flags: ignoreversion
Source: "3RDPARTY\*"; DestDir: "{app}\3RDPARTY"; Flags: ignoreversion
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[InstallDelete]
Type: files; Name: "{app}\VBA Sync Tool.exe"

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[INI]
Filename: "{app}\VBASync.ini"; Section: "General"; Key: "Language"; String: """{language}"""; Flags: createkeyifdoesntexist

[UninstallDelete]
Type: files; Name: "{app}\VBASync.ini"
