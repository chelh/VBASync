#define MyAppName "VBA Sync Tool"
#define MyAppVersion "1.2.0"
#define MyAppPublisher "Chelsea Hughes"
#define MyAppURL "https://github.com/chelh/VBASync"
#define MyAppExeName "VBA Sync Tool.exe"

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
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "french"; MessagesFile: "compiler:Languages\French.isl"

[Files]
Source: "src\VBASync\bin\Release\VBA Sync Tool.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "src\VBACompressionCodec.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "LICENSE.rtf"; DestDir: "{app}"; Flags: ignoreversion
Source: "3RDPARTY\DiffPlex.txt"; DestDir: "{app}\3RDPARTY"; Flags: ignoreversion
Source: "3RDPARTY\Ookii.Dialogs.txt"; DestDir: "{app}\3RDPARTY"; Flags: ignoreversion
Source: "3RDPARTY\OpenMCDF.txt"; DestDir: "{app}\3RDPARTY"; Flags: ignoreversion
Source: "3RDPARTY\SharpZipLib.txt"; DestDir: "{app}\3RDPARTY"; Flags: ignoreversion
Source: "3RDPARTY\VBACompressionCodec.txt"; DestDir: "{app}\3RDPARTY"; Flags: ignoreversion
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

