#define MyAppName "Gandalf OSINT"
#define MyAppExeName "GandalfOSINT.exe"
#define MyAppVersion GetEnv("GANDALF_VERSION")
#if MyAppVersion == ""
#define MyAppVersion "dev"
#endif

[Setup]
AppId={{7D99E702-2395-4E1F-93AC-16EB71089002}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher=Theotime Massard
DefaultDirName={localappdata}\Programs\GandalfOSINT
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=installer
OutputBaseFilename=GandalfOSINT-Setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
UninstallDisplayIcon={app}\{#MyAppExeName}
CloseApplications=yes
CloseApplicationsFilter={#MyAppExeName}
RestartApplications=no

[Languages]
Name: "french"; MessagesFile: "compiler:Languages\French.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent
