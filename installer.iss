#define AppName "Monitor Epson L395"
#define AppExe "MonitorEpsonL395.exe"
[Setup]
AppId={{C63B6FA4-A2F1-4B76-92AC-7C4C1C4EE795}
AppName={#AppName}
AppVersion=0.1.0
DefaultDirName={localappdata}\Programs\MonitorEpsonL395
DefaultGroupName={#AppName}
PrivilegesRequired=lowest
OutputDir=output
OutputBaseFilename=Monitor_Epson_L395_Setup
Compression=lzma
SolidCompression=yes
ArchitecturesAllowed=x64compatible
WizardStyle=modern
[Files]
Source: "publish\*"; DestDir: "{app}"; Flags: recursesubdirs ignoreversion
[Icons]
Name: "{autoprograms}\Monitor Epson L395"; Filename: "{app}\{#AppExe}"
Name: "{autodesktop}\Monitor Epson L395"; Filename: "{app}\{#AppExe}"
[Run]
Filename: "{app}\{#AppExe}"; Description: "Abrir Monitor Epson L395"; Flags: nowait postinstall skipifsilent
