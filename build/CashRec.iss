#ifndef PublishDir
  #define PublishDir "..\artifacts\publish"
#endif
#ifndef OutputDir
  #define OutputDir "..\artifacts\installer"
#endif
[Setup]
AppId={{761E54F2-4705-485F-A63B-8DDC75A6A96B}
AppName=Cash Rec
AppVersion=0.1.0
AppPublisher=Cash Rec
DefaultDirName={localappdata}\Programs\Cash Rec
DefaultGroupName=Cash Rec
OutputBaseFilename=CashRec_Setup
OutputDir={#OutputDir}
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=10.0
PrivilegesRequired=lowest
UninstallDisplayIcon={app}\CashRec.exe
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
DisableProgramGroupPage=yes
[Files]
Source: "{#PublishDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
[Icons]
Name: "{group}\Cash Rec"; Filename: "{app}\CashRec.exe"
Name: "{autodesktop}\Cash Rec"; Filename: "{app}\CashRec.exe"
[Run]
Filename: "{app}\CashRec.exe"; Description: "Launch Cash Rec"; Flags: nowait postinstall skipifsilent
