[Setup] 
AppName=ActivityLogger
AppVersion=1.0
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
DefaultDirName={autopf}\ActivityLogger
OutputDir=.
OutputBaseFilename=ActivityLogger_Setup

[Dirs]
Name: "{userappdata}\Local\ActivityLogger"

[Files]
Source: "dist\ActivityLogger.exe"; DestDir: "{app}"
Source: "README.md"; DestDir: "{app}"


[Icons]
Name: "{commonprograms}\ActivityLogger"; Filename: "{app}\ActivityLogger.exe"
Name: "{commondesktop}\ActivityLogger"; Filename: "{app}\ActivityLogger.exe"
Name: "{commonstartup}\ActivityLogger"; Filename: "{app}\ActivityLogger.exe"

[Run]
Filename: "{app}\ActivityLogger.exe"; Description: "{cm:LaunchProgram,ActivityLogger}"; Flags: nowait postinstall
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!
