[Setup] 
AppName=ActivityLogger
AppVersion=1.0
DefaultDirName={commonpf}\ActivityLogger
OutputDir=.
OutputBaseFilename=ActivityLogger_Setup

[Dirs]
Name: "{userappdata}\Local\ActivityLogger"

[Files]
Source: "dist\ActivityLogger.exe"; DestDir: "{app}"
Source: "README.md"; DestDir: "{app}"


[Icons]
Name: "{commondesktop}\ActivityLogger"; Filename: "{app}\ActivityLogger.exe"
Name: "{commonstartup}\ActivityLogger"; Filename: "{app}\ActivityLogger.exe"
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

