; -- Example1.iss --
; Demonstrates copying 3 files and creating an icon.

; SEE THE DOCUMENTATION FOR DETAILS ON CREATING .ISS SCRIPT FILES!

[Setup]
AppName=GTA2 Game Hunter
AppVersion=1.5992
WizardStyle=modern
DefaultDirName={autopf}\GTA2 Game Hunter
DefaultGroupName=GTA2 Game Hunter
UninstallDisplayIcon={app}\gta2gh.exe
Compression=lzma2
SolidCompression=yes
;OutputDir=userdocs:Inno Setup Examples Output

[Files]
Source: "gta2gh.exe"; DestDir: "{app}"
Source: "7za.exe"; DestDir: "{app}"
Source: "mswinsck.ocx"; DestDir: "{sys}"
Source: "richtx32.ocx"; DestDir: "{sys}"
Source: "mscomctl.ocx"; DestDir: "{sys}"


[Icons]
Name: "{group}\GTA2 Game Hunter"; Filename: "gta2gh.ico"
;Name: "{group}\GTA2 Game Hunter"; Filename: "{app}\gta2gh.exe"

[Code]
procedure InitializeWizard;
begin
RegisterServer(False, ExpandConstant('{sys}\mswinsck.ocx'), False);
RegisterServer(False, ExpandConstant('{sys}\richtx32.ocx'), False);
RegisterServer(False, ExpandConstant('{sys}\mscomctl.ocx'), False);
end;

[Run]
;Filename: "{app}\README.TXT"; Description: "View the README file"; Flags: postinstall shellexec skipifsilent
Filename: "{app}\gta2gh.exe"; Description: "Launch application"; Flags: postinstall nowait skipifsilent


