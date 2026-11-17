; Inno Setup script: turns dist\MouseColony (built by PyInstaller) into one Setup .exe.
; Built automatically by .github/workflows/build-installers.yml:
;   ISCC /DMyAppVersion=2.3.3 install\windows_installer.iss   ->   dist\MouseColony-2.3.3-Windows-Setup.exe
; Installs for the current user (no admin rights needed); an admin can choose "for all users".
#ifndef MyAppVersion
  #define MyAppVersion "0.0.0"
#endif

[Setup]
AppId={{8C4F2A1E-5B7D-4E2A-9C3F-1D6B8A0E7F42}
AppName=Mouse Colony
AppVersion={#MyAppVersion}
AppVerName=Mouse Colony {#MyAppVersion}
AppPublisher=Volgyi lab, University of Pecs
AppPublisherURL=https://jerjely.github.io/mouse-colony/
VersionInfoVersion={#MyAppVersion}
DefaultDirName={autopf}\Mouse Colony
DefaultGroupName=Mouse Colony
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
OutputDir=..\dist
OutputBaseFilename=MouseColony-{#MyAppVersion}-Windows-Setup
SetupIconFile=..\mousecolony.ico
UninstallDisplayIcon={app}\MouseColony.exe
UninstallDisplayName=Mouse Colony
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
CloseApplications=yes

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
Source: "..\dist\MouseColony\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[InstallDelete]
; an update replaces the program files completely (settings, backups and the workbook live elsewhere)
Type: filesandordirs; Name: "{app}\_internal"

[Icons]
Name: "{autoprograms}\Mouse Colony"; Filename: "{app}\MouseColony.exe"
Name: "{autodesktop}\Mouse Colony"; Filename: "{app}\MouseColony.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\MouseColony.exe"; Description: "{cm:LaunchProgram,Mouse Colony}"; Flags: nowait postinstall skipifsilent
