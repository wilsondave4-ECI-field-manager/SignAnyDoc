#define MyAppName "Doc Sign"
#define MyAppVersion "1.0.3"
#define MyAppPublisher "ECI Automation"
#define MyAppExeName "Doc Sign.exe"

[Setup]
AppId={{B841BB27-81CD-4A69-B14B-7C66B271C151}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\Programs\Doc Sign
DefaultGroupName=Doc Sign
OutputDir=installer-output
OutputBaseFilename=Doc Sign-Setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
UninstallDisplayIcon={app}\{#MyAppExeName}
; Do not use Inno Setup Restart Manager for this one-file PyInstaller app.
; It can falsely report Doc Sign as still using files even after uninstall/exit.
CloseApplications=no
RestartApplications=no

[Files]
Source: "dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "README.md"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\Doc Sign"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\Doc Sign"; Filename: "{app}\{#MyAppExeName}"

[Run]
Filename: "{sys}\netsh.exe"; Parameters: "advfirewall firewall delete rule name=""Doc Sign Local Signing"""; Flags: runhidden waituntilterminated
Filename: "{sys}\netsh.exe"; Parameters: "advfirewall firewall delete rule name=""Doc Sign Local Signing TCP 8765"""; Flags: runhidden waituntilterminated
Filename: "{sys}\netsh.exe"; Parameters: "advfirewall firewall add rule name=""Doc Sign Local Signing TCP 8765"" dir=in action=allow protocol=TCP localport=8765 profile=private program=""{app}\{#MyAppExeName}"" enable=yes"; Flags: runhidden waituntilterminated
Filename: "{app}\{#MyAppExeName}"; Description: "Launch Doc Sign"; Flags: nowait postinstall skipifsilent

[UninstallRun]
Filename: "{sys}\netsh.exe"; Parameters: "advfirewall firewall delete rule name=""Doc Sign Local Signing"""; Flags: runhidden waituntilterminated
Filename: "{sys}\netsh.exe"; Parameters: "advfirewall firewall delete rule name=""Doc Sign Local Signing TCP 8765"""; Flags: runhidden waituntilterminated
