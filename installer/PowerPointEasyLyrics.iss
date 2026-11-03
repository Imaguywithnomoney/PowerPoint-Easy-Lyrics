[Setup]
AppId={{2F33A8BE-6525-48E6-84AE-D137D94F450F}
AppName=PowerPoint Easy Lyrics
AppVersion=1.0.0
DefaultDirName={autopf}\PowerPoint Easy Lyrics
DefaultGroupName=PowerPoint Easy Lyrics
OutputDir=output
OutputBaseFilename=PowerPointEasyLyrics-Installer
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional tasks:"

[Files]
Source: "..\dist\PowerPointEasyLyrics.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\PowerPoint Easy Lyrics"; Filename: "{app}\PowerPointEasyLyrics.exe"
Name: "{autodesktop}\PowerPoint Easy Lyrics"; Filename: "{app}\PowerPointEasyLyrics.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\PowerPointEasyLyrics.exe"; Description: "Launch PowerPoint Easy Lyrics"; Flags: nowait postinstall skipifsilent
