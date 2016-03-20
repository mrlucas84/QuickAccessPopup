#define MyAppName "Quick Access Popup"
#define MyAppNameLower "quickaccesspopup"
#define MyAppPublisher "Jean Lalonde"
#define MyAppURL "http://wwww.QuickAccessPopup.com"
#define MyAppExeName "QuickAccessPopup.exe"
#define FPImportVersionFileName "ImportFPsettings-1_0-32-bit.exe"

#define MyAppVersion "v7.1.5"
#define MyVersionFileName "7_1_5"

[Setup]
; NOTE: The value of AppId uniquely identifies this application.
; Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{BE9D760B-0D64-40BD-9F24-B5B8AB90131B}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={pf}\{#MyAppName}
DefaultGroupName={#MyAppName}
LicenseFile=C:\Dropbox\AutoHotkey\QuickAccessPopup\Inno Setup files\licence.txt
OutputDir=C:\Dropbox\AutoHotkey\QuickAccessPopup\build\
OutputBaseFilename={#MyAppNameLower}-setup
SetupIconFile=C:\Dropbox\AutoHotkey\QuickAccessPopup\build\QuickAccessPopup-512.ico
Compression=lzma
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64
AppMutex={#MyAppName}Mutex

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "french"; MessagesFile: "compiler:Languages\French.isl"
Name: "german"; MessagesFile: "compiler:Languages\German.isl"
; Name: "dutch"; MessagesFile: "compiler:Languages\Dutch.isl"
; Name: "korean"; MessagesFile: "compiler:Languages\Korean.isl"
Name: "swedish"; MessagesFile: "compiler:Languages\Swedish.isl"
; Name: "italian"; MessagesFile: "compiler:Languages\Italian.isl"
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"
Name: "brazilportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Dirs]
; repository for files to be copied to "{userappdata}\{#MyAppName}" at first QAP execution with quickaccesspopup.ini and _temp subfolder
Name: "{commonappdata}\{#MyAppName}" 

[Files]
Source: "C:\Dropbox\AutoHotkey\QuickAccessPopup\build\QuickAccessPopup-{#MyVersionFileName}-64-bit.exe"; DestDir: "{app}"; DestName: "QuickAccessPopup.exe"; Check: IsWin64; Flags: 64bit ignoreversion
Source: "C:\Dropbox\AutoHotkey\QuickAccessPopup\build\QuickAccessPopup-{#MyVersionFileName}-32-bit.exe"; DestDir: "{app}"; DestName: "QuickAccessPopup.exe"; Check: "not IsWin64"; Flags: 32bit ignoreversion
Source: "C:\Dropbox\AutoHotkey\QuickAccessPopup\build\{#FPImportVersionFileName}"; DestDir: "{app}"; DestName: "ImportFPsettings.exe"
; Source: "C:\Dropbox\AutoHotkey\QuickAccessPopup\build\OSVersion.exe"; DestDir: "{app}"; DestName: "OSVersion.exe"
; Source: "C:\Dropbox\AutoHotkey\QuickAccessPopup\build\QAPconnect.ini"; DestDir: "{commonappdata}\{#MyAppName}"; DestName: "QAPconnect.ini" -> now created by QAP from a default template
Source: "C:\Dropbox\AutoHotkey\QuickAccessPopup\build\_do_not_remove_or_rename.txt"; DestDir: "{app}"; DestName: "_do_not_remove_or_rename.txt"; Flags: ignoreversion
Source: "C:\Dropbox\AutoHotkey\QuickAccessPopup\build\QuickAccessPopup-512.ico"; DestDir: "{app}"; DestName: "QuickAccessPopup.ico"; Flags: ignoreversion
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[INI]
Filename: "{commonappdata}\{#MyAppName}\{#MyAppNameLower}-setup.ini"; Section: "Global"; Key: "LanguageCode"; String: "EN"; Languages: english
Filename: "{commonappdata}\{#MyAppName}\{#MyAppNameLower}-setup.ini"; Section: "Global"; Key: "LanguageCode"; String: "FR"; Languages: french
Filename: "{commonappdata}\{#MyAppName}\{#MyAppNameLower}-setup.ini"; Section: "Global"; Key: "LanguageCode"; String: "DE"; Languages: german
; Filename: "{commonappdata}\{#MyAppName}\{#MyAppNameLower}-setup.ini"; Section: "Global"; Key: "LanguageCode"; String: "NL"; Languages: dutch
; Filename: "{commonappdata}\{#MyAppName}\{#MyAppNameLower}-setup.ini"; Section: "Global"; Key: "LanguageCode"; String: "KO"; Languages: korean
Filename: "{commonappdata}\{#MyAppName}\{#MyAppNameLower}-setup.ini"; Section: "Global"; Key: "LanguageCode"; String: "SV"; Languages: swedish
; Filename: "{commonappdata}\{#MyAppName}\{#MyAppNameLower}-setup.ini"; Section: "Global"; Key: "LanguageCode"; String: "IT"; Languages: italian
Filename: "{commonappdata}\{#MyAppName}\{#MyAppNameLower}-setup.ini"; Section: "Global"; Key: "LanguageCode"; String: "ES"; Languages: spanish
Filename: "{commonappdata}\{#MyAppName}\{#MyAppNameLower}-setup.ini"; Section: "Global"; Key: "LanguageCode"; String: "PT-BR"; Languages: brazilportuguese

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{commonappdata}\{#MyAppName}"
Name: "{group}\Import Folders Popup Settings"; Filename: "{app}\ImportFPsettings.exe"; WorkingDir: "{commonappdata}\{#MyAppName}"
; Name: "{group}\OS Version Info"; Filename: "{app}\OSVersion.exe"
Name: "{group}\{cm:ProgramOnTheWeb,{#MyAppName}}"; Filename: "{#MyAppURL}";
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"

[Run]
Filename: "{app}\ImportFPsettings.exe"; Flags: runhidden waituntilterminated; WorkingDir: "{commonappdata}\{#MyAppName}"; Parameters: "/calledfromsetup"; Tasks: importfpsettings
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; WorkingDir: "{commonappdata}\{#MyAppName}"; Flags: waituntilidle postinstall skipifsilent

[Tasks]
Name: importfpsettings; Description: "Import &Folders Popup settings and favorites (only for Folders Popup users)"; Flags: unchecked

[UninstallDelete]
Type: files; Name: "{userstartup}\{#MyAppName}.lnk"
