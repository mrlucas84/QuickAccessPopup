#define MyAppName "Quick Access Popup"
#define MyAppNameLower "quickaccesspopup"
#define MyAppPublisher "Jean Lalonde"
#define MyAppURL "http://www.QuickAccessPopup.com"
#define MyAppExeName "QuickAccessPopup.exe"
#define FPImportExeName "ImportFPsettings.exe"

#define MyAppVersion "v6.1.4 ALPHA"
#define MyVersionFileName "6_1_4-alpha"
#define FPImportVersionFileName "ImportFPsettings-0_3-ALPHA.exe"

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
OutputDir=C:\Dropbox\AutoHotkey\QuickAccessPopup\build-beta\
OutputBaseFilename={#MyAppNameLower}-setup-alpha
SetupIconFile=C:\Dropbox\AutoHotkey\QuickAccessPopup\QuickAccessPopup-ALPHA-red-512.ico
Compression=lzma
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64
AppMutex={#MyAppName}Mutex

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
; Name: "french"; MessagesFile: "compiler:Languages\French.isl"
; Name: "german"; MessagesFile: "compiler:Languages\German.isl"
; Name: "dutch"; MessagesFile: "compiler:Languages\Dutch.isl"
; Name: "korean"; MessagesFile: "compiler:Languages\Korean.isl"
; Name: "swedish"; MessagesFile: "compiler:Languages\Swedish.isl"
; Name: "italian"; MessagesFile: "compiler:Languages\Italian.isl"
; Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"
; Name: "brazilportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Files]
Source: "C:\Dropbox\AutoHotkey\QuickAccessPopup\build-beta\QuickAccessPopup-{#MyVersionFileName}-64-bit.exe"; DestDir: "{app}"; DestName: "QuickAccessPopup.exe"; Check: IsWin64; Flags: 64bit ignoreversion
; Source: "C:\Dropbox\AutoHotkey\QuickAccessPopup\build-beta\QuickAccessPopup-{#MyVersionFileName}-32-bit.exe"; DestDir: "{app}"; DestName: "QuickAccessPopup.exe"; Check: "not IsWin64"; Flags: 32bit ignoreversion
Source: "C:\Dropbox\AutoHotkey\QuickAccessPopup\build-beta\{#FPImportVersionFileName}"; DestDir: "{app}"; DestName: "ImportFPsettings.exe"
Source: "C:\Dropbox\AutoHotkey\QuickAccessPopup\build-beta\QAPconnect.ini"; DestDir: "{userappdata}"; DestName: "QAPconnect.ini"
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Dirs]
; for folderspopup.ini and _temp subfolder
Name: "{userappdata}\{#MyAppName}" 

[INI]
Filename: "{userappdata}\{#MyAppName}\{#MyAppName}-setup.ini"; Section: "Global"; Key: "LanguageCode"; String: "EN"; Languages: english
; Filename: "{userappdata}\{#MyAppName}\{#MyAppName}-setup.ini"; Section: "Global"; Key: "LanguageCode"; String: "FR"; Languages: french
; Filename: "{userappdata}\{#MyAppName}\{#MyAppName}-setup.ini"; Section: "Global"; Key: "LanguageCode"; String: "DE"; Languages: german
; Filename: "{userappdata}\{#MyAppName}\{#MyAppName}-setup.ini"; Section: "Global"; Key: "LanguageCode"; String: "NL"; Languages: dutch
; Filename: "{userappdata}\{#MyAppName}\{#MyAppName}-setup.ini"; Section: "Global"; Key: "LanguageCode"; String: "KO"; Languages: korean
; Filename: "{userappdata}\{#MyAppName}\{#MyAppName}-setup.ini"; Section: "Global"; Key: "LanguageCode"; String: "SV"; Languages: swedish
; Filename: "{userappdata}\{#MyAppName}\{#MyAppName}-setup.ini"; Section: "Global"; Key: "LanguageCode"; String: "IT"; Languages: italian
; Filename: "{userappdata}\{#MyAppName}\{#MyAppName}-setup.ini"; Section: "Global"; Key: "LanguageCode"; String: "ES"; Languages: spanish
; Filename: "{userappdata}\{#MyAppName}\{#MyAppName}-setup.ini"; Section: "Global"; Key: "LanguageCode"; String: "PT-BR"; Languages: brazilportuguese

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{userappdata}\{#MyAppName}"
Name: "{group}\Import Folders Popup Settings"; Filename: "{app}\ImportFPsettings.exe"; WorkingDir: "{userappdata}\{#MyAppName}"
Name: "{group}\{cm:ProgramOnTheWeb,{#MyAppName}}"; Filename: "{#MyAppURL}";
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"

[Run]
Filename: "{app}\ImportFPsettings.exe"; Flags: runhidden waituntilterminated; WorkingDir: "{userappdata}\{#MyAppName}"; Parameters: "/calledfromsetup"; Tasks: importfpsettings
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; WorkingDir: "{userappdata}\{#MyAppName}"; Flags: waituntilidle postinstall skipifsilent

[Tasks]
Name: importfpsettings; Description: "Import &Folders Popup settings and favorites (only for Folders Popup users)"; Flags: checkedonce

[UninstallDelete]
Type: files; Name: "{userstartup}\{#MyAppName}.lnk"


