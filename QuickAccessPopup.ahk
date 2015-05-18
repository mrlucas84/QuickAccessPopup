;===============================================
/*

Quick Access Popup
Written using AutoHotkey_L v1.1.09.03+ (http://ahkscript.org/)
By Jean Lalonde (JnLlnd on AHKScript.org forum)
	
Based on FoldersPopup from the same author
https://github.com/JnLlnd/FoldersPopup
initialy inspired by Robert Ryan's script DirMenu v2 (rbrtryn on AutoHotkey.com forum)
http://www.autohotkey.com/board/topic/91109-favorite-folders-popup-menu-with-gui/
who was maybe inspired by Savage's script FavoriteFolders
http://www.autohotkey.com/docs/scripts/FavoriteFolders.htm
or Rexx version Folder Menu
http://www.autohotkey.com/board/topic/13392-folder-menu-a-popup-menu-to-quickly-change-your-folders/


BUGS

TO-DO

LATER
-----
HELP
* Update links to QAP website in Help
* Update links to QAP reviews in Donate

LANGUAGE
* Replace or update occurences of "FoldersPopup" in language files

GUI
* update CleanUpBeforeExit to save current position


Version 6.0.2 alpha (2015-05-??)


Version: 6.0.1 alpha (2015-05-11)
* Replace "FoldersPopup" with "QuickAccessPopup"
* Update @Ahk2Exe-SetVersion with "6.0.1 alpha"
* Update strCurrentVersion with "6.0.1 alpha"
* Update @Ahk2Exe-SetDescription with "Most handy Windows launcher. Freeware!"
* Distinct variables strAppNameFile for "QuickAccessPopup" and strAppNameText for "Quick Access Popup"
* Update strCurrentBranch with "alpha"
* Adapt for alpha version without version checking for alpha branch
* Replace "FoldersPopup" with "QuickAccessPopup" in InitFileInstall, and language variable names
* Replace "strTempDir" with "g_strTempDir"

SEE PREVIOUS HISTORY on FoldersPopup's GitHub or in FoldersPopup.ahk file


VARIABLES NAMING CONVENTION
---------------------------

typNameOfVariable
^^^^^^^^^^^^^^^^^ description of the variable content, with name section from general to specific

typeNameOfVariable
^^^^ type of variable, str for strings, int for integers (any size), dbl for reals (not used in this app),
     arr for arrays, obj for objects, menu for menus, etc.
  
g_typNameOfVariable
^ g_ for global, nothing for local

*/ 
;========================================================================================================================
; --- COMPILER DIRECTIVES ---
;========================================================================================================================

; Doc: http://fincs.ahk4.net/Ahk2ExeDirectives.htm
; Note: prefix comma with `

;@Ahk2Exe-SetName Quick Access Popup
;@Ahk2Exe-SetDescription Quick Access Popup - Freeware launcher for Windows.
;@Ahk2Exe-SetVersion 6.0.1 alpha
;@Ahk2Exe-SetOrigFilename QuickAccessPopup.exe


;========================================================================================================================
; INITIALIZATION
;========================================================================================================================

#NoEnv
#SingleInstance force
#KeyHistory 0
ListLines, Off
DetectHiddenWindows, On
ComObjError(False) ; we will do our own error handling

; avoid error message when shortcut destination is missing
; see http://ahkscript.org/boards/viewtopic.php?f=5&t=4477&p=25239#p25236
DllCall("SetErrorMode", "uint", SEM_FAILCRITICALERRORS := 1)

; By default, the A_WorkingDir is A_ScriptDir.
; When the shortcut is created by Inno Setup, the working is set to the folder under {userappdata}.
; In portable mode, the user can set the working directory in his own Windows shortcut.
; If user enable "Run at startup", the "Start in:" shortcut option is set to the current A_WorkingDir.

; If A_WorkingDir equals A_ScriptDir and the file _do_not_remove_or_rename.txt is found in A_WorkingDir
; it means that QAP has been installed with the setup program but that it was launched directly in the
; Program Files directory instead of using the Start menu or Startup shortcuts. In this situation, we
; know that the working directory has not been set properly. The following lines will fix it.
if (A_WorkingDir = A_ScriptDir) and FileExist(A_WorkingDir . "\_do_not_remove_or_rename.txt")
	SetWorkingDir, %A_AppData%\QuickAccessPopup

; Force A_WorkingDir to A_ScriptDir if uncomplied (development environment)
;@Ahk2Exe-IgnoreBegin
; Start of code for development environment only - won't be compiled
; see http://fincs.ahk4.net/Ahk2ExeDirectives.htm
SetWorkingDir, %A_ScriptDir%
; to test user data directory: SetWorkingDir, %A_AppData%\QuickAccessPopup

ListLines, On
; / End of code for developement enviuronment only - won't be compiled
;@Ahk2Exe-IgnoreEnd

OnExit, CleanUpBeforeExit ; must be positioned before InitFileInstall to ensure deletion of temporary files

Gosub, InitFileInstall

Gosub, InitLanguageVariables

; --- Super-global variables

global g_strAppNameFile := "QuickAccessPopup"
global g_strAppNameText := "Quick Access Popup"
global g_strCurrentVersion := "6.0.1" ; "major.minor.bugs" or "major.minor.beta.release"
global g_strCurrentBranch := "alpha" ; "prod", "beta" or "alpha", always lowercase for filename
global g_strAppVersion := "v" . g_strCurrentVersion . (g_strCurrentBranch <> "prod" ? " " . g_strCurrentBranch : "")

global g_blnDiagMode := False
global g_strDiagFile := A_WorkingDir . "\" . g_strAppNameFile . "-DIAG.txt"
global g_strIniFile := A_WorkingDir . "\" . g_strAppNameFile . ".ini"
global g_blnMenuReady := false

global g_arrSubmenuStack := Object()
global g_arrSubmenuStackPosition := Object()
global g_objIconsFile := Object()
global g_objIconsIndex := Object()

global g_strMenuPathSeparator := ">"
global g_strGuiMenuSeparator := "----------------"
global g_strGuiMenuColumnBreak := "==="

if InStr(A_ScriptDir, A_Temp) ; must be positioned after g_strAppNameFile is created
; if the app runs from a zip file, the script directory is created under the system Temp folder
{
	Oops(lOopsZipFileError, g_strAppNameFile)
	ExitApp
}

;@Ahk2Exe-IgnoreBegin
; Start of code for developement environment only - won't be compiled
if (A_ComputerName = "JEAN-PC") ; for my home PC
	g_strIniFile := A_WorkingDir . "\" . g_strAppNameFile . "-HOME.ini"
else if InStr(A_ComputerName, "STIC") ; for my work hotkeys
	g_strIniFile := A_WorkingDir . "\" . g_strAppNameFile . "-WORK.ini"
; / End of code for developement environment only - won't be compiled
;@Ahk2Exe-IgnoreEnd

; Keep gosubs in this order
Gosub, InitSystemArrays
Gosub, InitLanguages
Gosub, InitLanguageArrays

###_D(1) ; ### REMOVE WHEN SCRIOPT PERSISTENT
ExitApp ; ### REMOVE WHEN SCRIOPT PERSISTENT

return


;========================================================================================================================
; INITIALIZATION
;========================================================================================================================

;-----------------------------------------------------------
InitFileInstall:
;-----------------------------------------------------------

g_strTempDir := A_WorkingDir . "\_temp"
FileCreateDir, %g_strTempDir%

FileInstall, FileInstall\QuickAccessPopup_LANG_DE.txt, %g_strTempDir%\QuickAccessPopup_LANG_DE.txt, 1
FileInstall, FileInstall\QuickAccessPopup_LANG_FR.txt, %g_strTempDir%\QuickAccessPopup_LANG_FR.txt, 1
FileInstall, FileInstall\QuickAccessPopup_LANG_NL.txt, %g_strTempDir%\QuickAccessPopup_LANG_NL.txt, 1
FileInstall, FileInstall\QuickAccessPopup_LANG_KO.txt, %g_strTempDir%\QuickAccessPopup_LANG_KO.txt, 1
FileInstall, FileInstall\QuickAccessPopup_LANG_SV.txt, %g_strTempDir%\QuickAccessPopup_LANG_SV.txt, 1
FileInstall, FileInstall\QuickAccessPopup_LANG_IT.txt, %g_strTempDir%\QuickAccessPopup_LANG_IT.txt, 1
FileInstall, FileInstall\QuickAccessPopup_LANG_ES.txt, %g_strTempDir%\QuickAccessPopup_LANG_ES.txt, 1
FileInstall, FileInstall\QuickAccessPopup_LANG_PT-BR.txt, %g_strTempDir%\QuickAccessPopup_LANG_PT-BR.txt, 1

FileInstall, FileInstall\default_browser_icon.html, %g_strTempDir%\default_browser_icon.html, 1

FileInstall, FileInstall\about-32.png, %g_strTempDir%\about-32.png
FileInstall, FileInstall\add_property-48.png, %g_strTempDir%\add_property-48.png
FileInstall, FileInstall\delete_property-48.png, %g_strTempDir%\delete_property-48.png
FileInstall, FileInstall\channel_mosaic-48.png, %g_strTempDir%\channel_mosaic-48.png
FileInstall, FileInstall\separator-26.png, %g_strTempDir%\separator-26.png
FileInstall, FileInstall\column-26.png, %g_strTempDir%\column-26.png
FileInstall, FileInstall\down_circular-26.png, %g_strTempDir%\down_circular-26.png
FileInstall, FileInstall\edit_property-48.png, %g_strTempDir%\edit_property-48.png
FileInstall, FileInstall\generic_sorting2-26-grey.png, %g_strTempDir%\generic_sorting2-26-grey.png
FileInstall, FileInstall\help-32.png, %g_strTempDir%\help-32.png
FileInstall, FileInstall\left-12.png, %g_strTempDir%\left-12.png
FileInstall, FileInstall\settings-32.png, %g_strTempDir%\settings-32.png
FileInstall, FileInstall\up-12.png, %g_strTempDir%\up-12.png
FileInstall, FileInstall\up_circular-26.png, %g_strTempDir%\up_circular-26.png

FileInstall, FileInstall\thumbs_up-32.png, %g_strTempDir%\thumbs_up-32.png
FileInstall, FileInstall\solutions-32.png, %g_strTempDir%\solutions-32.png
FileInstall, FileInstall\handshake-32.png, %g_strTempDir%\handshake-32.png
FileInstall, FileInstall\conference-32.png, %g_strTempDir%\conference-32.png
FileInstall, FileInstall\gift-32.png, %g_strTempDir%\gift-32.png

return
;-----------------------------------------------------------


;-----------------------------------------------------------
InitLanguageVariables:
;-----------------------------------------------------------

#Include %A_ScriptDir%\QuickAccessPopup_LANG.ahk

return
;-----------------------------------------------------------


;-----------------------------------------------------------
InitSystemArrays:
;-----------------------------------------------------------

; Hotkeys: ini names, hotkey variables name, default values, gosub label and Gui hotkey titles
strIniKeyNames := "LaunchHotkeyMouse|LaunchHotkeyKeyboard|NavigateHotkeyMouse|NavigateHotkeyKeyboard|PowerHotkeyMouse|PowerHotkeyKeyboard|SettingsHotkey"
StringSplit, g_arrIniKeyNames, strIniKeyNames, |
strHotkeyDefaults := "MButton|#a|+MButton|+#a|!MButton|!#a|+^s"
StringSplit, g_arrHotkeyDefaults, strHotkeyDefaults, |
; in QAP, strHotkeyLabels and g_arrHotkeyLabels = strIniKeyNames and g_arrIniKeyNames

g_strMouseButtons := "None|LButton|MButton|RButton|XButton1|XButton2|WheelUp|WheelDown|WheelLeft|WheelRight|"
; leave last | to enable default value on the last item
StringSplit, g_arrMouseButtons, g_strMouseButtons, |

; Icon files and index tested on Win 7 and Win 8.1. Not tested on Win 10.
strIconsMenus := "lMenuDesktop|lMenuDocuments|lMenuPictures|lMenuMyComputer|lMenuNetworkNeighborhood|lMenuControlPanel|lMenuRecycleBin"
	. "|menuRecentFolders|menuGroupDialog|menuGroupExplorer|lMenuSpecialFolders|lMenuGroup|lMenuFoldersInExplorer"
	. "|lMenuRecentFolders|lMenuSettings|lMenuAddThisFolder|lDonateMenu|Submenu|Network|UnknownDocument|Folder"
	. "|menuGroupSave|menuGroupLoad|lMenuDownloads|Templates|MyMusic|MyVideo|History|Favorites|Temporary|Winver"
	. "|Fonts|Application|Clipboard"
strIconsFile := "imageres|imageres|imageres|imageres|imageres|imageres|imageres"
			. "|imageres|imageres|imageres|imageres|shell32|imageres"
			. "|imageres|imageres|imageres|imageres|shell32|imageres|shell32|shell32"
			. "|shell32|shell32|imageres|shell32|imageres|imageres|shell32|shell32|shell32|winver"
			. "|shell32|shell32|shell32"
strIconsIndex := "106|189|68|105|115|23|50"
			. "|113|176|203|203|99|176"
			. "|113|110|217|208|298|29|1|4"
			. "|297|46|176|55|104|179|240|87|153|1"
			. "|39|304|261"

StringSplit, arrIconsFile, strIconsFile, |
StringSplit, arrIconsIndex, strIconsIndex, |

Loop, Parse, strIconsMenus, |
{
	g_objIconsFile[A_LoopField] := A_WinDir . "\System32\" . arrIconsFile%A_Index% . (arrIconsFile%A_Index% = "winver" ? ".exe" : ".dll")
	g_objIconsIndex[A_LoopField] := arrIconsIndex%A_Index%
}
; example: g_objIconsFile["lMenuPictures"] and g_objIconsIndex["lMenuPictures"]

strIniKeyNames := ""
strHotkeyDefaults := ""
strIconsMenus := ""
strIconsFile := ""
strIconsIndex := ""
arrIconsFile := ""
arrIconsIndex := ""

return
;-----------------------------------------------------------


;------------------------------------------------------------
InitLanguages:
;------------------------------------------------------------

IfNotExist, %g_strIniFile%
	; read language code from ini file created by the Inno Setup script in the user data folder
	IniRead, g_strLanguageCode, % A_WorkingDir . "\" . g_strAppNameFile . "-setup.ini", Global , LanguageCode, EN
else
	IniRead, g_strLanguageCode, %g_strIniFile%, Global, LanguageCode, EN

strLanguageFile := g_strTempDir . "\" . g_strAppNameFile . "_LANG_" . g_strLanguageCode . ".txt"
strReplacementForSemicolon := "!r4nd0mt3xt!" ; for non-comment semi-colons ";" escaped as ";;"

if FileExist(strLanguageFile)
{
	FileRead, strLanguageStrings, %strLanguageFile%
	Loop, Parse, strLanguageStrings, `n, `r
	{
		if (SubStr(A_LoopField, 1, 1) <> ";") ; skip comment lines
		{
			StringSplit, arrLanguageBit, A_LoopField, `t
			if SubStr(arrLanguageBit1, 1, 1) = "l"
				%arrLanguageBit1% := arrLanguageBit2
			StringReplace, %arrLanguageBit1%, %arrLanguageBit1%, ``n, `n, All
			
			if InStr(%arrLanguageBit1%, ";;") ; preserve escaped ; in string
				StringReplace, %arrLanguageBit1%, %arrLanguageBit1%, % ";;", %strReplacementForSemicolon%, A
			if InStr(%arrLanguageBit1%, ";")
				%arrLanguageBit1% := Trim(SubStr(%arrLanguageBit1%, 1, InStr(%arrLanguageBit1%, ";") - 1)) ; trim comment from ; and trim spaces and tabs
			if InStr(%arrLanguageBit1%, strReplacementForSemicolon) ; restore escaped ; in string
				StringReplace, %arrLanguageBit1%, %arrLanguageBit1%, %strReplacementForSemicolon%, % ";", A
		}
	}
}
else
	g_strLanguageCode := "EN"

strLanguageFile := ""
strReplacementForSemicolon := ""
strLanguageStrings := ""
arrLanguageBit := ""

return
;------------------------------------------------------------


;------------------------------------------------------------
InitLanguageArrays:
;------------------------------------------------------------
StringSplit, g_arrOptionsTitles, lOptionsTitles, |
strOptionsLanguageCodes := "EN|FR|DE|NL|KO|SV|IT|ES|PT-BR"
StringSplit, g_arrOptionsLanguageCodes, strOptionsLanguageCodes, |
StringSplit, g_arrOptionsLanguageLabels, lOptionsLanguageLabels, |

loop, %g_arrOptionsLanguageCodes0%
	if (g_arrOptionsLanguageCodes%A_Index% = g_strLanguageCode)
		{
			g_strLanguageLabel := g_arrOptionsLanguageLabels%A_Index%
			break
		}

lOptionsMouseButtonsText := lOptionsMouseNone . "|" . lOptionsMouseButtonsText ; use lOptionsMouseNone because this is displayed
StringSplit, g_arrMouseButtonsText, lOptionsMouseButtonsText, |

strOptionsLanguageCodes := ""

return
;------------------------------------------------------------


;-----------------------------------------------------------
CleanUpBeforeExit:
;-----------------------------------------------------------

FileRemoveDir, %g_strTempDir%, 1 ; Remove all files and subdirectories

if (g_blnDiagMode)
{
	MsgBox, 52, %g_strAppNameText%, % L(lDiagModeExit, g_strAppNameText, g_strDiagFile) . "`n`n" . lDiagModeIntro . "`n`n" . lDiagModeSee
	IfMsgBox, Yes
		Run, %g_strDiagFile%
}
ExitApp
;-----------------------------------------------------------


;========================================================================================================================
; END OF INITIALIZATION
;========================================================================================================================



;========================================================================================================================
; VARIOUS FUNCTIONS
;========================================================================================================================

;------------------------------------------------
L(strMessage, objVariables*)
;------------------------------------------------
{
	Loop
	{
		if InStr(strMessage, "~" . A_Index . "~")
			StringReplace, strMessage, strMessage, ~%A_Index%~, % objVariables[A_Index], A
 		else
			break
	}
	
	return strMessage
}
;------------------------------------------------


;------------------------------------------------
Oops(strMessage, objVariables*)
;------------------------------------------------
{
	Gui, 1:+OwnDialogs
	MsgBox, 48, % L(lOopsTitle, g_strAppNameText, g_strAppVersion), % L(strMessage, objVariables*)
}
; ------------------------------------------------


