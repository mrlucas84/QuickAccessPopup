;===============================================
/*

Import Settings from Folders Popup to Quick Access Popup
Written using AutoHotkey_L v1.1.09.03+ (http://ahkscript.org/)
By Jean Lalonde (JnLlnd on AHKScript.org forum)
	
DESCRIPTION

Convert the settings (options and favorites) from Folders Popup (folderspopup.ini) to Quick Access Popup

*/ 
;========================================================================================================================
; --- COMPILER DIRECTIVES ---
;========================================================================================================================

; Doc: http://fincs.ahk4.net/Ahk2ExeDirectives.htm
; Note: prefix comma with `

;@Ahk2Exe-SetName ImportFPsettings
;@Ahk2Exe-SetDescription Import settings from Folders Popup to Quick Access Popup
;@Ahk2Exe-SetVersion 0.1 alpha
;@Ahk2Exe-SetOrigFilename ImportFPsettings.exe


;========================================================================================================================
; INITIALIZATION
;========================================================================================================================

#NoEnv
#SingleInstance force
#KeyHistory 0
ListLines, Off

; Force A_WorkingDir to A_ScriptDir if uncomplied (development environment)
;@Ahk2Exe-IgnoreBegin
; Start of code for development environment only - won't be compiled
; see http://fincs.ahk4.net/Ahk2ExeDirectives.htm
SetWorkingDir, %A_ScriptDir%
ListLines, On
; / End of code for developement enviuronment only - won't be compiled
;@Ahk2Exe-IgnoreEnd

g_strParam := %0% ; "/calledfromsetup" if called from setup program

global g_strAppNameFile := "ImportFPsettings"
global g_strAppNameText := "Import Settings - FP to QAP"
global g_strCurrentVersion := "0.1" ; "major.minor.bugs" or "major.minor.beta.release"
global g_strCurrentBranch := "alpha" ; "prod", "beta" or "alpha", always lowercase for filename
global g_strAppVersion := "v" . g_strCurrentVersion . (g_strCurrentBranch <> "prod" ? " " . g_strCurrentBranch : "")

global g_strFPIniFile

global g_strQAPIniFile := A_WorkingDir . "\QuickAccessPopup.ini"
global g_intIniLine

global g_strFPStartupShortcut := A_Startup . "\FoldersPopup.lnk"
global g_strFPUserDataIniFile := A_AppData . "\FoldersPopup\FoldersPopup.ini"

; ### if (g_strParam = "/calledfromsetup")
if !StrLen(FileExist(g_strQAPIniFile))
{
	MsgBox, 16, %g_strAppNameText% %g_strAppVersion%, Error: the Quick Access Popup settings file could not be found:`n`n%g_strQAPIniFile%`n`nPlease, re-install Quick Access Popup. This task will be interrupted.
	ExitApp
}
; ### else ask QAP for ini file

MsgBox, 64, %g_strAppNameText% %g_strAppVersion%, This task will:`n`n- let you select the Folders Popup settings file`n`n- ask if you want to import the Folders Popup "Options" settings`n(these settings will replace existing or default QAP settings)`n`n- ask if you want to import the Folders Popup "Favorites" settings`n(these favorites will be added to existing QAP favorites)

if !FileExist(g_strFPStartupShortcut) ; if a FP shortcut exists in startup folder, use its app folder
{
	FileGetShortcut, %g_strFPStartupShortcut%, , g_strFPStartupWorkingDir
	g_strFPIniFile := g_strFPStartupWorkingDir . "\FoldersPopup.ini"
}

if !StrLen(g_strFPIniFile) ; else, see if we have a FP folder under appdata
	if StrLen(FileExist(g_strFPUserDataIniFile))
		g_strCandidateFPIniFile := g_strFPUserDataIniFile

Gosub, ConfirmFPIniFile ; confirm that we use this ini file

if !StrLen(g_strFPIniFile) ; if not, select a file
	Loop
	{
		FileSelectFile, g_strCandidateFPIniFile, 1, , Select the FoldersPopup.ini to import, *.ini
		if StrLen(g_strCandidateFPIniFile)
			Gosub, ConfirmFPIniFile
		else
			ExitApp
	}
	until StrLen(g_strFPIniFile)

blnImportOptions := "Yes"
MsgBox, 35, %g_strAppNameText% %g_strAppVersion%, Import "Options" settings?
IfMsgBox, Cancel
	ExitApp
IfMsgBox, No
	blnImportOptions := "No"
IfMsgBox, Yes
	gosub, ImportOptionsSettings

blnImportFavorites := "Yes"
MsgBox, 35, %g_strAppNameText% %g_strAppVersion%, Add "Favorites" to settings?
IfMsgBox, Cancel
	ExitApp
IfMsgBox, No
	blnImportFavorites := "No"
IfMsgBox, Yes
	gosub, ImportFavoritesSettings

MsgBox, , %g_strAppNameText% %g_strAppVersion%, Task finished.`n`nFolders Popup file:`n%g_strFPIniFile%`n`nQuick Access Popup file:`n%g_strQAPIniFile%`n`nImported "Options": %blnImportOptions%`nAdded "Favorites": %blnImportFavorites%

return
;-----------------------------------------------------------


;-----------------------------------------------------------
ConfirmFPIniFile:
;-----------------------------------------------------------

MsgBox, 35, %g_strAppNameText% %g_strAppVersion%, Import settings from Folders Popup file:`n%g_strCandidateFPIniFile%?
IfMsgBox, Cancel
	ExitApp
IfMsgBox, Yes
	g_strFPIniFile := g_strCandidateFPIniFile

return
;-----------------------------------------------------------


;-----------------------------------------------------------
ImportOptionsSettings:
;-----------------------------------------------------------

/*

ADAPT:
NavigateOrLaunchHotkeyMouse=MButton (FP = PopupHotkeyMouse)
NavigateOrLaunchHotkeyKeyboard=#A (FP = PopupHotkeyKeyboard)
PowerHotkeyMouseDefault=+MButton (FP = PopupHotkeyNewMouse)
PowerHotkeyKeyboardDefault=+#A (FP = PopupHotkeyNewKeyboard)

Read FP hotkeys
if we have hotkey for {Clipboard}, write to ini file
if we have hotkey for {Copy Favorite Location}, write to ini file
if we have hotkey for {Current Folders}, write to ini file
if we have hotkey for {Recent Folders}, write to ini file
(if we don't hae a hotkey, keep existing hotkey from QAP setup)

AS-IS:
[Global]
Check4Update
DefaultMenuBuilt
DiagMode
DisplayIcons
DisplayMenuShortcuts
DisplayTrayTip
HotkeyReminders
IconSize
LanguageCode
OpenMenuOnTaskbar
PopupFixPosition
PopupMenuPosition
RecentFolders
RememberSettingsPosition
SettingsPosition
Theme
DirectoryOpusPath
DirectoryOpusUseTabs
TotalCommanderPath
TotalCommanderUseTabs
FPconnectPath
AvailableThemes=Windows|Grey|Light Blue|Light Green|Light Red|Yellow

; read FP themes and import themes

*/

return
;-----------------------------------------------------------


;-----------------------------------------------------------
ImportFavoritesSettings:
;-----------------------------------------------------------

gosub, LoadFPFavorites
gosub, SaveQAPFavorites

return
;-----------------------------------------------------------


;-----------------------------------------------------------
LoadFPFavorites:
;-----------------------------------------------------------

g_objMainMenu := Object() ; object of menu structure entry point
g_objMenuIndex := Object() ; index of menu path used in Gui menu dropdown list
g_objMainMenu.MenuPath := "" ; empty name -  will be replaced with localized name
g_objMenuIndex.Insert(g_objMainMenu.MenuPath, g_objMainMenu) ; update the menu index

Loop
{
	IniRead, strLoadIniLine, %g_strFPIniFile%, Folders, Folder%A_Index%
	if (strLoadIniLine = "ERROR")
		Break
	
	strLoadIniLine := strLoadIniLine . "||||" ; additional "|" to make sure we have all empty items
	; 1 FavoriteName, 2 FavoriteLocation, 3 MenuName, 4 SubmenuFullName, 5 FavoriteType, 6 IconResource, 7 AppArguments, 8 AppWorkingDir
	StringSplit, arrThisFavorite, strLoadIniLine, |

	objLoadIniFavorite := Object() ; new menu item
	objLoadIniFavorite.FavoriteName := arrThisFavorite1 ; display name of this menu item
	objLoadIniFavorite.FavoriteLocation := arrThisFavorite2 ; path for this menu item
	StringReplace, arrThisFavorite3, arrThisFavorite3, >, % " > ", All
	objLoadIniFavorite.MenuName := Trim(arrThisFavorite3) ; parent menu of this menu item

	if StrLen(arrThisFavorite4)
	{
		; full name of the submenu, adding spaces before/after menu separator ">"
		StringReplace, arrThisFavorite4, arrThisFavorite4, >, % " > ", All
		objLoadIniFavorite.SubmenuFullName := Trim(arrThisFavorite4)
	}
	else
		objLoadIniFavorite.SubmenuFullName := ""
	
	if StrLen(arrThisFavorite5)
		objLoadIniFavorite.FavoriteType := arrThisFavorite5 ; FP types: "F" folder, "D" document, "U" URL, "A" application or "S" submenu
	else ; for upward compatibility from v1 and v2 ini files
		if StrLen(objLoadIniFavorite.SubmenuFullName)
			objLoadIniFavorite.FavoriteType := "S" ; "S" submenu
		else ; for upward compatibility from v1 ini files
			objLoadIniFavorite.FavoriteType := "F" ; "F" folder

	objLoadIniFavorite.FavoriteIconResource := arrThisFavorite6 ; icon resource in format "iconfile,iconindex"
	objLoadIniFavorite.FavoriteAppArguments := arrThisFavorite7 ; application arguments
	objLoadIniFavorite.FavoriteAppWorkingDir := arrThisFavorite8 ; application working directory
	
	; convert favorite types to QAP model and clean location for menu items
	if (objLoadIniFavorite.FavoriteType = "S") ; "S" submenu changed for "Menu"
	{
		objLoadIniFavorite.FavoriteType := "Menu"
		objLoadIniFavorite.FavoriteLocation := objLoadIniFavorite.SubmenuFullName ; full menu path without "Main" and spaces before/after ">"
	}
	else if (objLoadIniFavorite.FavoriteType = "F")
		objLoadIniFavorite.FavoriteType := "Folder"
	else if (objLoadIniFavorite.FavoriteType = "D")
		objLoadIniFavorite.FavoriteType := "Document"
	else if (objLoadIniFavorite.FavoriteType = "A")
		objLoadIniFavorite.FavoriteType := "Application"
	else if (objLoadIniFavorite.FavoriteType = "P") ; "P" sPecial changed for "Special"
		objLoadIniFavorite.FavoriteType := "Special"
	else if (objLoadIniFavorite.FavoriteType = "U")
		objLoadIniFavorite.FavoriteType := "URL"

	if (objLoadIniFavorite.FavoriteName = g_strGuiMenuSeparator)
	{
		objLoadIniFavorite.FavoriteType := "X"
		objLoadIniFavorite.FavoriteName := ""
		objLoadIniFavorite.FavoriteLocation := ""
	}
	
	if (SubStr(objLoadIniFavorite.FavoriteName, 1, 3) = g_strGuiMenuColumnBreak)
	{
		objLoadIniFavorite.FavoriteType := "K"
		objLoadIniFavorite.FavoriteName := ""
		objLoadIniFavorite.FavoriteLocation := ""
	}

	if (arrThisFavorite5 = "S") ; this is a submenu
	{
		objNewMenu := Object() ; create new menu object
		objNewMenu.MenuPath := objLoadIniFavorite.SubmenuFullName ; parent menu of this new menu item, without main menu name
		objLoadIniFavorite.SubMenu := objNewMenu ; pointer to the submenu
		g_objMenuIndex.Insert(objNewMenu.MenuPath, objNewMenu) ; add the new menu to the menu index
	}

	g_objMenuIndex[objLoadIniFavorite.MenuName].Insert(objLoadIniFavorite) ; insert favorite to its menu
}

return
;-----------------------------------------------------------


;-----------------------------------------------------------
SaveQAPFavorites:
;-----------------------------------------------------------

IniDelete, %g_strQAPIniFile%, Favorites

g_intIniLine := 1 ; reset counter before saving to ini file
RecursiveSaveFavoritesToIniFile(g_objMainMenu)

return
;-----------------------------------------------------------


;------------------------------------------------------------
RecursiveSaveFavoritesToIniFile(objCurrentMenu)
;------------------------------------------------------------
{
	Loop, % objCurrentMenu.MaxIndex()
	{
		IniWrite, % objCurrentMenu[A_Index].FavoriteType . "|" 
			. objCurrentMenu[A_Index].FavoriteName . "|" 
			. objCurrentMenu[A_Index].FavoriteLocation . "|" 
			. objCurrentMenu[A_Index].FavoriteIconResource . "|" 
			. objCurrentMenu[A_Index].FavoriteAppArguments . "|" 
			. objCurrentMenu[A_Index].FavoriteAppWorkingDir
			, %g_strQAPIniFile%, Favorites, Favorite%g_intIniLine%
		g_intIniLine += 1
		
		if (objCurrentMenu[A_Index].FavoriteType = "Menu")
			RecursiveSaveFavoritesToIniFile(objCurrentMenu[A_Index].SubMenu) ; RECURSIVE
	}
		
	IniWrite, Z, %g_strQAPIniFile%, Favorites, Favorite%g_intIniLine%
	g_intIniLine += 1
	
	return
}
;------------------------------------------------------------


