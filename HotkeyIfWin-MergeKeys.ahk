#NoEnv
#SingleInstance force

global g_strMouseHotkey := "MButton" ; shiftt-middle mouse button
global g_strKeyboardHotkey := "#A" ; Windows-A
global g_strMouseAlternateHotkey := "!MButton" ; shiftt-middle mouse button
global g_strKeyboardAlternateHotkey := "!#A" ; Windows-A

global g_strTargetWinId
global g_strTargetClass
global g_strTargetControl

global g_blnUseDirectoryOpus := 1
global g_blnUseTotalCommander := 0
global g_blnUseFPconnect := 0

global g_intCounterNavigate := 0
global g_intCounterLaunch := 0

global g_blnOpenMenuOnTaskbar := 1 ; from Options

global g_strExclusionClassList

g_strExclusionClassList := "SciTEWindow|Chrome_WidgetWin_1|" ; must end with |

; Hotkey, If, CanNavigate(A_ThisLabel, strTargetWinId, strTargetClass, strTargetControl)
Hotkey, If, CanNavigate(A_ThisHotkey)
Hotkey, %g_strMouseHotkey%, NavigateHotkeyMouse
Hotkey, %g_strKeyboardHotkey%, NavigateHotkeyKeyboard
Hotkey, If

Hotkey, If, CanLaunch(A_ThisHotkey)
Hotkey, %g_strMouseHotkey%, LaunchHotkeyMouse
Hotkey, %g_strKeyboardHotkey%, LaunchHotkeyKeyboard
Hotkey, If

Hotkey, %g_strMouseAlternateHotkey%, AskHotkeyMouse
Hotkey, %g_strKeyboardAlternateHotkey%, AskHotkeyKeyboard

; To popup menu when left click on the tray icon - See AHK_NOTIFYICON function below
OnMessage(0x404, "AHK_NOTIFYICON")

return


;------------------------------------------------------------
;------------------------------------------------------------
#If, CanNavigate(A_ThisHotkey)
; empty - act as a handle for the "Hotkey, If" confition
#If
;------------------------------------------------------------
;------------------------------------------------------------


;------------------------------------------------------------
;------------------------------------------------------------
#If, CanLaunch(A_ThisHotkey)
; empty - act as a handle for the "Hotkey, If" confition
#If
;------------------------------------------------------------
;------------------------------------------------------------


;------------------------------------------------------------
AHK_NOTIFYICON(wParam, lParam) 
; Adapted from Lexikos http://www.autohotkey.com/board/topic/11250-mouseover-trayicon-triggering-an-event/#entry153388
; To popup menu when left click on the tray icon - See the OnMessage command in the init section
;------------------------------------------------------------
{
	global blnClickOnTrayIcon
	
	if (lParam = 0x202) ; WM_LBUTTONUP
	{
		blnClickOnTrayIcon := 1
		; enable when in QAP: SetTimer, PopupMenuNewWindowMouse, -1
		return 0
	}
} 
;------------------------------------------------------------


;------------------------------------------------------------
NavigateHotkeyMouse: ; default MButton if CanNavigate
NavigateHotkeyKeyboard: ; default #A if CanNavigate
LaunchHotkeyMouse: ; default MButton if CanLaunch
LaunchHotkeyKeyboard: ; default #A if CanLaunch
AskHotkeyMouse: ; default !MButton if CanLaunch
AskHotkeyKeyboard: ; default !#A if CanLaunch
;------------------------------------------------------------

###_D("A_ThisLabel: " . A_ThisLabel . "`ng_strTargetWinId: " . g_strTargetWinId . "`ng_strTargetClass: " . g_strTargetClass . "`ng_strTargetControl: " . g_strTargetControl . "`nCounters (Navigate/Launch): " . g_intCounterNavigate . "/" . g_intCounterLaunch)

return
;------------------------------------------------------------



;------------------------------------------------------------
CanLaunch(strMouseOrKeyboard) ; SEE HotkeyIfWin.ahk to use Hotkey, If, Expression
;------------------------------------------------------------
{
	g_intCounterLaunch++
	
	if (strMouseOrKeyboard = g_strMouseHotkey)
	{
		MouseGetPos, , , g_strTargetWinId, g_strTargetControl
		WinGetClass g_strTargetClass, % "ahk_id " . g_strTargetWinId
		; TrayTip, Launch Mouse, %strMouseOrKeyboard%: %g_strMouseNavigateHotkey%`n%g_strTargetControl%`nList: %g_strExclusionClassList%`nClass: %g_strTargetClass%
	}
	else ; Keyboard
	{
		g_strTargetWinId := WinExist("A")
		g_strTargetControl := ""
		WinGetClass g_strTargetClass, % "ahk_id " . g_strTargetWinId
		; TrayTip, Launch Keyboard, %strMouseOrKeyboard% = %g_strKeyboardNavigateHotkey%`nList: %g_strExclusionClassList%`nClass: %g_strTargetClass%
	}

	return !InStr(g_strExclusionClassList, g_strTargetClass . "|")
}
;------------------------------------------------------------


;------------------------------------------------------------
CanNavigate(strMouseOrKeyboard) ; SEE HotkeyIfWin.ahk to use Hotkey, If, Expression
; "CabinetWClass" and "ExploreWClass" -> Explorer
; "ProgMan" -> Desktop
; "WorkerW" -> Desktop
; "ConsoleWindowClass" -> Console (CMD)
; "#32770" -> Dialog
; "bosa_sdm_" (...) -> Dialog MS Office under WinXP
;------------------------------------------------------------
{
	g_intCounterNavigate++
	
	if (strMouseOrKeyboard = g_strMouseHotkey)
	{
		MouseGetPos, , , g_strTargetWinId, g_strTargetControl
		WinGetClass g_strTargetClass, % "ahk_id " . g_strTargetWinId
		; TrayTip, Navigate Mouse, %strMouseOrKeyboard% = %g_strMouseNavigateHotkey% (%g_intCounter%)`n%g_strTargetWinId%`n%g_strTargetClass%`n%g_strTargetControl%
	}
	else ; Keyboard
	{
		g_strTargetWinId := WinExist("A")
		g_strTargetControl := ""
		WinGetClass g_strTargetClass, % "ahk_id " . g_strTargetWinId
		; TrayTip, Navigate Keyboard, %strMouseOrKeyboard% = %g_strKeyboardNavigateHotkey% (%g_intCounter%)`n%g_strTargetWinId%`n%g_strTargetClass%
	}

	blnCanOpenFavorite := WindowIsAnExplorer(g_strTargetClass) or WindowIsDesktop(g_strTargetClass) or WindowIsConsole(g_strTargetClass)
		or WindowIsDialog(g_strTargetClass, g_strTargetWinId)
		or (g_blnUseDirectoryOpus and WindowIsDirectoryOpus(g_strTargetClass))
		or (g_blnUseTotalCommander and WindowIsTotalCommander(g_strTargetClass))
		or (g_blnUseFPconnect and WindowIsFPconnect(g_strTargetWinId))
		or WindowIsQuickAccessPopup(g_strTargetClass)

	return blnCanOpenFavorite
}
;------------------------------------------------------------


;------------------------------------------------------------
WindowIsAnExplorer(strClass)
;------------------------------------------------------------
{
	return (strClass = "CabinetWClass") or (strClass = "ExploreWClass")
}
;------------------------------------------------------------


;------------------------------------------------------------
WindowIsDesktop(strClass)
;------------------------------------------------------------
{
	global g_blnOpenMenuOnTaskbar
	global blnClickOnTrayIcon
	
	blnWindowIsDesktop := (strClass = "ProgMan")
		or (strClass = "WorkerW")
		or (strClass = "Shell_TrayWnd" and (g_blnOpenMenuOnTaskbar or blnClickOnTrayIcon))
		or (strClass = "NotifyIconOverflowWindow")

	blnClickOnTrayIcon := 0
	; blnClickOnTrayIcon was turned on by AHK_NOTIFYICON
	; turn it off to avoid further clicks on taskbar to be accepted if g_blnOpenMenuOnTaskbar is off

	return blnWindowIsDesktop
}
;------------------------------------------------------------


;------------------------------------------------------------
WindowIsTray(strClass)
;------------------------------------------------------------
{
	return (strClass = "Shell_TrayWnd") or (strClass = "NotifyIconOverflowWindow")
}
;------------------------------------------------------------


;------------------------------------------------------------
WindowIsConsole(strClass)
;------------------------------------------------------------
{
	return (strClass = "ConsoleWindowClass")
}
;------------------------------------------------------------


;------------------------------------------------------------
WindowIsDialog(strClass, strWinId)
;------------------------------------------------------------
{
	return (strClass = "#32770") and !WindowIsTreeview(strWinId)
	
	; or InStr(strClass, "bosa_sdm_")
	; Removed 2014-09-27  (see http://code.jeanlalonde.ca/folderspopupv3archives/#comment-7912)
}
;------------------------------------------------------------


;------------------------------------------------------------
WindowIsTreeview(strWinId)
; Disable popup menu in folder select dialog boxes (like those displayed by FileSelectFolder)
; because their Edit1 control does not react as expected in NavigateDialog.
; Signature: contains both SysTreeView321 and SHBrowseForFolder controls (tested on Win7 only)
; but NOT 100% sure this is a unique signature...
;------------------------------------------------------------
{
	global g_strAppNameText
	
	WinGet, strControlsList, ControlList, ahk_id %strWinId%
	blnIsTreeView := InStr(strControlsList, "SysTreeView321") and InStr(strControlsList, "SHBrowseForFolder")
	if (blnIsTreeView)
		TrayTip, %lWindowIsTreeviewTitle%, % L(lWindowIsTreeviewText, g_strAppNameText), , 2
	
	return blnIsTreeView
}
;------------------------------------------------------------


;------------------------------------------------------------
WindowIsDirectoryOpus(strClass)
;------------------------------------------------------------
{
	return InStr(strClass, "dopus")
}
;------------------------------------------------------------


;------------------------------------------------------------
WindowIsTotalCommander(strClass)
;------------------------------------------------------------
{
	return InStr(strClass, "TTOTAL_CMD")
}
;------------------------------------------------------------


;------------------------------------------------------------
WindowIsFPconnect(strWinId)
;------------------------------------------------------------
{
	global g_strFPconnectAppFilename
	global g_strFPconnectTargetFilename

	if (strTargetWinId = 0)
		return false

	; get path and filename of the app controling window strWinId
	; first get process ID
    intPID := 0
    DllCall("GetWindowThreadProcessId", "UInt", strWinId, "UInt *", intPID)
	; get filename of process
    hProcess := DllCall("OpenProcess", "UInt", 0x400 | 0x10, "Int", False, "UInt", intPID)
    intPathLength = 260*2
    VarSetCapacity(strFCAppFile, intPathLength, 0)
    DllCall("Psapi.dll\GetModuleFileNameExW", "UInt", hProcess, "Int", 0, "Str", strFCAppFile, "UInt", intPathLength)
    DllCall("CloseHandle", "UInt", hProcess)
	
	; get filename only and compare with FPconnect filename or FPconnect target filename (see FPconnect doc)
	SplitPath, strFCAppFile, strFCAppFile
	return (strFCAppFile = g_strFPconnectAppFilename) or (strFCAppFile = g_strFPconnectTargetFilename)
}
;------------------------------------------------------------


;------------------------------------------------------------
WindowIsQuickAccessPopup(strClass)
; enabled only when compiled
;------------------------------------------------------------
{
	return (strClass = "JeanLalonde.ca")
}
;------------------------------------------------------------


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



/*
; REREAD THIS THREAD UNTIL THE END
; http://ahkscript.org/boards/viewtopic.php?p=46821&sid=f256384010b2206e3a31848428a68d7b#p46821

#NoEnv
#SingleInstance force

global counter := 0
var := "!MButton" ; alt-middle mouse button

Hotkey, If, OK(++counter)
Hotkey, %var%, MyLabel  ; Creates a hotkey that works every two calls
Hotkey, If

return

#If, OK(++counter)
#If

MyLabel:
MsgBox, %A_ThisHotkey%
return

OK(count)
{
	; MsgBox, %count%
	return Mod(count, 2)
}
*/

/* BK JL based on Hotkeyit
Hotkey, !e, MyLabel  ; Creates a hotkey that works in all windows

Hotkey, IfWinActive, ahk_class Notepad
Hotkey, ^e, MyLabel  ; Creates a hotkey that works only in Notepad.
Hotkey, IfWinActive

Hotkey, If, OK(++counter)
Hotkey, !^e, MyLabel  ; Creates a hotkey that works every two calls
Hotkey, If

return


#If, OK(++counter)
#If


MyLabel:
MsgBox, %A_ThisHotkey%
return


OK(count)
{
	; MsgBox, %count%
	return Mod(count, 2)
}

*/


/* FROM HotKeyIt:
#NoEnv
#SingleInstance force
global counter := 0
Hotkey, If, OK()
Hotkey, !^e, MyLabel  ; Creates a hotkey that works every two calls
Hotkey, If
return
#If OK()
#If
MyLabel:
MsgBox, %A_ThisHotkey%
return
OK(){
	return Mod(++counter, 2)
}
*/

/* FROM Lexikos
#NoEnv
#SingleInstance force
 
global counter := 0
 
^e::MyHotkey("Global")
#IfWinActive ahk_class Notepad
^e::MyHotkey("Notepad")
#If OK()
^e::MyHotkey("Alternates on/off")
#If
 
MyHotkey(t:="") {
	MsgBox, %A_ThisHotkey% %t%
}
 
OK() {
	return Mod(++counter, 2)
}
*/
