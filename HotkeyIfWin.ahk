; rereAd THIS THREAD UNTIL THE END
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

/* BK JL based on Hokeyit
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
