; Author:         Tom Hastings
; customize this template by editing "ShellNew\Template.ahk" in Windows folder
#SingleInstance
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, 2 ;any part of wintitle is detected

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;   Auto Execute Section ;;;;;;;;;  ;;;;;;;;;;;;;;;;;;;;;;    ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Gosub,  InitialSetup
Gosub, GUIStart
return


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;; Setting up vars: ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
InitialSetup:
IniRead, Button1Title, paths.ini , Button1Title, key
IniRead, Button2Title, paths.ini , Button2Title, key
IniRead, Button3Title, paths.ini , Button3Title, key
IniRead, Button4Title, paths.ini , Button4Title, key
IniRead, Button5Title, paths.ini , Button5Title, key
IniRead, Button6Title, paths.ini , Button6Title, key
IniRead, Button7Title, paths.ini , Button7Title, key
IniRead, Button8Title, paths.ini , Button8Title, key
IniRead, Button9Title, paths.ini , Button9Title, key
Return

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;GUI SECTION;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
GUIStart:
Gui, Font, s16  ; Set a large font size (32-point)., 
Gui, Add, Tab, x12 y10 w1190 h500 , Corporate Adult|Corporate Family|Fire Electric|School Charity Church|Wedding|Private|Mostly Jazz|Misc Other
Gui, Tab, Corporate Adult
; user can change button titles variable now
Gui, Add, Button, x42 y100 w200 h70 gButton1 ,  %Button1Title%  ;1
Gui, Add, GroupBox, x295 y80 w400 h300, Change Paths and Rename Buttons
Gui, Add, Button, x300 y120 w100 h40 gButtonCHANGE1 , CHANGE ; path change button1
Gui, Add, Button, x42 y220 w230 h90 gButton2 , %Button2Title% ;2
Gui, Add, Button, x300 y220 w100 h40 gButtonCHANGE2 , CHANGE ; path change button2
Gui, Tab, Corporate Family
Gui, Add, Button, x92 y120 w300 h110 gButton3 , %Button3Title% ;3
Gui, Add, Button, x500 y120 w100 h40 gButtonCHANGE3 , CHANGE ; path change button3
Gui, Tab, Fire Electric
Gui, Add, Button, x52 y110 w220 h80 gButton4 , %Button4Title% ;4
Gui, Add, Button, x300 y110 w100 h40 gButtonCHANGE4 , CHANGE ; path change button4
Gui, Tab, School Charity Church
Gui, Add, Button, x32 y90 w380 h180 gButton5 , %Button5Title% ;5
Gui, Add, Button, x500 y100 w100 h40 gButtonCHANGE5 , CHANGE ; path change button5
Gui, Tab, Wedding
Gui, Add, Button, x32 y80 w350 h70 gButton6 , %Button6Title% ;6
Gui, Add, Button, x500 y80 w100 h40 gButtonCHANGE6 , CHANGE ; path change button6
Gui, Add, Button, x22 y200 w330 h70 gButton7 , %Button7Title% ;7
Gui, Add, Button, x500 y200 w100 h40 gButtonCHANGE7 , CHANGE ; path change button7
Gui, Add, Button, x22 y320 w680 h90 gButton8 , %Button8Title%  ;8
Gui, Add, Button, x800 y320 w100 h40 gButtonCHANGE8 , CHANGE ; path change button8
Gui, Tab, Private
Gui, Add, Button, x52 y120 w370 h100 gButton9, %Button9Title% ;9
Gui, Add, Button, x500 y120 w100 h40 gButtonCHANGE9 , CHANGE ; path change button9
; Generated using SmartGUI Creator 4.0
Gui, Show, x40 y 51 h560 w1223, BIG TOP ENTERTAINMENT email template options
return

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;; BUTTONS SECTION  ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;; ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;


Button1:
IniRead, path, paths.ini , Button1_path, key ;this button mail merge path into variable "path"
gosub, MailMergeMacro
ExitApp
return

ButtonCHANGE1:
inikey = Button1_path ; FOR PATH WRITER FUNCTION TO KNOW WHERE YOU CAME FROM and which .ini setting to change
Gosub, PathChooser
sleep, 40
inikey = Button1Title ; ditto above this time button title gets changed
gosub, UserButtonChanger
If path != 					; if a mail merge is chosen (path variable is not blank)
	{
	Reload
	}
Else 
	{
	return					; if nothing chosen do nothing
	}
Return

Button2:
IniRead, path, paths.ini , Button2_path, key ;this button mail merge path into variable "path"
gosub, MailMergeMacro
ExitApp
return

ButtonCHANGE2:
inikey = Button2_path ; FOR PATH WRITER FUNCTION TO KNOW WHERE YOU CAME FROM and which .ini setting to change
Gosub, PathChooser
sleep, 40
inikey = Button2Title ; ditto above this time button title gets changed
gosub, UserButtonChanger
If path != 					; if a mail merge is chosen (path variable is not blank)
	{
	Reload
	}
Else 
	{
	return					; if nothing chosen do nothing
	}
Return

Button3:
IniRead, path, paths.ini , Button3_path, key ;this button mail merge path into variable "path"
gosub, MailMergeMacro
ExitApp
return

ButtonCHANGE3:
inikey = Button3_path ; FOR PATH WRITER FUNCTION TO KNOW WHERE YOU CAME FROM and which .ini setting to change
Gosub, PathChooser
sleep, 40
inikey = Button3Title ; ditto above this time button title gets changed
gosub, UserButtonChanger
If path != 					; if a mail merge is chosen (path variable is not blank)
	{
	Reload
	}
Else 
	{
	return					; if nothing chosen do nothing
	}
Return

Button4:
IniRead, path, paths.ini , Button4_path, key ;this button mail merge path into variable "path"
gosub, MailMergeMacro
ExitApp
return

ButtonCHANGE4:
inikey = Button4_path ; FOR PATH WRITER FUNCTION TO KNOW WHERE YOU CAME FROM and which .ini setting to change
Gosub, PathChooser
sleep, 40
inikey = Button4Title ; ditto above this time button title gets changed
gosub, UserButtonChanger
If path != 					; if a mail merge is chosen (path variable is not blank)
	{
	Reload
	}
Else 
	{
	return					; if nothing chosen do nothing
	}
Return

Button5:
IniRead, path, paths.ini , Button5_path, key ;this button mail merge path into variable "path"
gosub, MailMergeMacro
ExitApp
return

ButtonCHANGE5:
inikey = Button5_path ; FOR PATH WRITER FUNCTION TO KNOW WHERE YOU CAME FROM and which .ini setting to change
Gosub, PathChooser
sleep, 40
inikey = Button5Title ; ditto above this time button title gets changed
gosub, UserButtonChanger
If path != 					; if a mail merge is chosen (path variable is not blank)
	{
	Reload
	}
Else 
	{
	return					; if nothing chosen do nothing
	}
Return

Button6:
IniRead, path, paths.ini , Button6_path, key ;this button mail merge path into variable "path"
gosub, MailMergeMacro
ExitApp
return

ButtonCHANGE6:
inikey = Button6_path ; FOR PATH WRITER FUNCTION TO KNOW WHERE YOU CAME FROM and which .ini setting to change
Gosub, PathChooser
sleep, 40
inikey = Button6Title ; ditto above this time button title gets changed
gosub, UserButtonChanger
If path != 					; if a mail merge is chosen (path variable is not blank)
	{
	Reload
	}
Else 
	{
	return					; if nothing chosen do nothing
	}
Return

Button7:
IniRead, path, paths.ini , Button7_path, key ;this button mail merge path into variable "path"
gosub, MailMergeMacro
ExitApp
return

ButtonCHANGE7:
inikey = Button7_path ; FOR PATH WRITER FUNCTION TO KNOW WHERE YOU CAME FROM and which .ini setting to change
Gosub, PathChooser
sleep, 40
inikey = Button7Title ; ditto above this time button title gets changed
gosub, UserButtonChanger
If path != 					; if a mail merge is chosen (path variable is not blank)
	{
	Reload
	}
Else 
	{
	return					; if nothing chosen do nothing
	}
Return

Button8:
IniRead, path, paths.ini , Button8_path, key ;this button mail merge path into variable "path"
gosub, MailMergeMacro
ExitApp
return

ButtonCHANGE8:
inikey = Button8_path ; FOR PATH WRITER FUNCTION TO KNOW WHERE YOU CAME FROM and which .ini setting to change
Gosub, PathChooser
sleep, 40
inikey = Button8Title ; ditto above this time button title gets changed
gosub, UserButtonChanger
If path != 					; if a mail merge is chosen (path variable is not blank)
	{
	Reload
	}
Else 
	{
	return					; if nothing chosen do nothing
	}
Return

Button9:
IniRead, path, paths.ini , Button9_path, key ;this button mail merge path into variable "path"
gosub, MailMergeMacro
ExitApp
return

ButtonCHANGE9:
inikey = Button9_path ; FOR PATH WRITER FUNCTION TO KNOW WHERE YOU CAME FROM and which .ini setting to change
Gosub, PathChooser
sleep, 40
inikey = Button9Title ; ditto above this time button title gets changed
gosub, UserButtonChanger
If path != 					; if a mail merge is chosen (path variable is not blank)
	{
	Reload
	}
Else 
	{
	return					; if nothing chosen do nothing
	}
Return

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;   MAIN FUNCTIONS  ;;;;;;;;;;;;;;;;;;;;;;;;;;;; ;;;;;;;;;;;;;;;;;;;;;;;;; ;;;;;;;;;;;;;   ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

UserButtonChanger:
;INIWRITE BUTTON NAME
InputBox, path , choose the button name
Gosub, PathWriter
return


PathChooser: ; should set "path" variable to the mail merge path
FileSelectFile, path , 3, C:\Dropbox\Big Top Entertainment\Templates , Select the Mail Merge to use for this (variable here) option! , ; "path" variable is set to path of selected mail merge!
If path != 					; if a mail merge is chosen (path variable is not blank)
	{
	gosub, PathWriter
	}
Else 
	{
	return					; if nothing chosen do nothing
	}

return


PathWriter: ; should take variable "path" and iniwrite to the correct slot - shown by "inikey" var. 
IniWrite, %path%, paths.ini , %inikey%, key
return

MailMergeMacro:
xl :=   ComObj("Excel.Application")
;MsgBox % xl.ActiveCell.Row
clipboard = % xl.ActiveCell.Row 
EnvSub, Clipboard, 1 ;subtracts 1 
; paste and you get the active row!!!!!
sleep, 100
run, %path%
sleep, 200
;WinGet, active_id, ID, A
;WinMaximize, ahk_id %active_id% 		; test code - maximizes window
;MsgBox, The active window's ID is "%active_id%". ; test code
winwait, ahk_class #32770
send, y 
sleep, 200
WinGet, active_id, ID, A
;WinMaximize, ahk_id %active_id% ;test code - maximizes window
;MsgBox, The active window's ID is "%active_id%". test code
winwait, ahk_id %active_id% ;********************** using window id from WinGet *****************
WinMaximize, ahk_id %active_id%    ; use the window found above

;START OF NON-OPTIMAL CODE

SLEEP, 1000

WinWait, ahk_id %active_id%, 
IfWinNotActive, ahk_id %active_id%, , WinActivate, ahk_id %active_id%, 
WinWaitActive, ahk_id %active_id%, 
sleep, 400
MouseClick, left,  389,  40
Sleep, 100
MouseClick, left,  785,  68
Sleep, 1000
MouseClick, left,  785,  68
Send, {CTRLDOWN}v{CTRLUP}
Sleep, 100
MouseClick, left,  692,  89
Sleep, 200
ExitApp
*/

return


 
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;   END: FINAL CLEANUP ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;; ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

GuiClose:
ExitApp

Esc::ExitApp