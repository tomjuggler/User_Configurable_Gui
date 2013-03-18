; Author:         Tom Hastings
; customize this template by editing "ShellNew\Template.ahk" in Windows folder
#SingleInstance
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, 2 ;any part of wintitle is detected

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;      GUI Section Auto Execute       ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
; Setting up vars:
IniRead, CorpAdultEnqButtonTitle, paths.ini , CORPORATEADULTENQUIRY_BUTTONTITLE, key
Gosub, GUIStart
return



;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;GUI SECTION;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
GUIStart:
Gui, Font, s16  ; Set a large font size (32-point)., 
Gui, Add, Tab, x12 y10 w1190 h500 , Corporate Adult|Corporate Family|Fire Electric|School Charity Church|Wedding|Private|Mostly Jazz|Misc Other
Gui, Tab, Corporate Adult
; user can change button titles variable now
Gui, Add, Button, x42 y100 w200 h70 gButtonCORPORATEADULTENQUIRY ,  %CorpAdultEnqButtonTitle%  ;1
Gui, Add, GroupBox, x295 y80 w400 h300, Change Paths and Rename Buttons
Gui, Add, Button, x300 y120 w100 h40 , CHANGE ; path change button
Gui, Add, Button, x42 y220 w230 h90 , CORPORATE ADULT ENQUIRY (WITH TOM SPECIFIC INFO) ;2
Gui, Tab, Corporate Family
Gui, Add, Button, x92 y120 w300 h110 , CORPORATE FAMILY ENQUIRY ;3
Gui, Tab, Fire Electric
Gui, Add, Button, x52 y110 w220 h80 , CORPORATE FIRE AND ELECTRIC ENQUIRY ;4
Gui, Tab, School Charity Church
Gui, Add, Button, x32 y90 w380 h180 , SCHOOL CHARITY CHURCH FAMILY ENQUIRY QUOTE ;5
Gui, Tab, Wedding
Gui, Add, Button, x32 y80 w350 h70 , WEDDING ENQUIRY ;6
Gui, Add, Button, x22 y200 w330 h70 , WEDDING PLANNER ENQUIRY ;7
Gui, Add, Button, x22 y320 w680 h90 , WEDDING PLANNER ENQUIRY (ZONE 1 KIDS ENTERTAINMENT ONLY) ;8
Gui, Tab, Private
Gui, Add, Button, x52 y120 w370 h100 , PRIVATE ENQUIRY ;9
; Generated using SmartGUI Creator 4.0
Gui, Show, x40 y 51 h560 w1223, BIG TOP ENTERTAINMENT email template options
Return

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;    BUTTONS SECTION  ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;; ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

ButtonCORPORATEADULTENQUIRY:
IniRead, path, paths.ini , CORPORATEADULTENQUIRY_path, key ;path of mail merge into variable "path"
sleep, 50
Gosub, MailMergeMacro 
;MsgBox, corp adult enq
ExitApp ;close gui after launching mail merge template
Return

ButtonCHANGE:
inikey = CORPORATEADULTENQUIRY_path ; FOR PATH WRITER FUNCTION TO KNOW WHERE YOU CAME FROM and which .ini setting to change
Gosub, PathChooser
sleep, 40
inikey = CORPORATEADULTENQUIRY_BUTTONTITLE ; ditto above this time button title gets changed
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

ButtonCORPORATEADULTENQUIRY(WITHTOMSPECIFICINFO):
MsgBox, corp adult tom specific
Return

ButtonCORPORATEFAMILYENQUIRY:
MsgBox, CORPORATE FAMILY ENQUIRY
return

ButtonCORPORATEFIREANDELECTRICENQUIRY:
MsgBox, CORPORATE FIRE AND ELECTRIC ENQUIRY
return

ButtonSCHOOLCHARITYCHURCHFAMILYENQUIRYQUOTE:
MsgBox, CHURCH FAMILY ENQUIRY QUOTE
return

ButtonWEDDINGENQUIRY:
MsgBox, WEDDING ENQUIRY
return

ButtonWEDDINGPLANNERENQUIRY:
MsgBox, WEDDING PLANNER ENQUIRY
return

ButtonWEDDINGPLANNERENQUIRY(ZONE1KIDSENTERTAINMENTONLY):
MsgBox, WEDDING PLANNER ENQUIRY (ZONE 1 KIDS ENTERTAINMENT ONLY)
return

ButtonPRIVATEENQUIRY:
MsgBox, PRIVATE ENQUIRY
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

;Menu, tray, add  ; Creates a separator line.
;Menu, tray, add, Item1, MenuHandler  ; Creates a new menu item.
;return

;MenuHandler:
;MsgBox You selected %A_ThisMenuItem% from menu %A_ThisMenu%.
;return
 
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;   END: FINAL CLEANUP ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;; ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

GuiClose:
ExitApp

Esc::ExitApp