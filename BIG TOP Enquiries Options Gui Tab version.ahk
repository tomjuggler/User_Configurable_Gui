; Author:         Tom Hastings
; customize this template by editing "ShellNew\Template.ahk" in Windows folder
#SingleInstance
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, 2 ;any part of wintitle is detected

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;      GUI Section Auto Execute       ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

Gui, Font, s16  ; Set a large font size (32-point)., 
Gui, Add, Tab, x12 y10 w1190 h500 , Corporate Adult|Corporate Family|Fire Electric|School Charity Church|Wedding|Private|Mostly Jazz|Misc Other
Gui, Tab, Corporate Adult
Gui, Add, Button, x42 y100 w200 h70 ,  CORPORATE ADULT ENQUIRY ;1
Gui, Add, Button, x300 y100 w200 h80 , CHANGE CORP ADULT ENQ PATH ;path change
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
MsgBox, corp adult enq
;close gui after launching mail merge template?
Return

ButtonCHANGECORPADULTENQPATH:
Gosub, PathChooser
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


PathChooser: ; should set "path" variable to the mail merge path

gosub, PathWriter

return


PathWriter: ; should take variable "path" and iniwrite to the correct slot.  

return


;Menu, tray, add  ; Creates a separator line.
;Menu, tray, add, Item1, MenuHandler  ; Creates a new menu item.
;return

;MenuHandler:
;MsgBox You selected %A_ThisMenuItem% from menu %A_ThisMenu%.
;return
 


GuiClose:
ExitApp

Esc::ExitApp