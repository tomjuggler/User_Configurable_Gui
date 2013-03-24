
; Author:         Tom Hastings (tomjuggler)

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;Description;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
 /*
 GUI program in which the User can change the button title and function (which file the button opens) by use of an adjacent "change" button on GUI (saves to paths.ini file in script directory). Useful for frequently changed options and so that the programmer (me) doesn't have to change the code every time a word doc is moved or renamed. The MailMerge function is called upon pressing a button to automate the word mail merge process - at least that is what we use this script for - have removed the Word specific code here, I hope someone finds this useful. Any comments or suggestions welcome it's my first post 
 */


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
InitialSetup: ;load from paths.ini. Does this use loads of memory? doesn't seem to slow anything down...
IniRead, Button1Title, paths.ini , Button1Title, key
IniRead, Button2Title, paths.ini , Button2Title, key
IniRead, Button3Title, paths.ini , Button3Title, key
IniRead, Button4Title, paths.ini , Button4Title, key
IniRead, Button5Title, paths.ini , Button5Title, key
IniRead, Button6Title, paths.ini , Button6Title, key
IniRead, Button7Title, paths.ini , Button7Title, key
IniRead, Button8Title, paths.ini , Button8Title, key
IniRead, Button9Title, paths.ini , Button9Title, key
IniRead, Button10Title, paths.ini , Button10Title, key
IniRead, Button11Title, paths.ini , Button11Title, key
IniRead, Button12Title, paths.ini , Button12Title, key
IniRead, Button13Title, paths.ini , Button13Title, key
IniRead, Button14Title, paths.ini , Button14Title, key
IniRead, Button15Title, paths.ini , Button15Title, key
IniRead, Button16Title, paths.ini , Button16Title, key
IniRead, Button17Title, paths.ini , Button17Title, key
IniRead, Button18Title, paths.ini , Button18Title, key
IniRead, Button19Title, paths.ini , Button19Title, key
IniRead, Button20Title, paths.ini , Button20Title, key
IniRead, Button21Title, paths.ini , Button21Title, key
IniRead, Button19Title, paths.ini , Button22Title, key
IniRead, Button20Title, paths.ini , Button23Title, key
IniRead, Button21Title, paths.ini , Button24Title, key
Return

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;GUI SECTION;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
GUIStart:
Gui, Font, s16  ; Set a large font size (32-point)., 
Gui, Add, Tab, x12 y10 w1190 h500 , Tab 1|Tab 2|Tab 3|Tab 4|Tab 5|Tab 6|Tab 7|Tab 8 ; call these whatever suits you
Gui, Tab, Tab 1
Gui, Add, Button, x32 y80 w350 h70 gButton1 , %Button1Title% ;1
Gui, Add, Button, x500 y80 w100 h40 gButtonCHANGE1 , CHANGE ; path change button1
Gui, Add, Button, x22 y200 w330 h70 gButton2 , %Button2Title% ;2
Gui, Add, Button, x500 y200 w100 h40 gButtonCHANGE2 , CHANGE ; path change button2
Gui, Add, Button, x22 y320 w680 h90 gButton3 , %Button3Title%  ;3
Gui, Add, Button, x800 y320 w100 h40 gButtonCHANGE3 , CHANGE ; path change button3
Gui, Tab, Tab 2
Gui, Add, Button, x32 y80 w350 h70 gButton4 , %Button4Title% ;4
Gui, Add, Button, x500 y80 w100 h40 gButtonCHANGE4 , CHANGE ; path change button4
Gui, Add, Button, x22 y200 w330 h70 gButton5 , %Button5Title% ;5
Gui, Add, Button, x500 y200 w100 h40 gButtonCHANGE5 , CHANGE ; path change button5
Gui, Add, Button, x22 y320 w680 h90 gButton6 , %Button6Title%  ;6
Gui, Add, Button, x800 y320 w100 h40 gButtonCHANGE6 , CHANGE ; path change button6
Gui, Tab, Tab 3
Gui, Add, Button, x32 y80 w350 h70 gButton7 , %Button7Title% ;7
Gui, Add, Button, x500 y80 w100 h40 gButtonCHANGE7 , CHANGE ; path change button1
Gui, Add, Button, x22 y200 w330 h70 gButton8 , %Button8Title% ;8
Gui, Add, Button, x500 y200 w100 h40 gButtonCHANGE8 , CHANGE ; path change button2
Gui, Add, Button, x22 y320 w680 h90 gButton9 , %Button9Title%  ;9
Gui, Add, Button, x800 y320 w100 h40 gButtonCHANGE9 , CHANGE ; path change button3
Gui, Tab, Tab 4
Gui, Add, Button, x32 y80 w350 h70 gButton10 , %Button10Title% ;10
Gui, Add, Button, x500 y80 w100 h40 gButtonCHANGE10 , CHANGE ; path change button10
Gui, Add, Button, x22 y200 w330 h70 gButton11 , %Button11Title% ;11
Gui, Add, Button, x500 y200 w100 h40 gButtonCHANGE11, CHANGE ; path change button11
Gui, Add, Button, x22 y320 w680 h90 gButton12 , %Button8Title%  ;12
Gui, Add, Button, x800 y320 w100 h40 gButtonCHANGE12 , CHANGE ; path change button12
Gui, Tab, Tab 5
Gui, Add, Button, x32 y80 w350 h70 gButton13 , %Button1Title% ;13
Gui, Add, Button, x500 y80 w100 h40 gButtonCHANGE13 , CHANGE ; path change button13
Gui, Add, Button, x22 y200 w330 h70 gButton14, %Button2Title% ;14
Gui, Add, Button, x500 y200 w100 h40 gButtonCHANGE14, CHANGE ; path change button14
Gui, Add, Button, x22 y320 w680 h90 gButton15 , %Button3Title%  ;15
Gui, Add, Button, x800 y320 w100 h40 gButtonCHANGE15 , CHANGE ; path change button15
Gui, Tab, Tab 6
Gui, Add, Button, x32 y80 w350 h70 gButton16 , %Button1Title% ;16
Gui, Add, Button, x500 y80 w100 h40 gButtonCHANGE16 , CHANGE ; path change button16
Gui, Add, Button, x22 y200 w330 h70 gButton17 , %Button2Title% ;17
Gui, Add, Button, x500 y200 w100 h40 gButtonCHANGE17 , CHANGE ; path change button17
Gui, Add, Button, x22 y320 w680 h90 gButton18 , %Button3Title%  ;18
Gui, Add, Button, x800 y320 w100 h40 gButtonCHANGE18 , CHANGE ; path change button18
Gui, Tab, Tab 7
Gui, Add, Button, x32 y80 w350 h70 gButton19 , %Button1Title% ;19
Gui, Add, Button, x500 y80 w100 h40 gButtonCHANGE19, CHANGE ; path change button19
Gui, Add, Button, x22 y200 w330 h70 gButton20 , %Button2Title% ;20
Gui, Add, Button, x500 y200 w100 h40 gButtonCHANGE20, CHANGE ; path change button20
Gui, Add, Button, x22 y320 w680 h90 gButton21 , %Button3Title%  ;21
Gui, Add, Button, x800 y320 w100 h40 gButtonCHANGE21, CHANGE ; path change button21
Gui, Tab, Tab 8
Gui, Add, Button, x32 y80 w350 h70 gButton22 , %Button1Title% ;22
Gui, Add, Button, x500 y80 w100 h40 gButtonCHANGE22, CHANGE ; path change button22
Gui, Add, Button, x22 y200 w330 h70 gButton23 , %Button2Title% ;23
Gui, Add, Button, x500 y200 w100 h40 gButtonCHANGE23, CHANGE ; path change button23
Gui, Add, Button, x22 y320 w680 h90 gButton24 , %Button3Title%  ;24
Gui, Add, Button, x800 y320 w100 h40 gButtonCHANGE24, CHANGE ; path change button24
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

Button10:
IniRead, path, paths.ini , Button10_path, key ;this button mail merge path into variable "path"
gosub, MailMergeMacro
ExitApp
return

ButtonCHANGE10:
inikey = Button10_path ; FOR PATH WRITER FUNCTION TO KNOW WHERE YOU CAME FROM and which .ini setting to change
Gosub, PathChooser
sleep, 40
inikey = Button10Title ; ditto above this time button title gets changed
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

Button11:
IniRead, path, paths.ini , Button11_path, key ;this button mail merge path into variable "path"
gosub, MailMergeMacro
ExitApp
return

ButtonCHANGE11:
inikey = Button11_path ; FOR PATH WRITER FUNCTION TO KNOW WHERE YOU CAME FROM and which .ini setting to change
Gosub, PathChooser
sleep, 40
inikey = Button11Title ; ditto above this time button title gets changed
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

Button12:
IniRead, path, paths.ini , Button12_path, key ;this button mail merge path into variable "path"
gosub, MailMergeMacro
ExitApp
return

ButtonCHANGE12:
inikey = Button12_path ; FOR PATH WRITER FUNCTION TO KNOW WHERE YOU CAME FROM and which .ini setting to change
Gosub, PathChooser
sleep, 40
inikey = Button12Title ; ditto above this time button title gets changed
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

Button13:
IniRead, path, paths.ini , Button13_path, key ;this button mail merge path into variable "path"
gosub, MailMergeMacro
ExitApp
return

ButtonCHANGE13:
inikey = Button13_path ; FOR PATH WRITER FUNCTION TO KNOW WHERE YOU CAME FROM and which .ini setting to change
Gosub, PathChooser
sleep, 40
inikey = Button13Title ; ditto above this time button title gets changed
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

Button14:
IniRead, path, paths.ini , Button14_path, key ;this button mail merge path into variable "path"
gosub, MailMergeMacro
ExitApp
return

ButtonCHANGE14:
inikey = Button14_path ; FOR PATH WRITER FUNCTION TO KNOW WHERE YOU CAME FROM and which .ini setting to change
Gosub, PathChooser
sleep, 40
inikey = Button14Title ; ditto above this time button title gets changed
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

Button15:
IniRead, path, paths.ini , Button15_path, key ;this button mail merge path into variable "path"
gosub, MailMergeMacro
ExitApp
return

ButtonCHANGE15:
inikey = Button15_path ; FOR PATH WRITER FUNCTION TO KNOW WHERE YOU CAME FROM and which .ini setting to change
Gosub, PathChooser
sleep, 40
inikey = Button15Title ; ditto above this time button title gets changed
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

Button16:
IniRead, path, paths.ini , Button16_path, key ;this button mail merge path into variable "path"
gosub, MailMergeMacro
ExitApp
return

ButtonCHANGE16:
inikey = Button16_path ; FOR PATH WRITER FUNCTION TO KNOW WHERE YOU CAME FROM and which .ini setting to change
Gosub, PathChooser
sleep, 40
inikey = Button16Title ; ditto above this time button title gets changed
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

Button17:
IniRead, path, paths.ini , Button17_path, key ;this button mail merge path into variable "path"
gosub, MailMergeMacro
ExitApp
return

ButtonCHANGE17:
inikey = Button17_path ; FOR PATH WRITER FUNCTION TO KNOW WHERE YOU CAME FROM and which .ini setting to change
Gosub, PathChooser
sleep, 40
inikey = Button17Title ; ditto above this time button title gets changed
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

Button18:
IniRead, path, paths.ini , Button18_path, key ;this button mail merge path into variable "path"
gosub, MailMergeMacro
ExitApp
return

ButtonCHANGE18:
inikey = Button18_path ; FOR PATH WRITER FUNCTION TO KNOW WHERE YOU CAME FROM and which .ini setting to change
Gosub, PathChooser
sleep, 40
inikey = Button18Title ; ditto above this time button title gets changed
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

Button19:
IniRead, path, paths.ini , Button19_path, key ;this button mail merge path into variable "path"
gosub, MailMergeMacro
ExitApp
return

ButtonCHANGE19:
inikey = Button19_path ; FOR PATH WRITER FUNCTION TO KNOW WHERE YOU CAME FROM and which .ini setting to change
Gosub, PathChooser
sleep, 40
inikey = Button19Title ; ditto above this time button title gets changed
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

Button20:
IniRead, path, paths.ini , Button20_path, key ;this button mail merge path into variable "path"
gosub, MailMergeMacro
ExitApp
return

ButtonCHANGE20:
inikey = Button20_path ; FOR PATH WRITER FUNCTION TO KNOW WHERE YOU CAME FROM and which .ini setting to change
Gosub, PathChooser
sleep, 40
inikey = Button20Title ; ditto above this time button title gets changed
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

Button21:
IniRead, path, paths.ini , Button21_path, key ;this button mail merge path into variable "path"
gosub, MailMergeMacro
ExitApp
return

ButtonCHANGE21:
inikey = Button21_path ; FOR PATH WRITER FUNCTION TO KNOW WHERE YOU CAME FROM and which .ini setting to change
Gosub, PathChooser
sleep, 40
inikey = Button21Title ; ditto above this time button title gets changed
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

Button22:
IniRead, path, paths.ini , Button22_path, key ;this button mail merge path into variable "path"
gosub, MailMergeMacro
ExitApp
return

ButtonCHANGE22:
inikey = Button22_path ; FOR PATH WRITER FUNCTION TO KNOW WHERE YOU CAME FROM and which .ini setting to change
Gosub, PathChooser
sleep, 40
inikey = Button22Title ; ditto above this time button title gets changed
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

Button23:
IniRead, path, paths.ini , Button23_path, key ;this button mail merge path into variable "path"
gosub, MailMergeMacro
ExitApp
return

ButtonCHANGE23:
inikey = Button23_path ; FOR PATH WRITER FUNCTION TO KNOW WHERE YOU CAME FROM and which .ini setting to change
Gosub, PathChooser
sleep, 40
inikey = Button23Title ; ditto above this time button title gets changed
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

Button24:
IniRead, path, paths.ini , Button24_path, key ;this button mail merge path into variable "path"
gosub, MailMergeMacro
ExitApp
return

ButtonCHANGE24:
inikey = Button24_path ; FOR PATH WRITER FUNCTION TO KNOW WHERE YOU CAME FROM and which .ini setting to change
Gosub, PathChooser
sleep, 40
inikey = Button24Title ; ditto above this time button title gets changed
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
FileSelectFile, path , 3, C:\ , Select the File to open for this option! , ; "path" variable is set to path of selected mail merge!
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

MailMergeMacro: ; this function throws an error if no excel worksheet is open (not something that happens in our office)
sleep, 100
run, %path%
sleep, 200
WinGet, active_id, ID, A
winwait, ahk_id %active_id% ;********************** using window id from WinGet *****************
WinMaximize, ahk_id %active_id%    ; use the window found above
;add your own code here 
ExitApp
return


 
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;   END: FINAL CLEANUP ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;; ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

GuiClose:
ExitApp

Esc::ExitApp