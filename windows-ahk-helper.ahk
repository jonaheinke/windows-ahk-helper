#SingleInstance Force
#NoEnv
SendMode Input

;@Ahk2Exe-SetName			AutoHotKey Helper
;@Ahk2Exe-SetDescription	Description
;@Ahk2Exe-SetCompanyName	Company Name
;@Ahk2Exe-SetCopyright		Copyright \(c\) under MIT license, Jona Heinke, 2022-2024
;@Ahk2Exe-SetVersion		1.2

;toggles a window to be always on top
^Space::
	WinSet, AlwaysOnTop, Toggle, A

;prevent F1 in Adobe Acrobat Reader
#IfWinActive ahk_exe Acrobat.exe
F1::
	Return

;prevent F1 in folders and desktop
#IfWinActive ahk_exe Explorer.EXE
F1::
	Return

;makes Ctrl+Backspace work correctly in explorer
^BackSpace::
	Send {ShiftDown}{CtrlDown}{Left}{CtrlUp}{Left}{ShiftUp}{Del}

;replaces a selected image in Powerpoint with one of your choice and loads the filepath from clipboard
#IfWinActive ahk_exe POWERPNT.exe
^r::
	Send {AppsKey}bd^{v}{Enter}

GetSelectedText() {
	clipboard := ""
	Send ^c
	ClipWait, 1
	return clipboard
}

;create new e-mail from YouTrack ticket
;#IfWinActive ahk_exe chrome.exe
#IfWinActive ahk_exe firefox.exe
^ö:: ;Ctrl+Ö
	;read information from YouTrack ticket
	Send {F2}^+{Home} ;edit ticket, select title (Ctrl because of multiline titles)
	title := GetSelectedText()
	Send {Tab}^+{End} ;select description (Ctrl because of multiline descriptions)
	description := GetSelectedText()
	Send {End}+{Tab}{Esc} ;deselect desciption, tab backwards to select title, press Esc (only works when title is selected)

	;Test Output
	;MsgBox % "title = " . title
	;MsgBox % "description = " . description

	;Switch to Outlook Window
	if not WinExist("ahk_exe notepad.exe") { ;TODO: change to outlook.exe
		MsgBox "Outlook not detected!, Exiting..."
		Return
	}
	WinActivate
	Sleep 100 ;just to be sure

	;Create New E-Mail
	;Send ^+m
	Sleep 250

	;TODO: extract and paste e-mail
	Send E-Mail{Enter}

	;Send {Tab 2}
	Send Tab{Enter}Tab{Enter}
	SendRaw % title
	Send {Home}^{Del 3}Auswertung{Space} ;replace date with "Auswertung"
	Send {End}{Enter}Tab{Enter}
	;Send {Tab}
	Send Guten Tag ,{Enter}hier ist Ihre Auswertung vom{Space}

	;TODO: extract and paste name
	Send Kundenname

	Send {Space}vom{Space}

	;TODO: extract and paste date
	Send 4. Januar

	Send .