#SingleInstance Force
#NoEnv
SendMode Input

;@Ahk2Exe-SetName			AutoHotKey Helper For My Work
;@Ahk2Exe-SetDescription	Description
;@Ahk2Exe-SetCompanyName	Jona Heinke
;@Ahk2Exe-SetCopyright		Copyright under MIT license`, Jona Heinke`, 2022-2024
;@Ahk2Exe-SetVersion		1.2

;toggles a window to be always on top
^Space::
	WinSet, AlwaysOnTop, Toggle, A

;prevent F1 in multiple programs
#IfWinActive ahk_exe Acrobat.exe
F1::
#IfWinActive ahk_exe chrome.exe
F1::
#IfWinActive ahk_exe explorer.exe
F1::
#IfWinActive ahk_exe firefox.exe
F1::
	Return

;makes Ctrl+Backspace work correctly in explorer
#IfWinActive ahk_exe explorer.exe
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
#IfWinActive ahk_exe chrome.exe
^m:: ;Ctrl+M
	;read information from YouTrack ticket
	Send {Esc 2}{F2}{Home}^+{End} ;edit ticket, select title (Ctrl because of multiline titles)
	title := GetSelectedText()
	Send {Tab}^+{End} ;select description (Ctrl because of multiline descriptions)
	Sleep 100
	description := GetSelectedText()
	Send {End}+{Tab}{Esc 2} ;deselect desciption, tab backwards to select title, press Esc (only works when title is selected)

	;Extract E-Mail Address
	mail := ""
	;Loop, Parse, description, `n, `r
	;{
	;	line := A_LoopField
	;	if(RegExMatch(line, "i)(?<=Auftragsbestätigung: )[-_a-zA-Z0-9.+!%]*@[-_a-zA-Z0-9.]*", mail) > 0) {
	;		Break
	;	}
	;}
	RegExMatch(description, "im)(?<=Auftragsbestätigung: )[-_a-zA-Z0-9.+!%]*@[-_a-zA-Z0-9.]*", mail)
	
	;Extract publishing date
	pubdate := ""
	RegExMatch(description, "im)(?<=Erscheintag: )\d{2}\.\d{2}\.\d{4}", pubdate)

	customer := RegExReplace(StrSplit(description, "`n")[1], " ?[\d@].*", "")

	adtype := ""
	if(RegExMatch(title, "i)FBA", adtype) > 0) {
		adtype := "Social-Media "
	}
	if(RegExMatch(title, "i)PRA", adtype) > 0) {
		adtype := "Advertorial "
	}
	if(RegExMatch(title, "i)TKA", adtype) > 0) {
		adtype := "Teaser-Kachel "
	}
	if(RegExMatch(title, "i)MED", adtype) > 0) {
		adtype := "Medium Rectangle "
	}
	if(RegExMatch(title, "i)JdT", adtype) > 0) {
		adtype := "Job des Tages "
	}

	;Test Output
	;MsgBox % "title = " . title
	;MsgBox % "description = " . description
	;MsgBox % "mail = " . mail
	;MsgBox % "pubdate = " . mail
	;Return

	;Switch to Outlook Window
	if not WinExist("ahk_exe outlook.exe") {
	;if not WinExist("ahk_exe thunderbird.exe") {
		MsgBox "Outlook not detected!, Exiting..."
		Return
	}
	WinActivate
	Sleep 500 ;just to be sure

	;Create New E-Mail
	Send ^+m
	;Send ^n
	Sleep 500
	;TODO: extract and paste e-mail
	SendRaw % mail
	Send {Enter}{Tab 2}
	;@Ahk2Exe-IgnoreBegin
	Send +{Tab}
	;@Ahk2Exe-IgnoreEnd
	SendRaw % title
	Send {Home}^{Del 3}Auswertung{Space}{End}{Tab} ;replace date with "Auswertung"
	Send Guten Tag ,{Enter}hier ist Ihre{Space}
	SendRaw % adtype
	Send Auswertung von{Space}
	;TODO: extract and paste name
	SendRaw % customer
	Send {Space}vom{Space}
	;TODO: extract and paste date
	SendRaw % pubdate
	Send .