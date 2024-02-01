#SingleInstance Force
#NoEnv
SendMode Input

#Include functions.ahk

;@Ahk2Exe-SetName			AutoHotKey Helper For My Work
;@Ahk2Exe-SetDescription	Description
;@Ahk2Exe-SetCompanyName	Jona Heinke
;@Ahk2Exe-SetCopyright		Copyright under MIT license`, Jona Heinke`, 2022-2024
;@Ahk2Exe-SetVersion		1.4

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
	Return

;replaces a selected image in Powerpoint with one of your choice and loads the filepath from clipboard
#IfWinActive ahk_exe POWERPNT.exe
^r::
	Send {AppsKey}bd^{v}{Enter}
	Return

GetSelectedText() {
	clipboard := ""
	Send ^c
	ClipWait, 1
	Return clipboard
}

;create new e-mail from YouTrack ticket when pressing Ctrl+M
#IfWinActive ahk_exe chrome.exe
^m::
#IfWinActive ahk_exe firefox.exe
^m::
	; ------------------------------------------- retrieve path from clipboard ------------------------------------------- ;
	path := clipboard
	file := ""
	if(RegExMatch(path, "([a-zA-Z]:|[/\\]|\.)([/\\][^<>:/\\\|\?\*])*[/\\]?"))
		Loop, %path%\*.pdf
			if(RegExMatch(A_LoopFileName, "^20\d{2}-\d{2}-\d{2}")) {
				file := A_LoopFileName
				break
			}

	; ------------------------------------ retrieve title and desciption from YouTrack ----------------------------------- ;
	Sleep 200
	ResetModifierKeys()
	;read information from YouTrack ticket
	;Send {Esc 2}{F2}{Home}^+{End} ;edit ticket, select title (Ctrl because of multiline titles)
	Send {Esc 2}{F2}^a ;edit ticket, select title
	title := GetSelectedText()
	;Send {Tab}^+{End} ;select description (Ctrl because of multiline descriptions)
	Send {Tab}^a ;select description
	Sleep 100
	description := GetSelectedText()
	Send {End}+{Tab}{Esc 3} ;deselect desciption, tab backwards to select title, press Esc (only works when title is selected)

	; ---------------------------------- extract further data from title and description --------------------------------- ;
	;Extract subject for e-mail
	subject := RegExReplace(title, "\d{2}-\d{2}(?= )", "Auswertung")
	subject := RegExReplace(subject, "i) (auf|für) (freiepresse|fp|blick|bl|erzgebirge|erz)\.de\D*(\d\D+)*", " | ")
	subject := RegExReplace(subject, "(?<=\d{6})[\|/\s,]+", " ")
	subject := RegExReplace(subject, " \d+ ?T\.?", "")
	;Extract E-Mail Address
	mail := ""
	RegExMatch(description, "im)(?<=Auftragsbestätigung: )[-_a-zA-Z0-9.+!%]+@[-_a-zA-Z0-9.]+\.[a-zA-Z0-9]{2,}", mail)
	;Extract name and relationship
	name := ""
	RegExMatch(description, "im)(?<=Vermittler: )[-a-zA-ZßäöüÄÖÜ. ]+?(?=Vermittler Nr)", name)
	name_split := StrSplit(name, " ")
	first_name := name_split[1]
	last_name  := name_split[name_split.Length()]
	relationship := GetRelationshipByName(name)
	;Extract ad type
	adtype := GetAdType(title)
	;Extract customer name
	customer := RegExReplace(StrSplit(description, "`n")[1], " ?[\d@].*", "")
	;Extract publishing date
	day := ""
	month := ""
	RegExMatch(description, "im)(?<=Erscheintag: )\d{2}\.(?=\d{2}\.\d{4})", day)
	RegExMatch(description, "im)(?<=Erscheintag: \d{2}\.)\d{2}(?=\.\d{4})", month)
	pubdate := RegExReplace(day, "^0", "") . " " . GetMonth(month)

	; ---------------------------------------------------- Test Output --------------------------------------------------- ;
	;@Ahk2Exe-IgnoreBegin
	MsgBox % "title = " . title
	MsgBox % "subject = " . subject
	MsgBox % "description = " . description
	MsgBox % "mail = " . mail
	MsgBox % "name = " . name
	MsgBox % "relationship = " . relationship
	MsgBox % "adtype = " . adtype
	MsgBox % "customer = " . customer
	MsgBox % "pubdate = " . pubdate
	;Return
	;@Ahk2Exe-IgnoreEnd

	; ------------------------------------------ switch focus to outlook window ------------------------------------------ ;
	;Switch to Outlook Window
	;if not WinExist("ahk_exe outlook.exe") {
	if not WinExist("ahk_exe thunderbird.exe") {
	;if not WinExist("ahk_exe notepad.exe") {
		MsgBox "Outlook not detected!, Exiting..."
		Return
	}
	WinActivate
	Sleep 500
	ResetModifierKeys()

	; ------------------------------------------------- create new email ------------------------------------------------- ;
	;Shortcut for New E-Mail
	Send ^+m
	;Send ^n ;Shortcut in Thunderbird
	Sleep 500
	ResetModifierKeys()
	SendRaw % mail
	Send {Tab 2}
	;Send +{Tab}
	SendRaw % subject
	;Send {Home}^{Del 3}Auswertung {Tab} ;replace date with "Auswertung"
	;@Ahk2Exe-IgnoreBegin
	;Send {End}
	Send {Enter 2}
	;@Ahk2Exe-IgnoreEnd
	Send {Tab}
	Sleep 100
	if(relationship == 2) {
		Send Hallo{Space}
		SendRaw % first_name
		Send `,{Enter}
		Sleep 100
		Send hier ist deine{Space}
	} else {
		Send Guten Tag{Space}
		if(relationship == 1) {
			Send Frau{Space}
			SendRaw % last_name
		} else {
			Send Herr{Space}
			SendRaw % last_name
		}
		Send `,{Enter}
		Sleep 100
		Send hier ist Ihre{Space}
	}
	SendRaw % adtype
	Send Auswertung von{Space}
	SendRaw % customer
	Send {Space}vom{Space}
	Send ^{Left}{Del}v{End} ;remove large V that sometimes appears
	SendRaw % pubdate
	Send .

	; -------------------------------------------------- attach PDF file ------------------------------------------------- ;
	if(file != "") {
		;@Ahk2Exe-IgnoreBegin
		if not WinExist("ahk_exe thunderbird.exe") {
			MsgBox "Thunderbird not detected!, Exiting..."
			Return
		}
		WinActivate
		Sleep 500
		Send ^n
		Sleep 500
		Send ^+a
		Sleep 500
		;@Ahk2Exe-IgnoreEnd
		Send !i5s
		Sleep 500
		Send {F4}^a
		SendRaw % path
		Sleep 500
		Send {Enter}
		Sleep 250
		Send !n
		SendRaw % file
		Sleep 100
		Send {Enter}
		}

	; --------------------------------------------------- exit program --------------------------------------------------- ;
	ResetModifierKeys()
	Return