#Include names.ahk

GetAdType(title) {
	;Doubles
	if(RegExMatch(title, "i)BIB MPA", adtype) > 0) {
		Return "Billboard und Mobile Poster Ad "
	}
	if(RegExMatch(title, "i)HPA MPA", adtype) > 0) {
		Return "Halfpage und Mobile Poster Ad "
	}
	;Perferred Singles
	if(RegExMatch(title, "i)PRAP", adtype) > 0 or RegExMatch(title, "i)PROF", adtype) > 0) {
		Return "Profi-Advertorial "
	}
	;Singles
	if(RegExMatch(title, "i)BIB", adtype) > 0) {
		Return "Billboard "
	}
	if(RegExMatch(title, "i)EPI", adtype) > 0) {
		Return "E-Paper Interstitial "
	}
	if(RegExMatch(title, "i)FBA", adtype) > 0) {
		Return "Social-Media "
	}
	if(RegExMatch(title, "i)HPA", adtype) > 0) {
		Return "Halfpage Ad "
	}
	if(RegExMatch(title, "i)JdT", adtype) > 0) {
		Return "Job des Tages "
	}
	if(RegExMatch(title, "i)MED", adtype) > 0) {
		Return "Medium Rectangle "
	}
	if(RegExMatch(title, "i)MPA", adtype) > 0) {
		Return "Mobile Poster Ad "
	}
	if(RegExMatch(title, "i)PRA", adtype) > 0) {
		Return "Advertorial "
	}
	if(RegExMatch(title, "i)TKA", adtype) > 0) {
		Return "Teaser-Kachel "
	}
	Return ""
}

join(strArray) {
	s := ""
	for i,v in strArray
		s .= ", " . v
	return substr(s, 3)
}

ArrayHasValue(haystack, needle) {
	if(!isObject(haystack))
		return false
	if(haystack.Length()==0)
		return false
	for k,v in haystack
		if(v==needle)
			return true
	return false
}

GetRelationshipByName(name) {
	global friendly_names, names, general_names
	;MsgBox % join(names)
	if(ArrayHasValue(friendly_names, name)) {
		Return 2
	}
	first_name := StrSplit(name, " ")[1]
	;MsgBox % first_name . " " . StrLen(first_name)
	if(ArrayHasValue(names, first_name) or ArrayHasValue(general_names, first_name)) {
		Return 1
	}
	Return 0
}

GetMonth(month) {
	months := {01: "Januar", 02: "Februar", 03: "März", 04: "April", 05: "Mai", 06: "Juni", 07: "Juli", 08: "August", 09: "September", 10: "Oktober", 11: "November", 12: "Dezember"}
	if(months.HasKey(month)) {
		Return months[month]
	}
	Return ""
}

ResetModifierKeys() {
	Send {LShift Up}{RShift Up}{LControl Up}{RControl Up}{LAlt Up}{RAlt Up}{LWin Up}{RWin Up}
}