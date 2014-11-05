on writePlist(thisKey, thisData)
	do shell script "defaults write ~/.finder-info.plist \"" & thisKey & "\" \"" & thisData & "\""
end writePlist

tell application "Finder"
	set thisTarget to target of Finder window 1 as alias
	set theseBounds to bounds of Finder window 1
	set x1 to item 1 of theseBounds
	set y1 to item 2 of theseBounds
	set x2 to item 3 of theseBounds
	set y2 to item 4 of theseBounds
end tell

writePlist("target", thisTarget)
writePlist("x1", x1)
writePlist("y1", y1)
writePlist("x2", x2)
writePlist("y2", y2)