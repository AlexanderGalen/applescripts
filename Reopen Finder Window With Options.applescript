on readPlist(thisKey)
	do shell script "defaults read ~/.finder-info.plist \"" & thisKey & "\""
end readPlist

set thisTarget to readPlist("target")
set x1 to readPlist("x1")
set y1 to readPlist("y1")
set x2 to readPlist("x2")
set y2 to readPlist("y2")

tell application "Finder"
	set newWindow to make new Finder window
	set properties of newWindow to {target:thisTarget, bounds:{x1, y1, x2, y2}}
end tell