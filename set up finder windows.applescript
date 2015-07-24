set path1 to "HOM_Shortrun:~HOM Archive Jobs:" as alias
set path2 to "HOM_Shortrun:~HOM Active Jobs:" as alias
set path3 to "Macintosh HD:Users:Maggie:Documents:Temp PDFs:" as alias
set path4 to "ART DEPARTMENT-NEW:Proofs-Shortrun:" as alias


set window1 to ""
set window2 to ""
set window3 to ""
set window4 to ""

to NewWindow(WindowName, PathName, PositionValue, BoundsValue)
	tell application "Finder"
		set WindowName to make new Finder window
		set properties of WindowName to {target:PathName, position:PositionValue, bounds:BoundsValue, current view:column view, sidebar width:205}
	end tell
end NewWindow

NewWindow(window1, path1, {-1280, 78}, {-1280, 78, -645, 568})
NewWindow(window2, path2, {-635, 78}, {-635, 78, 0, 568})
NewWindow(window3, path3, {-1280, 602}, {-1280, 602, -645, 1080})
NewWindow(window4, path4, {-635, 602}, {-635, 602, 0, 1079})

tell application "Finder"
	set properties of Finder window 1 to {current view:list view}
end tell