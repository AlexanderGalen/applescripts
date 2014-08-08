tell application "QuarkXPress"
	set x to selection as text
	tell document 1
		set CurrentPage to page number of current page
	end tell
end tell
set VersionNumber to x + 1
tell application "Keyboard Maestro Engine"
	make variable with properties {name:"Version Number", value:VersionNumber}
	make variable with properties {name:"Current Page", value:CurrentPage}
end tell