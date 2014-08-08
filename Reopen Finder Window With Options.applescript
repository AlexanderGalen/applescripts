tell application "Keyboard Maestro Engine"
	set TempVar to make variable with properties {name:"theTarget"}
	set theTarget to value of TempVar
	set TempVar to make variable with properties {name:"xpos"}
	set xpos to value of TempVar
	set TempVar to make variable with properties {name:"ypos"}
	set ypos to value of TempVar
	set TempVar to make variable with properties {name:"tlbound"}
	set tlbound to value of TempVar
	set TempVar to make variable with properties {name:"trbound"}
	set trbound to value of TempVar
	set TempVar to make variable with properties {name:"blbound"}
	set blbound to value of TempVar
	set TempVar to make variable with properties {name:"brbound"}
	set brbound to value of TempVar
	--set TempVar to make variable with properties {name:"viewMode"}
	--set viewMode to value of TempVar
	set TempVar to make variable with properties {name:"SideBarWidth"}
	set SideBarWidth to value of TempVar
	set TempVar to make variable with properties {name:"statusBarVisibility"}
	set statusBarVisibility to value of TempVar
	set TempVar to make variable with properties {name:"toolBarVisiblity"}
	set toolBarVisibility to value of TempVar
	
end tell

tell application "Finder"
	set newWindow to make new Finder window
	tell newWindow
		set target to theTarget
		set thePosition to {"", ""}
		set item 1 of thePosition to xpos
		set item 2 of thePosition to ypos
		set position to thePosition
		set theBounds to {"", "", "", ""}
		set item 1 of theBounds to tlbound
		set item 2 of theBounds to trbound
		set item 3 of theBounds to blbound
		set item 4 of theBounds to brbound
		set bounds to theBounds
		set sidebar width to SideBarWidth
		set statusbar visible to statusBarVisibility
		set toolbar visible to toolBarVisibility
		(*if viewMode is "icon view" then
			set current view to icon view
		else if viewMode is "list view" then
			set current view to list view
		else if viewMode is "column view" then
			set current view to column view
		else if viewMode is "flow view" then
			set current view to flow view
		end if*)
	end tell
end tell