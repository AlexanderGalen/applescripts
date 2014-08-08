tell application "Finder"
	tell Finder window 1
		set theIndex to index
	end tell
end tell
if theIndex is 1 then
	tell application "Finder"
		tell Finder window 1
			set theTarget to target as string
			set thePosition to position
			set xpos to item 1 of thePosition
			set ypos to item 2 of thePosition
			set theBounds to bounds
			set tlbound to item 1 of theBounds
			set trbound to item 2 of theBounds
			set blbound to item 3 of theBounds
			set brbound to item 4 of theBounds
			--set viewMode to current view as string
			set SideBarWidth to sidebar width
			set statusBarVisibility to statusbar visible
			set toolBarVisiblity to toolbar visible
		end tell
		
	end tell
	tell application "Keyboard Maestro Engine"
		make variable with properties {name:"theTarget", value:theTarget}
		make variable with properties {name:"xpos", value:xpos}
		make variable with properties {name:"ypos", value:ypos}
		make variable with properties {name:"tlbound", value:tlbound}
		make variable with properties {name:"trbound", value:trbound}
		make variable with properties {name:"blbound", value:blbound}
		make variable with properties {name:"brbound", value:brbound}
		--make variable with properties {name:"viewMode", value:viewMode}
		make variable with properties {name:"SideBarWidth", value:SideBarWidth}
		make variable with properties {name:"statusBarVisibility", value:statusBarVisibility}
		make variable with properties {name:"toolBarVisiblity", value:toolBarVisiblity}
	end tell
	
else
	
	--tell application "System Events" to keystroke "w" using command down
	
end if