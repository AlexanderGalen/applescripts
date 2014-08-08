tell application "QuarkXPress"
	tell document 1
		if visible of layer 1 is false then
			set visible of layer 1 to true
		end if
		set missingimages to missing of every image
		if missingimages contains true then
			tell application "Keyboard Maestro Engine"
				make variable with properties {name:"Missing Images", value:true}
			end tell
		else
			tell application "Keyboard Maestro Engine"
				make variable with properties {name:"Missing Images", value:false}
			end tell
		end if
	end tell
end tell