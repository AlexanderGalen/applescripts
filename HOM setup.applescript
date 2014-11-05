tell application "QuarkXPress"
	tell document 1
		make new layer at beginning
		repeat
			try
				set grouped of group box 1 to false
				set selection to null
			on error
				exit repeat
			end try
		end repeat
		set allBoxes to generic boxes
		repeat with thisBox in allBoxes
			try
				set theColor to name of color of frame of thisBox
				if theColor is "Die Cut" then move thisBox to beginning of layer 1
			end try
		end repeat
		set visible of layer 1 to false
	end tell
end tell