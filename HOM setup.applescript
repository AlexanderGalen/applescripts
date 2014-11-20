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
				set theBounds to bounds of thisBox as list
				set y1 to item 1 of theBounds
				set y1 to (coerce y1 to real)
				if y1 is 0 then delete thisBox
			end try
		end repeat
		set visible of layer 1 to false
		try
			set jobNumber to name as string
			set jobNumber to characters 1 thru 6 of jobNumber as string
			set prevInfoText to story 1 of text box "prev" as string
			set prevJob to find text "[0-9]{6}" in prevInfoText with regexp and string result
			set prevDoc to "HOM_Shortrun:~HOM Active Jobs:" & jobNumber & ":" & prevJob & ":" & prevJob & ".HOM.qxp"
			open file prevDoc
		end try
	end tell
end tell
