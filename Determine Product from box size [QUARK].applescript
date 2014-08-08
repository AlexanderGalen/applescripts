to determineProduct()
	tell application "QuarkXPress"
		tell document 1
			set allBoxes to a reference to every generic box as list
		end tell
	end tell
	
	repeat with theBox in allBoxes
		tell application "QuarkXPress"
			
			--gets coordinates and sizing if the frame color is "die cut"
			--uses this info to determine what product it is
			
			if name of color of frame of theBox is "Die Cut" then
				set {y1, x1, y2, x2} to bounds of theBox as list
				set x1 to (coerce x1 to real)
				set y1 to (coerce y1 to real)
				set x2 to (coerce x2 to real)
				set y2 to (coerce y2 to real)
				set boxWidth to x2 - x1
				set boxHeight to y2 - y1
			end if
			
		end tell
		
		if ((boxWidth = 3.5) and (boxHeight = 9)) then
			return "QX"
		else if ((boxWidth = 3.5) and (boxHeight = 9)) then
			return "QS"
		else if ((boxWidth = 3.5) and (boxHeight = 2)) then
			return "BC"
		end if
		
	end repeat
	
end determineProduct

set productType to determineProduct()

--sets variable in Keyboard Maestro

tell application "Keyboard Maestro Engine"
	make variable with properties {name:"productType", value:productType}
end tell