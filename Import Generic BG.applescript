tell application "QuarkXPress"
	set productCode to uppercase text returned of (display dialog "Enter Product Code:" default answer "")
	set qxpName to name of Document 1
	set theType to characters 1 thru 2 of qxpName as string
	set theSize to characters 3 thru 4 of qxpName as string
	if theSize contains "Q" then set theSize to "Q"
end tell



set thePath to "/Volumes/HOM_shortrun/SUPERmergeIN/Custom\\ Magnet\\ Backgrounds/" & theType & "/HY" & theSize & "*" & productCode & ".eps"
set theBG to do shell script "find " & thePath

tell application "QuarkXPress"
	tell document 1
		set allBoxes to a reference to every generic box as list
		repeat with thisBox in allBoxes
			set {y1, x1, y2, x2} to bounds of thisBox as list
			set x1 to (coerce x1 to real)
			set x2 to (coerce x2 to real)
			set theWidth to (round (x2 - x1) * 100) / 100
			if theWidth > 0.5 then
				if name of color of thisBox is "Magenta" then
					set theBox to thisBox
					exit repeat
				end if
			end if
		end repeat
		--gets box sizing info & determines product from that
		set {y1, x1, y2, x2} to bounds of theBox as list
		set x1 to (coerce x1 to real)
		set y1 to (coerce y1 to real)
		set x2 to (coerce x2 to real)
		set y2 to (coerce y2 to real)
		set theWidth to (round (x2 - x1) * 100) / 100
		set theHeight to (round (y2 - y1) * 100) / 100
		
		set image 1 of theBox to alias theBG
		--sets fill color to none
		set color of theBox to "None"

	end tell
end tell