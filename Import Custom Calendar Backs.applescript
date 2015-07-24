set newFolder to "HOM:PRODUCTS:CALENDARS:2015 CALENDARS:2015 Custom Calendars:NEW:"
set i to 55
repeat 1 times
	tell application "Microsoft Excel"
		tell row i
			set productCode to string value of cell 1
			set theTemplate to string value of cell 2
			set textColor to string value of cell 5
		end tell
	end tell
	
	set thisBG to "HOM_Shortrun:SUPERmergeIN:Custom Magnet Backgrounds:CA:" & productCode & ".eps"
	set thisQXP to newFolder & productCode & ".qxp"
	set thisProject to theTemplate & "A.qxp"
	set thisTemplate to "HOM:PRODUCTS:CALENDARS:2015 CALENDARS:Custom Updates:" & theTemplate & "A.qxp"
	
	tell application "QuarkXPress"
		open file thisTemplate
		tell document 1
			set allBoxes to a reference to every generic box as list
			repeat with thisBox in allBoxes
				set {y1, x1, y2, x2} to bounds of thisBox as list
				set x1 to (coerce x1 to real)
				set x2 to (coerce x2 to real)
				set theWidth to (x2 - x1)
				set theWidth to format theWidth into "##.###"
				if theWidth is "5.25" or theWidth is "6.75" then
					log "width matched"
					set boxFound to true
					set image 1 of thisBox to alias thisBG
					exit repeat
				else
					set boxFound to false
					log "width didn't match"
				end if
			end repeat
			if boxFound is false then
				return "Picture box import unsucessful. Row number: " & i
			end if
			if textColor is "white" then
				tell application "QuarkXPress"
					set selection to text of story 1 of text box 1 of layout space 1 of project thisProject of application "QuarkXPress"
					set thisSelection to selection
					set color of thisSelection to "white"
					try
						set selection to text of story 1 of text box 2 of layout space 1 of project thisProject of application "QuarkXPress"
						set thisSelection to selection
						set color of thisSelection to "white"
					end try
				end tell
			end if
			save in thisQXP
			close
		end tell
	end tell
	set i to i + 1
end repeat

tell application "Microsoft Excel"
	close active workbook
end tell