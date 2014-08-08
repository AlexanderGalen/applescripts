tell application "Keyboard Maestro Engine"
	set tempvar to make new variable with properties {name:"Job Number"}
	set currentJob to value of tempvar
end tell

set theFile to "HOM_Shortrun:SUPERmergeIN:Custom Magnet Backgrounds:FB:FBFC14-"
tell application "System Events"
	set theinfo to text returned of (display dialog "input file info" default answer "")
end tell
if length of theinfo is 4 then
	set theinfo to "0" & theinfo
else if length of theinfo is 3 then
	set theinfo to "00" & theinfo
end if

set firstthree to characters 1 thru 3 of theinfo as string
if firstthree is "243" then
	set theFile to "MERGE CENTRAL:BYO Schedules:PDFs:" & currentJob & ".pdf"
else
	set theFile to theFile & theinfo & ".eps"
end if

tell application "QuarkXPress"
	tell document 1
		try
			set image 1 of selection to alias theFile
		on error
			tell application "System Events"
				display dialog "An Error Occured"
			end tell
			activate
			return
		end try
		activate
	end tell
end tell