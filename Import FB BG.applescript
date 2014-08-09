tell application "Keyboard Maestro Engine"
	set tempvar to make new variable with properties {name:"Job Number"}
	set currentJob to value of tempvar
	set tempvar to make variable with properties {name:"Previous Job Number"}
	set OldJN to value of tempvar
end tell

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
		
		--sets fill color to none
		set color of theBox to "None"
	end tell
end tell

if theHeight = 7.25 then
	set theFile to "HOM_Shortrun:SUPERmergeIN:Custom Magnet Backgrounds:FB:FBQS14-"
else
	set theFile to "HOM_Shortrun:SUPERmergeIN:Custom Magnet Backgrounds:FB:FBFC14-"
end if

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
			set image 1 of theBox to alias theFile
		on error
			tell application "System Events"
				display dialog "An Error Occured"
				activate
				return
			end tell
		end try
	end tell
end tell

--gets the path to the old job copied to new folder and saves it to Keyboard Maestro
set shellScriptPath to "find " & quoted form of "/Volumes/HOM_shortrun/~Hom Active Jobs/" & currentJob & "/" & OldJN & "*/" & OldJN & "*HOM.qxp" & " | head -n 1"
set copiedPath to do shell script "find " & shellScriptPath
set copiedPath to POSIX file copiedPath

--opens previous job
tell application "QuarkXPress"
	try
		open file copiedPath
	end try
	activate
end tell