tell application "Keyboard Maestro Engine"
	set tempvar to make variable with properties {name:"Job Number"}
	set jobNumber to value of tempvar
end tell
set activeJobs to "HOM_Shortrun:~HOM Active Jobs:"
set thisPrintFile to activeJobs & jobNumber & ":" & jobNumber & ".printfile.pdf"
set thisQuarkDoc to activeJobs & jobNumber & ":" & jobNumber & ".HOM.qxp"


tell application "QuarkXPress"
	set theSelection to selection
	set grouped of theSelection to true
	copy theSelection
	close document 1 without saving

	set {y1, x1, y2, x2} to bounds of theSelection as list
	set x1 to (coerce x1 to real)
	set x2 to (coerce x2 to real)
	set theWidth to (round (x2 - x1) * 100) / 100
	
	set templatesPath to "Resource:Templates:Shortrun Templates.New:X.Igen.SR Templates:"

	if theWidth is 4 then
		set product to "CHCP"
		set thisTemplate to templatesPath & "~CHCP.HouseShape CalendarPads:CHCP.House_28up Layout.qxp"
		set imposedFile to activeJobs & jobNumber & ":" & jobNumber & ".CHCP.print.pdf"
	else if theWidth is 3.75 then
		set product to "CCCP"
		set thisTemplate to templatesPath & "~CCP:CCP.Print.30up.qxp"
		set imposedFile to activeJobs & jobNumber & ":" & jobNumber & ".CCCP.print.pdf"
	else
		return "Sizing Incorrect"
	end if

	set destinationDoc to open file thisTemplate
	tell destinationDoc
		activate
		paste
		set bounds of group box 1 to {0, 0, 2.25, 4}
	end tell

	export layout space 1 of project 1 in thisPrintFile as "PDF" PDF output style "No Compression"
	save thisTemplate in thisQuarkDoc 
	close every project without saving

	open file theTemplate
	set allBoxes to every picture box of document 1
	repeat with theSelection in allBoxes
		if layername of theSelection is "Default" then
			set image 1 of theSelection to alias thisPrintFile
		end if
	end repeat
	export layout space 1 of project 1 in imposedFile as "PDF" PDF output style "No Compression"
	close document 1 without saving
end tell

tell application "Finder"
	delete thisPrintFile
end tell