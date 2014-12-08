--checks to make sure quark is open and a document is open and something is selected and provides user feedback

set documentOpen to false
set validSelection to false
set quarkRunning to false

on is_running(appName)
	tell application "System Events" to (name of processes) contains appName
end is_running

set quarkRunning to is_running("QuarkXPress")
if quarkRunning then
	tell application "QuarkXPress"
		try
			get document 1
			set documentOpen to true
			if selection is not null then set validSelection to true
		end try
	end tell
end if

if not (documentOpen and validSelection) then
	tell application "QuarkXPress"
		activate
		display alert "For this script to work, Quark must be running, have a document open, and have area which you would like to make a printfile with selected"
	end tell
	return
end if



tell application "QuarkXPress"
	tell document 1
		set missingImages to missing of every image
		set modifiedImages to modified of every image
		if (missingImages contains true or modifiedImages contains true) then
			display dialog "Document contains Images that are either unlinked or modified, please update them before running this script"
		end if
	end tell
	set theName to name of document 1 as string
	try
		set jobNumber to find text "[0-9]{6}" in theName with regexp and string result
	on error
		set jobNumber to text returned of (display dialog "Input Job Number Please" default answer "")
	end try

	set activeJobs to "HOM_Shortrun:~HOM Active Jobs:"
	set thisPrintFile to activeJobs & jobNumber & ":" & jobNumber & ".printfile.pdf"
	set thisQuarkDoc to activeJobs & jobNumber & ":" & jobNumber & ".1up.print.qxp"

	set theSelection to selection
	if class of theSelection is group box then
		set grouped of theSelection to true
		set isGroupBox to true
	else
		set isGroupBox to false
	end if

	copy theSelection
	set {y1, x1, y2, x2} to bounds of theSelection as list
	set x1 to (coerce x1 to real)
	set x2 to (coerce x2 to real)
	set y1 to (coerce y1 to real)
	set y2 to (coerce y2 to real)
	set theWidth to x2 - x1
	set theHeight to y2 - y1

	set impositionTemplatesPath to "Resource:Templates:Shortrun Templates.New:X.Igen.SR Templates:"
	set newDocProperties to {page height:theHeight, page width:theWidth}

	--sets imposition templates and finished imposed filename according to size of product.
	if theWidth is 4 then
		set impositionTemplate to impositionTemplatesPath & "~CACH.HouseShape CalendarPads:CHCP.House_28up Layout.qxp"
		set imposedFile to activeJobs & jobNumber & ":" & jobNumber & ".CHCP.print.pdf"
	else if theWidth is 3.75 then
		set impositionTemplate to impositionTemplatesPath & "~CACP:CCP.Print.30up.qxp"
		set imposedFile to activeJobs & jobNumber & ":" & jobNumber & ".CCCP.print.pdf"
	else
		display dialog "Sizing is neither a CCCP or CHCP"
		return
	end if

	close document 1 without saving
	make new document with properties newDocProperties

	tell document 1
		activate
		paste

		if isGroupBox then
			set bounds of group box 1 to {0, 0, theHeight, theWidth}
		else
			set bounds of picture box 1 to {0, 0, theHeight, theWidth}
		end if

	  print print output style "Proof"
	end tell

	export layout space 1 of project 1 in thisPrintFile as "PDF" PDF output style "No Compression"
	save document 1 in thisQuarkDoc
	close every project without saving

	open file impositionTemplate
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
