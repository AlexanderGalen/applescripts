tell application "QuarkXPress"
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
	set theWidth to (round (x2 - x1) * 100) / 100
	
	set impositionTemplatesPath to "Resource:Templates:Shortrun Templates.New:X.Igen.SR Templates:"
	
	if theWidth is 4 then
		set product to "CHCP"
		set impositionTemplate to impositionTemplatesPath & "~CHCP.HouseShape CalendarPads:CHCP.House_28up Layout.qxp"
		set imposedFile to activeJobs & jobNumber & ":" & jobNumber & ".CHCP.print.pdf"
		set newDocProperties to {page height:Çdata FXVM0000A200È, page width:Çdata FXHM00002001È}
	else if theWidth is 3.75 then
		set product to "CCCP"
		set impositionTemplate to impositionTemplatesPath & "~CCP:CCP.Print.30up.qxp"
		set imposedFile to activeJobs & jobNumber & ":" & jobNumber & ".CCCP.print.pdf"
		set newDocProperties to {page height:Çdata FXVM0000A200È, page width:Çdata FXHM00000E01È}
	else
		return "Sizing is neither a CCCP or CHCP"
	end if
	
	close document 1 without saving
	make new document with properties newDocProperties
	
	tell document 1
		activate
		paste
		if isGroupBox then
			if product is "CCCP" then
				set bounds of group box 1 to {0, 0, 2.25, 3.75}
			else
				set bounds of group box 1 to {0, 0, 2.25, 4}
			end if
		else
			if product is "CCCP" then
				set bounds of picture box 1 to {0, 0, 2.25, 3.75}
			else
				set bounds of picture box 1 to {0, 0, 2.25, 4}
			end if
		end if
		
		--print print output style "Proof"
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