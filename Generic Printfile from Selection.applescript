tell application "QuarkXPress"
	set theName to name of document 1 as string
	
	try
		set jobNumber to find text "[0-9]{6}" in theName with regexp and string result
	on error
		set jobNumber to text returned of (display dialog "Input Job Number Please" default answer "")
	end try
	
	if exists page 2 of document 1 then
		set multiPage to true
	else
		set multiPage to false
	end if
	
	
	
end tell

set i to 1
set activeJobs to "HOM_Shortrun:~HOM Active Jobs:"
set theCondition to true
set thisprintFile to activeJobs & jobNumber & ":" & jobNumber & ".printfile." & i & ".pdf"
set for4Over to "ART DEPARTMENT-NEW:FOR 4over:"

repeat while theCondition
	tell application "Finder"
		if exists thisprintFile then
			set theCondition to true
			set i to i + 1
			set thisprintFile to activeJobs & jobNumber & ":" & jobNumber & ".printfile." & i & ".pdf"
		else
			set theCondition to false
		end if
	end tell
end repeat

tell application "QuarkXPress"
	
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
	set y1 to (coerce y1 to real)
	set x2 to (coerce x2 to real)
	set y2 to (coerce y2 to real)
	
	set theWidth to x2 - x1
	set theHeight to y2 - y1
	
	set newDocProperties to {page height:theHeight, page width:theWidth}
	if not multiPage then close document 1 without saving
	make new document with properties newDocProperties
	
	tell document 1
		activate
		paste
		if isGroupBox then
			set bounds of group box 1 to {0, 0, theHeight, theWidth}
		else
			set bounds of picture box 1 to {0, 0, theHeight, theWidth}
		end if
	end tell
	
	export layout space 1 of project 1 in thisprintFile as "PDF" PDF output style "No Compression"
	if not multiPage then
		close every project without saving
	else
		close document 2 without saving
	end if
end tell

set posixPrintfile to quoted form of POSIX path of thisprintFile
set posixDestination to quoted form of POSIX path of for4Over
do shell script "cp -p " & posixPrintfile & " " & posixDestination
