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
		display alert "For this script to work, Quark must be running, have a document open, and have area which you would like to make a printfile with selected."
	end tell
	return
end if



tell application "QuarkXPress"
	tell document 1
		set missingImages to missing of every image
		set modifiedImages to modified of every image
		if (missingImages contains true or modifiedImages contains true) then
			display dialog "Document contains Images that are either unlinked or modified, please update them, then try running this script again."
		end if
	end tell
	
	set activeJobs to "HOM_Shortrun:~HOM Active Jobs:"
	
	set docName to name of document 1 as string
	set AppleScript's text item delimiters to "."
	set splitName to text items of docName
	set docName to items 1 thru ((length of splitName) - 2) of splitName as string
	set AppleScript's text item delimiters to {""}
		
	-- get path to job folder from document, or ask for it if document has no file on disk
	set jobPath to file path of document 1
	if jobPath is null then
		set jobNumber to text returned of (display dialog "Input job number please" default answer "")
		set docName to jobNumber
		set jobPath to activeJobs & jobNumber & ":" & jobNumber & ".HOM.qxp"
	else 
		set jobPath to jobPath as string
	end if
	
	-- get the parent folder to save into from jobPath
	set AppleScript's text item delimiters to ":"
	set splitPath to text items of jobPath
	set parentFolder to items 1 thru ((length of splitPath) - 1) of splitPath as string
	set AppleScript's text item delimiters to {""}
	
	-- check if job folder exists, and create it if not
	tell application "Finder"
		set folderExists to exists of parentFolder
		if not folderExists then
			
			set AppleScript's text item delimiters to ":"
			set splitPath to text items of parentFolder
			set parentParentFolder to items 1 thru ((length of splitPath) - 1) of splitPath as string
			set parentFolderName to item (length of splitPath) as string
			set AppleScript's text item delimiters to {""}
			
			make new folder at parentParentFolder with properties {name:parentFolderName}
			
		end if
			
	end tell	
	
	
	
	set thisPrintFile to parentFolder & ":" &  docName & ".printfile.pdf"
	set thisDestinationFile to parentFolder & ":" & docName
	
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
	
	if theWidth is not 4 and theWidth is not 3.75 then
		display dialog "Sizing is neither a CCCP or CHCP."
		return
	end if
	
	close document 1 without saving
	
	set newDocProperties to {page height:theHeight, page width:theWidth}
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
	
	export layout space 1 of project 1 in thisPrintFile as "PDF" PDF output style "No Compression"
	close document 1 without saving
	
end tell

if theWidth is 4 then
		set shape to "CACH"
else 
	set  shape to "CACC"
end if

set thisPosixPrintFile to quoted form of POSIX path of thisPrintFile
set thisPosixDestinationFile to quoted form of POSIX path of (thisDestinationFile & "." & shape & ".pdf")
log thisPosixDestinationFile
--set nodeScript to quoted form of "/Volumes/RESOURCE/Scripting/AAASpectacularAlex Scripts/Impose Calendar Pads.js"
set nodeScript to quoted form of "/Users/Alex/Scripts/Impose Calendar Pads.js"

do shell script "/usr/local/bin/node " & nodeScript & " " & thisPosixPrintFile & " " & thisPosixDestinationFile & " " & shape