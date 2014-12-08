--these subroutines take a mac path as parameters. they fail if there is nothing in the folder so it checks first and does nothing if folder is empty
--copy functions will replace existing files

on cp_all(source, destination)
	set source to quoted form of POSIX path of source
	set destination to quoted form of POSIX path of destination
	do shell script "cp -Rpf " & source & "* " & destination
end cp_all

on cp(source, destination)
	set source to quoted form of POSIX path of source
	set destination to quoted form of POSIX path of destination
	do shell script "cp -Rpf " & source & " " & destination
end cp

on rm_all(source)
	tell application "Finder"
		set theFiles to entire contents of folder source as alias list
	end tell
	if theFiles is not {} then
		set source to quoted form of POSIX path of source
		do shell script "rm -d -f " & source & "*"
	end if
end rm_all

on rm(source)
	set source to quoted form of POSIX path of source
	do shell script "rm -d -f " & source
end rm


--checks to make sure boxes are named in quark to be sure that the currently opened job came from a template with named boxes.
tell application "QuarkXPress"
	tell document 1
		set theseNames to name of every generic box as list
		if not (theseNames contains "static" or theseNames contains "Die Cut" or theseNames contains "Proof Sheet Info") then
			display dialog "Document does not contain named boxes. This likely means that this document was built before this script was created, and thus the scipt will not work on this document."
			return
		end if
	end tell
end tell


--gets current date
set {year:y, month:m, day:d} to (current date)
set m to m * 1
set thisdate to m & "/" & d & "/" & y as string
set proofsShortRun to "Art Department-New:Proofs-Shortrun:"

tell application "QuarkXPress"
	try
		set currentDoc to document 1
		set currentProj to project 1
	on error
		display dialog "Unable to get document 1 or project 1."
		return
	end try

	--this part is printing, and saving the proof
	tell currentDoc

		--checks to make sure no images are unlinked or modified.
		set missingImages to missing of every image
		set modifiedImages to modified of every image
		if missingImages contains true or modifiedImages contains true then
			display dialog "Some images are unlinked or modified. Update them, then run this script again."
			return
		end if

		--checks to make sure die line layer is on. expects die line to be on the topmost layer. will work fine if it is on same layer as rest of layout.
		if visible of layer 1 is false then
			set visible of layer 1 to true
		end if

		--changes date and version number in proof sheet. expects the text box containing all that info to be named "Proof Sheet Info"
		set word 5 of story 1 of text box "Proof Sheet Info" to thisdate
		set proofSheetInfoText to story 1 of text box "Proof Sheet Info"
		set AppleScript's text item delimiters to "."
		set oldVersionNumber to last text item of proofSheetInfoText
		set newVersionNumber to oldVersionNumber + 1
		set theseTextItems to text items of proofSheetInfoText
		set last item of theseTextItems to newVersionNumber
		set proofSheetInfoText to theseTextItems as string
		set story 1 of text box "Proof Sheet Info" to proofSheetInfoText
		set AppleScript's text item delimiters to ""

		--tries to save quark document. this happens as early as possible in the script in case it fails
		try
			save
		on error
			display dialog "Document failed to save. Save it manually, close it, then re-open and try again."
			return
		end try

		set qxpName to name
		set jobNumber to characters 1 thru 6 of qxpName as string
		set jobFolder to "HOM_Shortrun:~HOM Active Jobs:" & jobNumber & ":"

		--gets version number from the name of quark document
		if length of qxpName is 14 then --standard, non group order document name
			set proofNumber to 1
		else if length of qxpName is 16 then --group order with a single digit version number (<10)
			set proofNumber to character 8 of qxpName as string
		else if length of qxpName is 17 then --group order with a double digit version number (>10)
			set proofNumber to characters 8 thru 9 of qxpName as string
		else
			display dialog "Name of quark document does not match standard naming convention. This script will only work correctly with documents that follow the standard naming convention"
		end if

		set printFilePath to jobFolder & jobNumber & ".v" & proofNumber & ".printfile.pdf"
		set proofPath to jobFolder & jobNumber & ".v" & proofNumber & "." & newVersionNumber & ".pdf"
		--check if proof already exists and display dialog to ask user what to do if so
		tell application "Finder"
			if exists proofPath then
				tell application "QuarkXPress"
					display dialog "Proof file already exists in Job Folder. Overwrite?"
				end tell
				delete proofPath
			end if
		end tell
	end tell
end tell


--checks to see if prinfile already exists, then displays a dialog asking to overwrite or not.
tell application "Finder"
	set proofExists to exists of proofPath
end tell

if proofExists then
	display dialog "Proof Already exists in Job Folder. Overwrite?"
	rm(proofPath)
end if


tell application "QuarkXPress"

	print currentDoc print output style "Proof"
	export layout space 1 of currentProj in proofPath as "PDF" PDF output style "PDF Proof"

	--this part is saving the printfile.
	tell currentDoc
		--sets tool to selection tool and deselects all boxes to avoid including any unwanted boxes in printfile
		set tool mode to drag mode
		set selected of every generic box to false
		--selects every box that is not named. this should only be the product and any boxes manually created.
		--uses two separate try blocks because one of the two will often fail, and it should execute both select commands even if the first fails. the second wont execute if they are in the same block
		try
			set selected of every generic box whose name is "" to true
		end try
		try
			set selected of every generic box whose name is null to true
		end try
		set thisProduct to selection

		--checks if product is composed of multiple boxes or a single box. this is to ensure that later, when setting the bounds in the new document, if there are multiple, it sets the bounds of a "group box" and if there is only one, a "generic box"
		if class of thisProduct is group box then
			set grouped of thisProduct to true
			set isGroupBox to true
		else
			set isGroupBox to false
		end if

		copy thisProduct

		--gets sizing of the product for the creation of the printfile document
		set {y1, x1, y2, x2} to bounds of thisProduct as list
		set x1 to (coerce x1 to real)
		set x2 to (coerce x2 to real)
		set y1 to (coerce y1 to real)
		set y2 to (coerce y2 to real)
		set theWidth to x2 - x1
		set theHeight to y2 - y1
		close without saving
	end tell

	--makes new document for printfile using measurements
	set thisPrintFileDoc to make new document with properties {page height:theHeight, page width:theWidth}

	--pastes product in newly created document, and sets its bounds so that it is lined up in the top left corner
	tell document 1
		activate
		paste
		if isGroupBox then
			set bounds of group box 1 to {0, 0, theHeight, theWidth}
		else
			set bounds of generic box 1 to {0, 0, theHeight, theWidth}
		end if
	end tell

	--deletes printfile if it already exists so quark is free to export the new one.
	tell application "Finder"
		set existsPrintFile to exists of printFilePath
	end tell
end tell

if existsPrintFile then
	rm(printFilePath)
end if

tell application "QuarkXPress"
	--exports new printfile in job folder and closes document
	export layout space 1 of project 1 in printFilePath as "PDF" PDF output style "No Compression"
	close document 1 without saving

end tell

--copies proof to proofs shortrun folder
cp(proofPath, proofsShortRun)
