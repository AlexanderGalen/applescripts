tell application "QuarkXPress"
	tell document 1

		set missingImages to missing of every image
		set modifiedImages to modified of every image
		if (missingImages contains true or modifiedImages contains true) then
			display dialog "Document contains Images that are either unlinked or modified, please update them before running this script"
			return
		end if

		set docPath to file path as string
		set finishedPDFName to characters 1 thru ((length of docPath) - 4) of docPath & ".1.pdf" as string

	end tell

	set activeSpace to active layout space of project 1

	try
		export activeSpace in finishedPDFName as "PDF" PDF output style "PDF Proof"
	on error number 12
		display dialog "File Already Exists. Overwrite?"
		tell application "Finder"
			delete file finishedPDFName
		end tell
		export activeSpace in finishedPDFName as "PDF" PDF output style "PDF Proof"
	end try

	close document 1 without saving

end tell
