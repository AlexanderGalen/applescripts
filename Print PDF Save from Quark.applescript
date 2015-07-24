--this is to tell me when I no longer have to keep Quark in focus
--because this script runs after a macro that needs quark at the front
tell application "Finder"
	activate
end tell

tell application "QuarkXPress"
	
	--declare some variables with info about the document
	set Currentdoc to document 1
	set CurrentProj to project 1
	set qxpname to name of Currentdoc
	set jobnumber to characters 1 through 6 of qxpname as text
	set n to 1
	set pageQty to count pages of document 1
	
	
	--print document
	print document 1 print output style "Proof"
	
	
	--export individual pages as pdfs
	repeat while n ² pageQty
		set FinishedPDFName to "Macintosh HD:Users:Maggie:Documents:Temp PDFs:" & (jobnumber) & ".v" & n & ".1" & ".pdf"
		export layout space 1 of CurrentProj in FinishedPDFName as "PDF" PDF output style "PDF Proof" page range n
		set n to n + 1
	end repeat
	
	--in case quark fails in saving document, saves in temp pdfs
	--which then copies into the job folder, overwriting existing file
	try
		save Currentdoc
	on error
		save Currentdoc in "Macintosh HD:Users:Maggie:Documents:Temp PDFs:" & jobnumber & ".HOM.qxp"
	end try
	close Currentdoc
end tell