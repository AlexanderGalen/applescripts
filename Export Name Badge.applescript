tell application "QuarkXPress"
	
	--saves document and exports a pdf for proofing into the folder that the document is in
	tell document 1
		save
		set tool mode to drag mode
		set docPath to file path as string
		set docName to name as string
		set finishedPDFName to characters 1 thru ((length of docPath) - 4) of docPath & ".pdf" as string
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
	
	
	--loops through all layout spaces and saves out a group picture of each one.
	tell project 1
		set allSpaces to layout spaces
	end tell
	
	repeat with thisSpace in allSpaces
		set thisName to name of thisSpace
		tell project 1 to set active layout space to layout space thisName
		set finishedGPName to "HOM:PRODUCTS:NAME BADGES:Group Pictures:" & characters 1 thru ((length of docName) - 4) of docName & "-" & thisName & ".gp" as string
		
		tell document 1
			
			set selected of every generic box of layer "Default" to true
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
			
			set newDocProperties to {page height:theHeight, page width:theWidth}
			
		end tell
		
		make new document with properties newDocProperties
		
		tell document 1
			activate
			paste
			
			if isGroupBox then
				set bounds of group box 1 to {0, 0, theHeight, theWidth}
			else
				set bounds of picture box 1 to {0, 0, theHeight, theWidth}
			end if
			
			save in finishedGPName
			close
			
		end tell
		
	end repeat
	
	tell document 1
		close without saving
	end tell
end tell