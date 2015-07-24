tell application "Finder"
	activate
end tell

--gets value of variables from KM
tell application "Keyboard Maestro Engine"
	set TempVar to make variable with properties {name:"Version Number"}
	set VersionNumber to value of TempVar as integer
	set TempVar to make variable with properties {name:"Current Page"}
	set CurrentPage to value of TempVar as integer
end tell

tell application "QuarkXPress"
	set Currentdoc to document 1
	set CurrentProj to project 1
	set qxpname to get name of Currentdoc
	set chars to characters of qxpname
	set jobNumber to items 1 through 6 of chars as text
	print page CurrentPage of document 1 print output style "Proof"
	export layout space 1 of CurrentProj in "Macintosh HD:Users:Maggie:Documents:Temp PDFs:" & jobNumber & ".v" & CurrentPage & "." & VersionNumber & ".pdf" as "PDF" page range CurrentPage
	try
		save Currentdoc
	on error
		save Currentdoc in "Macintosh HD:Users:Maggie:Documents:Temp PDFs:" & jobNumber & ".HOM.qxp"
	end try
	close Currentdoc with saving
end tell