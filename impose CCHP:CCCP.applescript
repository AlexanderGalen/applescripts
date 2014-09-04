set product to "CHCP"
set jobNumber to "381131"
set oneUpFile to "HOM_Shortrun:~HOM Active Jobs:" & jobNumber & ":" & jobNumber & ".printfile.pdf"
set imposedFile to "HOM_Shortrun:~HOM Active Jobs:" & jobNumber & ":" & jobNumber & ".CHCP.print.pdf"

if product is "CCCP" then
	set theTemplate to "RESOURCE:TEMPLATES:Shortrun Templates.New:X.Igen.SR Templates:~CCP:CCP.Print.30up.qxp"
else if product is "CHCP" then
	set theTemplate to "Resource:Templates:Shortrun Templates.New:X.Igen.SR Templates:~CHCP.HouseShape CalendarPads:CHCP.House_28up Layout.qxp"
end if

tell application "QuarkXPress"
	open file theTemplate
	set allBoxes to every picture box of document 1
	repeat with theBox in allBoxes
		if layername of theBox is "Default" then
			set image 1 of theBox to alias oneUpFile
		end if
	end repeat
	export layout space 1 of project 1 in imposedFile as "PDF" PDF output style "No Compression"
	close document 1 without saving
end tell
