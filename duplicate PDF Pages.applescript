set sourceFolder to "HOM:PRODUCTS:NOTE CARD CAFE:SETS A6:¥ A6.SingleCards.New Size:Holiday:" as alias
tell application "Finder"
	set filelist to entire contents of sourceFolder as alias list
end tell

repeat with theFile in filelist
	tell application "Finder"
		set theName to name of theFile
		set theExtension to name extension of theFile
		set firstLine to characters 1 thru 12 of theName as string
		set secondLine to characters 14 thru ((length of theName) - 4) of theName as string
		set coverSheetText to firstLine & "\n" & secondLine
		--set FinishedFilePath to "Macintosh HD:Users:maggie:desktop:new:" & theName
		set FinishedFilePath to "HOM:PRODUCTS:NOTE CARD CAFE:SETS A6:¥ A6.SingleCards.New Size:New:" & theName
		--return FinishedFilePath
	end tell
	
	set coverPDF to "Macintosh HD:Users:maggie:desktop:temp.pdf"
	
	if theExtension is "pdf" then
		tell application "QuarkXPress"
			open file "Macintosh HD:Users:maggie:desktop:temp.qxp"
			tell document 1
				set text of text box 1 to coverSheetText
			end tell
			export layout space 1 of project 1 in coverPDF as "PDF"
			close project 1 without saving
		end tell
		
		tell application "Adobe Acrobat Pro"
			open file coverPDF
			tell document 1
				save to file FinishedFilePath
			end tell
			set theFile to theFile as string
			open file theFile
			set i to 1
			repeat 18 times
				insert pages document 1 after i from document 2 starting with 1 number of pages 1
				set i to i + 1
			end repeat
			close document 2
			save document 1
			close document 1
		end tell
		
		tell application "Finder"
			delete coverPDF
		end tell
		
	end if
	
	
end repeat


