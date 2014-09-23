set sourceFolder to "hom:products:note card cafe:sets a6:¥ A6.SingleCards.New Size:NonHoliday:"
tell application "Finder"
	set fileList to entire contents of folder sourceFolder
end tell

repeat with theFile in fileList
	tell application "Finder"
		set coverSheet to "hom:products:note card cafe:sets a6:¥ A6.SingleCards.New Size:Coversheet:A6 Coversheet Template.qxp" as alias
		set finishedCoverSheet to "HOM:PRODUCTS:NOTE CARD CAFE:SETS A6:¥ A6.SingleCards.New Size:Coversheet:coverSheet.pdf"
		set theName to name of theFile
		set theLength to length of theName
		set productCode to characters 1 thru 12 of theName as string
		set productName to characters 14 thru (theLength - 4) of theName as string
		set finishedFileName to "A6.36singles." & productCode & "." & productName & ".pdf"
		set finishedFilePath to "hom:products:note card cafe:sets a6:¥ A6.SingleCards.New Size:Non Holiday 19 Page files:" & finishedFileName
		set coverText to productCode & return & productName
	end tell
	
	
	set coverSheet to coverSheet as string
	tell application "QuarkXPress"
		open file coverSheet
		tell document 1
			set text of text box 1 to coverText
		end tell
		export layout space 1 of project 1 in finishedCoverSheet as "PDF"
		close document 1 without saving
	end tell
	
	tell application "Adobe Acrobat Pro"
		open file finishedCoverSheet
		set theFile to theFile as string
		open file theFile
		set r to 1
		repeat 18 times
			insert pages document 1 after 1 from document 2 starting with 1 number of pages 1
			set r to r + 1
		end repeat
		close document 2
		save document 1 to finishedFilePath
		close document 1
	end tell
	
	tell application "Finder"
		delete finishedCoverSheet
	end tell
end repeat