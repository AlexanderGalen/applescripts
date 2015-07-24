set serverPath to "Quark:HPS Assets:Doc Pool:Templates:Support:"

tell application "QuarkXPress"
	tell document 1
		set allPictureBoxes to every picture box whose file path of image 1 is not null
		repeat with thisBox in allPictureBoxes
			set filePath to file path of image 1 of thisBox
			set filePathString to filePath as string
			set first5 to characters 1 thru 5 of filePathString as string
			if first5 is not "Quark" and first5 is not "Macin" then
				try
					set filePath to filePath as alias
					tell application "Finder" to set fileName to name of filePath as string
					set newFilePath to serverPath & fileName
					set image 1 of thisBox to alias newFilePath
				on error
					display dialog "Some image was not relinked:\n\n" & filePath & "\n\nIt is probably not copied to the quark server."
				end try
			end if
		end repeat
	end tell
end tell