tell application "Finder"
	
	set productSize to "JU"
	set source_folder to "HOM:PRODUCTS:CALENDARS:2015 CALENDARS:W2P working:" & productSize & ":" as alias
	
	
	set File_List to (files of entire contents of source_folder) as alias list
	repeat with theItem in File_List
		set {name:FileName, name extension:fileExtension} to theItem
		if fileExtension is "psd" or fileExtension is "ai" then
			
			if fileExtension is "psd" then
				set thelength to length of FileName
				set noExtension to characters 1 thru (thelength - 4) of FileName as string
			else
				set thelength to length of FileName
				set noExtension to characters 1 thru (thelength - 3) of FileName as string
			end if
			
			tell application "Image Events"
				set currentImage to open theItem
				set theDimensions to dimensions of currentImage
				close currentImage
			end tell
			if item 1 of theDimensions < item 2 of theDimensions then
				set theOrientation to "Vertical"
			else
				set theOrientation to "Horizontal"
			end if
			
			
			if theOrientation is "Vertical" then
				set theTextDoc to "Macintosh HD:USers:maggie:desktop:CA" & productSize & "V.txt"
			else
				set theTextDoc to "Macintosh HD:USers:maggie:desktop:CA" & productSize & "H.txt"
			end if
			
			open for access file theTextDoc with write permission
			write (noExtension & return) to file theTextDoc starting at eof
			close access file theTextDoc
			
			
		end if
	end repeat
end tell